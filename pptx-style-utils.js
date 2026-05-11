(() => {
    function parseNumber(value) {
        if (value === undefined || value === null || value === '') return null;
        const num = Number(String(value).replace(',', '.'));
        return Number.isFinite(num) ? num : null;
    }

    function parseXml(xml) {
        const doc = new DOMParser().parseFromString(xml, 'application/xml');
        const parserError = doc.getElementsByTagName('parsererror')[0];
        if (parserError) throw new Error(parserError.textContent || 'Invalid XML in PPTX');
        return doc;
    }

    function getSlideEntries(zip) {
        const slideRegex = /^ppt\/slides\/slide(\d+)\.xml$/;
        const entries = [];
        zip.forEach((path) => {
            const m = path.match(slideRegex);
            if (m) entries.push({ path, num: parseInt(m[1], 10) });
        });
        entries.sort((a, b) => a.num - b.num);
        return entries;
    }

    function getTagName(element) {
        return element.tagName || element.nodeName || '';
    }

    function getFirstDescendant(element, tagName) {
        return element.getElementsByTagName(tagName)[0] || null;
    }

    function getDirectChild(element, tagName) {
        if (!element) return null;
        for (const child of Array.from(element.children || [])) {
            if (getTagName(child) === tagName) return child;
        }
        return null;
    }

    function getTextFromRun(run) {
        const tNodes = run.getElementsByTagName('a:t');
        let text = '';
        for (let i = 0; i < tNodes.length; i++) {
            text += tNodes[i].textContent || '';
        }
        return text;
    }

    function getTextFromParagraph(aParagraph) {
        const runs = aParagraph.getElementsByTagName('a:r');
        let text = '';
        for (let i = 0; i < runs.length; i++) {
            text += getTextFromRun(runs[i]);
        }
        return text;
    }

    function getFirstTextRun(paragraph) {
        const runs = Array.from(paragraph.getElementsByTagName('a:r'));
        return runs.find(run => getTextFromRun(run).trim()) || runs[0] || null;
    }

    function getElementAttrBool(element, attrName) {
        if (!element || !element.hasAttribute(attrName)) return '';
        const value = String(element.getAttribute(attrName) || '').toLowerCase();
        return value === '1' || value === 'true' ? '1' : '0';
    }

    function getFillColor(element) {
        const solidFill = getDirectChild(element, 'a:solidFill');
        if (!solidFill) return '';
        const srgbClr = getDirectChild(solidFill, 'a:srgbClr');
        if (srgbClr && srgbClr.getAttribute('val')) return `#${srgbClr.getAttribute('val')}`;
        const schemeClr = getDirectChild(solidFill, 'a:schemeClr');
        if (schemeClr && schemeClr.getAttribute('val')) return `scheme:${schemeClr.getAttribute('val')}`;
        return 'solidFill';
    }

    function getLineSpacing(paragraph) {
        const pPr = getDirectChild(paragraph, 'a:pPr');
        const lnSpc = pPr ? getDirectChild(pPr, 'a:lnSpc') : null;
        if (!lnSpc) return '';
        const spcPct = getDirectChild(lnSpc, 'a:spcPct');
        if (spcPct && spcPct.getAttribute('val')) {
            const value = parseNumber(spcPct.getAttribute('val'));
            return value ? `${value / 1000}%` : '';
        }
        const spcPts = getDirectChild(lnSpc, 'a:spcPts');
        if (spcPts && spcPts.getAttribute('val')) {
            const value = parseNumber(spcPts.getAttribute('val'));
            return value ? `${value / 100}pt` : '';
        }
        return '';
    }

    function fontInfoFromElement(element, fontSource = '') {
        const raw = element ? element.getAttribute('sz') : '';
        const size = parseNumber(raw);
        if (!size) return null;
        const info = { fontPt: size / 100 };
        if (fontSource) info.fontSource = fontSource;
        return info;
    }

    function getParagraphStyleFingerprint(paragraph, fontInfo) {
        const run = getFirstTextRun(paragraph);
        const rPr = run ? getDirectChild(run, 'a:rPr') : null;
        const lineSpacing = getLineSpacing(paragraph);
        const color = getFillColor(rPr);
        const bold = getElementAttrBool(rPr, 'b');
        const italic = getElementAttrBool(rPr, 'i');
        const fontPt = fontInfo.fontPt || '';
        const source = fontInfo.fontSource || '';
        const fallback = fontInfo.styleFallbackSource || '';
        const parts = [
            `pt=${fontPt}`,
            `src=${source}`,
            `fallback=${fallback}`,
            `line=${lineSpacing}`,
            `color=${color}`,
            `b=${bold}`,
            `i=${italic}`
        ];
        const labelParts = [
            fontPt ? `${fontPt} pt` : 'unknown pt',
            color ? `color ${color}` : ''
        ].filter(Boolean);

        return {
            styleKey: parts.join('|'),
            styleLabel: labelParts.join(' | '),
            styleTraits: { fontPt, source, fallback, lineSpacing, color, bold, italic }
        };
    }

    function getParagraphLevel(paragraph) {
        const pPr = getDirectChild(paragraph, 'a:pPr');
        const level = pPr ? parseInt(pPr.getAttribute('lvl'), 10) : 0;
        return Number.isFinite(level) && level >= 0 ? level + 1 : 1;
    }

    function findLevelDefRPr(listStyle, level) {
        if (!listStyle) return null;
        const lvlPr = getDirectChild(listStyle, `a:lvl${level}pPr`)
            || getDirectChild(listStyle, 'a:lvl1pPr');
        return lvlPr ? getDirectChild(lvlPr, 'a:defRPr') : null;
    }

    function fontInfoFromLevelStyle(listStyle, level, fontSource) {
        const defRPr = findLevelDefRPr(listStyle, level);
        return defRPr ? fontInfoFromElement(defRPr, fontSource) : null;
    }

    function getPlaceholderInfo(element) {
        const ph = getFirstDescendant(element, 'p:ph');
        if (!ph) return { type: '', idx: '', hasPlaceholder: false };
        return {
            type: ph.getAttribute('type') || '',
            idx: ph.getAttribute('idx') || '',
            hasPlaceholder: true
        };
    }

    function placeholderInfoToTextStyle(placeholderInfo) {
        if (!placeholderInfo.hasPlaceholder) return 'otherStyle';
        const value = String(placeholderInfo.type || '').toLowerCase();
        if (value === 'title' || value === 'ctrtitle') return 'titleStyle';
        if (!value || value === 'body' || value === 'subtitle' || value === 'obj') return 'bodyStyle';
        return 'otherStyle';
    }

    function getShapeListStyle(container) {
        const txBody = getDirectChild(container.element, 'p:txBody');
        return txBody ? getDirectChild(txBody, 'a:lstStyle') : null;
    }

    function getParagraphDirectFontInfo(paragraph, sourcePrefix) {
        const runs = Array.from(paragraph.getElementsByTagName('a:r'));
        for (const run of runs) {
            const text = getTextFromRun(run).trim();
            const rPr = getDirectChild(run, 'a:rPr');
            if (!text || !rPr || !rPr.getAttribute('sz')) continue;
            const info = fontInfoFromElement(rPr, `${sourcePrefix}:rPr`);
            if (info) return info;
        }

        const endParaRPr = getDirectChild(paragraph, 'a:endParaRPr');
        if (endParaRPr && endParaRPr.getAttribute('sz')) {
            const info = fontInfoFromElement(endParaRPr, `${sourcePrefix}:endParaRPr`);
            if (info) return info;
        }

        const pPr = getDirectChild(paragraph, 'a:pPr');
        const defRPr = pPr ? getDirectChild(pPr, 'a:defRPr') : null;
        if (defRPr && defRPr.getAttribute('sz')) {
            const info = fontInfoFromElement(defRPr, `${sourcePrefix}:defRPr`);
            if (info) return info;
        }

        return { fontPt: '', fontSource: 'missing' };
    }

    function getTemplateShapeFontInfo(shape, paragraphIndex, sourcePrefix) {
        if (!shape) return null;
        const paragraphs = Array.from(shape.getElementsByTagName('a:p'));
        const paragraph = paragraphs[paragraphIndex] || paragraphs[0];
        if (!paragraph) return null;
        const info = getParagraphDirectFontInfo(paragraph, sourcePrefix);
        return info.fontPt ? info : null;
    }

    function findPlaceholderShape(doc, placeholderInfo) {
        if (!doc || (!placeholderInfo.idx && !placeholderInfo.type)) return null;
        const shapes = Array.from(doc.getElementsByTagName('*')).filter(element => {
            const tagName = getTagName(element);
            return tagName === 'p:sp' || tagName === 'p:graphicFrame';
        });

        if (placeholderInfo.idx) {
            const matchByIdx = shapes.find(shape => getPlaceholderInfo(shape).idx === placeholderInfo.idx);
            if (matchByIdx) return matchByIdx;
        }

        if (placeholderInfo.type) {
            const wantedType = placeholderInfo.type.toLowerCase();
            return shapes.find(shape => getPlaceholderInfo(shape).type.toLowerCase() === wantedType) || null;
        }

        return null;
    }

    function normalizePartPath(path) {
        const parts = [];
        for (const part of String(path || '').split('/')) {
            if (!part || part === '.') continue;
            if (part === '..') {
                parts.pop();
            } else {
                parts.push(part);
            }
        }
        return parts.join('/');
    }

    function resolveRelationshipTarget(sourcePart, target) {
        if (!target || /^[a-z]+:/i.test(target)) return '';
        if (target.startsWith('/')) return normalizePartPath(target.slice(1));
        const base = sourcePart.split('/').slice(0, -1).join('/');
        return normalizePartPath(`${base}/${target}`);
    }

    async function readXmlPart(zip, path) {
        const file = zip.file(path);
        if (!file) return null;
        return parseXml(await file.async('string'));
    }

    async function getRelationships(zip, sourcePart) {
        const name = sourcePart.split('/').pop();
        const dir = sourcePart.split('/').slice(0, -1).join('/');
        const relsPath = `${dir}/_rels/${name}.rels`;
        const relsDoc = await readXmlPart(zip, relsPath);
        if (!relsDoc) return [];

        return Array.from(relsDoc.getElementsByTagName('Relationship')).map(rel => ({
            id: rel.getAttribute('Id') || '',
            type: rel.getAttribute('Type') || '',
            target: resolveRelationshipTarget(sourcePart, rel.getAttribute('Target') || '')
        })).filter(rel => rel.target);
    }

    async function getRelatedPart(zip, sourcePart, typeSuffix) {
        const rels = await getRelationships(zip, sourcePart);
        const rel = rels.find(item => item.type.endsWith(typeSuffix));
        return rel ? rel.target : '';
    }

    function parseDefaultTextStyle(doc) {
        const style = getFirstDescendant(doc, 'p:defaultTextStyle')
            || getFirstDescendant(doc, 'a:defaultTextStyle');
        return style || null;
    }

    function parseMasterTextStyles(doc) {
        const txStyles = getFirstDescendant(doc, 'p:txStyles');
        const result = {};
        if (!txStyles) return result;

        for (const styleName of ['titleStyle', 'bodyStyle', 'otherStyle']) {
            result[styleName] = getDirectChild(txStyles, `p:${styleName}`);
        }

        return result;
    }

    async function buildStyleContext(zip, slides) {
        const context = {
            slideRels: new Map(),
            layoutDocs: new Map(),
            masterStyles: new Map(),
            presentationDefaultStyle: null
        };

        const presentationDoc = await readXmlPart(zip, 'ppt/presentation.xml');
        context.presentationDefaultStyle = presentationDoc ? parseDefaultTextStyle(presentationDoc) : null;

        for (const slide of slides) {
            const layoutPath = await getRelatedPart(zip, slide.path, '/slideLayout');
            const masterPath = layoutPath ? await getRelatedPart(zip, layoutPath, '/slideMaster') : '';
            context.slideRels.set(slide.path, { layoutPath, masterPath });

            if (layoutPath && !context.layoutDocs.has(layoutPath)) {
                context.layoutDocs.set(layoutPath, await readXmlPart(zip, layoutPath));
            }

            if (masterPath && !context.masterStyles.has(masterPath)) {
                const masterDoc = await readXmlPart(zip, masterPath);
                context.masterStyles.set(masterPath, masterDoc ? parseMasterTextStyles(masterDoc) : {});
            }
        }

        return context;
    }

    function itemSafeParagraphIndex(container, paragraph) {
        return Math.max(0, Array.from(container.paragraphs || []).indexOf(paragraph));
    }

    function resolveInheritedFontInfo(paragraph, container, styleContext) {
        const level = getParagraphLevel(paragraph);
        const shapeListStyle = getShapeListStyle(container);
        const shapeInfo = fontInfoFromLevelStyle(shapeListStyle, level, `shape:lstStyle:lvl${level}`);
        if (shapeInfo) return shapeInfo;

        const rels = styleContext.slideRels.get(container.slide.path) || {};
        const placeholderInfo = getPlaceholderInfo(container.element);
        const layoutDoc = rels.layoutPath ? styleContext.layoutDocs.get(rels.layoutPath) : null;
        const layoutShape = findPlaceholderShape(layoutDoc, placeholderInfo);
        const layoutDirectInfo = getTemplateShapeFontInfo(layoutShape, itemSafeParagraphIndex(container, paragraph), `layout:${rels.layoutPath || 'unknown'}`);
        if (layoutDirectInfo) return layoutDirectInfo;

        const layoutListStyle = layoutShape ? getShapeListStyle({ element: layoutShape }) : null;
        const layoutInfo = fontInfoFromLevelStyle(layoutListStyle, level, `layout:${rels.layoutPath || 'unknown'}:lstStyle:lvl${level}`);
        if (layoutInfo) return layoutInfo;

        const textStyle = placeholderInfoToTextStyle(placeholderInfo);
        const masterStyles = rels.masterPath ? styleContext.masterStyles.get(rels.masterPath) : null;
        const masterStyle = masterStyles ? masterStyles[textStyle] : null;
        const masterInfo = fontInfoFromLevelStyle(masterStyle, level, `master:${rels.masterPath || 'unknown'}:${textStyle}:lvl${level}`);
        if (masterInfo) return masterInfo;

        const defaultInfo = fontInfoFromLevelStyle(styleContext.presentationDefaultStyle, level, `presentation:defaultTextStyle:lvl${level}`);
        if (defaultInfo) return defaultInfo;

        return { fontPt: '', fontSource: 'missing' };
    }

    function getParagraphFontInfo(paragraph, container, styleContext) {
        const inlineInfo = getParagraphDirectFontInfo(paragraph, 'slide');
        const inheritedInfo = resolveInheritedFontInfo(paragraph, container, styleContext);
        const styleFallbackSource = inheritedInfo.fontPt ? inheritedInfo.fontSource : '';

        if (inlineInfo.fontPt) {
            return {
                ...inlineInfo,
                styleFallbackSource
            };
        }

        return {
            ...inheritedInfo,
            styleFallbackSource
        };
    }

    function collectTextContainers(doc, slide) {
        const elements = Array.from(doc.getElementsByTagName('*'));
        const containers = [];

        for (const element of elements) {
            const tagName = getTagName(element);
            if (tagName !== 'p:sp' && tagName !== 'p:graphicFrame') continue;

            const cNvPr = getFirstDescendant(element, 'p:cNvPr');
            const paragraphs = Array.from(element.getElementsByTagName('a:p'));
            const nonEmptyParagraphs = paragraphs
                .map((paragraph, paragraphIndex) => ({
                    paragraph,
                    paragraphIndex,
                    text: getTextFromParagraph(paragraph)
                }))
                .filter(item => item.text.trim());

            if (nonEmptyParagraphs.length === 0) continue;

            containers.push({
                element,
                cNvPr,
                slide,
                containerType: tagName === 'p:graphicFrame' ? 'graphicFrame' : 'shape',
                shapeId: cNvPr ? (cNvPr.getAttribute('id') || '') : '',
                shapeName: cNvPr ? (cNvPr.getAttribute('name') || '') : '',
                paragraphs,
                nonEmptyParagraphs
            });
        }

        return containers;
    }

    async function extractStyleRowsFromPptx(arrayBuffer) {
        const zip = await JSZip.loadAsync(arrayBuffer);
        const slides = getSlideEntries(zip);
        const styleContext = await buildStyleContext(zip, slides);
        const styleRows = [];

        for (const slide of slides) {
            const file = zip.file(slide.path);
            if (!file) continue;

            const xml = await file.async('string');
            const doc = parseXml(xml);
            const containers = collectTextContainers(doc, slide);

            for (const container of containers) {
                const shapeName = container.shapeName.trim();
                const keyBase = shapeName || `unmarked_id${container.shapeId || 'unknown'}`;
                for (const item of container.nonEmptyParagraphs) {
                    const fontInfo = getParagraphFontInfo(item.paragraph, container, styleContext);
                    const styleInfo = getParagraphStyleFingerprint(item.paragraph, fontInfo);
                    styleRows.push({
                        key: `${keyBase}__p${item.paragraphIndex}`,
                        shapeName,
                        paragraphIndex: item.paragraphIndex,
                        textSample: item.text,
                        fontPt: fontInfo.fontPt || '',
                        fontSource: fontInfo.fontSource || '',
                        styleFallbackSource: fontInfo.styleFallbackSource || '',
                        styleKey: styleInfo.styleKey,
                        styleLabel: styleInfo.styleLabel,
                        styleTraits: styleInfo.styleTraits
                    });
                }
            }
        }

        return styleRows;
    }

    window.PptxStyleUtils = {
        extractStyleRowsFromPptx
    };
})();
