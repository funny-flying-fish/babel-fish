(() => {
    const PPTX_MIME = 'application/vnd.openxmlformats-officedocument.presentationml.presentation';
    const XLSX_MIME = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
    const METADATA_COLUMNS = ['key', 'slideNumber', 'shapeName', 'shapeId', 'containerType', 'paragraphIndex'];
    const STYLE_METRIC_SUFFIXES = ['chars', 'fontPt', 'fontRef', 'fontSource', 'styleFallbackSource', 'styleKey', 'styleLabel', 'densityRatio', 'fontScale'];
    const A_NS = 'http://schemas.openxmlformats.org/drawingml/2006/main';
    const TRANSLATION_NAME_RE = /^tr_[a-f0-9]{6,12}_[a-z][a-z0-9]*$/i;

    // DOM refs: Prepare / Markup PPTX
    const pptxMarkupDropzone = document.getElementById('pptxMarkupDropzone');
    const pptxMarkupInput = document.getElementById('pptxMarkupInput');
    const pptxMarkupFilename = document.getElementById('pptxMarkupFilename');
    const markupBtn = document.getElementById('markupBtn');
    const markupDownload = document.getElementById('markupDownload');
    const markupError = document.getElementById('markupError');

    // DOM refs: PPTX -> XLSX
    const pptxExtractDropzone = document.getElementById('pptxExtractDropzone');
    const pptxExtractInput = document.getElementById('pptxExtractInput');
    const pptxExtractFilename = document.getElementById('pptxExtractFilename');
    const sourceLanguage = document.getElementById('sourceLanguage');
    const extractBtn = document.getElementById('extractBtn');
    const extractDownload = document.getElementById('extractDownload');
    const extractError = document.getElementById('extractError');

    // DOM refs: Style controls / per-language PPTX generation
    const stylePptxDropzone = document.getElementById('stylePptxDropzone');
    const stylePptxInput = document.getElementById('stylePptxInput');
    const stylePptxFilename = document.getElementById('stylePptxFilename');
    const styleXlsxDropzone = document.getElementById('styleXlsxDropzone');
    const styleXlsxInput = document.getElementById('styleXlsxInput');
    const styleXlsxFilename = document.getElementById('styleXlsxFilename');
    const stylePresetControls = document.getElementById('stylePresetControls');
    const styleSavePresetBtn = document.getElementById('styleSavePresetBtn');
    const styleLoadPresetBtn = document.getElementById('styleLoadPresetBtn');
    const stylePresetInput = document.getElementById('stylePresetInput');
    const styleTableArea = document.getElementById('styleTableArea');
    const styleResult = document.getElementById('styleResult');
    const styleError = document.getElementById('styleError');

    let markupFile = null;
    let extractFile = null;
    let stylePptxBuffer = null;
    let stylePptxName = '';
    let styleTranslationData = null;
    let styleState = null;
    let pendingStylePreset = null;

    function triggerBlobDownload(blob, filename) {
        const url = URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = filename;
        a.style.display = 'none';
        document.body.appendChild(a);
        a.click();
        setTimeout(() => {
            document.body.removeChild(a);
            URL.revokeObjectURL(url);
        }, 200);
    }

    function createDownloadLink(blob, filename, label) {
        const link = document.createElement('a');
        link.href = '#';
        link.textContent = label || `Download ${filename}`;
        link.style.color = '#4da6ff';
        link.addEventListener('click', (e) => {
            e.preventDefault();
            triggerBlobDownload(blob, filename);
        });
        return link;
    }

    function readFileAsArrayBuffer(file) {
        return new Promise((resolve, reject) => {
            const reader = new FileReader();
            reader.onload = () => resolve(reader.result);
            reader.onerror = reject;
            reader.readAsArrayBuffer(file);
        });
    }

    function readFileAsText(file) {
        return new Promise((resolve, reject) => {
            const reader = new FileReader();
            reader.onload = () => resolve(reader.result);
            reader.onerror = reject;
            reader.readAsText(file);
        });
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

    function parseXml(xml) {
        const doc = new DOMParser().parseFromString(xml, 'application/xml');
        const parserError = doc.getElementsByTagName('parsererror')[0];
        if (parserError) throw new Error(parserError.textContent || 'Invalid XML in PPTX');
        return doc;
    }

    function getTextFromParagraph(aParagraph) {
        const runs = aParagraph.getElementsByTagName('a:r');
        let text = '';
        for (let i = 0; i < runs.length; i++) {
            const tNodes = runs[i].getElementsByTagName('a:t');
            for (let j = 0; j < tNodes.length; j++) {
                text += tNodes[j].textContent || '';
            }
        }
        return text;
    }

    function getTextFromRun(run) {
        const tNodes = run.getElementsByTagName('a:t');
        let text = '';
        for (let i = 0; i < tNodes.length; i++) {
            text += tNodes[i].textContent || '';
        }
        return text;
    }

    function countCharacters(value) {
        return Array.from(String(value || '')).length;
    }

    function parseNumber(value) {
        if (value === undefined || value === null || value === '') return null;
        const num = Number(String(value).replace(',', '.'));
        return Number.isFinite(num) ? num : null;
    }

    function roundToHalf(value) {
        return Math.round(value * 2) / 2;
    }

    function clampNumber(value, min, max) {
        let result = value;
        if (Number.isFinite(min)) result = Math.max(result, min);
        if (Number.isFinite(max)) result = Math.min(result, max);
        return result;
    }

    function metricColumnName(lang, suffix) {
        return `${lang}_${suffix}`;
    }

    function isStyleMetricColumn(name) {
        const value = String(name || '').trim();
        return STYLE_METRIC_SUFFIXES.some(suffix => value.endsWith(`_${suffix}`));
    }

    function isLanguageColumn(name) {
        const value = String(name || '').trim();
        const metadataColumns = METADATA_COLUMNS.map(column => column.toLowerCase());
        return Boolean(value) && !metadataColumns.includes(value.toLowerCase()) && !isStyleMetricColumn(value);
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

    function getTagName(element) {
        return element.tagName || element.nodeName || '';
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

    function isTranslationName(name) {
        return TRANSLATION_NAME_RE.test(String(name || '').trim());
    }

    function isDefaultPowerPointName(name) {
        const value = String(name || '').trim();
        if (!value) return true;
        return /^(TextBox|Text|Shape|Rectangle|Oval|Line|Picture|Graphic|Table|Content Placeholder|Title Placeholder|Subtitle Placeholder|Slide Number Placeholder)\s*\d*$/i.test(value);
    }

    function generateShortId(usedIds) {
        let id = '';
        do {
            if (window.crypto && window.crypto.getRandomValues) {
                const bytes = new Uint8Array(3);
                window.crypto.getRandomValues(bytes);
                id = Array.from(bytes).map(byte => byte.toString(16).padStart(2, '0')).join('');
            } else {
                id = Math.floor(Math.random() * 0xffffff).toString(16).padStart(6, '0');
            }
        } while (usedIds.has(id));
        usedIds.add(id);
        return id;
    }

    function getTranslationId(name) {
        const m = String(name || '').match(/^tr_([a-f0-9]{6,12})_/i);
        return m ? m[1].toLowerCase() : '';
    }

    function classifyRole(container) {
        if (container.containerType === 'graphicFrame') return 'table';

        const placeholder = getFirstDescendant(container.element, 'p:ph');
        const placeholderType = placeholder ? (placeholder.getAttribute('type') || '') : '';
        if (/title/i.test(placeholderType)) return 'title';
        if (/subTitle/i.test(placeholderType)) return 'subtitle';

        if (container.nonEmptyParagraphs.length > 1) return 'list';

        const text = container.nonEmptyParagraphs[0].text.trim();
        if (text.length <= 24 || /^[\d\s.,:%/+~-]+$/.test(text)) return 'label';

        return 'body';
    }

    function makeShapeName(container, usedIds) {
        const role = classifyRole(container);
        return `tr_${generateShortId(usedIds)}_${role}`;
    }

    function createWarning(type, message, details = {}) {
        return {
            severity: details.severity || 'warning',
            type,
            message,
            key: details.key || '',
            slideNumber: details.slideNumber || '',
            shapeName: details.shapeName || '',
            shapeId: details.shapeId || '',
            containerType: details.containerType || '',
            paragraphIndex: details.paragraphIndex ?? ''
        };
    }

    function createWorkbookBlob(rows, sheetName = 'Sheet1') {
        const ws = XLSX.utils.aoa_to_sheet(rows);
        const wb = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(wb, ws, sheetName);
        const data = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
        return new Blob([data], { type: XLSX_MIME });
    }

    function createWarningsBlob(warnings) {
        const rows = [[
            'severity',
            'type',
            'message',
            'key',
            'slideNumber',
            'shapeName',
            'shapeId',
            'containerType',
            'paragraphIndex'
        ]];

        for (const warning of warnings) {
            rows.push([
                warning.severity,
                warning.type,
                warning.message,
                warning.key,
                warning.slideNumber,
                warning.shapeName,
                warning.shapeId,
                warning.containerType,
                warning.paragraphIndex
            ]);
        }

        return createWorkbookBlob(rows, 'Warnings');
    }

    function renderResult(target, summaryLines, warnings, reportBaseName) {
        const info = document.createElement('div');
        info.style.cssText = 'color: #6bbd6b; font-size: 0.85rem; margin-top: 6px;';
        info.innerHTML = summaryLines.map(line => `<div>${line}</div>`).join('');
        target.appendChild(info);

        if (warnings.length === 0) return;

        const warningBox = document.createElement('div');
        warningBox.style.cssText = 'color: #e8b64c; font-size: 0.85rem; margin-top: 8px;';
        warningBox.appendChild(document.createTextNode(`${warnings.length} warning(s):`));

        const list = document.createElement('ul');
        list.style.cssText = 'margin: 6px 0 0 18px;';
        for (const warning of warnings.slice(0, 8)) {
            const item = document.createElement('li');
            item.textContent = warning.message;
            list.appendChild(item);
        }
        if (warnings.length > 8) {
            const item = document.createElement('li');
            item.textContent = `...and ${warnings.length - 8} more`;
            list.appendChild(item);
        }
        warningBox.appendChild(list);

        const reportBlob = createWarningsBlob(warnings);
        const reportLink = createDownloadLink(reportBlob, `${reportBaseName}_warnings.xlsx`, 'Download warning report');
        const reportWrap = document.createElement('div');
        reportWrap.style.marginTop = '6px';
        reportWrap.appendChild(reportLink);
        warningBox.appendChild(reportWrap);

        target.appendChild(warningBox);
    }

    async function markupPptxIds(arrayBuffer) {
        const zip = await JSZip.loadAsync(arrayBuffer);
        const slides = getSlideEntries(zip);
        const warnings = [];
        const usedIds = new Set();
        const seenNames = new Set();
        let containerCount = 0;
        let renamedCount = 0;
        let keptCount = 0;

        for (const slide of slides) {
            const file = zip.file(slide.path);
            if (!file) continue;

            const xml = await file.async('string');
            const doc = parseXml(xml);
            const containers = collectTextContainers(doc, slide);
            let modified = false;

            for (const container of containers) {
                containerCount++;

                if (!container.cNvPr) {
                    warnings.push(createWarning('missing-cNvPr', `Slide ${slide.num}: skipped text container without cNvPr.`, {
                        slideNumber: slide.num,
                        containerType: container.containerType
                    }));
                    continue;
                }

                const currentName = container.shapeName.trim();
                const currentId = getTranslationId(currentName);
                if (currentId) usedIds.add(currentId);

                const shouldKeep = isTranslationName(currentName) && !seenNames.has(currentName);
                if (shouldKeep) {
                    seenNames.add(currentName);
                    keptCount++;
                    continue;
                }

                if (isTranslationName(currentName) && seenNames.has(currentName)) {
                    warnings.push(createWarning('duplicate-shape-name', `Slide ${slide.num}: duplicate ${currentName} renamed.`, {
                        slideNumber: slide.num,
                        shapeName: currentName,
                        shapeId: container.shapeId,
                        containerType: container.containerType
                    }));
                } else if (!isDefaultPowerPointName(currentName)) {
                    warnings.push(createWarning('custom-name-replaced', `Slide ${slide.num}: custom name "${currentName}" replaced with translation ID.`, {
                        severity: 'info',
                        slideNumber: slide.num,
                        shapeName: currentName,
                        shapeId: container.shapeId,
                        containerType: container.containerType
                    }));
                }

                const newName = makeShapeName(container, usedIds);
                container.cNvPr.setAttribute('name', newName);
                seenNames.add(newName);
                renamedCount++;
                modified = true;
            }

            if (modified) {
                zip.file(slide.path, new XMLSerializer().serializeToString(doc));
            }
        }

        const blob = await zip.generateAsync({ type: 'blob', mimeType: PPTX_MIME });
        return { blob, warnings, containerCount, renamedCount, keptCount };
    }

    async function handleMarkup() {
        markupError.textContent = '';
        markupDownload.innerHTML = '';

        if (!markupFile) return;

        try {
            markupBtn.disabled = true;
            markupBtn.textContent = 'Preparing...';

            const buffer = await readFileAsArrayBuffer(markupFile);
            const result = await markupPptxIds(buffer);
            const baseName = markupFile.name.replace(/\.pptx$/i, '');
            const filename = `${baseName}_marked.pptx`;

            markupDownload.appendChild(createDownloadLink(result.blob, filename));
            renderResult(markupDownload, [
                `Marked ${result.renamedCount} of ${result.containerCount} text containers.`,
                `Kept ${result.keptCount} existing translation IDs.`
            ], result.warnings, `${baseName}_markup`);
        } catch (err) {
            markupError.textContent = 'Error: ' + err.message;
        } finally {
            markupBtn.disabled = false;
            markupBtn.textContent = 'Prepare PPTX IDs';
        }
    }

    async function extractTextFromPptx(arrayBuffer, lang) {
        const zip = await JSZip.loadAsync(arrayBuffer);
        const slides = getSlideEntries(zip);
        const rows = [['key', lang]];
        const warnings = [];
        const nameCounts = new Map();
        let containerCount = 0;

        for (const slide of slides) {
            const file = zip.file(slide.path);
            if (!file) continue;

            const xml = await file.async('string');
            const doc = parseXml(xml);
            const containers = collectTextContainers(doc, slide);

            for (const container of containers) {
                containerCount++;
                const shapeName = container.shapeName.trim();
                if (shapeName) {
                    nameCounts.set(shapeName, (nameCounts.get(shapeName) || 0) + 1);
                }

                if (!isTranslationName(shapeName)) {
                    warnings.push(createWarning('unmarked-shape', `Slide ${slide.num}: "${shapeName || '(unnamed)'}" is not a translation ID. Run Prepare PPTX IDs first.`, {
                        slideNumber: slide.num,
                        shapeName,
                        shapeId: container.shapeId,
                        containerType: container.containerType
                    }));
                }

                for (const item of container.nonEmptyParagraphs) {
                    const keyBase = shapeName || `unmarked_id${container.shapeId || 'unknown'}`;
                    rows.push([
                        `${keyBase}__p${item.paragraphIndex}`,
                        item.text
                    ]);
                }
            }
        }

        for (const [shapeName, count] of nameCounts.entries()) {
            if (count > 1) {
                warnings.push(createWarning('duplicate-shape-name', `Duplicate shapeName "${shapeName}" found ${count} times. Run Prepare PPTX IDs to normalize duplicates.`, {
                    shapeName
                }));
            }
        }

        return { rows, warnings, containerCount };
    }

    async function handleExtract() {
        extractError.textContent = '';
        extractDownload.innerHTML = '';

        if (!extractFile) return;

        try {
            extractBtn.disabled = true;
            extractBtn.textContent = 'Extracting...';

            const buffer = await readFileAsArrayBuffer(extractFile);
            const lang = sourceLanguage.value.trim() || 'EN';
            const { rows, warnings, containerCount } = await extractTextFromPptx(buffer, lang);

            const baseName = extractFile.name.replace(/\.pptx$/i, '');
            const blob = createWorkbookBlob(rows, 'Translations');
            extractDownload.appendChild(createDownloadLink(blob, `${baseName}.xlsx`));
            renderResult(extractDownload, [
                `Extracted ${rows.length - 1} paragraph entries from ${containerCount} text containers.`
            ], warnings, `${baseName}_extract`);
        } catch (err) {
            extractError.textContent = 'Error: ' + err.message;
        } finally {
            extractBtn.disabled = false;
            extractBtn.textContent = 'Extract text';
        }
    }

    function parseXlsxFile(file) {
        return new Promise((resolve, reject) => {
            const reader = new FileReader();
            reader.onload = (e) => {
                try {
                    const data = new Uint8Array(e.target.result);
                    const wb = XLSX.read(data, { type: 'array' });
                    const sheet = wb.Sheets[wb.SheetNames[0]];
                    resolve(XLSX.utils.sheet_to_json(sheet, { header: 1 }));
                } catch (err) {
                    reject(err);
                }
            };
            reader.onerror = reject;
            reader.readAsArrayBuffer(file);
        });
    }

    function buildHeaderIndex(headerRow) {
        const index = new Map();
        headerRow.forEach((value, i) => {
            const key = String(value || '').trim();
            if (!key) return;
            index.set(key, i);
            for (const column of METADATA_COLUMNS) {
                if (key.toLowerCase() === column.toLowerCase() && !index.has(column)) {
                    index.set(column, i);
                }
            }
        });
        return index;
    }

    function getLanguageColumns(headerRow) {
        const languages = [];
        for (let i = 0; i < headerRow.length; i++) {
            const code = String(headerRow[i] || '').trim();
            if (isLanguageColumn(code)) languages.push({ code, index: i });
        }
        return languages;
    }

    function getRowIdentity(row, headerIndex) {
        const key = String(row[headerIndex.get('key')] || '').trim();
        if (key) return key;

        const shapeName = String(row[headerIndex.get('shapeName')] || '').trim();
        const paragraphIndex = String(row[headerIndex.get('paragraphIndex')] ?? '').trim();
        return shapeName && paragraphIndex ? `${shapeName}__p${paragraphIndex}` : '';
    }

    function getOptionalCell(row, headerIndex, column) {
        const index = headerIndex.get(column);
        return index === undefined ? '' : row[index];
    }

    function parseTranslationKey(key) {
        const match = String(key || '').trim().match(/^(.+)__p(\d+)$/);
        if (!match) return { shapeName: '', paragraphIndex: NaN };
        return {
            shapeName: match[1],
            paragraphIndex: parseInt(match[2], 10)
        };
    }

    function buildRowLookup(data, headerIndex) {
        const lookup = new Map();
        for (let r = 1; r < data.length; r++) {
            const row = data[r] || [];
            const identity = getRowIdentity(row, headerIndex);
            if (identity && !lookup.has(identity)) lookup.set(identity, { row, rowNumber: r + 1 });
        }
        return lookup;
    }

    function buildStyleGroups(styleRows) {
        const groupsByKey = new Map();
        for (const row of styleRows) {
            if (!groupsByKey.has(row.styleKey)) {
                groupsByKey.set(row.styleKey, {
                    styleKey: row.styleKey,
                    styleLabel: row.styleLabel,
                    fontPt: row.fontPt,
                    count: 0,
                    examples: []
                });
            }
            const group = groupsByKey.get(row.styleKey);
            group.count++;
            if (group.examples.length < 3) {
                group.examples.push({
                    text: row.textSample,
                    traits: row.styleTraits || {}
                });
            }
        }

        return Array.from(groupsByKey.values()).sort((a, b) => {
            const aPt = parseNumber(a.fontPt) || 0;
            const bPt = parseNumber(b.fontPt) || 0;
            if (aPt !== bPt) return aPt - bPt;
            return a.styleLabel.localeCompare(b.styleLabel);
        });
    }

    function buildStyleState(styleRows, xlsxData = null) {
        const header = xlsxData && xlsxData.length ? xlsxData[0].map(value => String(value || '').trim()) : [];
        const languages = header.length ? getLanguageColumns(header) : [];
        return {
            styleRows,
            styleRowsByKey: new Map(styleRows.map(row => [row.key, row])),
            groups: buildStyleGroups(styleRows),
            languages,
            sizeInputs: new Map(),
            multiplierSelects: new Map()
        };
    }

    function createMultiplierSelect() {
        const select = document.createElement('select');
        for (let pct = 100; pct >= 50; pct -= 2) {
            const option = document.createElement('option');
            option.value = String(pct);
            option.textContent = `${pct}%`;
            select.appendChild(option);
        }
        return select;
    }

    function getStyleSizeForLanguage(languageCode, styleKey) {
        const languageInputs = styleState && styleState.sizeInputs.get(languageCode);
        if (!languageInputs) return '';
        const input = languageInputs.get(styleKey);
        return input ? parseNumber(input.value) : '';
    }

    function getStyleMultiplierForLanguage(languageCode) {
        const select = styleState && styleState.multiplierSelects.get(languageCode);
        return select ? (parseNumber(select.value) || 100) / 100 : 1;
    }

    function updateStylePresetButton() {
        if (!styleSavePresetBtn) return;
        styleSavePresetBtn.disabled = !(styleState && stylePptxBuffer && styleTranslationData && styleState.languages.length);
    }

    function recalculateStyleInputsForLanguage(languageCode) {
        const languageInputs = styleState && styleState.sizeInputs.get(languageCode);
        if (!languageInputs) return;

        const multiplier = getStyleMultiplierForLanguage(languageCode);
        for (const input of languageInputs.values()) {
            const baseFontPt = parseNumber(input.dataset.baseFontPt);
            input.value = baseFontPt ? roundToHalf(baseFontPt * multiplier) : '';
        }
    }

    function createStylePreset() {
        if (!styleState) return null;
        const languages = {};

        for (const language of styleState.languages) {
            const select = styleState.multiplierSelects.get(language.code);
            const languageInputs = styleState.sizeInputs.get(language.code) || new Map();
            const sizes = {};
            for (const [styleKey, input] of languageInputs.entries()) {
                const value = parseNumber(input.value);
                if (value) sizes[styleKey] = value;
            }
            languages[language.code] = {
                multiplier: select ? parseNumber(select.value) || 100 : 100,
                sizes
            };
        }

        return {
            version: 1,
            type: 'pptx-style-size-preset',
            languages
        };
    }

    function saveStylePreset() {
        const preset = createStylePreset();
        if (!preset) return;
        const blob = new Blob([JSON.stringify(preset, null, 2)], { type: 'application/json' });
        const baseName = stylePptxName ? stylePptxName.replace(/\.pptx$/i, '') : 'pptx';
        triggerBlobDownload(blob, `${baseName}_style-preset.json`);
    }

    function validateStylePreset(value) {
        if (!value || value.type !== 'pptx-style-size-preset' || value.version !== 1 || typeof value.languages !== 'object') {
            throw new Error('Invalid style preset JSON.');
        }
        return value;
    }

    function applyStylePreset(preset) {
        pendingStylePreset = preset;
        if (!styleState || !styleState.languages.length || !styleState.sizeInputs.size) return;

        for (const language of styleState.languages) {
            const languagePreset = preset.languages[language.code];
            if (!languagePreset) continue;

            const select = styleState.multiplierSelects.get(language.code);
            if (select && languagePreset.multiplier) {
                const value = String(languagePreset.multiplier);
                if (Array.from(select.options).some(option => option.value === value)) {
                    select.value = value;
                }
            }

            const languageInputs = styleState.sizeInputs.get(language.code);
            if (!languageInputs || !languagePreset.sizes) continue;
            for (const [styleKey, size] of Object.entries(languagePreset.sizes)) {
                const input = languageInputs.get(styleKey);
                const value = parseNumber(size);
                if (input && value) input.value = value;
            }
        }
    }

    async function loadStylePreset(file) {
        const text = await readFileAsText(file);
        const preset = validateStylePreset(JSON.parse(text));
        applyStylePreset(preset);
    }

    function getReadableTextColor(backgroundColor) {
        const value = String(backgroundColor || '').replace('#', '');
        if (!/^[0-9a-f]{6}$/i.test(value)) return '#111111';
        const r = parseInt(value.slice(0, 2), 16);
        const g = parseInt(value.slice(2, 4), 16);
        const b = parseInt(value.slice(4, 6), 16);
        const luminance = (0.299 * r + 0.587 * g + 0.114 * b) / 255;
        return luminance > 0.5 ? '#111111' : '#ffffff';
    }

    function styleTraitsToCss(traits = {}) {
        const css = [];
        if (traits.fontPt) css.push(`font-size: ${traits.fontPt}pt`);
        if (traits.color && /^#[0-9a-f]{6}$/i.test(traits.color)) css.push(`color: ${traits.color}`);
        if (traits.bold === '1') css.push('font-weight: 700');
        if (traits.italic === '1') css.push('font-style: italic');
        if (traits.lineSpacing) css.push(`line-height: ${traits.lineSpacing}`);
        return css.join('; ');
    }

    function getContrastingColor(hexColor) {
        const value = String(hexColor || '').replace('#', '');
        if (!/^[0-9a-f]{6}$/i.test(value)) return '#ffffff';
        const r = parseInt(value.slice(0, 2), 16);
        const g = parseInt(value.slice(2, 4), 16);
        const b = parseInt(value.slice(4, 6), 16);
        const luminance = (0.299 * r + 0.587 * g + 0.114 * b) / 255;
        return luminance > 0.5 ? '#111111' : '#ffffff';
    }

    function getPreviewBackground(group) {
        const firstColor = (group.examples || [])
            .map(example => example.traits && example.traits.color)
            .find(color => /^#[0-9a-f]{6}$/i.test(color || ''));
        return firstColor ? getContrastingColor(firstColor) : '#ffffff';
    }

    function formatStyleDetails(traits = {}) {
        const details = [];
        if (traits.fontPt) details.push(`Size: ${traits.fontPt}`);
        if (traits.bold === '1') details.push('Weight: Bold');
        if (traits.italic === '1') details.push('Style: Italic');
        if (traits.lineSpacing) details.push(`line-height: ${traits.lineSpacing}`);
        if (traits.source) details.push(`Source: ${traits.source}`);
        if (traits.fallback) details.push(`Fallback: ${traits.fallback}`);
        return details;
    }

    function createStylePreviewCard(group, styleNumber, onClose) {
        const card = document.createElement('div');
        const cardBackground = getPreviewBackground(group);
        const cardTextColor = getReadableTextColor(cardBackground);
        card.style.cssText = [
            'position: relative',
            'padding: 10px 34px 10px 10px',
            'border: 1px solid #555',
            'border-radius: 8px',
            `background: ${cardBackground}`,
            `color: ${cardTextColor}`,
            'box-shadow: 0 4px 14px rgba(0,0,0,0.18)'
        ].join('; ');

        const closeButton = document.createElement('button');
        closeButton.type = 'button';
        closeButton.textContent = '×';
        closeButton.setAttribute('aria-label', `Hide style ${styleNumber}`);
        closeButton.style.cssText = 'position: absolute; top: 6px; right: 8px; background: none; border: 0; color: inherit; font-size: 1.1rem; line-height: 1; cursor: pointer;';
        closeButton.addEventListener('click', onClose);
        card.appendChild(closeButton);

        const number = document.createElement('div');
        number.textContent = `#${styleNumber}`;
        number.style.cssText = 'font-weight: 700; font-size: 0.95rem; margin-bottom: 6px;';
        card.appendChild(number);

        const details = formatStyleDetails(group.examples[0] ? group.examples[0].traits : {});
        if (details.length) {
            const detailsEl = document.createElement('div');
            detailsEl.style.cssText = 'font-size: 0.8rem; opacity: 0.6; margin-bottom: 8px;';
            for (const detail of details) {
                const line = document.createElement('div');
                line.textContent = detail;
                detailsEl.appendChild(line);
            }
            card.appendChild(detailsEl);
        }

        for (const example of group.examples) {
            const item = document.createElement('div');
            item.textContent = example.text || '(empty text)';
            item.style.cssText = `margin-top: 6px; padding: 4px 6px; border-radius: 4px; background: ${cardBackground}; ${styleTraitsToCss(example.traits)}`;
            card.appendChild(item);
        }

        return card;
    }

    function createStylePreviewLabel(group, styleNumber, onSelect) {
        const label = document.createElement('button');
        label.type = 'button';
        label.textContent = group.styleLabel;
        label.style.cssText = 'background: none; border: 0; color: inherit; padding: 0; text-align: left; cursor: pointer; text-decoration: underline dotted;';

        const popup = document.createElement('div');
        const popupBackground = getPreviewBackground(group);
        const popupTextColor = getReadableTextColor(popupBackground);
        popup.style.cssText = [
            'display: none',
            'position: fixed',
            'z-index: 9999',
            'min-width: 260px',
            'max-width: 420px',
            'padding: 10px',
            'border: 1px solid #555',
            'border-radius: 8px',
            `background: ${popupBackground}`,
            `color: ${popupTextColor}`,
            'box-shadow: 0 8px 24px rgba(0,0,0,0.3)'
        ].join('; ');

        const details = formatStyleDetails(group.examples[0] ? group.examples[0].traits : {});
        if (details.length) {
            const detailsEl = document.createElement('div');
            detailsEl.style.cssText = 'font-size: 0.8rem; opacity: 0.6; margin-bottom: 8px;';
            for (const detail of details) {
                const line = document.createElement('div');
                line.textContent = detail;
                detailsEl.appendChild(line);
            }
            popup.appendChild(detailsEl);
        }

        for (const example of group.examples) {
            const item = document.createElement('div');
            item.textContent = example.text || '(empty text)';
            item.style.cssText = `margin-top: 6px; padding: 4px 6px; border-radius: 4px; background: ${popupBackground}; ${styleTraitsToCss(example.traits)}`;
            popup.appendChild(item);
        }

        document.body.appendChild(popup);
        const show = () => {
            const rect = label.getBoundingClientRect();
            popup.style.display = 'block';
            const popupRect = popup.getBoundingClientRect();
            const preferredLeft = rect.right + 10;
            const left = preferredLeft + popupRect.width + 12 > window.innerWidth
                ? Math.max(12, rect.left - popupRect.width - 10)
                : preferredLeft;
            const top = rect.top + popupRect.height + 12 > window.innerHeight
                ? Math.max(12, window.innerHeight - popupRect.height - 12)
                : rect.top;
            popup.style.left = `${Math.max(12, left)}px`;
            popup.style.top = `${top}px`;
        };
        const hide = () => { popup.style.display = 'none'; };
        label.addEventListener('mouseenter', show);
        label.addEventListener('mouseleave', hide);
        label.addEventListener('focus', show);
        label.addEventListener('blur', hide);
        label.addEventListener('click', (e) => {
            e.preventDefault();
            onSelect(group, styleNumber);
        });
        return label;
    }

    function buildStyleGenerationRows(data, language, options = {}) {
        if (!data.length) throw new Error('XLSX is empty.');

        const header = data[0].map(value => String(value || '').trim());
        const headerIndex = buildHeaderIndex(header);
        if (!headerIndex.has('key')) throw new Error('XLSX is missing required column "key".');

        const fontColumn = metricColumnName(language.code, 'fontPt');
        const fontColumnIndex = headerIndex.get(fontColumn);
        const rows = [];
        let skippedEmpty = 0;

        for (let r = 1; r < data.length; r++) {
            const sourceRow = data[r] || [];
            const value = sourceRow[language.index];
            if (value === undefined || value === null || !String(value).trim()) {
                skippedEmpty++;
                continue;
            }

            const key = String(getOptionalCell(sourceRow, headerIndex, 'key') || '').trim();
            const parsedKey = parseTranslationKey(key);
            const styleRow = options.styleRowsByKey ? options.styleRowsByKey.get(key) : null;
            const groupFontPt = styleRow ? getStyleSizeForLanguage(language.code, styleRow.styleKey) : '';
            const fallbackFontPt = fontColumnIndex === undefined ? '' : parseNumber(sourceRow[fontColumnIndex]);
            const baseFontPt = groupFontPt || fallbackFontPt || '';
            const fontPt = baseFontPt ? roundToHalf(baseFontPt) : '';

            rows.push({
                rowNumber: r + 1,
                key,
                slideNumber: String(getOptionalCell(sourceRow, headerIndex, 'slideNumber') || '').trim(),
                shapeName: String(getOptionalCell(sourceRow, headerIndex, 'shapeName') || parsedKey.shapeName || '').trim(),
                shapeId: String(getOptionalCell(sourceRow, headerIndex, 'shapeId') || '').trim(),
                containerType: String(getOptionalCell(sourceRow, headerIndex, 'containerType') || '').trim(),
                paragraphIndex: parsedKey.paragraphIndex,
                text: String(value),
                fontPt,
                fontRaw: fontPt ? String(fontPt) : '',
                fontColumn: fontColumnIndex === undefined ? 'style table' : fontColumn
            });
        }

        return { rows, skippedEmpty };
    }

    async function downloadStyledPresentation(language) {
        if (!stylePptxBuffer || !styleTranslationData || !styleState) return;
        styleError.textContent = '';
        if (styleResult) styleResult.innerHTML = '';

        try {
            const { rows, skippedEmpty } = buildStyleGenerationRows(styleTranslationData, language, {
                styleRowsByKey: styleState.styleRowsByKey
            });
            if (!rows.length) {
                styleError.textContent = `No translations found for ${language.code}.`;
                return;
            }

            const result = await injectTranslations(stylePptxBuffer.slice(0), rows);
            const baseName = stylePptxName.replace(/\.pptx$/i, '') || 'presentation';
            triggerBlobDownload(result.blob, `${baseName}_${language.code}.pptx`);
            if (styleResult) {
                renderResult(styleResult, [
                    `Generated ${language.code}: replaced ${result.replacedCount} of ${rows.length} translated paragraph entries.`,
                    `Applied ${result.fontSizeAppliedCount} style-table font size value(s).`,
                    `Skipped ${skippedEmpty} empty translation cells.`
                ], result.warnings, `${baseName}_${language.code}`);
            }
        } catch (err) {
            styleError.textContent = 'Error: ' + err.message;
        }
    }

    function renderStyleControls() {
        if (!styleTableArea) return;
        styleTableArea.innerHTML = '';
        if (styleResult) styleResult.innerHTML = '';
        if (stylePresetControls) stylePresetControls.style.display = 'none';
        updateStylePresetButton();

        if (!stylePptxBuffer) {
            styleTableArea.textContent = 'Load a marked PPTX file first.';
            return;
        }

        if (!styleTranslationData) {
            styleTableArea.textContent = 'Load an XLSX file with language columns to build the table.';
            return;
        }

        if (!styleState || !styleState.languages.length) {
            styleTableArea.textContent = 'No language columns found in the XLSX.';
            return;
        }

        if (!styleState.groups.length) {
            styleTableArea.textContent = 'No editable text styles found in the PPTX.';
            return;
        }

        if (stylePresetControls) stylePresetControls.style.display = 'flex';

        const layout = document.createElement('div');
        layout.style.cssText = 'display: flex; align-items: flex-start; gap: 16px; margin-top: 12px;';

        const tableWrap = document.createElement('div');
        tableWrap.style.cssText = 'overflow-x: auto; flex: 1 1 auto; min-width: 0;';

        const sidePanel = document.createElement('div');
        sidePanel.style.cssText = [
            'min-width: 300px',
            'flex: 0 0 440px',
            'max-height: 100vh',
            'position: fixed',
            'right: 0',
            'top: 0',
            'max-width: 450px',
            'overflow-y: auto',
            'border: 1px solid rgba(255,255,255,0.18)',
            'border-radius: 10px',
            'padding: 10px',
            'background: rgba(0,0,0,0.08)'
        ].join('; ');

        const sidePanelTitle = document.createElement('div');
        sidePanelTitle.textContent = 'Selected styles';
        sidePanelTitle.style.cssText = 'font-weight: 600; margin-bottom: 8px;';
        sidePanel.appendChild(sidePanelTitle);

        const emptySidePanel = document.createElement('div');
        emptySidePanel.textContent = 'Click a style name in the table to add it here.';
        emptySidePanel.style.cssText = 'font-size: 0.85rem; opacity: 0.65;';
        sidePanel.appendChild(emptySidePanel);

        const sidePanelCards = document.createElement('div');
        sidePanelCards.style.cssText = 'display: grid; gap: 10px;';
        sidePanel.appendChild(sidePanelCards);

        const selectedCards = new Map();
        const updateEmptySidePanel = () => {
            emptySidePanel.style.display = selectedCards.size ? 'none' : 'block';
        };
        const addStyleCard = (group, styleNumber) => {
            const existing = selectedCards.get(styleNumber);
            if (existing) {
                existing.scrollIntoView({ block: 'nearest' });
                return;
            }
            const card = createStylePreviewCard(group, styleNumber, () => {
                card.remove();
                selectedCards.delete(styleNumber);
                updateEmptySidePanel();
            });
            selectedCards.set(styleNumber, card);
            sidePanelCards.appendChild(card);
            updateEmptySidePanel();
        };

        const table = document.createElement('table');
        table.style.cssText = 'width: max-content; font-size: 0.85rem;';

        const thead = document.createElement('thead');
        const languageRow = document.createElement('tr');

        const multiplierRow = document.createElement('tr');

        styleState.sizeInputs = new Map();
        styleState.multiplierSelects = new Map();

        const numberHead = document.createElement('th');
        numberHead.textContent = '#';
        languageRow.appendChild(numberHead);

        const numberBlank = document.createElement('th');
        numberBlank.textContent = '';
        multiplierRow.appendChild(numberBlank);

        for (const language of styleState.languages) {
            const th = document.createElement('th');
            th.textContent = language.code;
            languageRow.appendChild(th);

            const multiplierCell = document.createElement('td');
            const select = createMultiplierSelect();
            styleState.multiplierSelects.set(language.code, select);
            select.addEventListener('change', () => recalculateStyleInputsForLanguage(language.code));
            multiplierCell.appendChild(select);
            multiplierRow.appendChild(multiplierCell);
            styleState.sizeInputs.set(language.code, new Map());
        }

        const styleHead = document.createElement('th');
        styleHead.textContent = 'Size / color';

        const usesHead = document.createElement('th');
        usesHead.textContent = 'Uses';
        languageRow.appendChild(usesHead);
        languageRow.appendChild(styleHead);

        const usesBlank = document.createElement('th');
        usesBlank.textContent = '';
        multiplierRow.appendChild(usesBlank);

        const multiplierHead = document.createElement('th');
        multiplierHead.textContent = 'Multiplier';
        multiplierRow.appendChild(multiplierHead);

        thead.appendChild(languageRow);
        thead.appendChild(multiplierRow);
        table.appendChild(thead);

        const tbody = document.createElement('tbody');
        for (const [index, group] of styleState.groups.entries()) {
            const styleNumber = index + 1;
            const row = document.createElement('tr');

            const numberCell = document.createElement('td');
            numberCell.textContent = String(styleNumber);
            numberCell.style.cssText = 'font-weight: 600; opacity: 0.8; text-align: right;';
            row.appendChild(numberCell);

            for (const language of styleState.languages) {
                const cell = document.createElement('td');
                const input = document.createElement('input');
                input.type = 'number';
                input.min = '1';
                input.step = '0.5';
                input.value = group.fontPt || '';
                input.dataset.baseFontPt = group.fontPt || '';
                input.style.width = '72px';
                cell.appendChild(input);
                styleState.sizeInputs.get(language.code).set(group.styleKey, input);
                row.appendChild(cell);
            }

            const usesCell = document.createElement('td');
            usesCell.textContent = String(group.count);
            row.appendChild(usesCell);

            const labelCell = document.createElement('td');
            labelCell.style.position = 'relative';
            labelCell.appendChild(createStylePreviewLabel(group, styleNumber, addStyleCard));
            row.appendChild(labelCell);

            tbody.appendChild(row);
        }
        table.appendChild(tbody);

        if (styleState.languages.length) {
            const tfoot = document.createElement('tfoot');
            const footerRow = document.createElement('tr');

            const numberCell = document.createElement('td');
            numberCell.textContent = '';
            footerRow.appendChild(numberCell);

            for (const language of styleState.languages) {
                const cell = document.createElement('td');
                const button = document.createElement('button');
                button.type = 'button';
                button.textContent = `↓ ${language.code}`;
                button.title = `Download ${language.code}`;
                button.style.cssText = 'display: inline-flex; align-items: center; gap: 6px; padding: 6px 10px; border-radius: 6px; border: 1px solid rgba(255,255,255,0.25); background: #4da6ff; color: #fff; font-weight: 600; cursor: pointer;';
                button.addEventListener('click', (e) => {
                    e.preventDefault();
                    downloadStyledPresentation(language);
                });
                cell.appendChild(button);
                footerRow.appendChild(cell);
            }

            const labelCell = document.createElement('td');
            labelCell.textContent = '';
            footerRow.appendChild(labelCell);

            const styleCell = document.createElement('td');
            styleCell.textContent = 'Download';
            footerRow.appendChild(styleCell);

            tfoot.appendChild(footerRow);
            table.appendChild(tfoot);
        }

        tableWrap.appendChild(table);
        layout.appendChild(tableWrap);
        layout.appendChild(sidePanel);
        styleTableArea.appendChild(layout);
    }

    async function rebuildStyleState() {
        if (!stylePptxBuffer) {
            styleState = null;
            renderStyleControls();
            return;
        }

        const styleRows = await window.PptxStyleUtils.extractStyleRowsFromPptx(stylePptxBuffer.slice(0));
        styleState = buildStyleState(styleRows, styleTranslationData);
        renderStyleControls();
        if (pendingStylePreset) applyStylePreset(pendingStylePreset);
        updateStylePresetButton();
    }

    function ensureDirectChild(element, tagName, namespaceUri, beforeNode = null) {
        let child = getDirectChild(element, tagName);
        if (child) return child;

        child = element.ownerDocument.createElementNS(namespaceUri, tagName);
        if (beforeNode) {
            element.insertBefore(child, beforeNode);
        } else {
            element.appendChild(child);
        }
        return child;
    }

    function applyParagraphFontSize(paragraph, run, fontPt) {
        if (!Number.isFinite(fontPt) || fontPt <= 0) return false;

        const sz = String(Math.round(fontPt * 100));
        const firstChild = run.firstElementChild || run.firstChild;
        const rPr = ensureDirectChild(run, 'a:rPr', A_NS, firstChild);
        rPr.setAttribute('sz', sz);

        const endParaRPr = ensureDirectChild(paragraph, 'a:endParaRPr', A_NS);
        endParaRPr.setAttribute('sz', sz);
        return true;
    }

    function replaceParagraphText(paragraph, newText, fontPt = null) {
        const runs = paragraph.getElementsByTagName('a:r');
        if (runs.length === 0) return false;

        const firstRunT = runs[0].getElementsByTagName('a:t')[0];
        if (!firstRunT) return false;

        firstRunT.textContent = newText;
        firstRunT.setAttribute('xml:space', 'preserve');

        for (let ri = 1; ri < runs.length; ri++) {
            const t = runs[ri].getElementsByTagName('a:t')[0];
            if (t) t.textContent = '';
        }

        applyParagraphFontSize(paragraph, runs[0], fontPt);
        return true;
    }

    async function injectTranslations(pptxBuffer, translationRows) {
        const zip = await JSZip.loadAsync(pptxBuffer);
        const slides = getSlideEntries(zip);
        const warnings = [];
        const slideDocs = new Map();
        const containersByName = new Map();
        let replacedCount = 0;
        let fontSizeAppliedCount = 0;

        for (const slide of slides) {
            const file = zip.file(slide.path);
            if (!file) continue;

            const xml = await file.async('string');
            const doc = parseXml(xml);
            slideDocs.set(slide.path, { doc, modified: false });

            for (const container of collectTextContainers(doc, slide)) {
                container.slidePath = slide.path;
                const shapeName = container.shapeName.trim();
                if (!shapeName) continue;
                if (!containersByName.has(shapeName)) containersByName.set(shapeName, []);
                containersByName.get(shapeName).push(container);
            }
        }

        for (const row of translationRows) {
            if (!row.key || !row.shapeName || Number.isNaN(row.paragraphIndex)) {
                warnings.push(createWarning('invalid-row', `XLSX row ${row.rowNumber}: missing or invalid key.`, {
                    key: row.key,
                    shapeName: row.shapeName,
                    paragraphIndex: Number.isNaN(row.paragraphIndex) ? '' : row.paragraphIndex
                }));
                continue;
            }

            const matches = containersByName.get(row.shapeName) || [];
            if (matches.length === 0) {
                warnings.push(createWarning('missing-shape', `XLSX row ${row.rowNumber}: shapeName "${row.shapeName}" was not found.`, row));
                continue;
            }

            if (matches.length > 1) {
                warnings.push(createWarning('duplicate-shape-name', `XLSX row ${row.rowNumber}: shapeName "${row.shapeName}" is duplicated in PPTX; skipped.`, row));
                continue;
            }

            const container = matches[0];
            if (row.slideNumber && String(container.slide.num) !== row.slideNumber) {
                warnings.push(createWarning('slide-number-changed', `XLSX row ${row.rowNumber}: "${row.shapeName}" moved from slide ${row.slideNumber} to slide ${container.slide.num}; text replaced by shapeName.`, {
                    ...row,
                    severity: 'info'
                }));
            }
            if (row.shapeId && container.shapeId !== row.shapeId) {
                warnings.push(createWarning('shape-id-changed', `XLSX row ${row.rowNumber}: "${row.shapeName}" shapeId changed from ${row.shapeId} to ${container.shapeId}; text replaced by shapeName.`, {
                    ...row,
                    severity: 'info'
                }));
            }

            const paragraph = container.paragraphs[row.paragraphIndex];
            if (!paragraph) {
                warnings.push(createWarning('missing-paragraph', `XLSX row ${row.rowNumber}: paragraph ${row.paragraphIndex} not found in "${row.shapeName}".`, row));
                continue;
            }

            if (row.fontRaw && (!Number.isFinite(row.fontPt) || row.fontPt <= 0)) {
                warnings.push(createWarning('invalid-font-size', `XLSX row ${row.rowNumber}: ${row.fontColumn} value "${row.fontRaw}" is not a valid font size; text was replaced without resizing.`, row));
            }

            const shouldApplyFont = Number.isFinite(row.fontPt) && row.fontPt > 0;
            if (!replaceParagraphText(paragraph, row.text, shouldApplyFont ? row.fontPt : null)) {
                warnings.push(createWarning('missing-text-run', `XLSX row ${row.rowNumber}: paragraph ${row.paragraphIndex} in "${row.shapeName}" has no editable text run.`, row));
                continue;
            }

            if (shouldApplyFont) fontSizeAppliedCount++;
            slideDocs.get(container.slidePath).modified = true;
            replacedCount++;
        }

        for (const [path, entry] of slideDocs.entries()) {
            if (entry.modified) {
                zip.file(path, new XMLSerializer().serializeToString(entry.doc));
            }
        }

        const blob = await zip.generateAsync({ type: 'blob', mimeType: PPTX_MIME });
        return { blob, replacedCount, fontSizeAppliedCount, warnings };
    }

    function setupSingleDropzone(dropzone, input, filenameEl, accept, onFile) {
        if (!dropzone || !input || !filenameEl) return;

        dropzone.addEventListener('click', (e) => {
            if (e.target.closest('.dropzone-filename')) return;
            input.click();
        });

        dropzone.addEventListener('dragenter', (e) => {
            e.preventDefault();
            dropzone.classList.add('dragover');
        });
        dropzone.addEventListener('dragover', (e) => {
            e.preventDefault();
            dropzone.classList.add('dragover');
        });
        dropzone.addEventListener('dragleave', (e) => {
            if (!dropzone.contains(e.relatedTarget)) dropzone.classList.remove('dragover');
        });

        dropzone.addEventListener('drop', (e) => {
            e.preventDefault();
            dropzone.classList.remove('dragover');
            const files = Array.from(e.dataTransfer.files).filter(f =>
                f.name.toLowerCase().endsWith(accept)
            );
            if (files.length > 0) onFile(files[0]);
        });

        input.addEventListener('change', () => {
            if (input.files.length > 0) onFile(input.files[0]);
            input.value = '';
        });
    }

    function showFilename(el, name) {
        el.textContent = name;
        el.style.display = name ? 'block' : 'none';
    }

    setupSingleDropzone(pptxMarkupDropzone, pptxMarkupInput, pptxMarkupFilename, '.pptx', (file) => {
        markupFile = file;
        showFilename(pptxMarkupFilename, file.name);
        markupBtn.disabled = false;
        markupDownload.innerHTML = '';
        markupError.textContent = '';
    });

    if (markupBtn) markupBtn.addEventListener('click', handleMarkup);

    setupSingleDropzone(pptxExtractDropzone, pptxExtractInput, pptxExtractFilename, '.pptx', (file) => {
        extractFile = file;
        showFilename(pptxExtractFilename, file.name);
        extractBtn.disabled = false;
        extractDownload.innerHTML = '';
        extractError.textContent = '';
    });

    extractBtn.addEventListener('click', handleExtract);

    setupSingleDropzone(stylePptxDropzone, stylePptxInput, stylePptxFilename, '.pptx', async (file) => {
        stylePptxName = file.name;
        showFilename(stylePptxFilename, file.name);
        styleError.textContent = '';
        if (styleResult) styleResult.innerHTML = '';

        try {
            stylePptxBuffer = await readFileAsArrayBuffer(file);
            await rebuildStyleState();
        } catch (err) {
            styleError.textContent = 'Error reading PPTX: ' + err.message;
            stylePptxBuffer = null;
            styleState = null;
            renderStyleControls();
        }
    });

    setupSingleDropzone(styleXlsxDropzone, styleXlsxInput, styleXlsxFilename, '.xlsx', async (file) => {
        showFilename(styleXlsxFilename, file.name);
        styleError.textContent = '';
        if (styleResult) styleResult.innerHTML = '';

        try {
            styleTranslationData = await parseXlsxFile(file);
            await rebuildStyleState();
        } catch (err) {
            styleError.textContent = 'Error reading XLSX: ' + err.message;
            styleTranslationData = null;
            await rebuildStyleState();
        }
    });

    if (styleSavePresetBtn) styleSavePresetBtn.addEventListener('click', saveStylePreset);
    if (styleLoadPresetBtn && stylePresetInput) {
        styleLoadPresetBtn.addEventListener('click', () => stylePresetInput.click());
        stylePresetInput.addEventListener('change', async () => {
            if (!stylePresetInput.files.length) return;
            styleError.textContent = '';
            try {
                await loadStylePreset(stylePresetInput.files[0]);
            } catch (err) {
                styleError.textContent = 'Error loading preset: ' + err.message;
            } finally {
                stylePresetInput.value = '';
            }
        });
    }
    renderStyleControls();
})();
