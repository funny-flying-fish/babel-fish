(() => {
    const PPTX_MIME = 'application/vnd.openxmlformats-officedocument.presentationml.presentation';
    const XLSX_MIME = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
    const METADATA_COLUMNS = ['key', 'slideNumber', 'shapeName', 'shapeId', 'containerType', 'paragraphIndex'];
    const STYLE_METRIC_SUFFIXES = ['chars', 'fontPt', 'fontRef', 'fontSource', 'styleFallbackSource', 'styleKey', 'styleLabel', 'densityRatio', 'fontScale'];
    const A_NS = 'http://schemas.openxmlformats.org/drawingml/2006/main';
    const TRANSLATION_NAME_RE = /^tr_[a-f0-9]{6,12}_[a-z][a-z0-9]*$/i;
    const AUTO_KEY_PARAGRAPH_RE = /^(.+?)__p(\d+)$/;
    const AUTO_KEY_HASH_RE = /^(.+?)__h([a-f0-9]+)(?:__(\d+))?$/i;
    const AUTO_KEY_MARKER_RE = /(__p\d+|__h[a-f0-9]+(?:__\d+)?)$/i;
    const TRANSLATIONS_TO_PPTX_CONFIG_URL = 'pptx-translations-pptx.json';
    const DEFAULT_TRANSLATIONS_TO_PPTX_CONFIG = {
        defaultLanguage: 'EN',
        preset: {
            version: 2,
            type: 'pptx-style-size-preset',
            fileSuffix: '_style-preset.json'
        },
        labels: {
            selectedStylesTitle: 'Selected styles',
            emptySelectedStyles: 'Click a style name in the table to add it here.',
            mergeStyles: 'Merge',
            unmergeStyles: 'Unmerge',
            slides: 'Slides',
            slide: 'Slide',
            sizeColorHeader: 'Size / color',
            usesHeader: 'Uses',
            multiplierHeader: 'Multiplier',
            downloadFooter: 'Download'
        },
        messages: {
            loadPptxFirst: 'Load a marked PPTX file first.',
            loadXlsxForExport: 'Load an XLSX file to enable export for additional languages.',
            noLanguageColumns: 'No language columns found in the XLSX.',
            noEditableStyles: 'No editable text styles found in the PPTX.',
            styleUtilitiesMissing: 'Style utilities are not loaded. Check that pptx-style-utils.js is included before pptx-converter.js.',
            richTextUtilitiesMissing: 'Rich text utilities are not loaded. Check that pptx-rich-text.js is included before pptx-converter.js.',
            invalidPreset: 'Invalid style preset JSON.'
        }
    };

    // DOM refs: Prepare / Markup PPTX
    const pptxMarkupDropzone = document.getElementById('pptxMarkupDropzone');
    const pptxMarkupInput = document.getElementById('pptxMarkupInput');
    const pptxMarkupFilename = document.getElementById('pptxMarkupFilename');
    const prepareOldXlsxDropzone = document.getElementById('prepareOldXlsxDropzone');
    const prepareOldXlsxInput = document.getElementById('prepareOldXlsxInput');
    const prepareOldXlsxFilename = document.getElementById('prepareOldXlsxFilename');
    const markupBtn = document.getElementById('markupBtn');
    const markupDownload = document.getElementById('markupDownload');
    const markupError = document.getElementById('markupError');

    const sourceLanguage = document.getElementById('sourceLanguage');

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
    let previousXlsxFile = null;
    let stylePptxBuffer = null;
    let stylePptxName = '';
    let styleTranslationData = null;
    let styleState = null;
    let pendingStylePreset = null;
    let translationsToPptxConfig = DEFAULT_TRANSLATIONS_TO_PPTX_CONFIG;

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

    function mergeTranslationsToPptxConfig(config) {
        const safeConfig = config && typeof config === 'object' ? config : {};
        return {
            ...DEFAULT_TRANSLATIONS_TO_PPTX_CONFIG,
            ...safeConfig,
            defaultLanguage: safeConfig.defaultLanguage || DEFAULT_TRANSLATIONS_TO_PPTX_CONFIG.defaultLanguage,
            preset: {
                ...DEFAULT_TRANSLATIONS_TO_PPTX_CONFIG.preset,
                ...(safeConfig.preset || {})
            },
            labels: {
                ...DEFAULT_TRANSLATIONS_TO_PPTX_CONFIG.labels,
                ...(safeConfig.labels || {})
            },
            messages: {
                ...DEFAULT_TRANSLATIONS_TO_PPTX_CONFIG.messages,
                ...(safeConfig.messages || {})
            }
        };
    }

    async function loadTranslationsToPptxConfig() {
        try {
            const response = await fetch(TRANSLATIONS_TO_PPTX_CONFIG_URL, { cache: 'no-store' });
            if (!response.ok) throw new Error(`HTTP ${response.status}`);
            translationsToPptxConfig = mergeTranslationsToPptxConfig(await response.json());
        } catch (err) {
            translationsToPptxConfig = DEFAULT_TRANSLATIONS_TO_PPTX_CONFIG;
        }
    }

    function translationsToPptxLabel(key) {
        return translationsToPptxConfig.labels[key] || DEFAULT_TRANSLATIONS_TO_PPTX_CONFIG.labels[key] || key;
    }

    function translationsToPptxMessage(key) {
        return translationsToPptxConfig.messages[key] || DEFAULT_TRANSLATIONS_TO_PPTX_CONFIG.messages[key] || key;
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

    function getEncodedParagraphText(paragraph) {
        const rich = window.PptxRichText;
        if (rich && typeof rich.encodeParagraph === 'function') {
            return rich.encodeParagraph(paragraph);
        }
        return {
            text: getTextFromParagraph(paragraph),
            hadMultipleRuns: false,
            hadStyledRuns: false
        };
    }

    function normalizeTextForKey(text) {
        const rich = window.PptxRichText;
        const plain = rich && typeof rich.plainTextForComparison === 'function'
            ? rich.plainTextForComparison(text)
            : String(text || '').trim();
        return plain.replace(/\s+/g, ' ').toLowerCase();
    }

    // FNV-1a 32-bit hash; we keep only 6 hex characters to stay compact in keys.
    function shortHashFromText(text) {
        const normalized = normalizeTextForKey(text);
        let hash = 0x811c9dc5;
        for (let i = 0; i < normalized.length; i++) {
            hash ^= normalized.charCodeAt(i);
            hash = Math.imul(hash, 0x01000193);
        }
        return (hash >>> 0).toString(16).padStart(8, '0').slice(0, 6);
    }

    function isAutoKey(key) {
        const value = String(key || '').trim();
        if (!value) return false;
        const pMatch = value.match(AUTO_KEY_PARAGRAPH_RE);
        if (pMatch && TRANSLATION_NAME_RE.test(pMatch[1])) return true;
        const hMatch = value.match(AUTO_KEY_HASH_RE);
        if (hMatch && TRANSLATION_NAME_RE.test(hMatch[1])) return true;
        return false;
    }

    function splitAutoMarker(key) {
        const value = String(key || '');
        const match = value.match(AUTO_KEY_MARKER_RE);
        if (!match) return { head: value, marker: '' };
        return { head: value.slice(0, value.length - match[1].length), marker: match[1] };
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
                .map((paragraph, paragraphIndex) => {
                    const encoded = getEncodedParagraphText(paragraph);
                    return {
                        paragraph,
                        paragraphIndex,
                        text: encoded.text,
                        richMeta: encoded
                    };
                })
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
        const warning = {
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
        if (details.text !== undefined) warning.text = details.text;
        if (details.migrate) warning.migrate = true;
        if (details.kind) warning.kind = details.kind;
        if (details.reason) warning.reason = details.reason;
        if (details.matchCount) warning.matchCount = details.matchCount;
        return warning;
    }

    function getPptxBaseName(filename) {
        return String(filename || '').replace(/\.pptx$/i, '');
    }

    function getMarkedPptxFilename(filename) {
        const baseName = getPptxBaseName(filename);
        if (/_marked$/i.test(baseName)) return `${baseName}.pptx`;
        return `${baseName}_marked.pptx`;
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

    function appendMessageList(box, messages, limit = 8) {
        const list = document.createElement('ul');
        list.style.cssText = 'margin: 6px 0 0 18px;';
        for (const entry of messages.slice(0, limit)) {
            const item = document.createElement('li');
            item.textContent = entry.message;
            list.appendChild(item);
        }
        if (messages.length > limit) {
            const item = document.createElement('li');
            item.textContent = `...and ${messages.length - limit} more`;
            list.appendChild(item);
        }
        box.appendChild(list);
    }

    function renderResult(target, summaryLines, messages, reportBaseName) {
        const info = document.createElement('div');
        info.style.cssText = 'color: #6bbd6b; font-size: 0.85rem; margin-top: 6px;';
        info.innerHTML = summaryLines.map(line => `<div>${line}</div>`).join('');
        target.appendChild(info);

        if (!messages.length) return;

        const warnings = messages.filter(entry => entry.severity !== 'info');
        const informational = messages.filter(entry => entry.severity === 'info');
        const reportHost = document.createElement('div');
        reportHost.style.marginTop = '8px';

        if (warnings.length) {
            const warningBox = document.createElement('div');
            warningBox.style.cssText = 'color: #e8b64c; font-size: 0.85rem;';
            warningBox.appendChild(document.createTextNode(`${warnings.length} warning(s):`));
            appendMessageList(warningBox, warnings);
            reportHost.appendChild(warningBox);
        }

        if (informational.length) {
            const infoBox = document.createElement('div');
            infoBox.style.cssText = [
                'color: #9aa0a6',
                'font-size: 0.85rem',
                warnings.length ? 'margin-top: 10px' : ''
            ].filter(Boolean).join('; ');
            infoBox.appendChild(document.createTextNode(`${informational.length} informational message(s):`));
            appendMessageList(infoBox, informational);
            reportHost.appendChild(infoBox);
        }

        const reportBlob = createWarningsBlob(messages);
        const reportLink = createDownloadLink(
            reportBlob,
            `${reportBaseName}_report.xlsx`,
            'Download full report'
        );
        const reportWrap = document.createElement('div');
        reportWrap.style.cssText = 'margin-top: 6px; font-size: 0.85rem;';
        reportWrap.appendChild(reportLink);
        reportHost.appendChild(reportWrap);
        target.appendChild(reportHost);
    }

    async function markupPptxIds(arrayBuffer, knownCustomNames) {
        const zip = await JSZip.loadAsync(arrayBuffer);
        const slides = getSlideEntries(zip);
        const warnings = [];
        const usedIds = new Set();
        const seenNames = new Set();
        const customNames = knownCustomNames instanceof Set ? knownCustomNames : new Set();
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

                const isKnownCustom = customNames.has(currentName);
                const isAcceptable = isTranslationName(currentName) || isKnownCustom;
                const shouldKeep = isAcceptable && !seenNames.has(currentName);
                if (shouldKeep) {
                    seenNames.add(currentName);
                    keptCount++;
                    continue;
                }

                if (isAcceptable && seenNames.has(currentName)) {
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

        const prepareLabel = 'Prepare PPTX';

        try {
            markupBtn.disabled = true;
            markupBtn.textContent = 'Preparing...';

            let oldRows = null;
            let knownCustomShapeNames = new Set();
            if (previousXlsxFile) {
                oldRows = await parseXlsxFile(previousXlsxFile);
                knownCustomShapeNames = extractCustomShapeNamesFromXlsx(oldRows);
            }

            const buffer = await readFileAsArrayBuffer(markupFile);
            const result = await markupPptxIds(buffer, knownCustomShapeNames);
            const baseName = getPptxBaseName(markupFile.name);
            const markedFilename = getMarkedPptxFilename(markupFile.name);

            const markedBuffer = await result.blob.arrayBuffer();
            const lang = sourceLanguage && sourceLanguage.value.trim() ? sourceLanguage.value.trim() : 'EN';
            const extractResult = await extractTextFromPptx(markedBuffer, lang, knownCustomShapeNames);

            let migrateSummary = null;
            if (oldRows) {
                migrateSummary = applyPreviousCustomKeys(
                    extractResult.entries,
                    oldRows,
                    lang,
                    extractResult.warnings
                );
            }

            renderKeyEditor({
                entries: extractResult.entries,
                lang,
                baseName,
                markedBlob: result.blob,
                markedFilename,
                markupResult: result,
                extractWarnings: extractResult.warnings,
                migrateSummary
            });
        } catch (err) {
            markupError.textContent = 'Error: ' + err.message;
        } finally {
            markupBtn.disabled = false;
            markupBtn.textContent = prepareLabel;
        }
    }

    async function extractTextFromPptx(arrayBuffer, lang, knownCustomNames) {
        const customNames = knownCustomNames instanceof Set ? knownCustomNames : new Set();
        const zip = await JSZip.loadAsync(arrayBuffer);
        const slides = getSlideEntries(zip);
        const entries = [];
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

                if (!isTranslationName(shapeName) && !customNames.has(shapeName)) {
                    warnings.push(createWarning('unmarked-shape', `Slide ${slide.num}: "${shapeName || '(unnamed)'}" is not a translation ID. Run Prepare PPTX first.`, {
                        slideNumber: slide.num,
                        shapeName,
                        shapeId: container.shapeId,
                        containerType: container.containerType
                    }));
                }

                const isTable = container.containerType === 'graphicFrame';
                const tableHashCounts = new Map();

                for (const item of container.nonEmptyParagraphs) {
                    const keyBase = shapeName || `unmarked_id${container.shapeId || 'unknown'}`;
                    if (item.richMeta && (item.richMeta.hadMultipleRuns || item.richMeta.hadStyledRuns)) {
                        warnings.push(createWarning('rich-text-exported', `Slide ${slide.num}: paragraph ${item.paragraphIndex} in "${shapeName || '(unnamed)'}" exported with inline markdown.`, {
                            severity: 'info',
                            slideNumber: slide.num,
                            shapeName,
                            shapeId: container.shapeId,
                            containerType: container.containerType,
                            paragraphIndex: item.paragraphIndex
                        }));
                    }

                    let autoKey;
                    if (isTable) {
                        const hash = shortHashFromText(item.text);
                        const count = (tableHashCounts.get(hash) || 0) + 1;
                        tableHashCounts.set(hash, count);
                        autoKey = count > 1
                            ? `${keyBase}__h${hash}__${count}`
                            : `${keyBase}__h${hash}`;
                    } else {
                        autoKey = `${keyBase}__p${item.paragraphIndex}`;
                    }

                    entries.push({
                        autoKey,
                        currentKey: autoKey,
                        shapeName,
                        slideNumber: String(slide.num),
                        containerType: container.containerType,
                        paragraphIndex: item.paragraphIndex,
                        text: item.text,
                        origin: 'auto'
                    });
                }
            }
        }

        for (const [shapeName, count] of nameCounts.entries()) {
            if (count > 1) {
                warnings.push(createWarning('duplicate-shape-name', `Duplicate shapeName "${shapeName}" found ${count} times. Run Prepare PPTX to normalize duplicates.`, {
                    shapeName
                }));
            }
        }

        return { entries, warnings, containerCount, lang };
    }

    function buildXlsxRowsFromEntries(entries, lang) {
        const rows = [['key', lang]];
        for (const entry of entries) {
            rows.push([entry.currentKey, entry.text]);
        }
        return rows;
    }

    function inferShapeNameFromKey(key) {
        if (!key) return '';
        const value = String(key).trim();
        if (!value) return '';
        const { head, marker } = splitAutoMarker(value);
        if (marker) return head;
        return value.replace(/_+\d+$/, '');
    }

    function normalizeLevelFragment(value) {
        return String(value || '').trim().replace(/_+$/, '');
    }

    function composeKey(parts) {
        const segs = [];
        if (parts.userPrefix) segs.push(parts.userPrefix);
        if (parts.userBasename) segs.push(parts.userBasename);
        if (parts.keyid) segs.push(parts.keyid);
        return segs.join('__') + (parts.marker || '');
    }

    function extractKeyid(entry) {
        const { head } = splitAutoMarker(entry.currentKey);
        let s = head;
        if (entry.userPrefix) {
            if (s === entry.userPrefix) return '';
            if (s.startsWith(entry.userPrefix + '__')) s = s.slice(entry.userPrefix.length + 2);
        }
        if (entry.userBasename) {
            if (s === entry.userBasename) return '';
            if (s.startsWith(entry.userBasename + '__')) s = s.slice(entry.userBasename.length + 2);
        }
        return s;
    }

    function stripAutoHashPrefix(value) {
        return String(value || '').replace(/^tr_[a-f0-9]{6,12}_/i, '');
    }

    function inferLevelsForMigratedEntries(entries) {
        const bySlide = new Map();
        for (const e of entries) {
            const k = e.slideNumber || '0';
            if (!bySlide.has(k)) bySlide.set(k, []);
            bySlide.get(k).push(e);
        }

        for (const slideEntries of bySlide.values()) {
            const preserved = slideEntries.filter(e => e.origin === 'preserved');
            if (!preserved.length) continue;

            const parsed = preserved.map(e => {
                const { head, marker } = splitAutoMarker(e.currentKey);
                return { entry: e, segments: head ? head.split('__') : [], marker };
            });

            let slidePrefix = null;
            if (parsed.length > 0 && parsed.every(p => p.segments.length >= 1)) {
                const firsts = new Set(parsed.map(p => p.segments[0]));
                if (firsts.size === 1) {
                    const candidate = parsed[0].segments[0];
                    const someHasMoreLevels = parsed.some(p => p.segments.length >= 2);
                    const multipleContainers = new Set(parsed.map(p => p.entry.shapeName || '(unnamed)')).size > 1;
                    if (someHasMoreLevels || multipleContainers || parsed.length === 1) {
                        slidePrefix = candidate;
                    }
                }
            }

            if (slidePrefix) {
                for (const p of parsed) {
                    if (p.segments[0] === slidePrefix) {
                        p.entry.userPrefix = slidePrefix;
                        p.segments.shift();
                    }
                }
                for (const e of slideEntries) {
                    if (e.origin === 'preserved') continue;
                    if (e.userPrefix === slidePrefix) continue;
                    const marker = splitAutoMarker(e.currentKey).marker;
                    let keyid = extractKeyid(e);
                    keyid = stripAutoHashPrefix(keyid);
                    if (keyid === slidePrefix) keyid = '';
                    else if (keyid.startsWith(slidePrefix + '__')) keyid = keyid.slice(slidePrefix.length + 2);
                    const oldKey = e.currentKey;
                    e.userPrefix = slidePrefix;
                    e.currentKey = composeKey({
                        userPrefix: slidePrefix,
                        userBasename: e.userBasename,
                        keyid,
                        marker
                    });
                    if (e.currentKey !== oldKey && e.origin === 'auto') e.origin = 'custom';
                }
            }

            const byContainer = new Map();
            for (const p of parsed) {
                const k = p.entry.shapeName || '(unnamed)';
                if (!byContainer.has(k)) byContainer.set(k, []);
                byContainer.get(k).push(p);
            }

            for (const [shapeName, cps] of byContainer.entries()) {
                let basename = null;
                if (cps.length > 0 && cps.every(p => p.segments.length >= 1)) {
                    const firsts = new Set(cps.map(p => p.segments[0]));
                    if (firsts.size === 1) basename = cps[0].segments[0];
                }
                if (!basename) continue;

                for (const p of cps) {
                    if (p.segments[0] === basename) {
                        p.entry.userBasename = basename;
                        p.segments.shift();
                    }
                }

                for (const e of slideEntries) {
                    if (e.origin === 'preserved') continue;
                    if ((e.shapeName || '(unnamed)') !== shapeName) continue;
                    if (e.userBasename === basename) continue;
                    if (!e.userPrefix) continue;
                    const marker = splitAutoMarker(e.currentKey).marker;
                    let keyid = extractKeyid(e);
                    if (keyid === basename) keyid = '';
                    else if (keyid.startsWith(basename + '__')) keyid = keyid.slice(basename.length + 2);
                    if (!e.keyidIsUserSet) {
                        keyid = stripAutoHashPrefix(keyid);
                        if (marker) keyid = '';
                    }
                    const oldKey = e.currentKey;
                    e.userBasename = basename;
                    e.currentKey = composeKey({
                        userPrefix: e.userPrefix,
                        userBasename: basename,
                        keyid,
                        marker
                    });
                    if (e.currentKey !== oldKey && e.origin === 'auto') e.origin = 'custom';
                }
            }
        }
    }

    function longestCommonPrefix(values) {
        if (!values.length) return '';
        let prefix = values[0];
        for (let i = 1; i < values.length; i++) {
            while (prefix && !values[i].startsWith(prefix)) {
                prefix = prefix.slice(0, -1);
            }
            if (!prefix) return '';
        }
        return prefix;
    }

    function deriveShapeNameFromEntries(entries) {
        const heads = entries
            .map(e => splitAutoMarker(e.currentKey).head)
            .filter(Boolean);
        if (!heads.length) return null;
        const unique = Array.from(new Set(heads));
        if (unique.length === 1) return unique[0];
        const stripped = Array.from(new Set(unique.map(h => h.replace(/_+\d+$/, ''))));
        if (stripped.length === 1) return stripped[0];
        return longestCommonPrefix(stripped).replace(/_+$/, '') || null;
    }

    function computeShapeRenames(entries) {
        const byShape = new Map();
        for (const entry of entries) {
            if (!entry.shapeName) continue;
            if (!byShape.has(entry.shapeName)) byShape.set(entry.shapeName, []);
            byShape.get(entry.shapeName).push(entry);
        }
        const renames = new Map();
        const usedNames = new Set();
        for (const [oldName, items] of byShape) {
            let newName = deriveShapeNameFromEntries(items);
            if (!newName || newName === oldName) continue;
            if (usedNames.has(newName)) {
                let suffix = 2;
                while (usedNames.has(`${newName}_${suffix}`)) suffix++;
                newName = `${newName}_${suffix}`;
            }
            usedNames.add(newName);
            renames.set(oldName, newName);
        }
        return renames;
    }

    async function applyShapeNamesToPptx(pptxBlob, renames) {
        if (!renames.size) return pptxBlob;
        const arrayBuffer = await pptxBlob.arrayBuffer();
        const zip = await JSZip.loadAsync(arrayBuffer);
        const slides = getSlideEntries(zip);
        for (const slide of slides) {
            const file = zip.file(slide.path);
            if (!file) continue;
            const xml = await file.async('string');
            const doc = parseXml(xml);
            const cNvPrs = doc.getElementsByTagName('p:cNvPr');
            let modified = false;
            for (let i = 0; i < cNvPrs.length; i++) {
                const el = cNvPrs[i];
                const currentName = el.getAttribute('name');
                if (renames.has(currentName)) {
                    el.setAttribute('name', renames.get(currentName));
                    modified = true;
                }
            }
            if (modified) {
                zip.file(slide.path, new XMLSerializer().serializeToString(doc));
            }
        }
        return await zip.generateAsync({ type: 'blob', mimeType: PPTX_MIME });
    }

    function extractCustomShapeNamesFromXlsx(rows) {
        const names = new Set();
        if (!rows || rows.length < 2) return names;
        const header = (rows[0] || []).map(v => String(v || '').trim());
        const headerIndex = buildHeaderIndex(header);

        if (headerIndex.has('shapeName')) {
            const idx = headerIndex.get('shapeName');
            for (let r = 1; r < rows.length; r++) {
                const name = String(((rows[r] || [])[idx]) || '').trim();
                if (name && !isTranslationName(name)) names.add(name);
            }
        }

        if (headerIndex.has('key')) {
            const idx = headerIndex.get('key');
            for (let r = 1; r < rows.length; r++) {
                const key = String(((rows[r] || [])[idx]) || '').trim();
                if (!key || isAutoKey(key)) continue;
                const candidate = inferShapeNameFromKey(key);
                if (candidate && !isTranslationName(candidate)) names.add(candidate);
            }
        }

        return names;
    }

    function applyPreviousCustomKeys(entries, oldXlsxRows, lang, warnings) {
        if (!oldXlsxRows || oldXlsxRows.length < 2) {
            return { migrated: 0, orphaned: 0, ambiguous: 0 };
        }

        const header = (oldXlsxRows[0] || []).map(value => String(value || '').trim());
        const headerIndex = buildHeaderIndex(header);
        if (!headerIndex.has('key')) {
            return { migrated: 0, orphaned: 0, ambiguous: 0 };
        }

        const languageColumns = getLanguageColumns(header);
        let langColumnIndex = null;
        const matchingLang = languageColumns.find(
            col => col.code.toLowerCase() === String(lang || '').toLowerCase()
        );
        if (matchingLang) langColumnIndex = matchingLang.index;
        else if (languageColumns.length) langColumnIndex = languageColumns[0].index;

        if (langColumnIndex === null) {
            return { migrated: 0, orphaned: 0, ambiguous: 0 };
        }

        const hasShapeCol = headerIndex.has('shapeName');
        const hasSlideCol = headerIndex.has('slideNumber');

        const entriesByShape = new Map();
        const entriesByText = new Map();
        for (const entry of entries) {
            if (entry.shapeName) {
                if (!entriesByShape.has(entry.shapeName)) entriesByShape.set(entry.shapeName, []);
                entriesByShape.get(entry.shapeName).push(entry);
            }
            const textKey = normalizeTextForKey(entry.text);
            if (!entriesByText.has(textKey)) entriesByText.set(textKey, []);
            entriesByText.get(textKey).push(entry);
        }

        const consumed = new Set();
        let migrated = 0;
        let orphaned = 0;
        let ambiguous = 0;

        for (let r = 1; r < oldXlsxRows.length; r++) {
            const oldRow = oldXlsxRows[r] || [];
            const oldKey = String(oldRow[headerIndex.get('key')] || '').trim();
            if (!oldKey || isAutoKey(oldKey)) continue;

            const oldShapeName = hasShapeCol
                ? String(oldRow[headerIndex.get('shapeName')] || '').trim()
                : inferShapeNameFromKey(oldKey);
            const oldSlide = hasSlideCol
                ? String(oldRow[headerIndex.get('slideNumber')] || '').trim()
                : '';
            const oldEn = String(oldRow[langColumnIndex] || '').trim();

            if (!oldEn) {
                orphaned++;
                warnings.push(createWarning('orphaned-custom-key', `Previous XLSX: custom key "${oldKey}" has empty source text.`, {
                    key: oldKey,
                    shapeName: oldShapeName,
                    slideNumber: oldSlide,
                    text: '',
                    migrate: true,
                    kind: 'orphaned',
                    reason: 'empty-text'
                }));
                continue;
            }

            const normEn = normalizeTextForKey(oldEn);
            const shapeCandidates = oldShapeName
                ? (entriesByShape.get(oldShapeName) || []).filter(e => normalizeTextForKey(e.text) === normEn)
                : [];
            const globalCandidates = entriesByText.get(normEn) || [];

            let pool = shapeCandidates.length ? shapeCandidates : globalCandidates;
            let available = pool.filter(e => !consumed.has(e));

            if (!available.length && shapeCandidates.length && globalCandidates.length) {
                available = globalCandidates.filter(e => !consumed.has(e));
                pool = globalCandidates;
            }

            if (!available.length) {
                orphaned++;
                const exists = pool.length > 0;
                warnings.push(createWarning('orphaned-custom-key',
                    exists
                        ? `Previous XLSX: custom key "${oldKey}" — all matching paragraphs already taken by earlier custom keys.`
                        : `Previous XLSX: custom key "${oldKey}" — source text not found (cell may be removed or edited).`,
                    {
                        key: oldKey,
                        shapeName: oldShapeName,
                        slideNumber: oldSlide,
                        text: oldEn,
                        migrate: true,
                        kind: 'orphaned',
                        reason: exists ? 'already-migrated' : 'source-not-found'
                    }
                ));
                continue;
            }

            const target = available[0];
            target.currentKey = oldKey;
            target.origin = 'preserved';
            target.keyidIsUserSet = true;
            consumed.add(target);
            migrated++;
        }

        inferLevelsForMigratedEntries(entries);
        return { migrated, orphaned, ambiguous };
    }

    function buildKeyDraft(entries) {
        const slides = [];
        const slidesByNum = new Map();
        const containersByPath = new Map();

        entries.forEach((entry, entryIndex) => {
            entry.id = entryIndex;
            if (entry.autoMarker === undefined) {
                entry.autoMarker = splitAutoMarker(entry.autoKey).marker;
            }
            if (entry.userPrefix === undefined) entry.userPrefix = '';
            if (entry.userBasename === undefined) entry.userBasename = '';
            if (entry.keyidIsUserSet === undefined) entry.keyidIsUserSet = false;
            const slideKey = entry.slideNumber || '0';
            let slide = slidesByNum.get(slideKey);
            if (!slide) {
                slide = {
                    num: entry.slideNumber,
                    label: entry.slideNumber ? `Slide ${entry.slideNumber}` : 'Unknown slide',
                    containers: [],
                    containersByName: new Map()
                };
                slidesByNum.set(slideKey, slide);
                slides.push(slide);
            }

            const containerKey = `${slideKey}::${entry.shapeName || '(unnamed)'}`;
            let container = containersByPath.get(containerKey);
            if (!container) {
                container = {
                    shapeName: entry.shapeName || '(unnamed)',
                    containerType: entry.containerType,
                    entries: []
                };
                containersByPath.set(containerKey, container);
                slide.containers.push(container);
                slide.containersByName.set(container.shapeName, container);
            }
            container.entries.push(entry);
        });

        return { slides, flatEntries: entries };
    }

    function isValidKeyFragment(value) {
        return /^[A-Za-z0-9_.]+$/.test(String(value || ''));
    }

    function renderKeyEditor(opts) {
        const {
            entries,
            lang,
            baseName,
            markedBlob,
            markedFilename,
            markupResult,
            extractWarnings,
            migrateSummary
        } = opts;

        const draft = buildKeyDraft(entries);

        markupDownload.innerHTML = '';

        const editor = document.createElement('div');
        editor.className = 'key-editor';

        const summary = document.createElement('div');
        summary.className = 'key-editor-summary';
        editor.appendChild(summary);

        const warningsHost = document.createElement('div');
        warningsHost.className = 'key-editor-warnings';
        editor.appendChild(warningsHost);

        const mainWrap = document.createElement('div');
        mainWrap.className = 'key-editor-main';

        const tree = document.createElement('div');
        tree.className = 'key-editor-tree';
        mainWrap.appendChild(tree);

        const downloadsWrap = document.createElement('div');
        downloadsWrap.className = 'key-editor-downloads';
        mainWrap.appendChild(downloadsWrap);

        editor.appendChild(mainWrap);

        markupDownload.appendChild(editor);

        function refreshSummary() {
            const total = draft.flatEntries.length;
            let customCount = 0;
            for (const entry of draft.flatEntries) {
                if (entry.currentKey !== entry.autoKey) customCount++;
            }
            const lines = [
                `<strong>${total}</strong> paragraph entries (${lang})`,
                `<strong>${customCount}</strong> custom keys`,
                `<strong>${draft.slides.length}</strong> slides`
            ];
            if (migrateSummary) {
                lines.push(`Migrated from previous XLSX: <strong>${migrateSummary.migrated}</strong>`);
                if (migrateSummary.orphaned) lines.push(`Orphaned: <strong>${migrateSummary.orphaned}</strong>`);
                if (migrateSummary.ambiguous) lines.push(`Ambiguous: <strong>${migrateSummary.ambiguous}</strong>`);
            }
            summary.innerHTML = lines.map(line => `<span>${line}</span>`).join('');
        }

        function updateRowClass(row, entry) {
            row.classList.toggle('is-custom', entry.currentKey !== entry.autoKey && entry.origin !== 'preserved');
            row.classList.toggle('is-preserved', entry.currentKey !== entry.autoKey && entry.origin === 'preserved');
            const resetBtn = row.querySelector('.key-reset');
            if (resetBtn) resetBtn.disabled = entry.currentKey === entry.autoKey;
        }

        const containerProjections = [];
        const basenameControls = [];

        function refreshContainerProjections() {
            for (const { container, span } of containerProjections) {
                const newName = deriveShapeNameFromEntries(container.entries);
                if (newName && newName !== container.shapeName) {
                    span.textContent = ` → ${newName}`;
                    span.classList.add('is-active');
                } else {
                    span.textContent = '';
                    span.classList.remove('is-active');
                }
            }
        }

        function refreshBasenameAvailability() {
            for (const ctrl of basenameControls) {
                const hasPrefix = ctrl.container.entries.every(e => Boolean(e.userPrefix));
                const hasBasename = ctrl.container.entries.some(e => Boolean(e.userBasename));
                ctrl.input.disabled = !hasPrefix;
                ctrl.applyBtn.disabled = !hasPrefix;
                ctrl.resetBtn.disabled = !hasBasename;
                ctrl.input.title = hasPrefix ? '' : 'Set a slide prefix first to keep basenames unique across slides.';
                ctrl.applyBtn.title = hasPrefix ? '' : 'Set a slide prefix first to keep basenames unique across slides.';
                ctrl.resetBtn.title = hasBasename ? 'Reset basename for this container' : 'No basename to reset';
            }
        }

        function refreshAllRowClasses() {
            tree.querySelectorAll('.key-editor-row').forEach(row => {
                const id = Number(row.dataset.entryId);
                const entry = draft.flatEntries[id];
                if (entry) {
                    if (row._syncRowParts) row._syncRowParts();
                    updateRowClass(row, entry);
                }
            });
            refreshContainerProjections();
            refreshBasenameAvailability();
        }

        function flashStatus(el, message, kind) {
            if (!el) return;
            el.textContent = message;
            el.classList.remove('is-ok', 'is-err');
            el.classList.add(kind === 'ok' ? 'is-ok' : 'is-err');
            clearTimeout(el._flashTimer);
            el._flashTimer = setTimeout(() => {
                el.textContent = '';
                el.classList.remove('is-ok', 'is-err');
            }, kind === 'ok' ? 2200 : 3500);
        }

        function getMarkerForEntry(entry) {
            return splitAutoMarker(entry.currentKey).marker;
        }

        function deduplicateSlideKeys(slide) {
            const allEntries = [];
            for (const c of slide.containers) {
                for (const e of c.entries) allEntries.push(e);
            }
            for (const entry of allEntries) {
                if (entry._dedupSuffix) {
                    const marker = splitAutoMarker(entry.currentKey).marker;
                    let keyid = extractKeyid(entry);
                    if (keyid.endsWith(entry._dedupSuffix)) {
                        keyid = keyid.slice(0, -entry._dedupSuffix.length);
                    }
                    entry._dedupSuffix = '';
                    entry.currentKey = composeKey({
                        userPrefix: entry.userPrefix,
                        userBasename: entry.userBasename,
                        keyid,
                        marker
                    });
                }
            }
            const usedKeys = new Set();
            for (const entry of allEntries) {
                let candidate = entry.currentKey;
                if (usedKeys.has(candidate)) {
                    const marker = splitAutoMarker(entry.currentKey).marker;
                    let keyid = extractKeyid(entry);
                    let n = 2;
                    while (usedKeys.has(candidate)) {
                        const suffix = '_' + n;
                        candidate = composeKey({
                            userPrefix: entry.userPrefix,
                            userBasename: entry.userBasename,
                            keyid: keyid + suffix,
                            marker
                        });
                        n++;
                    }
                    entry._dedupSuffix = '_' + (n - 1);
                    entry.currentKey = candidate;
                }
                usedKeys.add(candidate);
            }
        }

        function findSlideForContainer(container) {
            for (const s of draft.slides) {
                if (s.containers.includes(container)) return s;
            }
            return null;
        }

        function applyPrefixToSlide(slide, prefix) {
            if (!prefix) return { ok: false, error: 'Prefix is empty' };
            if (!isValidKeyFragment(prefix)) return { ok: false, error: 'Allowed: A-Z, a-z, 0-9, _ .' };
            markupError.textContent = '';
            const normalized = normalizeLevelFragment(prefix);
            if (!normalized) return { ok: false, error: 'Prefix is empty' };
            let count = 0;
            for (const container of slide.containers) {
                for (const entry of container.entries) {
                    const marker = getMarkerForEntry(entry);
                    let keyid = extractKeyid(entry);
                    keyid = stripAutoHashPrefix(keyid);
                    entry.userPrefix = normalized;
                    entry.currentKey = composeKey({
                        userPrefix: normalized,
                        userBasename: entry.userBasename,
                        keyid,
                        marker
                    });
                    if (entry.origin === 'auto') entry.origin = 'custom';
                    count++;
                }
            }
            deduplicateSlideKeys(slide);
            refreshAllRowClasses();
            refreshSummary();
            return { ok: true, count };
        }

        function applyBasenameToContainer(container, basename) {
            if (!basename) return { ok: false, error: 'Basename is empty' };
            if (!isValidKeyFragment(basename)) return { ok: false, error: 'Allowed: A-Z, a-z, 0-9, _ .' };
            const missingPrefix = container.entries.some(e => !e.userPrefix);
            if (missingPrefix) {
                return { ok: false, error: 'Set slide prefix first — basename alone may collide across slides.' };
            }
            markupError.textContent = '';
            const normalized = normalizeLevelFragment(basename);
            if (!normalized) return { ok: false, error: 'Basename is empty' };
            const total = container.entries.length;
            container.entries.forEach((entry, idx) => {
                const marker = getMarkerForEntry(entry);
                let keyid = extractKeyid(entry);
                if (entry.keyidIsUserSet) {
                    // Preserve user-typed keyid as-is
                } else {
                    keyid = stripAutoHashPrefix(keyid);
                    if (marker) {
                        keyid = '';
                    } else if (total > 1) {
                        keyid = String(idx + 1);
                    }
                }
                entry.userBasename = normalized;
                entry.currentKey = composeKey({
                    userPrefix: entry.userPrefix,
                    userBasename: normalized,
                    keyid,
                    marker
                });
                if (entry.origin === 'auto') entry.origin = 'custom';
            });
            const parentSlide = findSlideForContainer(container);
            if (parentSlide) deduplicateSlideKeys(parentSlide);
            refreshAllRowClasses();
            refreshSummary();
            return { ok: true, count: container.entries.length };
        }

        function resetBasenameInContainer(container) {
            let changed = 0;
            for (const entry of container.entries) {
                if (!entry.userBasename) continue;
                const marker = splitAutoMarker(entry.currentKey).marker;
                let keyid = extractKeyid(entry);
                if (!entry.keyidIsUserSet) {
                    keyid = stripAutoHashPrefix(splitAutoMarker(entry.autoKey).head);
                }
                entry.userBasename = '';
                entry.currentKey = composeKey({
                    userPrefix: entry.userPrefix,
                    userBasename: '',
                    keyid,
                    marker
                });
                if (entry.currentKey === entry.autoKey && !entry.userPrefix) {
                    entry.origin = 'auto';
                    entry.keyidIsUserSet = false;
                }
                changed++;
            }
            if (changed) {
                const parentSlide = findSlideForContainer(container);
                if (parentSlide) deduplicateSlideKeys(parentSlide);
                refreshAllRowClasses();
                refreshSummary();
            }
            return { ok: true, count: changed };
        }

        for (const slide of draft.slides) {
            const slideEl = document.createElement('div');
            slideEl.className = 'key-editor-slide is-collapsed';

            const header = document.createElement('div');
            header.className = 'key-editor-slide-header';

            const titleBtn = document.createElement('button');
            titleBtn.type = 'button';
            titleBtn.className = 'key-editor-slide-toggle';
            const totalInSlide = slide.containers.reduce((sum, c) => sum + c.entries.length, 0);
            titleBtn.textContent = `${slide.label} · ${slide.containers.length} container(s), ${totalInSlide} key(s)`;
            titleBtn.addEventListener('click', () => {
                slideEl.classList.toggle('is-collapsed');
            });
            header.appendChild(titleBtn);

            const prefixGroup = document.createElement('div');
            prefixGroup.className = 'key-editor-inline-actions';
            const prefixInput = document.createElement('input');
            prefixInput.type = 'text';
            prefixInput.placeholder = 'Prefix (e.g. s5_)';
            prefixInput.className = 'key-editor-inline-input';
            const prefixStatus = document.createElement('span');
            prefixStatus.className = 'key-editor-inline-status';
            const runPrefix = () => {
                const result = applyPrefixToSlide(slide, prefixInput.value.trim());
                if (result.ok) {
                    prefixInput.value = '';
                    flashStatus(prefixStatus, `Applied to ${result.count} key(s)`, 'ok');
                } else {
                    flashStatus(prefixStatus, result.error, 'err');
                }
            };
            prefixInput.addEventListener('keydown', (e) => {
                if (e.key === 'Enter') {
                    e.preventDefault();
                    runPrefix();
                }
            });
            const prefixBtn = document.createElement('button');
            prefixBtn.type = 'button';
            prefixBtn.textContent = 'Apply prefix';
            prefixBtn.addEventListener('click', runPrefix);
            prefixGroup.appendChild(prefixInput);
            prefixGroup.appendChild(prefixBtn);
            prefixGroup.appendChild(prefixStatus);
            header.appendChild(prefixGroup);

            slideEl.appendChild(header);

            const slideBody = document.createElement('div');
            slideBody.className = 'key-editor-slide-body';
            slideEl.appendChild(slideBody);

            for (const container of slide.containers) {
                const containerEl = document.createElement('div');
                containerEl.className = 'key-editor-container';

                const containerHeader = document.createElement('div');
                containerHeader.className = 'key-editor-container-header';

                const containerTitle = document.createElement('div');
                containerTitle.className = 'key-editor-container-title';
                const typeBadge = document.createElement('span');
                typeBadge.className = 'key-editor-container-type';
                typeBadge.textContent = container.containerType === 'graphicFrame' ? 'table' : 'shape';
                containerTitle.appendChild(typeBadge);
                const nameSpan = document.createElement('span');
                nameSpan.textContent = `${container.shapeName} · ${container.entries.length} paragraph(s)`;
                containerTitle.appendChild(nameSpan);
                const projectionSpan = document.createElement('span');
                projectionSpan.className = 'key-editor-container-projection';
                containerTitle.appendChild(projectionSpan);
                containerProjections.push({ container, span: projectionSpan });
                containerHeader.appendChild(containerTitle);

                const isTable = container.containerType === 'graphicFrame';
                const showBasename = isTable || container.entries.length > 1;
                if (showBasename) {
                    const basenameGroup = document.createElement('div');
                    basenameGroup.className = 'key-editor-inline-actions';
                    const basenameInput = document.createElement('input');
                    basenameInput.type = 'text';
                    basenameInput.placeholder = 'Basename (e.g. pricing)';
                    basenameInput.className = 'key-editor-inline-input';
                    const basenameStatus = document.createElement('span');
                    basenameStatus.className = 'key-editor-inline-status';
                    const runBasename = () => {
                        const result = applyBasenameToContainer(container, basenameInput.value.trim());
                        if (result.ok) {
                            basenameInput.value = '';
                            flashStatus(basenameStatus, `Applied to ${result.count} key(s)`, 'ok');
                        } else {
                            flashStatus(basenameStatus, result.error, 'err');
                        }
                    };
                    basenameInput.addEventListener('keydown', (e) => {
                        if (e.key === 'Enter') {
                            e.preventDefault();
                            runBasename();
                        }
                    });
                    const basenameBtn = document.createElement('button');
                    basenameBtn.type = 'button';
                    basenameBtn.textContent = 'Apply basename';
                    basenameBtn.addEventListener('click', runBasename);
                    const basenameResetBtn = document.createElement('button');
                    basenameResetBtn.type = 'button';
                    basenameResetBtn.className = 'key-editor-inline-reset';
                    basenameResetBtn.setAttribute('aria-label', 'Reset basename');
                    basenameResetBtn.textContent = '×';
                    basenameResetBtn.addEventListener('click', () => {
                        const result = resetBasenameInContainer(container);
                        basenameInput.value = '';
                        if (result.ok && result.count) {
                            flashStatus(basenameStatus, `Reset ${result.count} key(s)`, 'ok');
                        }
                    });
                    basenameGroup.appendChild(basenameInput);
                    basenameGroup.appendChild(basenameBtn);
                    basenameGroup.appendChild(basenameResetBtn);
                    basenameGroup.appendChild(basenameStatus);
                    containerHeader.appendChild(basenameGroup);
                    basenameControls.push({
                        container,
                        input: basenameInput,
                        applyBtn: basenameBtn,
                        resetBtn: basenameResetBtn
                    });
                }

                containerEl.appendChild(containerHeader);

                for (const entry of container.entries) {
                    const row = document.createElement('div');
                    row.className = 'key-editor-row';
                    row.dataset.entryId = String(entry.id);

                    const keyWrap = document.createElement('div');
                    keyWrap.className = 'key-input-wrap';

                    const prefixLabel = document.createElement('span');
                    prefixLabel.className = 'key-input-chip key-input-prefix';
                    keyWrap.appendChild(prefixLabel);

                    const basenameLabel = document.createElement('span');
                    basenameLabel.className = 'key-input-chip key-input-basename';
                    keyWrap.appendChild(basenameLabel);

                    const middleInput = document.createElement('input');
                    middleInput.type = 'text';
                    middleInput.className = 'key-input-middle';
                    middleInput.spellcheck = false;
                    keyWrap.appendChild(middleInput);

                    const markerLabel = document.createElement('span');
                    markerLabel.className = 'key-input-chip key-input-marker';
                    keyWrap.appendChild(markerLabel);

                    row.appendChild(keyWrap);

                    const enPreview = document.createElement('div');
                    enPreview.className = 'key-en';
                    enPreview.textContent = entry.text;
                    enPreview.title = entry.text;
                    row.appendChild(enPreview);

                    const resetBtn = document.createElement('button');
                    resetBtn.type = 'button';
                    resetBtn.className = 'key-reset';
                    resetBtn.textContent = 'Reset';
                    row.appendChild(resetBtn);

                    function computeRowParts() {
                        const marker = splitAutoMarker(entry.currentKey).marker;
                        const keyid = extractKeyid(entry);
                        return {
                            prefix: entry.userPrefix || '',
                            basename: entry.userBasename || '',
                            middle: keyid,
                            marker
                        };
                    }

                    function syncRowPartsToEntry() {
                        const { prefix, basename, middle, marker } = computeRowParts();
                        const prefixHasFollower = Boolean(basename) || Boolean(middle);
                        const basenameHasFollower = Boolean(middle);
                        prefixLabel.textContent = prefix ? (prefixHasFollower ? prefix + '__' : prefix) : '';
                        prefixLabel.classList.toggle('is-empty', !prefix);
                        basenameLabel.textContent = basename ? (basenameHasFollower ? basename + '__' : basename) : '';
                        basenameLabel.classList.toggle('is-empty', !basename);
                        markerLabel.textContent = marker;
                        markerLabel.classList.toggle('is-empty', !marker);
                        if (document.activeElement !== middleInput) {
                            middleInput.value = middle;
                        }
                    }

                    const dedupeEntrySlide = () => {
                        const s = findSlideForContainer(container);
                        if (s) deduplicateSlideKeys(s);
                        refreshAllRowClasses();
                    };

                    middleInput.addEventListener('input', () => {
                        const keyid = middleInput.value.trim().replace(/^_+|_+$/g, '');
                        const marker = splitAutoMarker(entry.currentKey).marker;
                        entry.keyidIsUserSet = true;
                        entry.currentKey = composeKey({
                            userPrefix: entry.userPrefix,
                            userBasename: entry.userBasename,
                            keyid,
                            marker
                        });
                        if (entry.currentKey !== entry.autoKey && entry.origin === 'auto') {
                            entry.origin = 'custom';
                        } else if (entry.currentKey === entry.autoKey) {
                            entry.userPrefix = '';
                            entry.userBasename = '';
                            entry.keyidIsUserSet = false;
                            entry.origin = 'auto';
                            syncRowPartsToEntry();
                        }
                        updateRowClass(row, entry);
                        refreshContainerProjections();
                        refreshBasenameAvailability();
                        refreshSummary();
                    });

                    middleInput.addEventListener('blur', () => {
                        dedupeEntrySlide();
                    });

                    resetBtn.addEventListener('click', () => {
                        entry.currentKey = entry.autoKey;
                        entry.userPrefix = '';
                        entry.userBasename = '';
                        entry.keyidIsUserSet = false;
                        entry.origin = 'auto';
                        syncRowPartsToEntry();
                        dedupeEntrySlide();
                        updateRowClass(row, entry);
                        refreshContainerProjections();
                        refreshBasenameAvailability();
                        refreshSummary();
                    });

                    row._syncRowParts = syncRowPartsToEntry;
                    syncRowPartsToEntry();
                    updateRowClass(row, entry);
                    containerEl.appendChild(row);
                }

                slideBody.appendChild(containerEl);
            }

            tree.appendChild(slideEl);
        }

        const downloadPptxBtn = document.createElement('button');
        downloadPptxBtn.type = 'button';
        downloadPptxBtn.textContent = 'Download PPTX';
        const pptxLabel = downloadPptxBtn.textContent;
        downloadPptxBtn.addEventListener('click', async () => {
            downloadPptxBtn.disabled = true;
            downloadPptxBtn.textContent = 'Preparing…';
            try {
                const renames = computeShapeRenames(draft.flatEntries);
                const blob = await applyShapeNamesToPptx(markedBlob, renames);
                triggerBlobDownload(blob, markedFilename);
            } catch (err) {
                markupError.textContent = 'Error: ' + err.message;
            } finally {
                downloadPptxBtn.disabled = false;
                downloadPptxBtn.textContent = pptxLabel;
            }
        });
        downloadsWrap.appendChild(downloadPptxBtn);

        const xlsxName = markedFilename.replace(/\.pptx$/i, '.xlsx');
        const downloadXlsxBtn = document.createElement('button');
        downloadXlsxBtn.type = 'button';
        downloadXlsxBtn.textContent = 'Download XLSX';
        downloadXlsxBtn.addEventListener('click', () => {
            const rows = buildXlsxRowsFromEntries(draft.flatEntries, lang);
            const blob = createWorkbookBlob(rows, 'Translations');
            triggerBlobDownload(blob, xlsxName);
        });
        downloadsWrap.appendChild(downloadXlsxBtn);

        const allMessages = [
            ...(markupResult ? markupResult.warnings : []),
            ...(extractWarnings || [])
        ];
        renderKeyEditorWarnings(warningsHost, allMessages);
        renderKeyEditorReport(markupDownload, [
            `Marked ${markupResult.renamedCount} of ${markupResult.containerCount} text containers.`,
            `Kept ${markupResult.keptCount} existing translation IDs.`,
            `Extracted ${draft.flatEntries.length} paragraph entries for ${lang}.`
        ], allMessages, `${baseName}_prepare`);

        refreshContainerProjections();
        refreshBasenameAvailability();
        refreshSummary();
    }

    function renderKeyEditorWarnings(host, messages) {
        host.innerHTML = '';
        if (!messages || !messages.length) return;

        const migrate = messages.filter(m => m.severity !== 'info' && m.migrate);
        const otherWarnings = messages.filter(m => m.severity !== 'info' && !m.migrate);

        if (migrate.length) host.appendChild(buildMigrateWarningsSection(migrate));
        if (otherWarnings.length) host.appendChild(buildPlainWarningsSection(otherWarnings));
    }

    function buildMigrateWarningsSection(items) {
        const wrap = document.createElement('div');
        wrap.className = 'key-editor-warnings-section';

        const title = document.createElement('div');
        title.className = 'key-editor-warnings-title';
        title.textContent = `Previous XLSX: ${items.length} key${items.length === 1 ? '' : 's'} could not be migrated`;
        wrap.appendChild(title);

        const table = document.createElement('table');
        table.className = 'key-editor-warnings-table';

        const thead = document.createElement('thead');
        const headRow = document.createElement('tr');
        for (const label of ['Old key', 'Source text', 'Shape', 'Slide', 'Issue']) {
            const th = document.createElement('th');
            th.textContent = label;
            headRow.appendChild(th);
        }
        thead.appendChild(headRow);
        table.appendChild(thead);

        const tbody = document.createElement('tbody');
        for (const w of items) {
            const tr = document.createElement('tr');
            const kindClass = w.kind === 'ambiguous' ? 'is-ambiguous' : 'is-orphaned';
            tr.classList.add(kindClass);

            const keyCell = document.createElement('td');
            const keyCode = document.createElement('code');
            keyCode.textContent = w.key || '';
            keyCell.appendChild(keyCode);
            tr.appendChild(keyCell);

            const textCell = document.createElement('td');
            textCell.className = 'key-editor-warnings-text';
            textCell.textContent = w.text || '—';
            if (w.text) textCell.title = w.text;
            tr.appendChild(textCell);

            const shapeCell = document.createElement('td');
            if (w.shapeName) {
                const shapeCode = document.createElement('code');
                shapeCode.textContent = w.shapeName;
                shapeCell.appendChild(shapeCode);
            } else {
                shapeCell.textContent = '—';
            }
            tr.appendChild(shapeCell);

            const slideCell = document.createElement('td');
            slideCell.textContent = w.slideNumber || '—';
            tr.appendChild(slideCell);

            const issueCell = document.createElement('td');
            issueCell.className = 'key-editor-warnings-issue ' + kindClass;
            issueCell.textContent = formatMigrateIssue(w);
            tr.appendChild(issueCell);

            tbody.appendChild(tr);
        }
        table.appendChild(tbody);
        wrap.appendChild(table);
        return wrap;
    }

    function formatMigrateIssue(w) {
        if (w.kind === 'ambiguous') {
            if (w.reason === 'matches-multiple') return `Matches ${w.matchCount} paragraphs — skipped`;
            if (w.reason === 'conflict') return 'Conflicts with another migrated key — skipped';
            return 'Ambiguous — skipped';
        }
        if (w.reason === 'source-not-found') return 'Source text not found (cell removed or edited)';
        if (w.reason === 'already-migrated') return 'All matching paragraphs already taken';
        if (w.reason === 'no-shape-name') return 'No shapeName column in old XLSX';
        if (w.reason === 'empty-text') return 'Empty source text in old XLSX';
        return 'Orphaned';
    }

    function buildPlainWarningsSection(items) {
        const wrap = document.createElement('div');
        wrap.className = 'key-editor-warnings-section';

        const title = document.createElement('div');
        title.className = 'key-editor-warnings-title';
        title.textContent = `${items.length} warning${items.length === 1 ? '' : 's'}`;
        wrap.appendChild(title);

        const list = document.createElement('ul');
        list.className = 'key-editor-warnings-list';
        const limit = 12;
        for (const w of items.slice(0, limit)) {
            const li = document.createElement('li');
            li.textContent = w.message;
            list.appendChild(li);
        }
        if (items.length > limit) {
            const more = document.createElement('li');
            more.className = 'key-editor-warnings-more';
            more.textContent = `...and ${items.length - limit} more (see full report)`;
            list.appendChild(more);
        }
        wrap.appendChild(list);
        return wrap;
    }

    function renderKeyEditorReport(target, summaryLines, messages, reportBaseName) {
        const info = document.createElement('div');
        info.style.cssText = 'color: #6bbd6b; font-size: 0.85rem; margin-top: 12px;';
        info.innerHTML = summaryLines.map(line => `<div>${line}</div>`).join('');
        target.appendChild(info);

        const informational = messages.filter(m => m.severity === 'info');
        const warnings = messages.filter(m => m.severity !== 'info');

        if (informational.length) {
            const infoBox = document.createElement('div');
            infoBox.style.cssText = 'color: #9aa0a6; font-size: 0.85rem; margin-top: 8px;';
            infoBox.appendChild(document.createTextNode(`${informational.length} informational message(s):`));
            appendMessageList(infoBox, informational);
            target.appendChild(infoBox);
        }

        if (warnings.length || informational.length) {
            const reportBlob = createWarningsBlob(messages);
            const reportLink = createDownloadLink(
                reportBlob,
                `${reportBaseName}_report.xlsx`,
                'Download full report'
            );
            const reportWrap = document.createElement('div');
            reportWrap.style.cssText = 'margin-top: 6px; font-size: 0.85rem;';
            reportWrap.appendChild(reportLink);
            target.appendChild(reportWrap);
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

    function getStyleLanguages(xlsxData) {
        if (xlsxData && xlsxData.length) {
            const header = xlsxData[0].map(value => String(value || '').trim());
            const languages = getLanguageColumns(header);
            if (languages.length) return languages;
            return [];
        }
        const defaultLanguage = translationsToPptxConfig.defaultLanguage || DEFAULT_TRANSLATIONS_TO_PPTX_CONFIG.defaultLanguage;
        return [{ code: defaultLanguage, index: null }];
    }

    function getRowIdentity(row, headerIndex) {
        const key = String(row[headerIndex.get('key')] || '').trim();
        if (key) return key;

        const shapeName = String(row[headerIndex.get('shapeName')] || '').trim();
        const paragraphIndex = String(row[headerIndex.get('paragraphIndex')] ?? '').trim();
        if (shapeName && paragraphIndex) return `${shapeName}__p${paragraphIndex}`;
        return '';
    }

    function getOptionalCell(row, headerIndex, column) {
        const index = headerIndex.get(column);
        return index === undefined ? '' : row[index];
    }

    function parseTranslationKey(key) {
        const value = String(key || '').trim();

        const pMatch = value.match(AUTO_KEY_PARAGRAPH_RE);
        if (pMatch) {
            return {
                shapeName: pMatch[1],
                paragraphIndex: parseInt(pMatch[2], 10),
                enHash: '',
                collisionIndex: 0,
                isCustom: false
            };
        }

        const hMatch = value.match(AUTO_KEY_HASH_RE);
        if (hMatch) {
            return {
                shapeName: hMatch[1],
                paragraphIndex: NaN,
                enHash: hMatch[2].toLowerCase(),
                collisionIndex: hMatch[3] ? parseInt(hMatch[3], 10) : 1,
                isCustom: false
            };
        }

        return {
            shapeName: '',
            paragraphIndex: NaN,
            enHash: '',
            collisionIndex: 0,
            isCustom: Boolean(value)
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
                    slideNumbers: new Set(),
                    examples: []
                });
            }
            const group = groupsByKey.get(row.styleKey);
            group.count++;
            if (row.slideNumber) group.slideNumbers.add(String(row.slideNumber));
            if (group.examples.length < 3) {
                group.examples.push({
                    text: row.textSample,
                    slideNumber: row.slideNumber ? String(row.slideNumber) : '',
                    traits: row.styleTraits || {}
                });
            }
        }

        return Array.from(groupsByKey.values()).map(group => ({
            ...group,
            slideNumbers: Array.from(group.slideNumbers).sort((a, b) => (parseNumber(a) || 0) - (parseNumber(b) || 0))
        })).sort((a, b) => {
            const aPt = parseNumber(a.fontPt) || 0;
            const bPt = parseNumber(b.fontPt) || 0;
            if (aPt !== bPt) return aPt - bPt;
            return a.styleLabel.localeCompare(b.styleLabel);
        });
    }

    function buildStyleState(styleRows, xlsxData = null) {
        const languages = getStyleLanguages(xlsxData);
        const groups = buildStyleGroups(styleRows);
        return {
            styleRows,
            styleRowsByKey: new Map(styleRows.map(row => [row.key, row])),
            groups,
            groupsByKey: new Map(groups.map(group => [group.styleKey, group])),
            languages,
            sizeInputs: new Map(),
            multiplierSelects: new Map(),
            selectedStyleKeys: [],
            styleAliases: new Map()
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

    function resolveStyleKey(styleKey) {
        if (!styleState || !styleKey) return styleKey;
        let resolvedKey = styleKey;
        const seen = new Set();
        while (styleState.styleAliases.has(resolvedKey) && !seen.has(resolvedKey)) {
            seen.add(resolvedKey);
            resolvedKey = styleState.styleAliases.get(resolvedKey);
        }
        return resolvedKey;
    }

    function getMergeReferenceKey(styleKey) {
        const resolvedKey = resolveStyleKey(styleKey);
        return resolvedKey && resolvedKey !== styleKey ? resolvedKey : '';
    }

    function getMergedAliasKeys(referenceKey) {
        if (!styleState || !referenceKey) return [];
        return Array.from(styleState.styleAliases.entries())
            .filter(([, targetKey]) => resolveStyleKey(targetKey) === referenceKey)
            .map(([aliasKey]) => aliasKey);
    }

    function setSelectedStyleKey(styleKey, checked) {
        if (!styleState) return;
        const selected = styleState.selectedStyleKeys.filter(key => key !== styleKey);
        if (checked) selected.push(styleKey);
        styleState.selectedStyleKeys = selected;
    }

    function mergeSelectedStyles() {
        if (!styleState || styleState.selectedStyleKeys.length < 2) return;
        const currentPreset = createStylePreset();
        const referenceKey = resolveStyleKey(styleState.selectedStyleKeys[0]);
        for (const styleKey of styleState.selectedStyleKeys.slice(1)) {
            const aliasKey = resolveStyleKey(styleKey) === referenceKey ? styleKey : resolveStyleKey(styleKey);
            if (aliasKey && aliasKey !== referenceKey) styleState.styleAliases.set(aliasKey, referenceKey);
        }
        styleState.selectedStyleKeys = [];
        if (currentPreset) currentPreset.styleAliases = styleAliasesToObject();
        renderStyleControls();
        if (currentPreset) applyStylePresetValues(currentPreset);
    }

    function unmergeSelectedStyles() {
        if (!styleState || !styleState.selectedStyleKeys.length) return;
        const currentPreset = createStylePreset();
        for (const styleKey of styleState.selectedStyleKeys) {
            styleState.styleAliases.delete(styleKey);
            for (const aliasKey of getMergedAliasKeys(styleKey)) {
                styleState.styleAliases.delete(aliasKey);
            }
        }
        styleState.selectedStyleKeys = [];
        if (currentPreset) currentPreset.styleAliases = styleAliasesToObject();
        renderStyleControls();
        if (currentPreset) applyStylePresetValues(currentPreset);
    }

    function normalizeStyleAliases(aliases = {}) {
        const result = new Map();
        if (!styleState || !aliases || typeof aliases !== 'object') return result;
        for (const [aliasKey, referenceKey] of Object.entries(aliases)) {
            if (!styleState.groupsByKey.has(aliasKey) || !styleState.groupsByKey.has(referenceKey)) continue;
            if (aliasKey !== referenceKey) result.set(aliasKey, referenceKey);
        }
        return result;
    }

    function styleAliasesToObject() {
        if (!styleState) return {};
        return Object.fromEntries(styleState.styleAliases.entries());
    }

    function getStyleOverrideForLanguage(languageCode, styleKey, fallbackFontPt = '') {
        if (!styleState || !styleKey) return null;
        const resolvedKey = resolveStyleKey(styleKey);
        const group = styleState.groupsByKey.get(resolvedKey);
        if (!group) return null;
        const example = group.examples[0] || {};
        const traits = { ...(example.traits || {}) };
        const groupFontPt = getStyleSizeForLanguage(languageCode, resolvedKey);
        const fontPt = groupFontPt || parseNumber(traits.fontPt) || fallbackFontPt || '';
        if (fontPt) traits.fontPt = fontPt;
        return {
            styleKey,
            resolvedStyleKey: resolvedKey,
            traits,
            fontPt
        };
    }

    function getStyleMultiplierForLanguage(languageCode) {
        const select = styleState && styleState.multiplierSelects.get(languageCode);
        return select ? (parseNumber(select.value) || 100) / 100 : 1;
    }

    function updateStylePresetButton() {
        if (!styleSavePresetBtn) return;
        styleSavePresetBtn.disabled = !(styleState && stylePptxBuffer && styleState.languages.length);
    }

    function captureStyleUiState() {
        if (!styleState) return null;
        const snapshot = {
            styleAliases: styleAliasesToObject(),
            selectedStyleKeys: [...styleState.selectedStyleKeys],
            languages: {}
        };
        for (const language of styleState.languages) {
            const languageInputs = styleState.sizeInputs.get(language.code);
            if (!languageInputs) continue;
            snapshot.languages[language.code] = {};
            for (const [styleKey, input] of languageInputs.entries()) {
                snapshot.languages[language.code][styleKey] = input.value;
            }
        }
        return snapshot;
    }

    function restoreStyleUiState(snapshot) {
        if (!snapshot || !styleState) return;
        styleState.styleAliases = normalizeStyleAliases(snapshot.styleAliases || {});
        styleState.selectedStyleKeys = (snapshot.selectedStyleKeys || [])
            .filter(styleKey => styleState.groupsByKey.has(styleKey));
        for (const language of styleState.languages) {
            const savedSizes = snapshot.languages[language.code];
            const languageInputs = styleState.sizeInputs.get(language.code);
            if (!savedSizes || !languageInputs) continue;
            for (const [styleKey, value] of Object.entries(savedSizes)) {
                const input = languageInputs.get(styleKey);
                if (input) input.value = value;
            }
        }
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
            version: translationsToPptxConfig.preset.version,
            type: translationsToPptxConfig.preset.type,
            styleAliases: styleAliasesToObject(),
            languages
        };
    }

    function saveStylePreset() {
        const preset = createStylePreset();
        if (!preset) return;
        const blob = new Blob([JSON.stringify(preset, null, 2)], { type: 'application/json' });
        const baseName = stylePptxName ? stylePptxName.replace(/\.pptx$/i, '') : 'pptx';
        triggerBlobDownload(blob, `${baseName}${translationsToPptxConfig.preset.fileSuffix}`);
    }

    function validateStylePreset(value) {
        const supportedVersions = [1, translationsToPptxConfig.preset.version];
        if (!value || value.type !== translationsToPptxConfig.preset.type || !supportedVersions.includes(value.version) || typeof value.languages !== 'object') {
            throw new Error(translationsToPptxMessage('invalidPreset'));
        }
        return value;
    }

    function applyStylePresetValues(preset) {
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

    function applyStylePreset(preset) {
        pendingStylePreset = preset;
        if (!styleState) return;
        styleState.styleAliases = normalizeStyleAliases(preset.styleAliases || {});
        renderStyleControls();
        applyStylePresetValues(preset);
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
        else if (traits.colorHex && /^#[0-9a-f]{6}$/i.test(traits.colorHex)) css.push(`color: ${traits.colorHex}`);
        if (traits.bold === '1') css.push('font-weight: 700');
        if (traits.italic === '1') css.push('font-style: italic');
        if (traits.caps === 'all') css.push('text-transform: uppercase');
        if (traits.lineSpacing) css.push(`line-height: ${traits.lineSpacing}`);
        return css.join('; ');
    }

    function getGroupTraits(group) {
        return group.examples[0] ? group.examples[0].traits || {} : {};
    }

    function formatFontStyles(traits = {}) {
        const fontStyles = [];
        if (traits.italic === '1') fontStyles.push('Italic');
        if (traits.caps === 'all') fontStyles.push('Uppercase');
        if (traits.caps === 'small') fontStyles.push('Small caps');
        return fontStyles.join(', ');
    }

    function formatColorDisplay(traits = {}) {
        const raw = traits.color ? String(traits.color).replace(/^color\s+/i, '') : '';
        const hex = traits.colorHex ? String(traits.colorHex).toUpperCase() : '';
        if (raw.startsWith('scheme:') && hex) return `${raw} (${hex})`;
        if (hex && /^#[0-9a-f]{6}$/i.test(hex)) return hex;
        return raw;
    }

    function getStyleTableValues(group) {
        const traits = getGroupTraits(group);
        return {
            size: traits.fontPt ? String(traits.fontPt) : '',
            weight: traits.bold === '1' ? 'Bold' : '',
            color: formatColorDisplay(traits),
            lineHeight: traits.lineSpacing ? String(traits.lineSpacing) : '',
            style: formatFontStyles(traits),
            font: traits.fontFamily ? String(traits.fontFamily) : ''
        };
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
            .map(example => {
                const traits = example.traits || {};
                return traits.colorHex || traits.color;
            })
            .find(color => /^#[0-9a-f]{6}$/i.test(color || ''));
        return firstColor ? getContrastingColor(firstColor) : '#ffffff';
    }

    function formatSlideNumbers(slideNumbers = []) {
        const numbers = Array.from(new Set(slideNumbers.filter(Boolean)));
        if (!numbers.length) return '';
        return `${translationsToPptxLabel('slides')}: ${numbers.join(', ')}`;
    }

    function formatStyleDetails(traits = {}, group = null) {
        const details = [];
        if (traits.fontPt) details.push(`Size: ${traits.fontPt}`);
        if (traits.bold === '1') details.push('Weight: Bold');
        if (traits.italic === '1') details.push('Style: Italic');
        if (traits.caps === 'all') details.push('Style: Uppercase');
        if (traits.caps === 'small') details.push('Style: Small caps');
        if (traits.lineSpacing) details.push(`line-height: ${traits.lineSpacing}`);
        const colorDetail = formatColorDisplay(traits);
        if (colorDetail) details.push(`Color: ${colorDetail}`);
        if (traits.fontFamily) details.push(`Font: ${traits.fontFamily}`);
        const slides = group ? formatSlideNumbers(group.slideNumbers) : '';
        if (slides) details.push(slides);
        if (traits.source) details.push(`Source: ${traits.source}`);
        if (traits.fallback) details.push(`Fallback: ${traits.fallback}`);
        return details;
    }

    function appendStyleExample(parent, example, background) {
        const item = document.createElement('div');
        item.style.cssText = `margin-top: 6px; padding: 4px 6px; border-radius: 4px; background: ${background}; ${styleTraitsToCss(example.traits)}`;

        if (example.slideNumber) {
            const slideMeta = document.createElement('div');
            slideMeta.textContent = `${translationsToPptxLabel('slide')}: ${example.slideNumber}`;
            slideMeta.style.cssText = 'font-size: 0.8rem; opacity: 0.6; margin-bottom: 2px;';
            item.appendChild(slideMeta);
        }

        const text = document.createElement('div');
        text.textContent = example.text || '(empty text)';
        item.appendChild(text);
        parent.appendChild(item);
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

        const details = formatStyleDetails(group.examples[0] ? group.examples[0].traits : {}, group);
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

        for (const example of group.examples) appendStyleExample(card, example, cardBackground);

        return card;
    }

    function buildStylePreviewPopup(group, styleNumber) {
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

        const details = formatStyleDetails(group.examples[0] ? group.examples[0].traits : {}, group);
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

        for (const example of group.examples) appendStyleExample(popup, example, popupBackground);
        document.body.appendChild(popup);

        let hideTimer = null;
        const cancelHide = () => {
            if (hideTimer) {
                clearTimeout(hideTimer);
                hideTimer = null;
            }
        };
        const scheduleHide = () => {
            cancelHide();
            hideTimer = setTimeout(() => {
                popup.style.display = 'none';
                hideTimer = null;
            }, 120);
        };

        popup.addEventListener('mouseenter', cancelHide);
        popup.addEventListener('mouseleave', scheduleHide);

        return {
            show(trigger) {
                if (!trigger) return;
                cancelHide();
                const rect = trigger.getBoundingClientRect();
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
            },
            hide: scheduleHide,
            cancelHide,
            destroy() {
                cancelHide();
                popup.remove();
            }
        };
    }

    function createStylePreviewButton(labelText, popupApi, onSelect) {
        const label = document.createElement('button');
        label.type = 'button';
        label.textContent = labelText;
        label.style.cssText = 'background: none; border: 0; color: inherit; padding: 0; text-align: left; cursor: pointer; text-decoration: underline dotted; white-space: nowrap;';
        label.addEventListener('mouseenter', () => popupApi.show(label));
        label.addEventListener('mouseleave', popupApi.hide);
        label.addEventListener('focus', () => {
            popupApi.cancelHide();
            popupApi.show(label);
        });
        label.addEventListener('blur', popupApi.hide);
        label.addEventListener('click', (e) => {
            e.preventDefault();
            onSelect();
        });
        return label;
    }

    function createStylePreviewLabel(group, styleNumber, onSelect, labelText = '') {
        const popupApi = buildStylePreviewPopup(group, styleNumber);
        return createStylePreviewButton(
            labelText || group.styleLabel || `#${styleNumber}`,
            popupApi,
            () => onSelect(group, styleNumber)
        );
    }

    function buildStyleGenerationRows(data, language, options = {}) {
        if (!data.length) throw new Error('XLSX is empty.');

        const header = data[0].map(value => String(value || '').trim());
        const headerIndex = buildHeaderIndex(header);
        if (!headerIndex.has('key')) throw new Error('XLSX is missing required column "key".');

        const fontColumn = metricColumnName(language.code, 'fontPt');
        const fontColumnIndex = headerIndex.get(fontColumn);

        // Detect the source-language column for routing custom keys.
        // The Prepare PPTX step writes the source language first (typically EN);
        // we prefer a language column different from the target.
        const languageColumns = getLanguageColumns(header);
        let sourceColumnIndex = null;
        const preferred = languageColumns.find(col => col.code !== language.code);
        if (preferred) sourceColumnIndex = preferred.index;
        else if (languageColumns.length) sourceColumnIndex = languageColumns[0].index;

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
            const fallbackFontPt = fontColumnIndex === undefined ? '' : parseNumber(sourceRow[fontColumnIndex]);
            const styleOverride = styleRow
                ? getStyleOverrideForLanguage(language.code, styleRow.styleKey, fallbackFontPt)
                : null;
            const baseFontPt = styleOverride ? styleOverride.fontPt : fallbackFontPt || '';
            const fontPt = baseFontPt ? roundToHalf(baseFontPt) : '';
            if (styleOverride && fontPt) styleOverride.fontPt = fontPt;
            if (styleOverride && fontPt) styleOverride.traits.fontPt = fontPt;

            rows.push({
                rowNumber: r + 1,
                key,
                slideNumber: String(getOptionalCell(sourceRow, headerIndex, 'slideNumber') || '').trim(),
                shapeName: String(getOptionalCell(sourceRow, headerIndex, 'shapeName') || parsedKey.shapeName || '').trim(),
                shapeId: String(getOptionalCell(sourceRow, headerIndex, 'shapeId') || '').trim(),
                containerType: String(getOptionalCell(sourceRow, headerIndex, 'containerType') || '').trim(),
                paragraphIndex: parsedKey.paragraphIndex,
                enHash: parsedKey.enHash,
                collisionIndex: parsedKey.collisionIndex,
                isCustom: parsedKey.isCustom,
                sourceText: sourceColumnIndex !== null ? String(sourceRow[sourceColumnIndex] || '') : '',
                text: String(value),
                fontPt,
                fontRaw: fontPt ? String(fontPt) : '',
                fontColumn: fontColumnIndex === undefined ? 'style table' : fontColumn,
                styleOverride
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
                    `Applied ${result.fontSizeAppliedCount} style override value(s).`,
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
            styleTableArea.textContent = translationsToPptxMessage('loadPptxFirst');
            return;
        }

        if (!styleState || !styleState.languages.length) {
            styleTableArea.textContent = translationsToPptxMessage('noLanguageColumns');
            return;
        }

        if (!styleState.groups.length) {
            styleTableArea.textContent = translationsToPptxMessage('noEditableStyles');
            return;
        }

        if (stylePresetControls) stylePresetControls.style.display = 'flex';

        const layout = document.createElement('div');
        layout.style.cssText = 'display: flex; align-items: flex-start; gap: 16px; margin-top: 12px;';

        const tableWrap = document.createElement('div');
        tableWrap.style.cssText = 'overflow-x: auto; flex: 1 1 auto; min-width: 0;';

        if (!styleTranslationData) {
            const xlsxHint = document.createElement('div');
            xlsxHint.textContent = translationsToPptxMessage('loadXlsxForExport');
            xlsxHint.style.cssText = 'font-size: 0.85rem; opacity: 0.7; margin-bottom: 8px;';
            tableWrap.appendChild(xlsxHint);
        }

        const mergeToolbar = document.createElement('div');
        mergeToolbar.style.cssText = [
            'display: none',
            'position: fixed',
            'z-index: 10000',
            'align-items: center',
            'gap: 8px',
            'padding: 6px',
            'border: 1px solid rgba(255,255,255,0.18)',
            'border-radius: 8px',
            'background: rgba(20,20,20,0.92)',
            'box-shadow: 0 8px 24px rgba(0,0,0,0.28)'
        ].join('; ');

        const mergeButton = document.createElement('button');
        mergeButton.type = 'button';
        mergeButton.textContent = translationsToPptxLabel('mergeStyles');
        mergeButton.style.cssText = 'padding: 6px 10px; border-radius: 6px; border: 1px solid rgba(255,255,255,0.25); background: #2f8f46; color: #fff; font-weight: 600; cursor: pointer;';
        mergeButton.addEventListener('click', (e) => {
            e.preventDefault();
            mergeSelectedStyles();
        });

        const unmergeButton = document.createElement('button');
        unmergeButton.type = 'button';
        unmergeButton.textContent = translationsToPptxLabel('unmergeStyles');
        unmergeButton.style.cssText = 'padding: 6px 10px; border-radius: 6px; border: 1px solid rgba(255,255,255,0.25); background: transparent; color: inherit; font-weight: 600; cursor: pointer;';
        unmergeButton.addEventListener('click', (e) => {
            e.preventDefault();
            unmergeSelectedStyles();
        });

        const mergeHint = document.createElement('span');
        mergeHint.style.cssText = 'font-size: 0.8rem; opacity: 0.7;';

        mergeToolbar.appendChild(mergeButton);
        mergeToolbar.appendChild(unmergeButton);
        mergeToolbar.appendChild(mergeHint);
        tableWrap.appendChild(mergeToolbar);

        const styleLabelCells = new Map();
        const usesCells = new Map();
        const updateMergeToolbar = () => {
            const selectedCount = styleState.selectedStyleKeys.length;
            const hasMergeSelection = selectedCount >= 2;
            const hasUnmergeSelection = styleState.selectedStyleKeys.some(key =>
                styleState.styleAliases.has(key) || getMergedAliasKeys(key).length
            );
            mergeToolbar.style.display = selectedCount ? 'flex' : 'none';
            mergeButton.style.display = hasMergeSelection ? 'inline-flex' : 'none';
            unmergeButton.style.display = hasUnmergeSelection ? 'inline-flex' : 'none';
            const referenceKey = selectedCount ? resolveStyleKey(styleState.selectedStyleKeys[0]) : '';
            const referenceIndex = referenceKey
                ? styleState.groups.findIndex(group => group.styleKey === referenceKey) + 1
                : 0;
            mergeHint.textContent = hasMergeSelection && referenceIndex
                ? `Reference: #${referenceIndex}`
                : `${selectedCount} selected`;

            if (!selectedCount) return;
            const anchor = usesCells.get(styleState.selectedStyleKeys[0])
                || styleLabelCells.get(styleState.selectedStyleKeys[0]);
            if (!anchor) return;
            const rect = anchor.getBoundingClientRect();
            const toolbarRect = mergeToolbar.getBoundingClientRect();
            const preferredLeft = rect.right + 8;
            const maxLeft = Math.max(12, window.innerWidth - toolbarRect.width - 12);
            const left = Math.max(12, Math.min(preferredLeft, maxLeft));
            const top = Math.min(
                Math.max(12, rect.top + (rect.height - toolbarRect.height) / 2),
                window.innerHeight - toolbarRect.height - 12
            );
            mergeToolbar.style.left = `${left}px`;
            mergeToolbar.style.top = `${top}px`;
        };

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
        sidePanelTitle.textContent = translationsToPptxLabel('selectedStylesTitle');
        sidePanelTitle.style.cssText = 'font-weight: 600; margin-bottom: 8px;';
        sidePanel.appendChild(sidePanelTitle);

        const emptySidePanel = document.createElement('div');
        emptySidePanel.textContent = translationsToPptxLabel('emptySelectedStyles');
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

        const styleHeads = ['Size', 'Weight', 'Color', 'Line-height', 'Style', 'Font'];

        const selectHead = document.createElement('th');
        selectHead.textContent = '';

        const usesHead = document.createElement('th');
        usesHead.textContent = translationsToPptxLabel('usesHeader');
        for (const heading of styleHeads) {
            const styleHead = document.createElement('th');
            styleHead.textContent = heading;
            languageRow.appendChild(styleHead);
        }
        languageRow.appendChild(selectHead);
        languageRow.appendChild(usesHead);

        for (let i = 0; i < styleHeads.length; i++) {
            const styleBlank = document.createElement('th');
            styleBlank.textContent = '';
            multiplierRow.appendChild(styleBlank);
        }

        const selectBlank = document.createElement('th');
        selectBlank.textContent = '';
        multiplierRow.appendChild(selectBlank);

        const usesBlank = document.createElement('th');
        usesBlank.textContent = '';
        multiplierRow.appendChild(usesBlank);

        thead.appendChild(languageRow);
        thead.appendChild(multiplierRow);
        table.appendChild(thead);

        const tbody = document.createElement('tbody');
        for (const [index, group] of styleState.groups.entries()) {
            const styleNumber = index + 1;
            const row = document.createElement('tr');
            const mergeReferenceKey = getMergeReferenceKey(group.styleKey);
            const isReference = !mergeReferenceKey && getMergedAliasKeys(group.styleKey).length > 0;
            const isAlias = Boolean(mergeReferenceKey);
            if (isReference) {
                row.style.background = 'rgba(47, 200, 85, 0.34)';
                row.style.boxShadow = 'inset 4px 0 0 #35d46d';
            } else if (isAlias) {
                row.style.background = 'rgba(47, 143, 70, 0.18)';
                if (isAlias) row.style.opacity = '0.5';
            }

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

            const styleValues = getStyleTableValues(group);
            const styleFields = [
                styleValues.size,
                styleValues.weight,
                styleValues.color,
                styleValues.lineHeight,
                styleValues.style,
                styleValues.font
            ];
            const popupApi = buildStylePreviewPopup(group, styleNumber);
            const openStyleCard = () => addStyleCard(group, styleNumber);

            for (const [fieldIndex, value] of styleFields.entries()) {
                const labelCell = document.createElement('td');
                labelCell.style.position = 'relative';
                labelCell.appendChild(createStylePreviewButton(value || '-', popupApi, openStyleCard));
                if (fieldIndex === 0) {
                    styleLabelCells.set(group.styleKey, labelCell);
                    if (isAlias) {
                        const reference = document.createElement('div');
                        reference.textContent = `Merged into #${styleState.groups.findIndex(item => item.styleKey === mergeReferenceKey) + 1}`;
                        reference.style.cssText = 'font-size: 0.75rem; opacity: 0.75; margin-top: 3px;';
                        labelCell.appendChild(reference);
                    }
                }
                row.appendChild(labelCell);
            }

            const selectCell = document.createElement('td');
            selectCell.style.cssText = 'text-align: center;';
            const checkbox = document.createElement('input');
            checkbox.type = 'checkbox';
            checkbox.checked = styleState.selectedStyleKeys.includes(group.styleKey);
            checkbox.title = isAlias
                ? `Merged into #${styleState.groups.findIndex(item => item.styleKey === mergeReferenceKey) + 1}`
                : '';
            checkbox.addEventListener('change', () => {
                setSelectedStyleKey(group.styleKey, checkbox.checked);
                updateMergeToolbar();
            });
            selectCell.appendChild(checkbox);
            row.appendChild(selectCell);

            const usesCell = document.createElement('td');
            usesCell.textContent = String(group.count);
            usesCells.set(group.styleKey, usesCell);
            row.appendChild(usesCell);

            tbody.appendChild(row);
        }
        table.appendChild(tbody);

        if (styleState.languages.length && styleTranslationData && styleState.languages.some(language => language.index !== null)) {
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

            const styleCell = document.createElement('td');
            styleCell.textContent = translationsToPptxLabel('downloadFooter');
            footerRow.appendChild(styleCell);

            for (let i = 1; i < styleHeads.length; i++) {
                const blankStyleCell = document.createElement('td');
                blankStyleCell.textContent = '';
                footerRow.appendChild(blankStyleCell);
            }

            const selectCell = document.createElement('td');
            selectCell.textContent = '';
            footerRow.appendChild(selectCell);

            const usesCell = document.createElement('td');
            usesCell.textContent = '';
            footerRow.appendChild(usesCell);

            tfoot.appendChild(footerRow);
            table.appendChild(tfoot);
        }

        tableWrap.appendChild(table);
        updateMergeToolbar();
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

        if (!window.PptxStyleUtils || typeof window.PptxStyleUtils.extractStyleRowsFromPptx !== 'function') {
            throw new Error(translationsToPptxMessage('styleUtilitiesMissing'));
        }
        const uiSnapshot = captureStyleUiState();
        const styleRows = await window.PptxStyleUtils.extractStyleRowsFromPptx(stylePptxBuffer.slice(0));
        styleState = buildStyleState(styleRows, styleTranslationData);
        if (uiSnapshot) {
            styleState.styleAliases = normalizeStyleAliases(uiSnapshot.styleAliases || {});
            styleState.selectedStyleKeys = (uiSnapshot.selectedStyleKeys || [])
                .filter(styleKey => styleState.groupsByKey.has(styleKey));
        }
        renderStyleControls();
        if (uiSnapshot) restoreStyleUiState(uiSnapshot);
        if (pendingStylePreset) applyStylePresetValues(pendingStylePreset);
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

    function removeDirectChildren(element, tagName) {
        for (const child of Array.from(element.children || [])) {
            if (getTagName(child) === tagName) element.removeChild(child);
        }
    }

    function setRunBooleanAttr(rPr, attrName, value) {
        if (value === '1' || value === true) {
            rPr.setAttribute(attrName, '1');
        } else if (value === '0' || value === false || value === '') {
            rPr.setAttribute(attrName, '0');
        }
    }

    function applyRunColor(rPr, color) {
        const value = String(color || '').replace('#', '');
        if (!/^[0-9a-f]{6}$/i.test(value)) return false;

        removeDirectChildren(rPr, 'a:solidFill');
        const solidFill = rPr.ownerDocument.createElementNS(A_NS, 'a:solidFill');
        const srgbClr = rPr.ownerDocument.createElementNS(A_NS, 'a:srgbClr');
        srgbClr.setAttribute('val', value.toUpperCase());
        solidFill.appendChild(srgbClr);
        rPr.appendChild(solidFill);
        return true;
    }

    function applyRunSchemeColor(rPr, schemeVal) {
        const value = String(schemeVal || '').trim();
        if (!value) return false;

        removeDirectChildren(rPr, 'a:solidFill');
        const solidFill = rPr.ownerDocument.createElementNS(A_NS, 'a:solidFill');
        const schemeClr = rPr.ownerDocument.createElementNS(A_NS, 'a:schemeClr');
        schemeClr.setAttribute('val', value);
        solidFill.appendChild(schemeClr);
        rPr.appendChild(solidFill);
        return true;
    }

    function applyRunFillColor(rPr, traits = {}) {
        const scheme = traits.colorScheme
            || (String(traits.color || '').startsWith('scheme:') ? String(traits.color).slice(7) : '');
        if (scheme) return applyRunSchemeColor(rPr, scheme);

        const hex = traits.colorHex || traits.color;
        return applyRunColor(rPr, hex);
    }

    function applyRunFontFamily(rPr, fontFamily) {
        const typeface = String(fontFamily || '').trim();
        if (!typeface) return false;

        for (const tagName of ['a:latin', 'a:cs', 'a:ea']) {
            const fontEl = ensureDirectChild(rPr, tagName, A_NS);
            fontEl.setAttribute('typeface', typeface);
        }
        return true;
    }

    function applyParagraphLineSpacing(paragraph, lineSpacing) {
        const value = String(lineSpacing || '').trim();
        if (!value) return false;

        const percentMatch = value.match(/^([\d.]+)%$/);
        const pointMatch = value.match(/^([\d.]+)pt$/i);
        if (!percentMatch && !pointMatch) return false;

        const pPr = ensureDirectChild(paragraph, 'a:pPr', A_NS, paragraph.firstElementChild || paragraph.firstChild);
        removeDirectChildren(pPr, 'a:lnSpc');

        const lnSpc = paragraph.ownerDocument.createElementNS(A_NS, 'a:lnSpc');
        if (percentMatch) {
            const percent = parseNumber(percentMatch[1]);
            if (!percent) return false;
            const spcPct = paragraph.ownerDocument.createElementNS(A_NS, 'a:spcPct');
            spcPct.setAttribute('val', String(Math.round(percent * 1000)));
            lnSpc.appendChild(spcPct);
        } else {
            const points = parseNumber(pointMatch[1]);
            if (!points) return false;
            const spcPts = paragraph.ownerDocument.createElementNS(A_NS, 'a:spcPts');
            spcPts.setAttribute('val', String(Math.round(points * 100)));
            lnSpc.appendChild(spcPts);
        }

        pPr.appendChild(lnSpc);
        return true;
    }

    function applyParagraphStyleOverride(paragraph, run, styleOverride = null) {
        const traits = styleOverride && styleOverride.traits ? styleOverride.traits : {};
        const fontPt = styleOverride ? styleOverride.fontPt : null;
        let applied = false;

        const firstChild = run.firstElementChild || run.firstChild;
        const rPr = ensureDirectChild(run, 'a:rPr', A_NS, firstChild);

        if (Number.isFinite(fontPt) && fontPt > 0) {
            const sz = String(Math.round(fontPt * 100));
            rPr.setAttribute('sz', sz);

            const endParaRPr = ensureDirectChild(paragraph, 'a:endParaRPr', A_NS);
            endParaRPr.setAttribute('sz', sz);
            applied = true;
        }

        setRunBooleanAttr(rPr, 'b', traits.bold);
        setRunBooleanAttr(rPr, 'i', traits.italic);
        if (traits.bold !== undefined || traits.italic !== undefined) applied = true;
        if (traits.caps) {
            rPr.setAttribute('cap', traits.caps);
            applied = true;
        }
        if (applyRunFillColor(rPr, traits)) applied = true;
        if (applyRunFontFamily(rPr, traits.fontFamily)) applied = true;
        if (applyParagraphLineSpacing(paragraph, traits.lineSpacing)) applied = true;

        return applied;
    }

    function mergeSegmentTraits(baseTraits, segment) {
        const traits = { ...(baseTraits || {}) };
        const baseBold = traits.bold === '1' || traits.bold === true;
        const baseItalic = traits.italic === '1' || traits.italic === true;
        const mergedBold = baseBold || Boolean(segment.bold);
        const mergedItalic = baseItalic || Boolean(segment.italic);
        traits.bold = mergedBold ? '1' : '0';
        traits.italic = mergedItalic ? '1' : '0';
        return traits;
    }

    function replaceParagraphText(paragraph, newText, styleOverride = null) {
        const runs = paragraph.getElementsByTagName('a:r');
        if (runs.length === 0) return { success: false, parseWarnings: [] };

        const rich = window.PptxRichText;
        const baseTraits = styleOverride && styleOverride.traits ? styleOverride.traits : {};
        const parseWarnings = [];

        if (rich && typeof rich.parseMarkdown === 'function') {
            const parsed = rich.parseMarkdown(newText);

            for (const warning of parsed.warnings) {
                parseWarnings.push(warning);
            }

            if (parsed.usedFallback) {
                const plainText = rich.unescapePlainText
                    ? rich.unescapePlainText(parsed.segments[0].text)
                    : parsed.segments[0].text;
                return replaceParagraphPlain(paragraph, runs, plainText, styleOverride, parseWarnings);
            }

            if (rich.shouldUseRichPath(newText, parsed) && typeof rich.applyToParagraph === 'function') {
                const applied = rich.applyToParagraph(paragraph, parsed.segments, {
                    templateRun: runs[0],
                    styleRun: (p, run, segment) => {
                        const mergedTraits = mergeSegmentTraits(baseTraits, segment);
                        const override = styleOverride
                            ? { ...styleOverride, traits: mergedTraits }
                            : { traits: mergedTraits };
                        applyParagraphStyleOverride(p, run, override);
                    }
                });
                return { success: applied, parseWarnings };
            }
        }

        const plainText = rich && typeof rich.unescapePlainText === 'function'
            ? rich.unescapePlainText(newText)
            : newText;
        return replaceParagraphPlain(paragraph, runs, plainText, styleOverride, parseWarnings);
    }

    function replaceParagraphPlain(paragraph, runs, newText, styleOverride, parseWarnings = []) {
        const firstRunT = runs[0].getElementsByTagName('a:t')[0];
        if (!firstRunT) return { success: false, parseWarnings };

        firstRunT.textContent = newText;
        firstRunT.setAttribute('xml:space', 'preserve');

        for (let ri = 1; ri < runs.length; ri++) {
            const t = runs[ri].getElementsByTagName('a:t')[0];
            if (t) t.textContent = '';
        }

        applyParagraphStyleOverride(paragraph, runs[0], styleOverride);
        return { success: true, parseWarnings };
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
            const hasParagraphIndex = Number.isFinite(row.paragraphIndex);
            const hasHashRoute = Boolean(row.enHash);
            const hasCustomRoute = Boolean(row.isCustom);

            if (!row.key || !row.shapeName || (!hasParagraphIndex && !hasHashRoute && !hasCustomRoute)) {
                warnings.push(createWarning('invalid-row', `XLSX row ${row.rowNumber}: missing or invalid key.`, {
                    key: row.key,
                    shapeName: row.shapeName,
                    paragraphIndex: hasParagraphIndex ? row.paragraphIndex : ''
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

            let paragraph = null;
            let resolvedParagraphIndex = hasParagraphIndex ? row.paragraphIndex : -1;

            if (hasHashRoute) {
                const candidates = container.nonEmptyParagraphs.filter(
                    p => shortHashFromText(p.text) === row.enHash
                );
                const targetIndex = Math.max(0, (row.collisionIndex || 1) - 1);
                if (!candidates.length || !candidates[targetIndex]) {
                    warnings.push(createWarning('text-not-found-by-hash', `XLSX row ${row.rowNumber}: no paragraph in "${row.shapeName}" matches the EN hash. Re-run Prepare PPTX.`, row));
                    continue;
                }
                paragraph = candidates[targetIndex].paragraph;
                resolvedParagraphIndex = candidates[targetIndex].paragraphIndex;
            } else if (hasCustomRoute) {
                const anchorText = row.sourceText && String(row.sourceText).trim()
                    ? row.sourceText
                    : row.text;
                const normAnchor = normalizeTextForKey(anchorText);
                if (!normAnchor) {
                    warnings.push(createWarning('custom-key-text-not-found', `XLSX row ${row.rowNumber}: custom key "${row.key}" has no source text to match; skipped.`, row));
                    continue;
                }
                const candidates = container.nonEmptyParagraphs.filter(
                    p => normalizeTextForKey(p.text) === normAnchor
                );
                if (candidates.length === 0) {
                    warnings.push(createWarning('custom-key-text-not-found', `XLSX row ${row.rowNumber}: custom key "${row.key}" — source text not found in "${row.shapeName}".`, row));
                    continue;
                }
                if (candidates.length > 1) {
                    warnings.push(createWarning('custom-key-ambiguous-en', `XLSX row ${row.rowNumber}: custom key "${row.key}" matches ${candidates.length} paragraphs in "${row.shapeName}"; skipped.`, row));
                    continue;
                }
                paragraph = candidates[0].paragraph;
                resolvedParagraphIndex = candidates[0].paragraphIndex;
            } else {
                paragraph = container.paragraphs[row.paragraphIndex];
                if (!paragraph) {
                    warnings.push(createWarning('missing-paragraph', `XLSX row ${row.rowNumber}: paragraph ${row.paragraphIndex} not found in "${row.shapeName}".`, row));
                    continue;
                }
            }

            if (row.fontRaw && (!Number.isFinite(row.fontPt) || row.fontPt <= 0)) {
                warnings.push(createWarning('invalid-font-size', `XLSX row ${row.rowNumber}: ${row.fontColumn} value "${row.fontRaw}" is not a valid font size; text was replaced without resizing.`, row));
            }

            const shouldApplyFont = Number.isFinite(row.fontPt) && row.fontPt > 0;
            const replaceResult = replaceParagraphText(
                paragraph,
                row.text,
                row.styleOverride || (shouldApplyFont ? { fontPt: row.fontPt, traits: { fontPt: row.fontPt } } : null)
            );
            for (const parseWarning of replaceResult.parseWarnings) {
                warnings.push(createWarning(parseWarning.type || 'invalid-rich-text', `${parseWarning.message} (XLSX row ${row.rowNumber}).`, row));
            }
            if (!replaceResult.success) {
                warnings.push(createWarning('missing-text-run', `XLSX row ${row.rowNumber}: paragraph ${resolvedParagraphIndex} in "${row.shapeName}" has no editable text run.`, row));
                continue;
            }

            if (row.styleOverride || shouldApplyFont) fontSizeAppliedCount++;
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

    async function init() {
        await loadTranslationsToPptxConfig();

        setupSingleDropzone(pptxMarkupDropzone, pptxMarkupInput, pptxMarkupFilename, '.pptx', (file) => {
            markupFile = file;
            showFilename(pptxMarkupFilename, file.name);
            markupBtn.disabled = false;
            markupDownload.innerHTML = '';
            markupError.textContent = '';
        });

        setupSingleDropzone(prepareOldXlsxDropzone, prepareOldXlsxInput, prepareOldXlsxFilename, '.xlsx', (file) => {
            previousXlsxFile = file;
            showFilename(prepareOldXlsxFilename, file.name);
            markupDownload.innerHTML = '';
            markupError.textContent = '';
        });

        if (markupBtn) markupBtn.addEventListener('click', handleMarkup);

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
    }

    init();
})();
