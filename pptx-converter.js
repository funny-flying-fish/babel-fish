(() => {
    const PPTX_MIME = 'application/vnd.openxmlformats-officedocument.presentationml.presentation';
    const XLSX_MIME = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
    const METADATA_COLUMNS = ['key', 'slideNumber', 'shapeName', 'shapeId', 'containerType', 'paragraphIndex'];
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

    // DOM refs: XLSX -> PPTX
    const xlsxInjectDropzone = document.getElementById('xlsxInjectDropzone');
    const xlsxInjectInput = document.getElementById('xlsxInjectInput');
    const xlsxInjectFilename = document.getElementById('xlsxInjectFilename');
    const pptxInjectDropzone = document.getElementById('pptxInjectDropzone');
    const pptxInjectInput = document.getElementById('pptxInjectInput');
    const pptxInjectFilename = document.getElementById('pptxInjectFilename');
    const langSelect = document.getElementById('langSelect');
    const injectBtn = document.getElementById('injectBtn');
    const injectDownload = document.getElementById('injectDownload');
    const injectError = document.getElementById('injectError');

    let markupFile = null;
    let extractFile = null;
    let injectPptxBuffer = null;
    let injectPptxName = '';
    let translationData = null;

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

    function getFirstDescendant(element, tagName) {
        return element.getElementsByTagName(tagName)[0] || null;
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
        const rows = [[...METADATA_COLUMNS, lang]];
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
                        slide.num,
                        shapeName,
                        container.shapeId,
                        container.containerType,
                        item.paragraphIndex,
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
            extractDownload.appendChild(createDownloadLink(blob, `${baseName}_translations.xlsx`));
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
            if (key) index.set(key, i);
        });
        return index;
    }

    function populateLanguageSelect(headerRow) {
        langSelect.innerHTML = '';
        const metadataSet = new Set(METADATA_COLUMNS);

        for (let i = 0; i < headerRow.length; i++) {
            const code = String(headerRow[i] || '').trim();
            if (!code || metadataSet.has(code)) continue;

            const opt = document.createElement('option');
            opt.value = i;
            opt.textContent = code;
            langSelect.appendChild(opt);
        }

        if (langSelect.options.length === 0) {
            const opt = document.createElement('option');
            opt.value = '';
            opt.textContent = 'No language columns found';
            langSelect.appendChild(opt);
        }
        langSelect.disabled = langSelect.options.length === 0 || !langSelect.options[0].value;
    }

    function buildTranslationRows(data, colIndex) {
        if (!data.length) throw new Error('XLSX is empty.');

        const header = data[0].map(value => String(value || '').trim());
        const headerIndex = buildHeaderIndex(header);
        const required = ['shapeName', 'paragraphIndex'];

        for (const column of required) {
            if (!headerIndex.has(column)) {
                throw new Error(`XLSX is missing required column "${column}". Re-export translations with the new PPTX converter.`);
            }
        }

        const rows = [];
        let skippedEmpty = 0;
        for (let r = 1; r < data.length; r++) {
            const sourceRow = data[r] || [];
            const value = sourceRow[colIndex];
            if (value === undefined || value === null || !String(value).trim()) {
                skippedEmpty++;
                continue;
            }

            rows.push({
                rowNumber: r + 1,
                key: String(sourceRow[headerIndex.get('key')] || '').trim(),
                slideNumber: String(sourceRow[headerIndex.get('slideNumber')] || '').trim(),
                shapeName: String(sourceRow[headerIndex.get('shapeName')] || '').trim(),
                shapeId: String(sourceRow[headerIndex.get('shapeId')] || '').trim(),
                containerType: String(sourceRow[headerIndex.get('containerType')] || '').trim(),
                paragraphIndex: parseInt(sourceRow[headerIndex.get('paragraphIndex')], 10),
                text: String(value)
            });
        }

        return { rows, skippedEmpty };
    }

    function replaceParagraphText(paragraph, newText) {
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

        return true;
    }

    async function injectTranslations(pptxBuffer, translationRows) {
        const zip = await JSZip.loadAsync(pptxBuffer);
        const slides = getSlideEntries(zip);
        const warnings = [];
        const slideDocs = new Map();
        const containersByName = new Map();
        let replacedCount = 0;

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
            if (!row.shapeName || Number.isNaN(row.paragraphIndex)) {
                warnings.push(createWarning('invalid-row', `XLSX row ${row.rowNumber}: missing shapeName or paragraphIndex.`, {
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

            if (!replaceParagraphText(paragraph, row.text)) {
                warnings.push(createWarning('missing-text-run', `XLSX row ${row.rowNumber}: paragraph ${row.paragraphIndex} in "${row.shapeName}" has no editable text run.`, row));
                continue;
            }

            slideDocs.get(container.slidePath).modified = true;
            replacedCount++;
        }

        for (const [path, entry] of slideDocs.entries()) {
            if (entry.modified) {
                zip.file(path, new XMLSerializer().serializeToString(entry.doc));
            }
        }

        const blob = await zip.generateAsync({ type: 'blob', mimeType: PPTX_MIME });
        return { blob, replacedCount, warnings };
    }

    async function handleInject() {
        injectError.textContent = '';
        injectDownload.innerHTML = '';

        if (!translationData || !injectPptxBuffer) return;

        const colIndex = parseInt(langSelect.value, 10);
        if (Number.isNaN(colIndex)) {
            injectError.textContent = 'Please select a language.';
            return;
        }

        try {
            injectBtn.disabled = true;
            injectBtn.textContent = 'Converting...';

            const { rows, skippedEmpty } = buildTranslationRows(translationData, colIndex);
            if (rows.length === 0) {
                injectError.textContent = 'No translations found for the selected language.';
                return;
            }

            const { blob, replacedCount, warnings } = await injectTranslations(injectPptxBuffer.slice(0), rows);
            const langCode = String(translationData[0][colIndex] || 'translated').trim();
            const baseName = injectPptxName.replace(/\.pptx$/i, '');
            const filename = `${baseName}_${langCode}.pptx`;

            injectDownload.appendChild(createDownloadLink(blob, filename));
            renderResult(injectDownload, [
                `Replaced ${replacedCount} of ${rows.length} translated paragraph entries.`,
                `Skipped ${skippedEmpty} empty translation cells.`
            ], warnings, `${baseName}_${langCode}`);
        } catch (err) {
            injectError.textContent = 'Error: ' + err.message;
        } finally {
            injectBtn.disabled = false;
            injectBtn.textContent = 'Convert';
            updateInjectButton();
        }
    }

    function updateInjectButton() {
        injectBtn.disabled = !(translationData && injectPptxBuffer && langSelect.value);
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

    setupSingleDropzone(xlsxInjectDropzone, xlsxInjectInput, xlsxInjectFilename, '.xlsx', async (file) => {
        showFilename(xlsxInjectFilename, file.name);
        injectDownload.innerHTML = '';
        injectError.textContent = '';

        try {
            translationData = await parseXlsxFile(file);
            if (translationData.length > 0) {
                populateLanguageSelect(translationData[0]);
            }
        } catch (err) {
            injectError.textContent = 'Error reading XLSX: ' + err.message;
            translationData = null;
        }
        updateInjectButton();
    });

    setupSingleDropzone(pptxInjectDropzone, pptxInjectInput, pptxInjectFilename, '.pptx', async (file) => {
        showFilename(pptxInjectFilename, file.name);
        injectDownload.innerHTML = '';
        injectError.textContent = '';
        injectPptxName = file.name;

        try {
            injectPptxBuffer = await readFileAsArrayBuffer(file);
        } catch (err) {
            injectError.textContent = 'Error reading PPTX: ' + err.message;
            injectPptxBuffer = null;
        }
        updateInjectButton();
    });

    langSelect.addEventListener('change', () => {
        injectDownload.innerHTML = '';
        updateInjectButton();
    });

    injectBtn.addEventListener('click', handleInject);
})();
