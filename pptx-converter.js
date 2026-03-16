(() => {
    // ── DOM refs ─────────────────────────────────────────────────────
    // Tab 1: PPTX → XLSX
    const pptxExtractDropzone = document.getElementById('pptxExtractDropzone');
    const pptxExtractInput = document.getElementById('pptxExtractInput');
    const pptxExtractFilename = document.getElementById('pptxExtractFilename');
    const sourceLanguage = document.getElementById('sourceLanguage');
    const extractBtn = document.getElementById('extractBtn');
    const extractDownload = document.getElementById('extractDownload');
    const extractError = document.getElementById('extractError');

    // Tab 2: XLSX → PPTX
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

    // ── State ────────────────────────────────────────────────────────
    let extractFile = null;
    let injectXlsxFile = null;
    let injectPptxBuffer = null;
    let injectPptxName = '';
    let translationData = null; // parsed 2D array from XLSX

    // ── Helpers ──────────────────────────────────────────────────────

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

    // ── Tab 1: PPTX → XLSX extraction ───────────────────────────────

    async function extractTextFromPptx(arrayBuffer, lang) {
        const zip = await JSZip.loadAsync(arrayBuffer);
        const slides = getSlideEntries(zip);
        const rows = [['', lang]];

        for (const slide of slides) {
            const xml = await zip.file(slide.path).async('string');
            const doc = new DOMParser().parseFromString(xml, 'application/xml');
            const shapes = doc.getElementsByTagName('p:sp');

            for (let si = 0; si < shapes.length; si++) {
                const paragraphs = shapes[si].getElementsByTagName('a:p');
                for (let pi = 0; pi < paragraphs.length; pi++) {
                    const text = getTextFromParagraph(paragraphs[pi]);
                    if (!text.trim()) continue;
                    const key = `s${slide.num}_sh${si}_p${pi}`;
                    rows.push([key, text]);
                }
            }
        }

        return rows;
    }

    function downloadXlsxFromRows(rows, filename) {
        const ws = XLSX.utils.aoa_to_sheet(rows);
        const wb = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(wb, ws, 'Sheet1');
        const data = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
        const blob = new Blob([data], {
            type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        });
        const url = URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = filename;
        a.textContent = `Download ${filename}`;
        a.style.color = '#4da6ff';
        return a;
    }

    async function handleExtract() {
        extractError.textContent = '';
        extractDownload.innerHTML = '';

        if (!extractFile) return;

        try {
            extractBtn.disabled = true;
            extractBtn.textContent = 'Extracting…';

            const buffer = await readFileAsArrayBuffer(extractFile);
            const lang = sourceLanguage.value.trim() || 'en_EN';
            const rows = await extractTextFromPptx(buffer, lang);

            const baseName = extractFile.name.replace(/\.pptx$/i, '');
            const link = downloadXlsxFromRows(rows, `${baseName}_translations.xlsx`);
            extractDownload.appendChild(link);

            const info = document.createElement('div');
            info.style.cssText = 'color: #6bbd6b; font-size: 0.85rem; margin-top: 4px;';
            info.textContent = `Extracted ${rows.length - 1} text entries from ${extractFile.name}`;
            extractDownload.appendChild(info);
        } catch (err) {
            extractError.textContent = 'Error: ' + err.message;
        } finally {
            extractBtn.disabled = false;
            extractBtn.textContent = 'Extract text';
        }
    }

    // ── Tab 2: XLSX → PPTX injection ────────────────────────────────

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

    function populateLanguageSelect(headerRow) {
        langSelect.innerHTML = '';
        for (let i = 1; i < headerRow.length; i++) {
            const code = String(headerRow[i] || '').trim();
            if (!code) continue;
            const opt = document.createElement('option');
            opt.value = i;
            opt.textContent = code;
            langSelect.appendChild(opt);
        }
        langSelect.disabled = langSelect.options.length === 0;
    }

    function buildTranslationMap(data, colIndex) {
        const map = new Map();
        for (let r = 1; r < data.length; r++) {
            const key = String(data[r][0] || '').trim();
            const value = data[r][colIndex];
            if (key && value !== undefined && value !== null && String(value).trim()) {
                map.set(key, String(value));
            }
        }
        return map;
    }

    async function injectTranslations(pptxBuffer, translationMap) {
        const zip = await JSZip.loadAsync(pptxBuffer);
        const slides = getSlideEntries(zip);
        let replacedCount = 0;

        for (const slide of slides) {
            const xml = await zip.file(slide.path).async('string');
            const doc = new DOMParser().parseFromString(xml, 'application/xml');
            const shapes = doc.getElementsByTagName('p:sp');
            let modified = false;

            for (let si = 0; si < shapes.length; si++) {
                const paragraphs = shapes[si].getElementsByTagName('a:p');
                for (let pi = 0; pi < paragraphs.length; pi++) {
                    const key = `s${slide.num}_sh${si}_p${pi}`;
                    const newText = translationMap.get(key);
                    if (newText === undefined) continue;

                    const runs = paragraphs[pi].getElementsByTagName('a:r');
                    if (runs.length === 0) continue;

                    // Put full text into first run, empty the rest
                    const firstRunT = runs[0].getElementsByTagName('a:t')[0];
                    if (firstRunT) {
                        firstRunT.textContent = newText;
                        // Preserve space attribute for leading/trailing spaces
                        firstRunT.setAttribute('xml:space', 'preserve');
                    }

                    for (let ri = 1; ri < runs.length; ri++) {
                        const t = runs[ri].getElementsByTagName('a:t')[0];
                        if (t) t.textContent = '';
                    }

                    modified = true;
                    replacedCount++;
                }
            }

            if (modified) {
                const serializer = new XMLSerializer();
                const newXml = serializer.serializeToString(doc);
                zip.file(slide.path, newXml);
            }
        }

        const blob = await zip.generateAsync({ type: 'blob', mimeType: 'application/vnd.openxmlformats-officedocument.presentationml.presentation' });
        return { blob, replacedCount };
    }

    async function handleInject() {
        injectError.textContent = '';
        injectDownload.innerHTML = '';

        if (!translationData || !injectPptxBuffer) return;

        const colIndex = parseInt(langSelect.value, 10);
        if (isNaN(colIndex)) {
            injectError.textContent = 'Please select a language.';
            return;
        }

        try {
            injectBtn.disabled = true;
            injectBtn.textContent = 'Converting…';

            const translationMap = buildTranslationMap(translationData, colIndex);
            if (translationMap.size === 0) {
                injectError.textContent = 'No translations found for the selected language.';
                return;
            }

            const { blob, replacedCount } = await injectTranslations(injectPptxBuffer.slice(0), translationMap);

            const langCode = String(translationData[0][colIndex] || 'translated').trim();
            const baseName = injectPptxName.replace(/\.pptx$/i, '');
            const filename = `${baseName}_${langCode}.pptx`;
            const url = URL.createObjectURL(blob);
            const a = document.createElement('a');
            a.href = url;
            a.download = filename;
            a.textContent = `Download ${filename}`;
            a.style.color = '#4da6ff';
            injectDownload.appendChild(a);

            const info = document.createElement('div');
            info.style.cssText = 'color: #6bbd6b; font-size: 0.85rem; margin-top: 4px;';
            info.textContent = `Replaced ${replacedCount} of ${translationMap.size} entries`;
            injectDownload.appendChild(info);
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

    // ── Dropzone setup ───────────────────────────────────────────────

    function setupSingleDropzone(dropzone, input, filenameEl, accept, onFile) {
        dropzone.addEventListener('click', (e) => {
            if (e.target.closest('.dropzone-filename')) return;
            input.click();
        });

        dropzone.addEventListener('dragenter', (e) => { e.preventDefault(); dropzone.classList.add('dragover'); });
        dropzone.addEventListener('dragover', (e) => { e.preventDefault(); dropzone.classList.add('dragover'); });
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

    // Tab 1 dropzone
    setupSingleDropzone(pptxExtractDropzone, pptxExtractInput, pptxExtractFilename, '.pptx', (file) => {
        extractFile = file;
        showFilename(pptxExtractFilename, file.name);
        extractBtn.disabled = false;
        extractDownload.innerHTML = '';
        extractError.textContent = '';
    });

    extractBtn.addEventListener('click', handleExtract);

    // Tab 2 dropzones
    setupSingleDropzone(xlsxInjectDropzone, xlsxInjectInput, xlsxInjectFilename, '.xlsx', async (file) => {
        injectXlsxFile = file;
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
