(() => {
    const XLSX_MIME = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
    const MISSING_TRANSLATION_MARK = '-';

    const mergeNewXlsxDropzone = document.getElementById('mergeNewXlsxDropzone');
    const mergeNewXlsxInput = document.getElementById('mergeNewXlsxInput');
    const mergeNewXlsxFilename = document.getElementById('mergeNewXlsxFilename');
    const mergeOldXlsxDropzone = document.getElementById('mergeOldXlsxDropzone');
    const mergeOldXlsxInput = document.getElementById('mergeOldXlsxInput');
    const mergeOldXlsxFilename = document.getElementById('mergeOldXlsxFilename');
    const mergeXlsxBtn = document.getElementById('mergeXlsxBtn');
    const mergeXlsxResult = document.getElementById('mergeXlsxResult');
    const mergeXlsxError = document.getElementById('mergeXlsxError');

    let newXlsxFile = null;
    let oldXlsxFile = null;

    function normalizeHeader(value) {
        return String(value || '').trim().toLowerCase();
    }

    function normalizeKey(value) {
        return String(value || '').trim();
    }

    function cellText(value) {
        return value === undefined || value === null ? '' : String(value);
    }

    function readXlsxFile(file) {
        return new Promise((resolve, reject) => {
            const reader = new FileReader();
            reader.onload = (event) => {
                try {
                    const data = new Uint8Array(event.target.result);
                    const workbook = XLSX.read(data, { type: 'array' });
                    const sheet = workbook.Sheets[workbook.SheetNames[0]];
                    resolve(XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '' }));
                } catch (err) {
                    reject(err);
                }
            };
            reader.onerror = reject;
            reader.readAsArrayBuffer(file);
        });
    }

    function createWorkbookBlob(rows, sheetName = 'Translations') {
        const worksheet = XLSX.utils.aoa_to_sheet(rows);
        const workbook = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(workbook, worksheet, sheetName);
        const data = XLSX.write(workbook, { bookType: 'xlsx', type: 'array' });
        return new Blob([data], { type: XLSX_MIME });
    }

    function triggerBlobDownload(blob, filename) {
        const url = URL.createObjectURL(blob);
        const link = document.createElement('a');
        link.href = url;
        link.download = filename;
        link.style.display = 'none';
        document.body.appendChild(link);
        link.click();
        setTimeout(() => {
            document.body.removeChild(link);
            URL.revokeObjectURL(url);
        }, 200);
    }

    function createDownloadLink(blob, filename) {
        const link = document.createElement('a');
        link.href = '#';
        link.textContent = `Download ${filename}`;
        link.addEventListener('click', (event) => {
            event.preventDefault();
            triggerBlobDownload(blob, filename);
        });
        return link;
    }

    function showFilename(element, name) {
        if (!element) return;
        element.textContent = name;
        element.style.display = name ? 'block' : 'none';
    }

    function updateMergeButton() {
        if (!mergeXlsxBtn) return;
        mergeXlsxBtn.disabled = !(newXlsxFile && oldXlsxFile);
    }

    function setupSingleDropzone(dropzone, input, filenameEl, accept, onFile) {
        if (!dropzone || !input || !filenameEl) return;

        dropzone.addEventListener('click', (event) => {
            if (event.target.closest('.dropzone-filename')) return;
            input.click();
        });

        dropzone.addEventListener('dragenter', (event) => {
            event.preventDefault();
            dropzone.classList.add('dragover');
        });
        dropzone.addEventListener('dragover', (event) => {
            event.preventDefault();
            dropzone.classList.add('dragover');
        });
        dropzone.addEventListener('dragleave', (event) => {
            if (!dropzone.contains(event.relatedTarget)) dropzone.classList.remove('dragover');
        });
        dropzone.addEventListener('drop', (event) => {
            event.preventDefault();
            dropzone.classList.remove('dragover');
            const files = Array.from(event.dataTransfer.files).filter(file =>
                file.name.toLowerCase().endsWith(accept)
            );
            if (files.length > 0) onFile(files[0]);
        });

        input.addEventListener('change', () => {
            if (input.files.length > 0) onFile(input.files[0]);
            input.value = '';
        });
    }

    function buildHeaderIndex(headerRow, label) {
        const index = new Map();
        headerRow.forEach((value, columnIndex) => {
            const key = normalizeHeader(value);
            if (!key || index.has(key)) return;
            index.set(key, columnIndex);
        });

        if (!index.has('key')) throw new Error(`${label} XLSX is missing required column "key".`);
        if (!index.has('en')) throw new Error(`${label} XLSX is missing required column "EN".`);
        return index;
    }

    function getOldTranslationColumns(oldHeader, newHeaderIndex) {
        return oldHeader
            .map((name, index) => ({ name: String(name || '').trim(), index }))
            .filter(column => column.name && normalizeHeader(column.name) !== 'key' && normalizeHeader(column.name) !== 'en')
            .map(column => ({
                ...column,
                outputIndex: newHeaderIndex.get(normalizeHeader(column.name))
            }));
    }

    function buildOldLookup(oldRows, oldHeaderIndex) {
        const lookup = new Map();
        let duplicateKeys = 0;

        for (let rowIndex = 1; rowIndex < oldRows.length; rowIndex++) {
            const row = oldRows[rowIndex] || [];
            const key = normalizeKey(row[oldHeaderIndex.get('key')]);
            if (!key) continue;
            if (lookup.has(key)) {
                duplicateKeys++;
                continue;
            }
            lookup.set(key, row);
        }

        return { lookup, duplicateKeys };
    }

    function buildOutputHeader(newHeader, oldTranslationColumns) {
        const outputHeader = newHeader.slice();
        const outputHeaderIndex = new Map();

        outputHeader.forEach((name, index) => {
            const key = normalizeHeader(name);
            if (key && !outputHeaderIndex.has(key)) outputHeaderIndex.set(key, index);
        });

        for (const column of oldTranslationColumns) {
            const key = normalizeHeader(column.name);
            if (outputHeaderIndex.has(key)) {
                column.outputIndex = outputHeaderIndex.get(key);
                continue;
            }

            column.outputIndex = outputHeader.length;
            outputHeaderIndex.set(key, column.outputIndex);
            outputHeader.push(column.name);
        }

        return outputHeader;
    }

    function mergeXlsxRows(newRows, oldRows) {
        if (!newRows.length) throw new Error('New XLSX is empty.');
        if (!oldRows.length) throw new Error('Old XLSX is empty.');

        const newHeader = (newRows[0] || []).map(value => String(value || '').trim());
        const oldHeader = (oldRows[0] || []).map(value => String(value || '').trim());
        const newHeaderIndex = buildHeaderIndex(newHeader, 'New');
        const oldHeaderIndex = buildHeaderIndex(oldHeader, 'Old');
        const oldTranslationColumns = getOldTranslationColumns(oldHeader, newHeaderIndex);

        if (!oldTranslationColumns.length) {
            throw new Error('Old XLSX has no translation columns to merge.');
        }

        const outputHeader = buildOutputHeader(newHeader, oldTranslationColumns);
        const { lookup: oldLookup, duplicateKeys } = buildOldLookup(oldRows, oldHeaderIndex);
        const outputRows = [outputHeader];
        let matchedRows = 0;
        let missingKeys = 0;
        let changedEnRows = 0;
        let copiedTranslationCells = 0;

        for (let rowIndex = 1; rowIndex < newRows.length; rowIndex++) {
            const newRow = newRows[rowIndex] || [];
            const outputRow = outputHeader.map((_, columnIndex) => newRow[columnIndex] ?? '');
            const key = normalizeKey(newRow[newHeaderIndex.get('key')]);
            const oldRow = key ? oldLookup.get(key) : null;
            const enMatches = oldRow && cellText(newRow[newHeaderIndex.get('en')]) === cellText(oldRow[oldHeaderIndex.get('en')]);

            if (!oldRow) {
                missingKeys++;
            } else if (!enMatches) {
                changedEnRows++;
            } else {
                matchedRows++;
            }

            for (const column of oldTranslationColumns) {
                const value = enMatches ? oldRow[column.index] : MISSING_TRANSLATION_MARK;
                outputRow[column.outputIndex] = value;
                if (enMatches && value !== undefined && value !== null && String(value) !== '') {
                    copiedTranslationCells++;
                }
            }

            outputRows.push(outputRow);
        }

        return {
            rows: outputRows,
            summary: {
                processedRows: Math.max(0, newRows.length - 1),
                matchedRows,
                missingKeys,
                changedEnRows,
                duplicateKeys,
                copiedTranslationCells,
                translationColumnCount: oldTranslationColumns.length
            }
        };
    }

    function renderSummary(summary) {
        if (!mergeXlsxResult) return;

        const info = document.createElement('div');
        info.style.cssText = 'color: #6bbd6b; font-size: 0.85rem; margin-top: 6px;';
        info.innerHTML = [
            `Processed ${summary.processedRows} row(s).`,
            `Matched and copied translations for ${summary.matchedRows} row(s).`,
            `Copied ${summary.copiedTranslationCells} non-empty translation cell(s) across ${summary.translationColumnCount} old translation column(s).`,
            `Missing key in old XLSX: ${summary.missingKeys} row(s).`,
            `Changed EN text: ${summary.changedEnRows} row(s).`
        ].map(line => `<div>${line}</div>`).join('');
        mergeXlsxResult.appendChild(info);

        if (summary.duplicateKeys > 0) {
            const warning = document.createElement('div');
            warning.style.cssText = 'color: #e8b64c; font-size: 0.85rem; margin-top: 8px;';
            warning.textContent = `Ignored ${summary.duplicateKeys} duplicate key row(s) in old XLSX.`;
            mergeXlsxResult.appendChild(warning);
        }
    }

    async function handleMergeXlsx() {
        if (!newXlsxFile || !oldXlsxFile) return;

        mergeXlsxError.textContent = '';
        mergeXlsxResult.innerHTML = '';

        try {
            mergeXlsxBtn.disabled = true;
            mergeXlsxBtn.textContent = 'Merging...';

            const [newRows, oldRows] = await Promise.all([
                readXlsxFile(newXlsxFile),
                readXlsxFile(oldXlsxFile)
            ]);
            const { rows, summary } = mergeXlsxRows(newRows, oldRows);
            const baseName = newXlsxFile.name.replace(/\.xlsx$/i, '');
            const filename = `${baseName}_merged.xlsx`;
            const blob = createWorkbookBlob(rows);

            mergeXlsxResult.appendChild(createDownloadLink(blob, filename));
            renderSummary(summary);
        } catch (err) {
            mergeXlsxError.textContent = 'Error: ' + err.message;
        } finally {
            updateMergeButton();
            mergeXlsxBtn.textContent = 'Merge';
        }
    }

    setupSingleDropzone(mergeNewXlsxDropzone, mergeNewXlsxInput, mergeNewXlsxFilename, '.xlsx', (file) => {
        newXlsxFile = file;
        showFilename(mergeNewXlsxFilename, file.name);
        if (mergeXlsxResult) mergeXlsxResult.innerHTML = '';
        if (mergeXlsxError) mergeXlsxError.textContent = '';
        updateMergeButton();
    });

    setupSingleDropzone(mergeOldXlsxDropzone, mergeOldXlsxInput, mergeOldXlsxFilename, '.xlsx', (file) => {
        oldXlsxFile = file;
        showFilename(mergeOldXlsxFilename, file.name);
        if (mergeXlsxResult) mergeXlsxResult.innerHTML = '';
        if (mergeXlsxError) mergeXlsxError.textContent = '';
        updateMergeButton();
    });

    if (mergeXlsxBtn) mergeXlsxBtn.addEventListener('click', handleMergeXlsx);
})();
