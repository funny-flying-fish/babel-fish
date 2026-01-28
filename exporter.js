/**
 * Lokalise Export Module
 * Handles XLSX parsing, diff computation, and API operations for Lokalise sync.
 */

(() => {
    // === Constants ===
    // Use local proxy to avoid CORS issues
    const LOKALISE_API_BASE = '/api/lokalise';
    const RATE_LIMIT_MS = 170; // ~6 req/sec with margin
    const PAGE_LIMIT = 500;

    // === DOM Elements ===
    const apiToken = document.getElementById('apiToken');
    const loadProjectsBtn = document.getElementById('loadProjectsBtn');
    const projectSelect = document.getElementById('projectSelect');
    const xlsxInput = document.getElementById('xlsxInput');
    const opCreate = document.getElementById('opCreate');
    const opUpdate = document.getElementById('opUpdate');
    const opDelete = document.getElementById('opDelete');
    const previewBtn = document.getElementById('previewBtn');
    const runExportBtn = document.getElementById('runExportBtn');
    const statusLine = document.getElementById('statusLine');
    const previewSection = document.getElementById('previewSection');
    const previewContent = document.getElementById('previewContent');
    const logTitle = document.getElementById('logTitle');
    const logCount = document.getElementById('logCount');
    const logEmpty = document.getElementById('logEmpty');
    const logList = document.getElementById('logList');

    // === State ===
    let xlsxFile = null;
    let fileKeys = new Map(); // key_name -> source_text
    let lokaliseKeys = new Map(); // key_name -> { key_id, translation_id, language_iso, source_text }
    let diffOps = { create: [], update: [], delete: [] };
    let logEntries = [];
    let isRunning = false;

    // === Rate Limiter ===
    let lastRequestTime = 0;

    async function rateLimitedFetch(url, options) {
        const now = Date.now();
        const timeSinceLastRequest = now - lastRequestTime;
        if (timeSinceLastRequest < RATE_LIMIT_MS) {
            await new Promise(resolve => setTimeout(resolve, RATE_LIMIT_MS - timeSinceLastRequest));
        }
        lastRequestTime = Date.now();
        return fetch(url, options);
    }

    // === Logging ===
    function resetLog() {
        logEntries = [];
        renderLog();
    }

    function addLog(type, message, details = null) {
        logEntries.push({ type, message, details, timestamp: new Date() });
        renderLog();
    }

    function escapeHtml(value) {
        return String(value)
            .replace(/&/g, '&amp;')
            .replace(/</g, '&lt;')
            .replace(/>/g, '&gt;')
            .replace(/"/g, '&quot;')
            .replace(/'/g, '&#39;');
    }

    function renderLog() {
        logList.innerHTML = '';
        if (logEntries.length === 0) {
            logEmpty.style.display = 'block';
            logCount.textContent = '0 operations';
            return;
        }
        logEmpty.style.display = 'none';
        logCount.textContent = `${logEntries.length} operations`;

        logEntries.forEach((entry) => {
            const item = document.createElement('li');
            let typeClass = '';
            if (entry.type === 'error') typeClass = 'log-error';
            else if (entry.type === 'success') typeClass = 'log-success';
            else if (entry.type === 'create') typeClass = 'log-after';
            else if (entry.type === 'update') typeClass = 'log-before';
            else if (entry.type === 'delete') typeClass = 'log-error';

            let html = `<span class="${typeClass}">[${entry.type.toUpperCase()}]</span> ${escapeHtml(entry.message)}`;
            if (entry.details) {
                html += ` <span class="log-rule">(${escapeHtml(entry.details)})</span>`;
            }
            item.innerHTML = html;
            logList.appendChild(item);
        });

        // Scroll to bottom
        logList.scrollTop = logList.scrollHeight;
    }

    // === Status ===
    function setStatus(message, type = '') {
        statusLine.textContent = message;
        statusLine.className = 'status-line' + (type ? ' ' + type : '');
    }

    // === API Functions ===
    async function apiRequest(endpoint, options = {}) {
        const token = apiToken.value.trim();
        if (!token) {
            throw new Error('API token is required');
        }

        const response = await rateLimitedFetch(`${LOKALISE_API_BASE}${endpoint}`, {
            ...options,
            headers: {
                'X-Api-Token': token,
                'Content-Type': 'application/json',
                ...options.headers
            }
        });

        if (!response.ok) {
            const errorData = await response.json().catch(() => ({}));
            throw new Error(errorData.error?.message || `API error: ${response.status}`);
        }

        return response.json();
    }

    async function loadProjects() {
        setStatus('Loading projects...');
        loadProjectsBtn.disabled = true;

        try {
            const projects = [];
            let page = 1;
            let hasMore = true;

            while (hasMore) {
                const data = await apiRequest(`/projects?page=${page}&limit=${PAGE_LIMIT}`);
                projects.push(...data.projects);
                hasMore = data.projects.length === PAGE_LIMIT;
                page++;
            }

            projectSelect.innerHTML = '<option value="">? Select project ?</option>';
            projects.forEach(project => {
                const option = document.createElement('option');
                option.value = project.project_id;
                option.textContent = project.name;
                projectSelect.appendChild(option);
            });

            projectSelect.disabled = false;
            setStatus(`Loaded ${projects.length} projects`, 'success');
            addLog('info', `Loaded ${projects.length} projects from Lokalise`);
        } catch (err) {
            setStatus(`Error: ${err.message}`, 'error');
            addLog('error', `Failed to load projects: ${err.message}`);
        } finally {
            loadProjectsBtn.disabled = false;
        }
    }

    async function loadLokaliseKeys(projectId) {
        setStatus('Loading keys from Lokalise...');
        lokaliseKeys.clear();

        try {
            let page = 1;
            let hasMore = true;
            let totalLoaded = 0;

            while (hasMore) {
                const data = await apiRequest(
                    `/projects/${projectId}/keys?page=${page}&limit=${PAGE_LIMIT}&include_translations=1`
                );

                data.keys.forEach(key => {
                    const enTranslation = key.translations?.find(t => 
                        t.language_iso === 'en' || t.language_iso === 'en_US' || t.language_iso === 'en_GB'
                    );
                    const keyName = typeof key.key_name === 'object' ? key.key_name.web : key.key_name;
                    lokaliseKeys.set(keyName, {
                        key_id: key.key_id,
                        translation_id: enTranslation?.translation_id || null,
                        language_iso: enTranslation?.language_iso || 'en',
                        source_text: enTranslation?.translation || ''
                    });
                });

                totalLoaded += data.keys.length;
                hasMore = data.keys.length === PAGE_LIMIT;
                page++;

                setStatus(`Loading keys... ${totalLoaded} loaded`);
            }

            addLog('info', `Loaded ${lokaliseKeys.size} keys from Lokalise`);
            return true;
        } catch (err) {
            setStatus(`Error loading keys: ${err.message}`, 'error');
            addLog('error', `Failed to load keys: ${err.message}`);
            return false;
        }
    }

    // === XLSX Parsing ===
    function parseXLSX(file) {
        return new Promise((resolve, reject) => {
            const reader = new FileReader();
            reader.onload = (e) => {
                try {
                    const data = new Uint8Array(e.target.result);
                    const workbook = XLSX.read(data, { type: 'array' });
                    const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
                    const jsonData = XLSX.utils.sheet_to_json(firstSheet, { header: 1 });
                    resolve(jsonData);
                } catch (err) {
                    reject(err);
                }
            };
            reader.onerror = reject;
            reader.readAsArrayBuffer(file);
        });
    }

    async function parseFileKeys() {
        if (!xlsxFile) {
            throw new Error('No file selected');
        }

        const data = await parseXLSX(xlsxFile);
        fileKeys.clear();

        // Skip header row, parse key | source
        for (let i = 1; i < data.length; i++) {
            const row = data[i];
            if (!row || !row[0]) continue;

            const keyName = String(row[0]).trim();
            const sourceText = row[1] !== undefined ? String(row[1]).trim() : '';

            if (keyName) {
                fileKeys.set(keyName, sourceText);
            }
        }

        addLog('info', `Parsed ${fileKeys.size} keys from XLSX file`);
        return fileKeys.size;
    }

    // === Diff Computation ===
    function computeDiff() {
        diffOps = { create: [], update: [], delete: [] };

        // Find keys to create and update
        for (const [keyName, sourceText] of fileKeys) {
            if (lokaliseKeys.has(keyName)) {
                const existing = lokaliseKeys.get(keyName);
                if (existing.source_text !== sourceText) {
                    diffOps.update.push({
                        key_name: keyName,
                        key_id: existing.key_id,
                        translation_id: existing.translation_id,
                        language_iso: existing.language_iso,
                        old_source: existing.source_text,
                        new_source: sourceText
                    });
                }
            } else {
                diffOps.create.push({
                    key_name: keyName,
                    source_text: sourceText
                });
            }
        }

        // Find keys to delete
        for (const [keyName, data] of lokaliseKeys) {
            if (!fileKeys.has(keyName)) {
                diffOps.delete.push({
                    key_name: keyName,
                    key_id: data.key_id,
                    source_text: data.source_text
                });
            }
        }

        return diffOps;
    }

    // === Preview Rendering ===
    function renderPreview() {
        const hasCreate = opCreate.checked && diffOps.create.length > 0;
        const hasUpdate = opUpdate.checked && diffOps.update.length > 0;
        const hasDelete = opDelete.checked && diffOps.delete.length > 0;

        if (!hasCreate && !hasUpdate && !hasDelete) {
            previewContent.innerHTML = '<p class="preview-empty">No operations to perform based on current settings.</p>';
            previewSection.style.display = 'block';
            return;
        }

        let html = '';

        if (hasCreate) {
            html += `<div class="preview-group">
                <div class="preview-group-title create">Create (${diffOps.create.length})</div>
                <ul class="preview-list">`;
            diffOps.create.slice(0, 50).forEach(op => {
                html += `<li><strong>${escapeHtml(op.key_name)}</strong>: "${escapeHtml(op.source_text.substring(0, 100))}${op.source_text.length > 100 ? '...' : ''}"</li>`;
            });
            if (diffOps.create.length > 50) {
                html += `<li>... and ${diffOps.create.length - 50} more</li>`;
            }
            html += '</ul></div>';
        }

        if (hasUpdate) {
            html += `<div class="preview-group">
                <div class="preview-group-title update">Update Source (${diffOps.update.length})</div>
                <ul class="preview-list">`;
            diffOps.update.slice(0, 50).forEach(op => {
                const oldText = op.old_source.length > 60 ? op.old_source.substring(0, 60) + '...' : op.old_source;
                const newText = op.new_source.length > 60 ? op.new_source.substring(0, 60) + '...' : op.new_source;
                html += `<li><strong>${escapeHtml(op.key_name)}</strong>: <span class="log-before">"${escapeHtml(oldText)}"</span> → <span class="log-after">"${escapeHtml(newText)}"</span></li>`;
            });
            if (diffOps.update.length > 50) {
                html += `<li>... and ${diffOps.update.length - 50} more</li>`;
            }
            html += '</ul></div>';
        }

        if (hasDelete) {
            html += `<div class="preview-group">
                <div class="preview-group-title delete">Delete (${diffOps.delete.length})</div>
                <ul class="preview-list">`;
            diffOps.delete.slice(0, 50).forEach(op => {
                html += `<li><strong>${escapeHtml(op.key_name)}</strong></li>`;
            });
            if (diffOps.delete.length > 50) {
                html += `<li>... and ${diffOps.delete.length - 50} more</li>`;
            }
            html += '</ul></div>';
        }

        previewContent.innerHTML = html;
        previewSection.style.display = 'block';
    }

    // === API Operations ===
    async function createKeys(projectId, keysToCreate) {
        if (keysToCreate.length === 0) return;

        // Lokalise allows up to 500 keys per request
        const batchSize = 500;
        let created = 0;

        for (let i = 0; i < keysToCreate.length; i += batchSize) {
            const batch = keysToCreate.slice(i, i + batchSize);
            const payload = {
                keys: batch.map(op => ({
                    key_name: op.key_name,
                    platforms: ['web'],
                    translations: [{
                        language_iso: 'en',
                        translation: op.source_text
                    }]
                }))
            };

            try {
                await apiRequest(`/projects/${projectId}/keys`, {
                    method: 'POST',
                    body: JSON.stringify(payload)
                });

                created += batch.length;
                batch.forEach(op => {
                    addLog('create', `Created key: ${op.key_name}`);
                });
                setStatus(`Creating keys... ${created}/${keysToCreate.length}`);
            } catch (err) {
                addLog('error', `Failed to create batch: ${err.message}`);
                throw err;
            }
        }

        return created;
    }

    async function updateKeys(projectId, keysToUpdate) {
        if (keysToUpdate.length === 0) return;

        let updated = 0;

        // Update translations via Translations API (requires translation_id)
        for (const op of keysToUpdate) {
            try {
                if (!op.translation_id) {
                    addLog('error', `No translation_id for ${op.key_name}, skipping`);
                    continue;
                }

                // Use Translations API directly
                await apiRequest(`/projects/${projectId}/translations/${op.translation_id}`, {
                    method: 'PUT',
                    body: JSON.stringify({
                        translation: op.new_source
                    })
                });

                updated++;
                addLog('update', `Updated source for: ${op.key_name}`);
                setStatus(`Updating keys... ${updated}/${keysToUpdate.length}`);
            } catch (err) {
                addLog('error', `Failed to update ${op.key_name}: ${err.message}`);
            }
        }

        return updated;
    }

    async function deleteKeys(projectId, keysToDelete) {
        if (keysToDelete.length === 0) return;

        // Lokalise allows batch delete
        const batchSize = 500;
        let deleted = 0;

        for (let i = 0; i < keysToDelete.length; i += batchSize) {
            const batch = keysToDelete.slice(i, i + batchSize);
            const keyIds = batch.map(op => op.key_id);

            try {
                await apiRequest(`/projects/${projectId}/keys`, {
                    method: 'DELETE',
                    body: JSON.stringify({ keys: keyIds })
                });

                deleted += batch.length;
                batch.forEach(op => {
                    addLog('delete', `Deleted key: ${op.key_name}`);
                });
                setStatus(`Deleting keys... ${deleted}/${keysToDelete.length}`);
            } catch (err) {
                addLog('error', `Failed to delete batch: ${err.message}`);
                throw err;
            }
        }

        return deleted;
    }

    async function verifyKeys(projectId, keyNames) {
        // Fetch current state of specified keys from Lokalise
        try {
            setStatus('Verifying changes...');
            
            // Load all keys (we need to find the ones we just modified)
            const data = await apiRequest(
                `/projects/${projectId}/keys?limit=${PAGE_LIMIT}&include_translations=1`
            );

            const keyNameSet = new Set(keyNames);
            
            data.keys.forEach(key => {
                const keyName = typeof key.key_name === 'object' ? key.key_name.web : key.key_name;
                
                if (keyNameSet.has(keyName)) {
                    const enTranslation = key.translations?.find(t => 
                        t.language_iso === 'en' || t.language_iso === 'en_US' || t.language_iso === 'en_GB'
                    );
                    const sourceText = enTranslation?.translation || '(empty)';
                    addLog('info', `${keyName}: "${sourceText}"`);
                }
            });
            
            setStatus('Verification complete', 'success');
        } catch (err) {
            addLog('error', `Verification failed: ${err.message}`);
        }
    }

    // === Event Handlers ===
    async function handlePreview() {
        if (isRunning) return;
        isRunning = true;
        previewBtn.disabled = true;
        runExportBtn.disabled = true;
        resetLog();

        try {
            const projectId = projectSelect.value;
            if (!projectId) {
                throw new Error('Please select a project');
            }

            await parseFileKeys();
            const loaded = await loadLokaliseKeys(projectId);
            if (!loaded) {
                throw new Error('Failed to load keys from Lokalise');
            }

            computeDiff();
            renderPreview();

            const totalOps = 
                (opCreate.checked ? diffOps.create.length : 0) +
                (opUpdate.checked ? diffOps.update.length : 0) +
                (opDelete.checked ? diffOps.delete.length : 0);

            setStatus(`Preview ready: ${totalOps} operations`, 'success');
            runExportBtn.disabled = totalOps === 0;
        } catch (err) {
            setStatus(`Error: ${err.message}`, 'error');
            addLog('error', err.message);
        } finally {
            isRunning = false;
            previewBtn.disabled = false;
        }
    }

    async function handleRunExport() {
        if (isRunning) return;

        const totalOps = 
            (opCreate.checked ? diffOps.create.length : 0) +
            (opUpdate.checked ? diffOps.update.length : 0) +
            (opDelete.checked ? diffOps.delete.length : 0);

        if (totalOps === 0) {
            setStatus('No operations to perform', 'error');
            return;
        }

        if (!confirm(`This will perform ${totalOps} operations on Lokalise. Continue?`)) {
            return;
        }

        isRunning = true;
        previewBtn.disabled = true;
        runExportBtn.disabled = true;
        resetLog();

        const projectId = projectSelect.value;
        let totalCreated = 0;
        let totalUpdated = 0;
        let totalDeleted = 0;

        try {
            addLog('info', 'Starting export...');

            if (opCreate.checked && diffOps.create.length > 0) {
                totalCreated = await createKeys(projectId, diffOps.create);
            }

            if (opUpdate.checked && diffOps.update.length > 0) {
                totalUpdated = await updateKeys(projectId, diffOps.update);
            }

            if (opDelete.checked && diffOps.delete.length > 0) {
                totalDeleted = await deleteKeys(projectId, diffOps.delete);
            }

            const summary = [];
            if (totalCreated > 0) summary.push(`${totalCreated} created`);
            if (totalUpdated > 0) summary.push(`${totalUpdated} updated`);
            if (totalDeleted > 0) summary.push(`${totalDeleted} deleted`);

            setStatus(`Export complete: ${summary.join(', ')}`, 'success');
            addLog('success', `Export complete: ${summary.join(', ')}`);

            // Verify: fetch updated keys from Lokalise and show current values
            const keysToVerify = [
                ...diffOps.create.map(op => op.key_name),
                ...diffOps.update.map(op => op.key_name)
            ];
            
            if (keysToVerify.length > 0) {
                addLog('info', '--- Verification: fetching current values from Lokalise ---');
                await verifyKeys(projectId, keysToVerify);
            }

            // Clear diff after successful export
            diffOps = { create: [], update: [], delete: [] };
            previewSection.style.display = 'none';
        } catch (err) {
            setStatus(`Export failed: ${err.message}`, 'error');
            addLog('error', `Export failed: ${err.message}`);
        } finally {
            isRunning = false;
            previewBtn.disabled = false;
            runExportBtn.disabled = true;
        }
    }

    function updateButtonStates() {
        const hasToken = apiToken.value.trim().length > 0;
        const hasProject = projectSelect.value !== '';
        const hasFile = xlsxFile !== null;

        loadProjectsBtn.disabled = !hasToken;
        previewBtn.disabled = !hasToken || !hasProject || !hasFile || isRunning;
    }

    // === Initialize ===
    loadProjectsBtn.addEventListener('click', loadProjects);
    previewBtn.addEventListener('click', handlePreview);
    runExportBtn.addEventListener('click', handleRunExport);

    xlsxInput.addEventListener('change', (e) => {
        xlsxFile = e.target.files[0] || null;
        updateButtonStates();
        // Reset preview when file changes
        previewSection.style.display = 'none';
        runExportBtn.disabled = true;
    });

    apiToken.addEventListener('input', updateButtonStates);
    projectSelect.addEventListener('change', () => {
        updateButtonStates();
        // Reset preview when project changes
        previewSection.style.display = 'none';
        runExportBtn.disabled = true;
    });

    // Initial state
    updateButtonStates();
    resetLog();
})();
