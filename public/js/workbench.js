document.addEventListener('DOMContentLoaded', () => {
    // --- STATE MANAGEMENT ---
    const state = {
        datasets: [],
        activeDatasetIndex: 0,
        // --- Robust Validator State ---
        dictionary: {
            dictionaryName: 'My Data Dictionary',
            sheetsData: {},
            ruleCategories: {}
        },
        lastValidationResult: null,
        ignoredErrors: new Set(),
        currentEditingCategory: null,
        pendingAssignments: []
    };

    const ROW_LIMIT = 50000;

    // --- DOM ELEMENT REFERENCES ---
    const fileUploadInput = document.getElementById('file-upload');
    const welcomeView = document.getElementById('welcome-view');
    const dataView = document.getElementById('data-view');
    const actionsContainer = document.getElementById('actions-container');
    const tableContainer = document.getElementById('data-table-container');
    const tableTitle = document.getElementById('table-title');
    const statusBar = document.getElementById('status-bar');
    const loaderOverlay = document.getElementById('loader-overlay');
    const downloadBtn = document.getElementById('download-btn');
    const loadedFilesList = document.getElementById('loaded-files-list');
    
    // Standard Config Modal
    const configModal = document.getElementById('config-modal');
    const modalTitle = document.getElementById('modal-title');
    const modalBody = document.getElementById('modal-body');
    let modalConfirmBtn = document.getElementById('modal-confirm-btn');
    const modalCancelBtn = document.getElementById('modal-cancel-btn');
    const modalCloseBtn = document.getElementById('modal-close-btn');

    // --- INITIALIZATION ---
    fileUploadInput.addEventListener('change', handleFileUpload);
    downloadBtn.addEventListener('click', handleDownload);
    [modalCancelBtn, modalCloseBtn].forEach(btn => btn.addEventListener('click', () => closeModal('config-modal')));
    
    initializeValidator(); // Initialize the new robust validator

    // --- FILE HANDLING ---
    async function handleFileUpload(event) {
        const files = Array.from(event.target.files);
        if (files.length === 0) return;
        showLoader(true);
        state.datasets = [];
        for (const file of files) {
            try {
                const data = await readFile(file);
                const workbook = XLSX.read(data, { type: 'array', cellDates: true });
                if (workbook.SheetNames.length === 1) {
                    const jsonData = XLSX.utils.sheet_to_json(workbook.Sheets[workbook.SheetNames[0]]);
                    state.datasets.push({ name: file.name, data: jsonData, headers: Object.keys(jsonData[0] || {}), workbook: workbook });
                } else {
                    workbook.SheetNames.forEach(sheetName => {
                        const jsonData = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName]);
                        state.datasets.push({ name: `${file.name} - ${sheetName}`, data: jsonData, headers: Object.keys(jsonData[0] || {}), workbook: workbook });
                    });
                }
            } catch (error) {
                console.error("Error processing file:", file.name, error);
                alert(`Could not process file: ${file.name}`);
            }
        }
        state.activeDatasetIndex = 0;
        updateUI();
        showLoader(false);
    }

    function readFile(file) {
        return new Promise((resolve, reject) => {
            const reader = new FileReader();
            reader.onload = (e) => resolve(new Uint8Array(e.target.result));
            reader.onerror = (e) => reject(new Error("File reading error"));
            reader.readAsArrayBuffer(file);
        });
    }

    // --- UI RENDERING & HELPERS ---
    function updateUI() {
        if (state.datasets.length === 0) {
            welcomeView.style.display = 'flex';
            dataView.classList.add('hidden');
            actionsContainer.style.display = 'none';
            loadedFilesList.innerHTML = '';
            document.getElementById('validation-results-view').classList.add('hidden');
        } else {
            welcomeView.style.display = 'none';
            actionsContainer.style.display = 'block';
            renderLoadedFilesList();
            renderActiveDataset();
        }
    }
    
    function showLoader(show) { loaderOverlay.style.display = show ? 'flex' : 'none'; }
    function getActiveDataset() { return state.datasets[state.activeDatasetIndex]; }

    function renderLoadedFilesList() {
        loadedFilesList.innerHTML = '';
        state.datasets.forEach((ds, index) => {
            const item = document.createElement('div');
            item.className = `p-2 rounded-md cursor-pointer text-sm transition-colors`;
            if (index === state.activeDatasetIndex) {
                item.classList.add('bg-indigo-700', 'text-white', 'font-bold');
            } else {
                item.classList.add('text-slate-300', 'hover:bg-slate-700');
            }
            item.textContent = ds.name;
            item.onclick = () => {
                state.activeDatasetIndex = index;
                updateUI();
            };
            loadedFilesList.appendChild(item);
        });
    }

    function renderActiveDataset() {
        const activeDataset = getActiveDataset();
        if (!activeDataset) return;
        tableTitle.textContent = activeDataset.name;
        renderDataTable(activeDataset.data, activeDataset.headers);
        statusBar.textContent = `Displaying ${activeDataset.data.length.toLocaleString()} rows and ${activeDataset.headers.length} columns. (Preview of first 200 rows)`;
        dataView.classList.remove('hidden');
        document.getElementById('validation-results-view').classList.add('hidden');
    }

    function renderDataTable(data, headers) {
        const table = document.createElement('table');
        const thead = table.createTHead();
        const headerRow = thead.insertRow();
        headers.forEach(h => {
            const th = document.createElement('th');
            th.textContent = h;
            headerRow.appendChild(th);
        });
        const tbody = table.createTBody();
        data.slice(0, 200).forEach(row => {
            const tr = tbody.insertRow();
            headers.forEach(header => {
                 const td = tr.insertCell();
                 td.textContent = row[header] ?? '';
            });
        });
        tableContainer.innerHTML = '';
        tableContainer.appendChild(table);
    }
    
    // --- STANDARD MODALS ---
    function showConfigModal(title, content, onConfirm) {
        modalTitle.textContent = title;
        modalBody.innerHTML = content;
        configModal.style.display = 'flex';
        const newConfirmBtn = modalConfirmBtn.cloneNode(true);
        modalConfirmBtn.parentNode.replaceChild(newConfirmBtn, modalConfirmBtn);
        newConfirmBtn.addEventListener('click', onConfirm);
        modalConfirmBtn = newConfirmBtn;
    }

    function generateColumnCheckboxes(headers) { return headers.map(h => `<label class="flex items-center p-2 rounded hover:bg-slate-100"><input type="checkbox" class="h-4 w-4 rounded mr-2" data-column-name="${h}"><span class="text-sm">${h}</span></label>`).join(''); }
    function generateColumnSelect(headers, id) { return `<select id="${id}" class="w-full p-2 border rounded mt-1">${headers.map(h => `<option value="${h}">${h}</option>`).join('')}</select>`; }
    
    // --- STANDARD ACTIONS ---
    function addNewDataset(name, data, headers) {
        state.datasets.push({ name, data, headers });
        state.activeDatasetIndex = state.datasets.length - 1;
        updateUI();
    }
    document.getElementById('action-trim-whitespace').addEventListener('click', () => {
        const activeDS = getActiveDataset();
        if (!activeDS) return;
        const content = `<p class="text-sm mb-4">Select the column to trim.</p><label for="trim-column" class="block text-sm font-semibold">Column:</label>${generateColumnSelect(activeDS.headers, 'config-column')}`;
        showConfigModal('Trim Whitespace', content, () => {
            const column = document.getElementById('config-column').value;
            showLoader(true);
            setTimeout(() => {
                getActiveDataset().data.forEach(row => { if (typeof row[column] === 'string') row[column] = row[column].trim(); });
                renderActiveDataset();
                showLoader(false); closeModal('config-modal');
            }, 50);
        });
    });
    document.getElementById('action-anonymize').addEventListener('click', () => {
        const activeDS = getActiveDataset();
        if (!activeDS) return;
        const types = [{ v: 'NONE', l: 'Do Not Anonymize' }, { v: 'FULL_NAME', l: 'Full Name' }, { v: 'FIRST_NAME', l: 'First Name' }, { v: 'LAST_NAME', l: 'Last Name' }, { v: 'EMAIL', l: 'Email' }, { v: 'PHONE', l: 'Phone' }];
        const content = activeDS.headers.map(h => `<div class="grid grid-cols-2 gap-4 items-center border-b pb-2 mb-2"><label class="font-semibold truncate" title="${h}">${h}</label><select data-header="${h}" class="column-mapper w-full p-2 border rounded">${types.map(t => `<option value="${t.v}">${t.l}</option>`).join('')}</select></div>`).join('');
        showConfigModal('Anonymize Personal Information', content, () => {
            const mappings = Array.from(document.querySelectorAll('.column-mapper')).filter(s => s.value !== 'NONE').map(s => ({ header: s.dataset.header, type: s.value }));
            if (mappings.length === 0) return alert('Please select at least one column to anonymize.');
            showLoader(true);
            setTimeout(() => {
                const fake = { FIRST: ['Alex', 'Jordan', 'Casey', 'Taylor'], LAST: ['Smith', 'Jones', 'Williams', 'Brown'], FULL: () => `${fake.FIRST[Math.floor(Math.random()*4)]} ${fake.LAST[Math.floor(Math.random()*4)]}`, EMAIL: () => `user${Math.floor(1000+Math.random()*9000)}@example.com`, PHONE: () => `(555) ${Math.floor(100+Math.random()*900)}-${Math.floor(1000+Math.random()*9000)}` };
                const anonData = activeDS.data.map(row => {
                    const newRow = { ...row };
                    mappings.forEach(m => {
                        if (newRow[m.header] !== undefined) newRow[m.header] = { FULL_NAME: fake.FULL(), FIRST_NAME: fake.FIRST[Math.floor(Math.random()*4)], LAST_NAME: fake.LAST[Math.floor(Math.random()*4)], EMAIL: fake.EMAIL(), PHONE: fake.PHONE() }[m.type];
                    });
                    return newRow;
                });
                addNewDataset(`Anonymized - ${activeDS.name}`, anonData, activeDS.headers);
                showLoader(false); closeModal('config-modal');
            }, 50);
        });
    });
    document.getElementById('action-extract-columns').addEventListener('click', () => {
        const activeDS = getActiveDataset();
        if (!activeDS) return;
        const content = `<p class="text-sm mb-4">Select columns to keep.</p><div class="space-y-2 max-h-96 overflow-y-auto">${generateColumnCheckboxes(activeDS.headers)}</div>`;
        showConfigModal('Extract Columns', content, () => {
            const selected = Array.from(document.querySelectorAll('#modal-body input:checked')).map(cb => cb.dataset.columnName);
            if (selected.length === 0) return alert('Please select at least one column.');
            showLoader(true);
            setTimeout(() => {
                const newData = activeDS.data.map(row => selected.reduce((obj, key) => (obj[key] = row[key], obj), {}));
                addNewDataset(`Extracted - ${activeDS.name}`, newData, selected);
                showLoader(false); closeModal('config-modal');
            }, 50);
        });
    });
    document.getElementById('action-stack-sheets').addEventListener('click', () => {
        if (state.datasets.length < 2) return alert("Please load at least two files to stack.");
        const content = `<p class="text-sm mb-4">This will combine all ${state.datasets.length} currently loaded datasets into a single master sheet. Columns will be matched by header name.</p>`;
        showConfigModal('Stack All Sheets', content, () => {
            showLoader(true);
            setTimeout(() => {
                const allData = state.datasets.flatMap(ds => ds.data);
                const allHeaders = [...new Set(state.datasets.flatMap(ds => ds.headers))];
                addNewDataset(`Stacked - ${state.datasets.length} files`, allData, allHeaders);
                showLoader(false); closeModal('config-modal');
            }, 50);
        });
    });
    document.getElementById('action-merge-files').addEventListener('click', () => {
        if (state.datasets.length < 2) return alert("Upload at least two files to merge.");
        const generateDatasetSelect = (id) => `<select id="${id}" class="w-full p-2 border rounded mt-1">${state.datasets.map((ds, i) => `<option value="${i}">${ds.name}</option>`).join('')}</select>`;
        const content = `<div class="grid grid-cols-2 gap-4"><div class="border-r pr-4"><label class="block text-sm font-semibold">Left Table (Primary)</label>${generateDatasetSelect('config-ds1')}<label class="block text-sm font-semibold mt-2">Key Column</label><select id="config-key1" class="w-full p-2 border rounded mt-1"></select></div><div><label class="block text-sm font-semibold">Right Table (to join)</label>${generateDatasetSelect('config-ds2')}<label class="block text-sm font-semibold mt-2">Key Column</label><select id="config-key2" class="w-full p-2 border rounded mt-1"></select></div></div>`;
        const populateKeys = () => { ['1', '2'].forEach(n => { const ds_idx = document.getElementById(`config-ds${n}`).value; document.getElementById(`config-key${n}`).innerHTML = state.datasets[ds_idx].headers.map(h => `<option value="${h}">${h}</option>`).join(''); }); };
        showConfigModal('Merge Files (Left Join)', content, () => {
            const ds1_idx = document.getElementById('config-ds1').value;
            const ds2_idx = document.getElementById('config-ds2').value;
            const key1 = document.getElementById('config-key1').value;
            const key2 = document.getElementById('config-key2').value;
            showLoader(true);
            setTimeout(() => {
                const ds1 = state.datasets[ds1_idx];
                const ds2 = state.datasets[ds2_idx];
                const map2 = new Map(ds2.data.map(row => [row[key2], row]));
                const mergedData = ds1.data.map(row1 => ({ ...row1, ...(map2.get(row1[key1]) || {}) }));
                const newHeaders = [...new Set([...ds1.headers, ...ds2.headers])];
                addNewDataset(`Merged - ${ds1.name} & ${ds2.name}`, mergedData, newHeaders);
                showLoader(false); closeModal('config-modal');
            }, 50);
        });
        populateKeys();
        document.getElementById('config-ds1').onchange = populateKeys;
        document.getElementById('config-ds2').onchange = populateKeys;
    });
    document.getElementById('action-find-duplicates').addEventListener('click', () => {
        const activeDS = getActiveDataset();
        if (!activeDS) return;
        const content = `<p class="text-sm mb-4">Select columns to check for duplicates.</p><div class="space-y-2 max-h-96 overflow-y-auto">${generateColumnCheckboxes(activeDS.headers)}</div>`;
        showConfigModal('Find Duplicates', content, () => {
            const selected = Array.from(document.querySelectorAll('#modal-body input:checked')).map(cb => cb.dataset.columnName);
            if (selected.length === 0) return alert('Select at least one column.');
            showLoader(true);
            setTimeout(() => {
                const seen = new Map();
                const duplicates = [];
                activeDS.data.forEach(row => {
                    const key = selected.map(col => row[col]).join('||');
                    if (seen.has(key)) {
                        if (seen.get(key).first) { duplicates.push(seen.get(key).row); seen.get(key).first = false; }
                        duplicates.push(row);
                    } else { seen.set(key, { row: row, first: true }); }
                });
                if(duplicates.length > 0){
                    addNewDataset(`Duplicates - ${activeDS.name}`, duplicates, activeDS.headers);
                } else {
                    alert("No duplicates found based on the selected columns.");
                }
                showLoader(false); closeModal('config-modal');
            }, 50);
        });
    });
    document.getElementById('action-compare-sheets').addEventListener('click', () => {
        if (state.datasets.length < 2) return alert("Upload at least two files to compare.");
        const generateDatasetSelect = (id) => `<select id="${id}" class="w-full p-2 border rounded mt-1">${state.datasets.map((ds, i) => `<option value="${i}">${ds.name}</option>`).join('')}</select>`;
        const content = `<div class="grid grid-cols-2 gap-4"><div class="border-r pr-4"><label class="block text-sm font-semibold">Original / Old File</label>${generateDatasetSelect('config-ds1')}<label class="block text-sm font-semibold mt-2">Unique ID Column</label><select id="config-key1" class="w-full p-2 border rounded mt-1"></select></div><div><label class="block text-sm font-semibold">New / Updated File</label>${generateDatasetSelect('config-ds2')}<label class="block text-sm font-semibold mt-2">Unique ID Column</label><select id="config-key2" class="w-full p-2 border rounded mt-1"></select></div></div>`;
        const populateKeys = () => { ['1', '2'].forEach(n => { const ds_idx = document.getElementById(`config-ds${n}`).value; document.getElementById(`config-key${n}`).innerHTML = state.datasets[ds_idx].headers.map(h => `<option value="${h}">${h}</option>`).join(''); }); };
        showConfigModal('Compare Sheets', content, () => {
            const ds1_idx = document.getElementById('config-ds1').value, ds2_idx = document.getElementById('config-ds2').value;
            const key1 = document.getElementById('config-key1').value, key2 = document.getElementById('config-key2').value;
            showLoader(true);
            setTimeout(() => {
                const ds1 = state.datasets[ds1_idx], ds2 = state.datasets[ds2_idx];
                const map1 = new Map(ds1.data.map(row => [row[key1], row]));
                const map2 = new Map(ds2.data.map(row => [row[key2], row]));
                const results = [];
                const allHeaders = [...new Set([...ds1.headers, ...ds2.headers])];
                map2.forEach((row2, key) => {
                    const row1 = map1.get(key);
                    if (!row1) { results.push({ Status: 'Added', ...row2 }); } 
                    else {
                        let isModified = false;
                        for (const h of allHeaders) { if (String(row1[h] ?? '') !== String(row2[h] ?? '')) isModified = true; }
                        if (isModified) results.push({ Status: 'Modified', ...row2 });
                    }
                    map1.delete(key);
                });
                map1.forEach(row1 => results.push({ Status: 'Deleted', ...row1 }));
                addNewDataset(`Comparison - ${ds1.name} vs ${ds2.name}`, results, ['Status', ...allHeaders]);
                showLoader(false); closeModal('config-modal');
            }, 50);
        });
        populateKeys();
        document.getElementById('config-ds1').onchange = populateKeys;
        document.getElementById('config-ds2').onchange = populateKeys;
    });
    document.getElementById('action-split-file').addEventListener('click', () => {
        const activeDS = getActiveDataset();
        if (!activeDS) return;
        const content = `<p class="text-sm mb-4">Split the active dataset into multiple CSV files.</p><label for="config-rows" class="block text-sm font-semibold">Rows Per File</label><input type="number" id="config-rows" value="100000" min="1" class="w-full p-2 border rounded mt-1">`;
        showConfigModal('Split File by Row Count', content, () => {
            const rowsPerFile = parseInt(document.getElementById('config-rows').value, 10);
            if (isNaN(rowsPerFile) || rowsPerFile < 1) return alert("Invalid number.");
            showLoader(true);
            setTimeout(() => {
                const zip = new JSZip();
                for (let i = 0, f = 1; i < activeDS.data.length; i += rowsPerFile, f++) {
                    const chunk = activeDS.data.slice(i, i + rowsPerFile);
                    const ws = XLSX.utils.json_to_sheet(chunk);
                    zip.file(`split_${f}.csv`, XLSX.utils.sheet_to_csv(ws));
                }
                zip.generateAsync({ type: 'blob' }).then(c => { saveAs(c, `Split_${activeDS.name}.zip`); showLoader(false); closeModal('config-modal'); });
            }, 50);
        });
    });
    
    // --- DOWNLOADING ---
    function handleDownload() {
        const activeDataset = getActiveDataset();
        if (!activeDataset) return;
        showLoader(true);
        setTimeout(() => {
            const ws = XLSX.utils.json_to_sheet(activeDataset.data);
            const wb = XLSX.utils.book_new();
            XLSX.utils.book_append_sheet(wb, ws, "Result");
            XLSX.writeFile(wb, `Processed_${activeDataset.name.replace(/[^a-z0-9]/gi, '_').toLowerCase()}.xlsx`);
            showLoader(false);
        }, 50);
    }

    // ===================================================================================
    // --- ROBUST VALIDATOR LOGIC ---
    // ===================================================================================
    function initializeValidator() {
        // Load dictionary from Local Storage instead of API
        const savedDict = localStorage.getItem('validator_dictionary_v2');
        if (savedDict) {
            state.dictionary = JSON.parse(savedDict);
        }

        // Hook up main UI buttons
        document.getElementById('validator-run-btn').addEventListener('click', analyzeFileWorkflow);
        document.getElementById('validator-manage-dict-btn').addEventListener('click', openDictionaryEditor);
        document.getElementById('validator-manage-cat-btn').addEventListener('click', openRuleCategoryEditor);
        
        // Hook up modal buttons
        document.getElementById('builder_saveAndCloseBtn').addEventListener('click', saveAndCloseBuilder);
        document.getElementById('categoryEditorCloseBtn').addEventListener('click', handleCategoryEditorClose);
        document.getElementById('restore-input').addEventListener('change', handleRestoreUpload);
    }
    
    function saveRobustDictionary() {
        try {
            localStorage.setItem('validator_dictionary_v2', JSON.stringify(state.dictionary));
        } catch (error) {
            console.error("Failed to save dictionary to local storage:", error);
            alert("Could not save dictionary. Your browser's storage might be full.");
        }
    }

    // --- Global Functions (made available by being in the top-level scope) ---
    window.closeModal = (id) => document.getElementById(id).classList.add('hidden');

    window.openDictionaryEditor = function() {
        const activeDS = getActiveDataset();
        const headers = activeDS ? activeDS.headers : [];
        document.getElementById('builder_dictionaryName').value = state.dictionary.dictionaryName;
        renderBuilderTable(headers);
        document.getElementById('dictionaryModal').classList.remove('hidden');
    }

    window.openRuleCategoryEditor = function() {
        state.currentEditingCategory = null;
        document.getElementById('ruleCategoryEditorContainer').classList.add('hidden');
        renderRuleCategoryList();
        document.getElementById('ruleCategoryModal').classList.remove('hidden');
    }

    window.downloadFullPdfDictionary = function() {
        const dictionary = state.dictionary;
        if (!dictionary || !dictionary.sheetsData || Object.keys(dictionary.sheetsData).length === 0) { alert("No dictionary stored to download."); return; }
        const { jsPDF } = window.jspdf; const doc = new jsPDF('landscape');
        const tableBody = []; const tableHead = ['Column Name', 'Rule Category', 'Description', 'Defined Rules'];
        const sortedKeys = Object.keys(dictionary.sheetsData).sort((a,b) => dictionary.sheetsData[a]['Column Name'].localeCompare(dictionary.sheetsData[b]['Column Name']));
        sortedKeys.forEach(headerKey => {
            const ruleData = dictionary.sheetsData[headerKey];
            let rulesToDisplay = [];
            if (ruleData.category && dictionary.ruleCategories[ruleData.category]) { rulesToDisplay = dictionary.ruleCategories[ruleData.category].rules; }
            else if (ruleData.validation_rules) { rulesToDisplay = ruleData.validation_rules; }
            const rulesText = (rulesToDisplay || []).map(r => `${r.type}${r.value ? `: ${r.value}` : ''}`).join('\n');
            tableBody.push([ ruleData['Column Name'], ruleData.category || 'N/A', ruleData.description || '', rulesText ]);
        });
        doc.setFontSize(18); doc.text("Full Data Dictionary Report", 14, 22);
        doc.setFontSize(11); doc.text(`Dictionary: ${dictionary.dictionaryName}`, 14, 30);
        doc.text(`Generated: ${new Date().toLocaleString()}`, 14, 36);
        doc.autoTable({ startY: 40, head: [tableHead], body: tableBody, theme: 'grid', headStyles: { fillColor: [41, 128, 185] }, styles: { fontSize: 8, cellPadding: 1.5 }, columnStyles: { 0: { cellWidth: 40 }, 1: { cellWidth: 40 }, 2: { cellWidth: 'auto' }, 3: { cellWidth: 65 } } });
        doc.save(`${dictionary.dictionaryName.replace(/[^a-z0-9]/gi,'_')}.pdf`);
    }

    window.handleRestoreUpload = function(event) {
        const file = event.target.files[0];
        if (!file) return;
        const reader = new FileReader();
        reader.onload = (e) => {
            try {
                const backupData = JSON.parse(e.target.result);
                const dictionaryToRestore = backupData.dictionary || backupData; // For flexibility
                if (!dictionaryToRestore || !dictionaryToRestore.dictionaryName || !dictionaryToRestore.sheetsData) {
                    throw new Error("Backup file is missing a valid dictionary structure.");
                }
                if (confirm(`This will overwrite your current dictionary with "${dictionaryToRestore.dictionaryName}" from the JSON file. Are you sure?`)) {
                    state.dictionary = dictionaryToRestore;
                    saveRobustDictionary();
                    alert("Restore successful! Your dictionary has been updated.");
                    event.target.value = '';
                }
            } catch(err) {
                alert(`Error restoring file: ${err.message}`);
                event.target.value = '';
            }
        };
        reader.readAsText(file);
    }
    
    function analyzeFileWorkflow() { 
        document.getElementById('validation-results-view').classList.add('hidden');
        dataView.classList.remove('hidden'); // Ensure data view is visible
        
        state.lastValidationResult = null; 
        state.ignoredErrors.clear();
        const activeDS = getActiveDataset();
        if (!activeDS) { alert('Please load a file first.'); return; }
        
        showLoader(true);
        
        setTimeout(() => { // Use timeout to allow loader to render
            const fileHeaders = activeDS.headers;
            const dictHeaders = getAllHeadersFromDictionary(state.dictionary);
            const missingHeaders = fileHeaders.filter(header => !dictHeaders.has(header.toUpperCase()));
            
            if (missingHeaders.length > 0) {
                // If new columns are found, render them into a new section INSIDE the validation results view
                displaySchemaMismatchUI(missingHeaders);
            } else {
                runDataValidation(activeDS, state.dictionary);
            }
            showLoader(false);
        }, 50);
    }
            
    function displaySchemaMismatchUI(missingHeaders) {
        if (!state.dictionary.sheetsData) state.dictionary.sheetsData = {};
        const existingColumns = Object.values(state.dictionary.sheetsData).map(col => col['Column Name']).filter(Boolean).sort((a, b) => a.localeCompare(b));
        const copyOptions = existingColumns.map(name => `<option value="${name}">${name}</option>`).join('');
        let schemaHTML = `<div class="bg-white p-6 md:p-8 rounded-xl shadow-lg"><div class="p-4 border-l-4 border-yellow-400 bg-yellow-50 rounded-r-lg"><h3 class="text-lg font-bold text-yellow-800">New Columns Detected</h3><p class="text-sm text-yellow-700 mt-2">New columns were found. Please define their rules below. Columns without a description will be excluded.</p></div><div class="mt-4 space-y-4">`;
        const categoryOptions = Object.keys(state.dictionary.ruleCategories || {}).sort().map(catName => `<option value="${catName}">${catName}</option>`).join('');
        missingHeaders.forEach((header, index) => { schemaHTML += `<div class="p-3 bg-slate-50 rounded-md" data-missing-header="${header}"><div class="grid grid-cols-1 md:grid-cols-2 gap-x-6 gap-y-4 items-start"><div><label class="block text-sm font-semibold text-slate-800">${header}</label><select class="category-select w-full mt-1 p-2 border rounded-md" onchange="handleCategoryChoice(this, ${index})"><option value="">-- Select a Category --</option><option value="__CREATE_NEW__" class="font-bold text-blue-600">** Create New **</option>${categoryOptions}</select></div><div><label class="block text-sm font-medium text-gray-700">Description</label><textarea oninput="autoAdjustTextarea(this)" class="description-input w-full mt-1 p-2 border rounded-md" placeholder="Enter description..."></textarea><select class="w-full mt-2 p-1.5 border rounded-md bg-gray-50 text-sm" onchange="copySettingsFromExisting(this)"><option value="">-- Or, copy settings from existing --</option>${copyOptions}</select></div></div><div id="inline-editor-${index}" class="inline-editor hidden mt-4 p-4 border-t border-slate-200"><div class="mb-2"><label class="block text-sm font-medium">New Category Name:</label><input type="text" class="new-category-name w-full p-2 border rounded-md" placeholder="Enter Unique Name..."></div><div class="mb-2"><label class="block text-sm font-medium">Defined Rules:</label><div class="inline-rules-display flex flex-wrap gap-2 p-2 min-h-[40px] border rounded-md bg-white"></div></div><div><label class="block text-sm font-medium">Add New Rule:</label><div class="flex gap-2 items-center"><select class="inline-rule-type p-1 border rounded-md bg-white flex-1"><option value="REQUIRED">REQUIRED</option><option value="NOT REQUIRED">NOT REQUIRED</option><option value="ALLOWED_VALUES">ALLOWED_VALUES</option><option value="REGEX">REGEX</option><option value="VALID_DATE">VALID_DATE</option><option value="NO_FUTURE_DATE">NO_FUTURE_DATE</option></select><input type="text" placeholder="Value..." class="inline-rule-value p-1 border rounded-md flex-1"><button type="button" class="py-1 px-3 bg-blue-500 text-white rounded-md text-sm" onclick="addRuleToInlineCategory(this)">+</button></div></div></div></div>`; });
        schemaHTML += `</div><div class="text-center mt-6"><button id="workflow_proceedBtn" class="px-6 py-2 bg-green-600 hover:bg-green-700 text-white font-bold rounded-md">Apply & Continue</button></div></div>`;
        document.getElementById('validation-results-view').innerHTML = schemaHTML;
        dataView.classList.add('hidden');
        document.getElementById('validation-results-view').classList.remove('hidden');
        document.getElementById('workflow_proceedBtn').addEventListener('click', handleMappingProceed);
    }
    window.copySettingsFromExisting = function(selectElement) { const selectedColumnName = selectElement.value; if (!selectedColumnName) return; const columnData = Object.values(state.dictionary.sheetsData).find(col => col['Column Name'] === selectedColumnName); if (columnData) { const container = selectElement.closest('[data-missing-header]'); const descriptionInput = container.querySelector('.description-input'); const categorySelect = container.querySelector('.category-select'); descriptionInput.value = columnData.description || ''; autoAdjustTextarea(descriptionInput); if (columnData.category && categorySelect.querySelector(`option[value="${columnData.category}"]`)) { categorySelect.value = columnData.category; const inlineEditor = container.querySelector('.inline-editor'); if (inlineEditor) { inlineEditor.classList.add('hidden'); } } } selectElement.value = ''; }
    window.handleCategoryChoice = function(selectElement, index) { document.getElementById(`inline-editor-${index}`).classList.toggle('hidden', selectElement.value !== '__CREATE_NEW__'); }
    window.addRuleToInlineCategory = function(button) { const editor = button.closest('.inline-editor'); const type = editor.querySelector('.inline-rule-type').value; const valueInput = editor.querySelector('.inline-rule-value'); const value = valueInput.value.trim(); if (['ALLOWED_VALUES', 'REGEX'].includes(type) && !value) { alert("Value is required for this rule type."); return; } const display = editor.querySelector('.inline-rules-display'); const pill = document.createElement('span'); pill.className = 'rule-pill'; pill.dataset.type = type; pill.dataset.value = value; pill.innerHTML = `${type}: ${value || 'N/A'} <button type="button" class="delete-btn ml-2" onclick="this.parentElement.remove()">x</button>`; display.appendChild(pill); valueInput.value = ''; }
    function handleMappingProceed() { const assignments = []; const newCategories = {}; const allCategoryNames = new Set(Object.keys(state.dictionary.ruleCategories)); let hasError = false; const skippedColumns = []; document.querySelectorAll('[data-missing-header]').forEach(div => { const header = div.dataset.missingHeader; const categorySelect = div.querySelector('.category-select'); const descriptionInput = div.querySelector('.description-input'); if (categorySelect.value) { const description = descriptionInput.value.trim(); if (!description) { skippedColumns.push(header); return; } const assignment = { header, description }; if (categorySelect.value === '__CREATE_NEW__') { const editor = div.querySelector('.inline-editor'); const newCatName = editor.querySelector('.new-category-name').value.trim(); if (!newCatName) { alert(`Please enter a name for the new category for column "${header}".`); hasError = true; return; } if (allCategoryNames.has(newCatName)) { alert(`The category name "${newCatName}" already exists. Please choose a unique name.`); hasError = true; return; } assignment.category = newCatName; const rules = []; editor.querySelectorAll('.inline-rules-display .rule-pill').forEach(pill => { rules.push({ type: pill.dataset.type, value: pill.dataset.value, message: `Validation failed for rule ${pill.dataset.type}` }); }); newCategories[newCatName] = { rules }; allCategoryNames.add(newCatName); } else { assignment.category = categorySelect.value; } assignments.push(assignment); } else { skippedColumns.push(header); } }); if (hasError) return; const continueProcessing = () => { if (Object.keys(newCategories).length > 0) { state.dictionary.ruleCategories = { ...state.dictionary.ruleCategories, ...newCategories }; } state.pendingAssignments = assignments; state.lastValidationResult = { ...(state.lastValidationResult || {}), skippedColumns: skippedColumns }; applyPendingAssignments(); saveRobustDictionary(); commitMappingsAndValidate(); }; if (skippedColumns.length > 0) { if (confirm(`Some columns have no description and will be excluded. Continue?`)) { continueProcessing(); } } else { continueProcessing(); } }
    function applyPendingAssignments() { state.pendingAssignments.forEach(assignment => { const { header, category, description } = assignment; const headerKey = header.toUpperCase(); state.dictionary.sheetsData[headerKey] = { "Column Name": header, category: category, description: description, validation_rules: [] }; }); state.pendingAssignments = []; }
    function commitMappingsAndValidate() { document.getElementById('validation-results-view').classList.add('hidden'); showLoader(true); setTimeout(() => { runDataValidation(getActiveDataset(), state.dictionary); showLoader(false); }, 50); }
    function renderBuilderTable(headers) { const tbody = document.querySelector('#rulesTable tbody'); tbody.innerHTML = ''; const dictData = state.dictionary.sheetsData; const categoryOptions = Object.keys(state.dictionary.ruleCategories).sort().map(catName => `<option value="${catName}">${catName}</option>`).join(''); const allHeaders = [...new Set(headers.concat(Object.values(dictData).map(d => d["Column Name"])))].filter(Boolean).sort((a,b) => a.localeCompare(b)); allHeaders.forEach(header => { const headerKey = header.toUpperCase(); const columnData = Object.values(dictData).find(d => d["Column Name"].toUpperCase() === headerKey); if(!columnData) return; const row = tbody.insertRow(); let rulesToDisplay = []; if (columnData.category && state.dictionary.ruleCategories[columnData.category]) { rulesToDisplay = state.dictionary.ruleCategories[columnData.category].rules; } else if (columnData.validation_rules) { rulesToDisplay = columnData.validation_rules; } let rulesHTML = '<div class="flex flex-wrap gap-1">'; if(rulesToDisplay) { rulesToDisplay.forEach((rule, index) => { const deleteAction = columnData.category ? '' : `<button class="delete-btn ml-2" onclick="deleteRule('${headerKey}', ${index})">x</button>`; rulesHTML += `<span class="rule-pill">${rule.type}: ${rule.value || 'N/A'} ${deleteAction}</span>`; }); } rulesHTML += '</div>'; const addRuleHTML = `<div class="flex gap-2 items-center"><select class="p-1 border rounded-md bg-white w-1/3"><option value="REQUIRED">REQUIRED</option><option value="NOT REQUIRED">NOT REQUIRED</option><option value="ALLOWED_VALUES">ALLOWED_VALUES</option><option value="REGEX">REGEX</option><option value="VALID_DATE">VALID_DATE</option><option value="NO_FUTURE_DATE">NO_FUTURE_DATE</option></select><input type="text" placeholder="Value..." class="p-1 border rounded-md w-1/3"><button class="py-1 px-3 bg-blue-500 text-white rounded-md text-sm" onclick="addRule(this, '${headerKey}')">+</button></div>`; const categoryHTML = `<td><select onchange="applyRuleCategory('${headerKey}', this.value)" class="p-1 border rounded-md w-full"><option value="">-- None --</option>${categoryOptions}</select></td>`; const descriptionHTML = `<td><textarea oninput="autoAdjustTextarea(this)" onchange="updateDescription('${headerKey}', this.value)" class="description-input p-1 border rounded-md w-full">${columnData.description || ''}</textarea></td>`; row.innerHTML = `<td class="font-semibold align-middle">${header}<button class="delete-btn ml-4" onclick="deleteColumn('${headerKey}')">Delete</button></td>${categoryHTML}${descriptionHTML}<td>${rulesHTML}</td><td>${addRuleHTML}</td>`; row.querySelector('select').value = columnData.category || ''; autoAdjustTextarea(row.querySelector('textarea')); }); }
    window.applyRuleCategory = function(headerKey, categoryName) { const columnData = Object.values(state.dictionary.sheetsData).find(d=>d["Column Name"].toUpperCase() === headerKey); if (!columnData) return; if (categoryName === "") { columnData.validation_rules = []; columnData.category = ""; } else { const categoryTemplate = state.dictionary.ruleCategories[categoryName]; if (categoryTemplate && confirm(`This will replace existing rules for "${columnData['Column Name']}" with the "${categoryName}" category. Are you sure?`)) { columnData.validation_rules = []; columnData.category = categoryName; } } renderBuilderTable(Object.values(state.dictionary.sheetsData).map(r => r["Column Name"])); }
    window.updateDescription = function(headerKey, newDescription) { const col = Object.values(state.dictionary.sheetsData).find(d=>d["Column Name"].toUpperCase() === headerKey); if (col) col.description = newDescription; }
    window.addRule = function(buttonElement, headerKey) { const col = Object.values(state.dictionary.sheetsData).find(d=>d["Column Name"].toUpperCase() === headerKey); if(!col) return; const addRuleContainer = buttonElement.closest('div'); const select = addRuleContainer.querySelector('select'); const input = addRuleContainer.querySelector('input[type="text"]'); const ruleType = select.value, ruleValue = input.value.trim(); if (['ALLOWED_VALUES', 'REGEX'].includes(ruleType) && !ruleValue) { alert('Rule Value is required.'); return; } if(!col.validation_rules) col.validation_rules = []; col.validation_rules.push({ type: ruleType, value: ruleValue, message: `Validation failed for rule ${ruleType}` }); col.category = ''; renderBuilderTable(Object.values(state.dictionary.sheetsData).map(r => r["Column Name"])); }
    window.deleteRule = function(headerKey, ruleIndex) { const col = Object.values(state.dictionary.sheetsData).find(d=>d["Column Name"].toUpperCase() === headerKey); if(col) { col.validation_rules.splice(ruleIndex, 1); col.category = ''; renderBuilderTable(Object.values(state.dictionary.sheetsData).map(r => r["Column Name"])); } }
    window.deleteColumn = function(headerKey) { const colName = Object.values(state.dictionary.sheetsData).find(d=>d["Column Name"].toUpperCase() === headerKey)["Column Name"]; if(confirm(`Delete the column "${colName}"?`)){ const keyToDelete = Object.keys(state.dictionary.sheetsData).find(k => k.toUpperCase() === headerKey); if(keyToDelete) delete state.dictionary.sheetsData[keyToDelete]; renderBuilderTable(Object.values(state.dictionary.sheetsData).map(r => r["Column Name"])); } }
    function saveAndCloseBuilder() { state.dictionary.dictionaryName = document.getElementById('builder_dictionaryName').value; saveRobustDictionary(); closeModal('dictionaryModal'); }
    function handleCategoryEditorClose() { saveRobustDictionary(); closeModal('ruleCategoryModal'); if (state.pendingAssignments.length > 0) { applyPendingAssignments(); saveRobustDictionary(); commitMappingsAndValidate(); } }
    function renderRuleCategoryList() { const listEl = document.getElementById('ruleCategoryList'); listEl.innerHTML = ''; const sortedCategories = Object.keys(state.dictionary.ruleCategories).sort(); sortedCategories.forEach(catName => { const item = document.createElement('div'); item.className = 'category-list-item'; item.textContent = catName; if (catName === state.currentEditingCategory) item.classList.add('selected'); item.onclick = () => selectRuleCategoryToEdit(catName); listEl.appendChild(item); }); }
    function selectRuleCategoryToEdit(categoryName) { state.currentEditingCategory = categoryName; renderRuleCategoryList(); const editorEl = document.getElementById('ruleCategoryEditorContainer'); editorEl.classList.remove('hidden'); document.getElementById('categoryEditorTitle').textContent = `Editing: ${categoryName}`; renderCategoryRulesDisplay(); }
    function renderCategoryRulesDisplay() { const displayEl = document.getElementById('categoryRulesDisplay'); displayEl.innerHTML = ''; const category = state.dictionary.ruleCategories[state.currentEditingCategory]; if (!category || !category.rules) return; category.rules.forEach((rule, index) => { const pill = document.createElement('span'); pill.className = 'rule-pill'; pill.innerHTML = `${rule.type}: ${rule.value || 'N/A'} <button class="delete-btn ml-2" onclick="deleteRuleFromCategory(${index})">x</button>`; displayEl.appendChild(pill); }); }
    window.addNewRuleCategory = function() { const name = prompt("Enter a name for the new rule category:"); if (name && !state.dictionary.ruleCategories[name]) { state.dictionary.ruleCategories[name] = { rules: [] }; selectRuleCategoryToEdit(name); } else if (name) { alert("A category with this name already exists."); } }
    window.addRuleToCategory = function() { const catName = state.currentEditingCategory; if (!catName) return; const type = document.getElementById('categoryRuleType').value; const valueInput = document.getElementById('categoryRuleValue'); const value = valueInput.value.trim(); if (['ALLOWED_VALUES', 'REGEX'].includes(type) && !value) { alert("Value is required for this rule type."); return; } state.dictionary.ruleCategories[catName].rules.push({ type, value, message: `Validation failed for rule ${type}` }); valueInput.value = ''; renderCategoryRulesDisplay(); }
    window.deleteRuleFromCategory = function(index) { const catName = state.currentEditingCategory; if (catName) { state.dictionary.ruleCategories[catName].rules.splice(index, 1); renderCategoryRulesDisplay(); } }
    window.deleteRuleCategory = function() { const catName = state.currentEditingCategory; if (catName && confirm(`Delete the "${catName}" category?`)) { delete state.dictionary.ruleCategories[catName]; state.currentEditingCategory = null; document.getElementById('ruleCategoryEditorContainer').classList.add('hidden'); renderRuleCategoryList(); } }
    function parseDate(value) { if (value === null || value === undefined || String(value).trim() === '') return null; if (typeof value === 'number' && value > 1 && value <= 2958465) { return new Date((value - 25569) * 86400 * 1000); } const date = new Date(String(value)); return isNaN(date.getTime()) ? null : date; }
    
    function runDataValidation(activeDS, dictionary) {
        const { workbook, name: fileName } = activeDS;
        const analysisResults = {};
        const overallStats = { customIssueCount: 0, duplicateRowCount: 0, totalProcessedCells: 0, totalProcessedRows: 0, totalOriginalRows: 0 };
        // For now, duplicate method is hardcoded. Can be made into a UI option.
        const duplicateMethod = "thorough"; 

        for (const sheetName of workbook.SheetNames) {
            const worksheet = workbook.Sheets[sheetName];
            const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: null, raw: true });
            if (jsonData.length < 2) continue;
            
            const sheetHeader = jsonData[0].map(h => String(h || '').trim());
            let sheetDataRows = jsonData.slice(1);
            overallStats.totalOriginalRows += sheetDataRows.length;
            if (sheetDataRows.length > ROW_LIMIT) {
                sheetDataRows = sheetDataRows.slice(0, ROW_LIMIT);
            }
            
            const currentSheetIssues = { customValidation: {}, duplicateRows: [], sheetHeaders: sheetHeader };

            if (duplicateMethod !== 'none') {
                const seenAsDuplicate = new Set();
                for (let i = 0; i < sheetDataRows.length; i++) {
                    if (seenAsDuplicate.has(i)) continue;
                    for (let j = i + 1; j < sheetDataRows.length; j++) {
                        if (seenAsDuplicate.has(j)) continue;
                        const rowA = sheetDataRows[i], rowB = sheetDataRows[j];
                        if (duplicateMethod === 'fast') {
                            if (rowA.join('~!~') === rowB.join('~!~')) {
                                currentSheetIssues.duplicateRows.push({ rowNumber1: i + 2, rowData1: rowA, rowNumber2: j + 2, rowData2: rowB, matchPercentage: 100 });
                                seenAsDuplicate.add(j);
                            }
                        } else { // thorough
                            let matchingCells = 0;
                            for (let k = 0; k < rowA.length; k++) {
                                if (String(rowA[k] || '').trim() === String(rowB[k] || '').trim()) {
                                    matchingCells++;
                                }
                            }
                            const matchPercentage = (matchingCells / rowA.length) * 100;
                            if (matchPercentage > 50) {
                                currentSheetIssues.duplicateRows.push({ rowNumber1: i + 2, rowData1: rowA, rowNumber2: j + 2, rowData2: rowB, matchPercentage: matchPercentage.toFixed(0) });
                                seenAsDuplicate.add(j);
                            }
                        }
                    }
                }
                overallStats.duplicateRowCount += currentSheetIssues.duplicateRows.length;
            }

            sheetHeader.forEach((header, colIndex) => {
                const headerKey = header.toUpperCase();
                let ruleset = Object.values(dictionary.sheetsData).find(d => d["Column Name"].toUpperCase() === headerKey);
                if (!ruleset) return;

                let finalRules = [];
                if (ruleset.category && dictionary.ruleCategories[ruleset.category]) {
                    finalRules = dictionary.ruleCategories[ruleset.category].rules;
                } else if (ruleset.validation_rules) {
                    finalRules = ruleset.validation_rules;
                }
                if (!finalRules || finalRules.length === 0) return;

                currentSheetIssues.customValidation[header] = [];
                const otherRules = finalRules.filter(r => r.type !== 'ALLOWED_VALUES');
                const allowedValuesRules = finalRules.filter(r => r.type === 'ALLOWED_VALUES');
                if (allowedValuesRules.length > 0) {
                    const combined = new Set(allowedValuesRules.flatMap(r => String(r.value).split(',')).map(v => v.trim().toLowerCase()));
                    otherRules.push({ type: 'ALLOWED_VALUES_COMBINED', value: combined, message: `Value not in allowed list for ${header}` });
                }

                otherRules.forEach(rule => {
                    sheetDataRows.forEach((row, rowIndex) => {
                        overallStats.totalProcessedCells++;
                        const cellValue = row[colIndex];
                        const isCellEmpty = (cellValue === undefined || cellValue === null || String(cellValue).trim() === '');
                        if (rule.type === 'NOT REQUIRED' || (rule.type !== 'REQUIRED' && rule.type !== 'NO_FUTURE_DATE' && rule.type !== 'VALID_DATE' && isCellEmpty)) return;
                        
                        let isValid = true;
                        switch (rule.type) {
                            case 'REQUIRED': isValid = !isCellEmpty; break;
                            case 'ALLOWED_VALUES_COMBINED': isValid = isCellEmpty || rule.value.has(String(cellValue).trim().toLowerCase()); break;
                            case 'REGEX': isValid = isCellEmpty || new RegExp(rule.value, 'i').test(String(cellValue)); break;
                            case 'VALID_DATE':
                                if (isCellEmpty) { isValid = true; break; }
                                isValid = parseDate(cellValue) !== null;
                                if(!isValid) rule.message = "Value is not a recognizable date.";
                                break;
                            case 'NO_FUTURE_DATE':
                                if (isCellEmpty) { isValid = true; break; }
                                const parsedDate = parseDate(cellValue);
                                if (!parsedDate) {
                                    isValid = false; rule.message = "Value is not a recognizable date.";
                                } else {
                                    const today = new Date();
                                    today.setHours(23, 59, 59, 999);
                                    if (parsedDate > today) {
                                        isValid = false; rule.message = `Date ${parsedDate.toLocaleDateString()} is in the future.`;
                                    } else {
                                        isValid = true;
                                    }
                                }
                                break;
                        }
                        if (!isValid) {
                            currentSheetIssues.customValidation[header].push({ row: rowIndex + 2, value: cellValue, type: rule.type.replace('_COMBINED', ''), message: rule.message });
                        }
                    });
                });
                if(currentSheetIssues.customValidation[header]) overallStats.customIssueCount += currentSheetIssues.customValidation[header].length;
            });
            analysisResults[sheetName] = currentSheetIssues;
        }
        state.lastValidationResult = { ...(state.lastValidationResult || {}), fileName, results: analysisResults, stats: overallStats, headers: activeDS.headers };
        displayValidationReport();
    }

    function displayValidationReport() {
        const { fileName, results, stats } = state.lastValidationResult;
        const view = document.getElementById('validation-results-view');
        dataView.classList.add('hidden'); // Hide the standard table
        view.classList.remove('hidden');

        let html = `<div class="bg-white p-6 md:p-8 rounded-xl shadow-lg h-full flex flex-col"><h2 class="text-2xl font-bold text-slate-700 mb-4">Data Validation Report</h2>`;
        if (stats.totalOriginalRows > ROW_LIMIT) {
            html += `<div class="p-4 mb-4 border-l-4 border-blue-400 bg-blue-50 rounded-r-lg"><h3 class="text-sm font-bold text-blue-800">Note: File contains >${ROW_LIMIT.toLocaleString()} rows. Analysis is on the first ${ROW_LIMIT.toLocaleString()} rows.</h3></div>`;
        }
        html += `<p class="text-sm text-gray-500 mb-4">File: ${fileName}</p><div id="report-summary" class="bg-slate-50 p-4 rounded-lg mb-6"></div><div class="flex-grow overflow-y-auto">`;

        for(const sheetName in results) {
            html += `<div class="mb-6"><h3 class="text-xl font-bold text-slate-700 mb-3 border-b pb-2">Sheet: ${sheetName}</h3><ul class="space-y-2 text-sm">`;
            const issues = results[sheetName];
            let hasIssues = false;
            
            if (issues.duplicateRows.length > 0) {
                hasIssues = true;
                const issueKey = `dupe_${sheetName}`;
                let glimpseHTML = '<div class="space-y-4">';
                const sheetHeaders = issues.sheetHeaders || [];
                issues.duplicateRows.slice(0, 20).forEach(d => {
                    glimpseHTML += `<div class="p-2 border rounded-md bg-white"><p class="font-semibold text-sm">Row ${d.rowNumber1} is ${d.matchPercentage}% similar to Row ${d.rowNumber2}</p><table class="w-full text-xs mt-2 duplicate-glimpse-table"><thead><tr><th class="w-1/3">Header</th><th class="w-1/3">Row ${d.rowNumber1} Data</th><th class="w-1/3">Row ${d.rowNumber2} Data</th></tr></thead><tbody>`;
                    for(let i=0; i < Math.max(d.rowData1.length, d.rowData2.length); i++) {
                        const headerName = sheetHeaders[i] || `(Col ${i+1})`;
                        const val1 = d.rowData1[i] != null ? String(d.rowData1[i]) : '';
                        const val2 = d.rowData2[i] != null ? String(d.rowData2[i]) : '';
                        const highlight = val1.trim() !== val2.trim() ? 'bg-yellow-100' : '';
                        glimpseHTML += `<tr class="${highlight}"><td class="font-semibold text-gray-600">${headerName}</td><td>${val1}</td><td>${val2}</td></tr>`;
                    }
                    glimpseHTML += `</tbody></table></div>`;
                });
                glimpseHTML += '</div>';
                html += `<li><div class="flex items-center"><label class="inline-flex items-center cursor-pointer"><input type="checkbox" onchange="toggleIgnore('${issueKey}')" class="form-checkbox h-4 w-4 mr-2"><strong>Similar Rows Found:</strong> ${issues.duplicateRows.length} pairs.</label><button class="ml-2 text-blue-500 text-xs" onclick="toggleErrorDetails(this)">(show)</button></div><div class="error-details hidden">${glimpseHTML}</div></li>`;
            }

            for (const col in issues.customValidation) {
                if (issues.customValidation[col].length > 0) {
                    hasIssues = true;
                    const issueKey = `val_${sheetName}_${col}`;
                    const colI = issues.customValidation[col];
                    html += `<li><div class="flex items-center"><label class="inline-flex items-center cursor-pointer"><input type="checkbox" onchange="toggleIgnore('${issueKey}')" class="form-checkbox h-4 w-4 mr-2"><strong>Column "${col}":</strong> ${colI.length} issues.</label><button class="ml-2 text-blue-500 text-xs" onclick="toggleErrorDetails(this)">(show)</button></div><div class="error-details hidden"><ul class="list-disc list-inside">${colI.slice(0, 50).map(i => `<li>Row ${i.row}: Failed '${i.type}' (Value: "${i.value}"). Message: ${i.message}</li>`).join('')}</ul></div></li>`;
                }
            }
            if (!hasIssues) html += `<li class="text-green-700 font-semibold">No issues found in this sheet.</li>`;
            html += `</ul></div>`;
        }
        html += `</div><div class="flex flex-col md:flex-row gap-4 mt-6 border-t pt-6"><button id="downloadReportBtn" class="flex-1 py-3 px-4 rounded-lg text-lg font-medium text-white bg-green-600 hover:bg-green-700">Download Scored Report</button><button onclick="downloadSubsetPdf()" class="flex-1 py-3 px-4 rounded-lg text-lg font-medium text-white bg-gray-600 hover:bg-gray-700">Download Subset PDF</button></div></div>`;
        view.innerHTML = html;
        document.getElementById('downloadReportBtn').addEventListener('click', downloadScoredReport);
        recalculateScores();
    }
    window.autoAdjustTextarea = function(element) { element.style.height = 'auto'; element.style.height = (element.scrollHeight) + 'px'; }
    window.toggleErrorDetails = function(button){const d=button.closest('li').querySelector('.error-details');const i=d.classList.contains('hidden');d.classList.toggle('hidden',!i);button.textContent=i?'(hide)':'(show)'}
    window.toggleIgnore = function(key){if(state.ignoredErrors.has(key)){state.ignoredErrors.delete(key)}else{state.ignoredErrors.add(key)}document.querySelector(`input[onchange="toggleIgnore('${key}')"]`).closest('li').classList.toggle('issue-ignored',state.ignoredErrors.has(key));recalculateScores()}
    function recalculateScores(){const{results,stats}=state.lastValidationResult;let ignoredCustomCount=0,ignoredDupeCount=0;state.ignoredErrors.forEach(key=>{if(key.startsWith('dupe_'))ignoredDupeCount+=results[key.substring(5)].duplicateRows.length;else if(key.startsWith('val_')){const parts=key.split('_');if(results[parts[1]] && results[parts[1]].customValidation[parts.slice(2).join('_')]) {ignoredCustomCount+=results[parts[1]].customValidation[parts.slice(2).join('_')].length;}}});const effectiveCustomCount=stats.customIssueCount-ignoredCustomCount;const effectiveDupeCount=stats.duplicateRowCount-ignoredDupeCount;const totalErrors=effectiveCustomCount+effectiveDupeCount;const totalChecks=stats.totalProcessedCells+stats.totalProcessedRows;const cleanRate=totalChecks>0?Math.max(0,((totalChecks-totalErrors)/totalChecks)*100):100;const passStatus=cleanRate>=95;const statusClass=passStatus?'text-green-600 font-bold':'text-red-600 font-bold';document.getElementById('report-summary').innerHTML=`<p><strong>Overall Status:</strong> <span class="${statusClass}">${passStatus?'Pass':'Fail'}</span> (Clean Rate: ${cleanRate.toFixed(2)}%, Threshold: 95%)</p><p>Showing ${effectiveCustomCount} validation issues and ${effectiveDupeCount} similar row pairs.</p>`}
    function downloadScoredReport() {
        if (!state.lastValidationResult) return;
        const activeDS = getActiveDataset();
        const { fileName, results, stats, headers } = state.lastValidationResult;
        const newWb = XLSX.utils.book_new();
        let ignoredCustomCount=0, ignoredDupeCount=0;
        state.ignoredErrors.forEach(key => {
            if (key.startsWith('dupe_')) {
                ignoredDupeCount += results[key.substring(5)].duplicateRows.length;
            } else if (key.startsWith('val_')) {
                const parts = key.split('_');
                ignoredCustomCount += results[parts[1]].customValidation[parts.slice(2).join('_')].length;
            }
        });
        const effectiveCustomCount = stats.customIssueCount - ignoredCustomCount;
        const effectiveDupeCount = stats.duplicateRowCount - ignoredDupeCount;
        const totalErrors = effectiveCustomCount + effectiveDupeCount;
        const totalChecks = stats.totalProcessedCells + stats.totalProcessedRows;
        const cleanRate = totalChecks > 0 ? Math.max(0, ((totalChecks - totalErrors) / totalChecks) * 100) : 100;
        const summaryData = [
            ["Data Validation Summary Report"], [], ["File Name", fileName], ["Dictionary Name", state.dictionary.dictionaryName], ["Timestamp", new Date().toLocaleString()], [],
            ["Overall Status", cleanRate >= 95 ? "Pass" : "Fail"], ["Final Clean Rate", `${cleanRate.toFixed(2)}%`], ["Effective Validation Issues", effectiveCustomCount], ["Effective Similar Row Pairs", effectiveDupeCount]
        ];
        const summaryWs = XLSX.utils.aoa_to_sheet(summaryData);
        XLSX.utils.book_append_sheet(newWb, summaryWs, "Validation Summary");

        const duplicateData = []; let dupesFound = false;
        for (const sheetName in results) {
            if (results[sheetName].duplicateRows.length > 0 && !state.ignoredErrors.has(`dupe_${sheetName}`)) {
                dupesFound = true;
                results[sheetName].duplicateRows.forEach(d => {
                    duplicateData.push([sheetName, d.rowNumber1, `${d.matchPercentage}%`, d.rowNumber2, ...d.rowData1]);
                    duplicateData.push([sheetName, d.rowNumber2, `${d.matchPercentage}%`, d.rowNumber1, ...d.rowData2]);
                });
            }
        }
        if(dupesFound) {
            const dupesWs = XLSX.utils.aoa_to_sheet([["Original Sheet", "Row Number", "Match %", "Matched With Row", ...headers], ...duplicateData]);
            XLSX.utils.book_append_sheet(newWb, dupesWs, "Duplicate Rows");
        }

        const errorDetailsByRow = {};
        for (const sheetName in results) {
            errorDetailsByRow[sheetName] = {};
            for (const col in results[sheetName].customValidation) {
                if (!state.ignoredErrors.has(`val_${sheetName}_${col}`)) {
                    results[sheetName].customValidation[col].forEach(issue => {
                        const existing = errorDetailsByRow[sheetName][issue.row] || '';
                        const newError = `Column '${col}': ${issue.message || issue.type}`;
                        errorDetailsByRow[sheetName][issue.row] = existing ? `${existing}; ${newError}` : newError;
                    });
                }
            }
        }
        
        const errorData = []; let errorsFound = false;
        activeDS.workbook.SheetNames.forEach(sheetName => {
            const errorRowNumbers = Object.keys(errorDetailsByRow[sheetName] || {}).map(Number);
            if (errorRowNumbers.length > 0) {
                errorsFound = true;
                const originalData = XLSX.utils.sheet_to_json(activeDS.workbook.Sheets[sheetName], { header: 1, raw: true });
                errorRowNumbers.sort((a,b) => a - b).forEach(rowNum => {
                    if(originalData[rowNum - 1]) {
                        errorData.push([sheetName, errorDetailsByRow[sheetName][rowNum], ...originalData[rowNum - 1]]);
                    }
                });
            }
        });
        if (errorsFound) {
            const errorsWs = XLSX.utils.aoa_to_sheet([["Original Sheet Name", "Error Details", ...headers], ...errorData]);
            XLSX.utils.book_append_sheet(newWb, errorsWs, "Validation Errors");
        }

        activeDS.workbook.SheetNames.forEach(sheetName => {
            XLSX.utils.book_append_sheet(newWb, activeDS.workbook.Sheets[sheetName], `Original_${sheetName}`);
        });

        const desiredOrder = ["Validation Summary", "Duplicate Rows", "Validation Errors"];
        const reorderedSheets = desiredOrder.filter(name => newWb.SheetNames.includes(name));
        newWb.SheetNames.forEach(name => { if (!reorderedSheets.includes(name)) reorderedSheets.push(name); });
        newWb.SheetNames = reorderedSheets;
        XLSX.writeFile(newWb, `Validated_${fileName}`);
    }
    
    window.downloadSubsetPdf = function() {
        if (!state.dictionary || !state.lastValidationResult) { alert("Analyze a file first."); return; }
        const { jsPDF } = window.jspdf;
        const doc = new jsPDF('landscape');
        const { headers: fileHeaders, skippedColumns } = state.lastValidationResult;
        const headersToKeep = new Set(fileHeaders.map(h => h.toUpperCase()));
        const tableBody = [], tableHead = ['Column Name', 'Rule Category', 'Description', 'Defined Rules'];
        
        fileHeaders.forEach(header => {
            const headerKey = header.toUpperCase();
            if (headersToKeep.has(headerKey) && (!skippedColumns || !skippedColumns.includes(header))) {
                const ruleData = Object.values(state.dictionary.sheetsData).find(d=>d["Column Name"].toUpperCase() === headerKey);
                if (ruleData) {
                    let rulesToDisplay = [];
                    if (ruleData.category && state.dictionary.ruleCategories[ruleData.category]) {
                        rulesToDisplay = state.dictionary.ruleCategories[ruleData.category].rules;
                    } else if (ruleData.validation_rules) {
                        rulesToDisplay = ruleData.validation_rules;
                    }
                    const rulesText = (rulesToDisplay || []).map(r => `${r.type}${r.value ? `: ${r.value}` : ''}`).join('\n');
                    tableBody.push([ ruleData['Column Name'], ruleData.category || 'N/A', ruleData.description || '', rulesText ]);
                }
            }
        });

        doc.setFontSize(18); doc.text("Data Dictionary Subset Report", 14, 22);
        doc.setFontSize(11);
        doc.text(`Dictionary: ${state.dictionary.dictionaryName}`, 14, 30);
        doc.text(`Source File: ${state.lastValidationResult.fileName}`, 14, 36);
        doc.autoTable({ startY: 40, head: [tableHead], body: tableBody, theme: 'grid', headStyles: { fillColor: [22, 160, 133] }, styles: { fontSize: 8, cellPadding: 1.5 }, columnStyles: { 0: { cellWidth: 40 }, 1: { cellWidth: 40 }, 2: { cellWidth: 'auto' }, 3: { cellWidth: 65 } } });
        
        if (skippedColumns && skippedColumns.length > 0) {
            doc.addPage();
            doc.setFontSize(18); doc.text("Columns Excluded From Dictionary", 14, 22);
            doc.setFontSize(11); doc.text("The following columns were not saved to the dictionary because no description was provided.", 14, 30);
            const skippedTableBody = skippedColumns.map(col => [col]);
            doc.autoTable({ startY: 40, head: [['Excluded Column Name']], body: skippedTableBody, theme: 'grid' });
        }
        doc.save(`Subset_Dictionary_For_${state.lastValidationResult.fileName}.pdf`);
    }

    function getAllHeadersFromDictionary(dictionary) {
        return dictionary && dictionary.sheetsData ? new Set(Object.values(dictionary.sheetsData).map(item => item["Column Name"].toUpperCase())) : new Set();
    }
});
