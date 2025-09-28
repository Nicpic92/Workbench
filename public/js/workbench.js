document.addEventListener('DOMContentLoaded', () => {
    // --- STATE MANAGEMENT ---
    const state = {
        datasets: [],
        activeDatasetIndex: 0,
        dictionaries: {},
    };

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
    const configModal = document.getElementById('config-modal');
    const modalTitle = document.getElementById('modal-title');
    const modalBody = document.getElementById('modal-body');
    let modalConfirmBtn = document.getElementById('modal-confirm-btn');
    const modalCancelBtn = document.getElementById('modal-cancel-btn');
    const modalCloseBtn = document.getElementById('modal-close-btn');

    // --- INITIALIZATION ---
    fileUploadInput.addEventListener('change', handleFileUpload);
    downloadBtn.addEventListener('click', handleDownload);
    [modalCancelBtn, modalCloseBtn].forEach(btn => btn.addEventListener('click', hideModal));
    loadDictionaries();

    // --- FILE HANDLING ---
    async function handleFileUpload(event) {
        const files = Array.from(event.target.files);
        if (files.length === 0) return;
        showLoader(true, 'Reading files...');
        state.datasets = [];
        for (const file of files) {
            try {
                const data = await readFile(file);
                const workbook = XLSX.read(data, { type: 'array', cellDates: true });
                if (workbook.SheetNames.length === 1) {
                    const jsonData = XLSX.utils.sheet_to_json(workbook.Sheets[workbook.SheetNames[0]]);
                    state.datasets.push({ name: file.name, data: jsonData, headers: Object.keys(jsonData[0] || {}) });
                } else {
                    workbook.SheetNames.forEach(sheetName => {
                        const jsonData = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName]);
                        state.datasets.push({ name: `${file.name} - ${sheetName}`, data: jsonData, headers: Object.keys(jsonData[0] || {}) });
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
            dataView.style.display = 'none';
            actionsContainer.style.display = 'none';
            loadedFilesList.innerHTML = '';
        } else {
            welcomeView.style.display = 'none';
            dataView.style.display = 'flex';
            actionsContainer.style.display = 'block';
            renderLoadedFilesList();
            renderActiveDataset();
        }
    }

    function renderLoadedFilesList() {
        loadedFilesList.innerHTML = `<h3 class="font-semibold text-gray-700 mb-2 border-t pt-4">Loaded Datasets</h3>`;
        state.datasets.forEach((ds, index) => {
            const isActive = index === state.activeDatasetIndex;
            const item = document.createElement('div');
            item.className = `p-2 rounded cursor-pointer ${isActive ? 'bg-indigo-100 font-bold' : 'hover:bg-gray-100'}`;
            item.textContent = ds.name;
            item.onclick = () => { state.activeDatasetIndex = index; updateUI(); };
            loadedFilesList.appendChild(item);
        });
    }

    function renderActiveDataset() {
        const activeDataset = getActiveDataset();
        if (!activeDataset) return;
        tableTitle.textContent = activeDataset.name;
        renderDataTable(activeDataset.data, activeDataset.headers);
        statusBar.textContent = `Displaying ${activeDataset.data.length.toLocaleString()} rows and ${activeDataset.headers.length} columns. (Preview of first 200 rows)`;
    }

    function renderDataTable(data, headers) {
        const table = document.createElement('table');
        table.className = 'min-w-full divide-y divide-gray-200';
        const thead = table.createTHead();
        thead.className = 'bg-gray-50';
        const headerRow = thead.insertRow();
        headers.forEach(h => {
            const th = document.createElement('th');
            th.className = 'px-4 py-3 text-left text-xs font-medium text-gray-500 uppercase tracking-wider bg-gray-50';
            th.textContent = h;
            headerRow.appendChild(th);
        });
        const tbody = table.createTBody();
        tbody.className = 'bg-white divide-y divide-gray-200';
        data.slice(0, 200).forEach(row => {
            const tr = tbody.insertRow();
            tr.className = 'hover:bg-gray-50';
            headers.forEach(header => {
                 const td = tr.insertCell();
                 td.className = 'px-4 py-2 whitespace-nowrap text-sm text-gray-700';
                 td.textContent = row[header] ?? '';
            });
        });
        tableContainer.innerHTML = '';
        tableContainer.appendChild(table);
    }
    
    function showLoader(show, message = '') {
        loaderOverlay.style.display = show ? 'flex' : 'none';
    }

    // --- MODAL & CONFIGURATION ---
    function showModal(title, content, onConfirm) {
        modalTitle.textContent = title;
        modalBody.innerHTML = content;
        configModal.style.display = 'flex';
        const newConfirmBtn = modalConfirmBtn.cloneNode(true);
        modalConfirmBtn.parentNode.replaceChild(newConfirmBtn, modalConfirmBtn);
        newConfirmBtn.addEventListener('click', onConfirm);
        modalConfirmBtn = newConfirmBtn;
    }

    function hideModal() { configModal.style.display = 'none'; }
    function generateColumnCheckboxes(headers) { return headers.map(h => `<label class="flex items-center p-2 rounded hover:bg-gray-100"><input type="checkbox" class="h-4 w-4 rounded mr-2" data-column-name="${h}"><span class="text-sm">${h}</span></label>`).join(''); }
    function generateColumnSelect(headers, id) { return `<select id="${id}" class="w-full p-2 border rounded mt-1">${headers.map(h => `<option value="${h}">${h}</option>`).join('')}</select>`; }
    function generateDatasetSelect(id) { return `<select id="${id}" class="w-full p-2 border rounded mt-1">${state.datasets.map((ds, i) => `<option value="${i}">${ds.name}</option>`).join('')}</select>`; }

    // --- ACTION IMPLEMENTATIONS ---
    function getActiveDataset() { return state.datasets[state.activeDatasetIndex]; }
    function addNewDataset(name, data, headers) {
        state.datasets.push({ name, data, headers });
        state.activeDatasetIndex = state.datasets.length - 1;
        updateUI();
    }
    
    // --- ACTION EVENT LISTENERS ---
    document.getElementById('action-trim-whitespace').addEventListener('click', () => {
        const activeDS = getActiveDataset();
        if (!activeDS) return alert("Please load a file first.");
        const content = `<p class="text-sm mb-4">Select the column to trim.</p><label for="trim-column" class="block text-sm font-semibold">Column:</label>${generateColumnSelect(activeDS.headers, 'config-column')}`;
        showModal('Trim Whitespace', content, () => {
            const column = document.getElementById('config-column').value;
            showLoader(true);
            setTimeout(() => {
                getActiveDataset().data.forEach(row => { if (typeof row[column] === 'string') row[column] = row[column].trim(); });
                renderActiveDataset();
                showLoader(false); hideModal();
            }, 50);
        });
    });

    document.getElementById('action-anonymize').addEventListener('click', () => {
        const activeDS = getActiveDataset();
        if (!activeDS) return alert("Please load a file first.");
        const types = [{ v: 'NONE', l: 'Do Not Anonymize' }, { v: 'FULL_NAME', l: 'Full Name' }, { v: 'FIRST_NAME', l: 'First Name' }, { v: 'LAST_NAME', l: 'Last Name' }, { v: 'EMAIL', l: 'Email' }, { v: 'PHONE', l: 'Phone' }];
        const content = activeDS.headers.map(h => `<div class="grid grid-cols-2 gap-4 items-center border-b pb-2 mb-2"><label class="font-semibold truncate" title="${h}">${h}</label><select data-header="${h}" class="column-mapper w-full p-2 border rounded">${types.map(t => `<option value="${t.v}">${t.l}</option>`).join('')}</select></div>`).join('');
        showModal('Anonymize Personal Information', content, () => {
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
                showLoader(false); hideModal();
            }, 50);
        });
    });

    document.getElementById('action-extract-columns').addEventListener('click', () => {
        const activeDS = getActiveDataset();
        if (!activeDS) return alert("Please load a file first.");
        const content = `<p class="text-sm mb-4">Select columns to keep.</p><div class="space-y-2 max-h-96 overflow-y-auto">${generateColumnCheckboxes(activeDS.headers)}</div>`;
        showModal('Extract Columns', content, () => {
            const selected = Array.from(modalBody.querySelectorAll('input:checked')).map(cb => cb.dataset.columnName);
            if (selected.length === 0) return alert('Please select at least one column.');
            showLoader(true);
            setTimeout(() => {
                const newData = activeDS.data.map(row => selected.reduce((obj, key) => (obj[key] = row[key], obj), {}));
                addNewDataset(`Extracted - ${activeDS.name}`, newData, selected);
                showLoader(false); hideModal();
            }, 50);
        });
    });

    document.getElementById('action-stack-sheets').addEventListener('click', () => {
        if (state.datasets.length < 2) return alert("Please load at least two files to stack.");
        const content = `<p class="text-sm mb-4">This will combine all ${state.datasets.length} currently loaded datasets into a single master sheet. Columns will be matched by header name.</p>`;
        showModal('Stack All Sheets', content, () => {
            showLoader(true);
            setTimeout(() => {
                const allData = state.datasets.flatMap(ds => ds.data);
                const allHeaders = [...new Set(state.datasets.flatMap(ds => ds.headers))];
                addNewDataset(`Stacked - ${state.datasets.length} files`, allData, allHeaders);
                showLoader(false); hideModal();
            }, 50);
        });
    });
    
    document.getElementById('action-merge-files').addEventListener('click', () => {
        if (state.datasets.length < 2) return alert("Upload at least two files to use merge.");
        const content = `<div class="grid grid-cols-2 gap-4"><div class="border-r pr-4"><label class="block text-sm font-semibold">Left Table (Primary)</label>${generateDatasetSelect('config-ds1')}<label class="block text-sm font-semibold mt-2">Key Column</label><select id="config-key1" class="w-full p-2 border rounded mt-1"></select></div><div><label class="block text-sm font-semibold">Right Table (to join)</label>${generateDatasetSelect('config-ds2')}<label class="block text-sm font-semibold mt-2">Key Column</label><select id="config-key2" class="w-full p-2 border rounded mt-1"></select></div></div>`;
        const populateKeys = () => { ['1', '2'].forEach(n => { const ds_idx = document.getElementById(`config-ds${n}`).value; document.getElementById(`config-key${n}`).innerHTML = state.datasets[ds_idx].headers.map(h => `<option value="${h}">${h}</option>`).join(''); }); };
        showModal('Merge Files (Left Join)', content, () => {
            const ds1_idx = document.getElementById('config-ds1').value, ds2_idx = document.getElementById('config-ds2').value;
            const key1 = document.getElementById('config-key1').value, key2 = document.getElementById('config-key2').value;
            showLoader(true);
            setTimeout(() => {
                const ds1 = state.datasets[ds1_idx], ds2 = state.datasets[ds2_idx];
                const map2 = new Map(ds2.data.map(row => [row[key2], row]));
                const mergedData = ds1.data.map(row1 => ({ ...row1, ...(map2.get(row1[key1]) || {}) }));
                const newHeaders = [...new Set([...ds1.headers, ...ds2.headers])];
                addNewDataset(`Merged - ${ds1.name} & ${ds2.name}`, mergedData, newHeaders);
                showLoader(false); hideModal();
            }, 50);
        });
        populateKeys();
        document.getElementById('config-ds1').onchange = populateKeys;
        document.getElementById('config-ds2').onchange = populateKeys;
    });

    document.getElementById('action-find-duplicates').addEventListener('click', () => {
        const activeDS = getActiveDataset();
        if (!activeDS) return alert("Please load a file first.");
        const content = `<p class="text-sm mb-4">Select columns to check for duplicates.</p><div class="space-y-2 max-h-96 overflow-y-auto">${generateColumnCheckboxes(activeDS.headers)}</div>`;
        showModal('Find Duplicates', content, () => {
            const selected = Array.from(modalBody.querySelectorAll('input:checked')).map(cb => cb.dataset.columnName);
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
                showLoader(false); hideModal();
            }, 50);
        });
    });

    document.getElementById('action-compare-sheets').addEventListener('click', () => {
        if (state.datasets.length < 2) return alert("Upload at least two files to compare.");
        const content = `<div class="grid grid-cols-2 gap-4"><div class="border-r pr-4"><label class="block text-sm font-semibold">Original / Old File</label>${generateDatasetSelect('config-ds1')}<label class="block text-sm font-semibold mt-2">Unique ID Column</label><select id="config-key1" class="w-full p-2 border rounded mt-1"></select></div><div><label class="block text-sm font-semibold">New / Updated File</label>${generateDatasetSelect('config-ds2')}<label class="block text-sm font-semibold mt-2">Unique ID Column</label><select id="config-key2" class="w-full p-2 border rounded mt-1"></select></div></div>`;
        const populateKeys = () => { ['1', '2'].forEach(n => { const ds_idx = document.getElementById(`config-ds${n}`).value; document.getElementById(`config-key${n}`).innerHTML = state.datasets[ds_idx].headers.map(h => `<option value="${h}">${h}</option>`).join(''); }); };
        showModal('Compare Sheets', content, () => {
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
                showLoader(false); hideModal();
            }, 50);
        });
        populateKeys();
        document.getElementById('config-ds1').onchange = populateKeys;
        document.getElementById('config-ds2').onchange = populateKeys;
    });

    document.getElementById('action-split-file').addEventListener('click', () => {
        const activeDS = getActiveDataset();
        if (!activeDS) return alert("Please load a file first.");
        const content = `<p class="text-sm mb-4">Split the active dataset into multiple CSV files.</p><label for="config-rows" class="block text-sm font-semibold">Rows Per File</label><input type="number" id="config-rows" value="100000" min="1" class="w-full p-2 border rounded mt-1">`;
        showModal('Split File by Row Count', content, () => {
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
                zip.generateAsync({ type: 'blob' }).then(c => { saveAs(c, `Split_${activeDS.name}.zip`); showLoader(false); hideModal(); });
            }, 50);
        });
    });
    
    // --- DICTIONARY & VALIDATOR MODULE ---
    function loadDictionaries() { state.dictionaries = JSON.parse(localStorage.getItem('spreadsim_dictionaries') || '{}'); }
    function saveDictionaries() { localStorage.setItem('spreadsim_dictionaries', JSON.stringify(state.dictionaries)); }

    document.getElementById('manage-dictionaries-btn').addEventListener('click', () => {
        let content = `<div class="flex justify-between items-center mb-4"><div><label for="dictionary-select" class="block text-sm font-semibold">Select:</label><select id="dictionary-select" class="p-2 border rounded-md mt-1"></select></div><div><button id="new-dictionary-btn" class="text-sm bg-green-500 text-white py-1 px-3 rounded hover:bg-green-600">New</button><button id="delete-dictionary-btn" class="text-sm bg-red-500 text-white py-1 px-3 rounded hover:bg-red-600 ml-2">Delete</button></div></div><div id="dictionary-rules-container" class="space-y-2"></div><button id="add-rule-btn" class="mt-4 text-indigo-600 font-semibold hover:text-indigo-800">+ Add Rule</button>`;
        showModal('Manage Data Dictionaries', content, () => {
            const name = document.getElementById('dictionary-select').value; if (!name) return;
            state.dictionaries[name] = { rules: Array.from(document.querySelectorAll('#dictionary-rules-container > div')).map(d => ({ column: d.querySelector('.rule-column').value, type: d.querySelector('.rule-type').value, value: d.querySelector('.rule-value').value })) };
            saveDictionaries(); hideModal();
        });
        const select = document.getElementById('dictionary-select');
        const populateSelect = () => { select.innerHTML = Object.keys(state.dictionaries).map(n => `<option value="${n}">${n}</option>`).join(''); };
        const renderRules = name => {
            const container = document.getElementById('dictionary-rules-container');
            if (!name || !state.dictionaries[name]) { container.innerHTML = ''; return; }
            const rules = state.dictionaries[name].rules;
            const opts = [{ v: 'REQUIRED', l: 'Not Empty' }, { v: 'REGEX', l: 'Matches Pattern' }, { v: 'ALLOWED_VALUES', l: 'Is One Of' }];
            container.innerHTML = rules.map(r => `<div class="grid grid-cols-4 gap-2 items-end p-2 border rounded bg-gray-50"><div class="col-span-2"><label class="text-xs font-medium">Column</label><input type="text" value="${r.column}" class="rule-column w-full p-1 border rounded mt-1"></div><div><label class="text-xs font-medium">Rule</label><select class="rule-type w-full p-1 border rounded mt-1">${opts.map(o => `<option value="${o.v}" ${r.type===o.v?'selected':''}>${o.l}</option>`).join('')}</select></div><div class="flex items-center gap-2"><div class="flex-grow"><label class="text-xs font-medium">Value</label><input type="text" value="${r.value||''}" class="rule-value w-full p-1 border rounded mt-1"></div><button class="text-red-500 font-semibold" onclick="this.parentElement.parentElement.remove()">X</button></div></div>`).join('');
        };
        modalBody.onclick = e => {
            if (e.target.id === 'new-dictionary-btn') { const n = prompt("Dictionary name:"); if (n && !state.dictionaries[n]) { state.dictionaries[n] = { rules: [] }; populateSelect(); select.value = n; renderRules(n); } } 
            else if (e.target.id === 'delete-dictionary-btn') { const n = select.value; if (n && confirm(`Delete "${n}"?`)) { delete state.dictionaries[n]; saveDictionaries(); populateSelect(); renderRules(select.value); } } 
            else if (e.target.id === 'add-rule-btn') { const n = select.value; if (n) { state.dictionaries[n].rules.push({ column: '', type: 'REQUIRED', value: '' }); renderRules(n); } }
        };
        select.onchange = () => renderRules(select.value);
        populateSelect(); renderRules(select.value);
    });

    document.getElementById('action-validate-data').addEventListener('click', () => {
        if (Object.keys(state.dictionaries).length === 0) return alert("No dictionaries found. Create one first.");
        const activeDS = getActiveDataset();
        if (!activeDS) return alert("Please load a file first.");
        const content = `<p class="text-sm mb-4">Select a dictionary to validate against.</p><label for="validator-dict-select" class="block text-sm font-semibold">Dictionary:</label><select id="validator-dict-select" class="w-full p-2 border rounded mt-1">${Object.keys(state.dictionaries).map(n=>`<option value="${n}">${n}</option>`).join('')}</select>`;
        showModal('Validate Data', content, () => {
            const dictName = document.getElementById('validator-dict-select').value;
            const dictionary = state.dictionaries[dictName];
            showLoader(true);
            setTimeout(() => {
                const errors = [];
                activeDS.data.forEach((row, index) => {
                    dictionary.rules.forEach(rule => {
                        const value = row[rule.column]; let isValid = true, msg = '';
                        switch (rule.type) {
                            case 'REQUIRED': if (value == null || String(value).trim() === '') { isValid = false; msg = 'Is empty'; } break;
                            case 'REGEX': try { if (!new RegExp(rule.value).test(value)) { isValid = false; msg = `Doesn't match pattern`; } } catch (e) { } break;
                            case 'ALLOWED_VALUES': if (!rule.value.split(',').map(v => v.trim()).includes(String(value))) { isValid = false; msg = `Not in allowed list`; } break;
                        }
                        if (!isValid) errors.push({ 'Row': index + 2, 'Column': rule.column, 'Value': value, 'Error': msg });
                    });
                });
                if (errors.length > 0) addNewDataset(`Validation Errors - ${activeDS.name}`, errors, ['Row', 'Column', 'Value', 'Error']);
                else alert('Validation complete. No errors found!');
                showLoader(false); hideModal();
            }, 50);
        });
    });

    // --- CLAIM STATUS REPORT MODULE (Corrected Logic) ---
    document.getElementById('action-claim-status-report').addEventListener('click', () => {
        if (state.datasets.length < 1) return alert("Please upload at least one report file.");
        const presets = {
            solis: { label: 'Clean Age (Q):', cleanAgeCol: 'Q', claimStatusCol: 'I', payerCol: 'A', dsnpCol: 'X', claimTypeCol: 'B', totalChargesCol: 'S', dateCols: 'E,O,P', notesCol: 'AA', claimNumberCol: 'C' },
            liberty: { label: 'Age (R):', cleanAgeCol: 'R', claimStatusCol: 'I', payerCol: 'A', dsnpCol: 'Y', claimTypeCol: 'B', totalChargesCol: 'T', dateCols: 'E,O,P', notesCol: 'AA', claimNumberCol: 'C' },
            secur: { label: 'Clean Age (Q):', cleanAgeCol: 'Q', claimStatusCol: 'I', payerCol: 'A', dsnpCol: 'Y', claimTypeCol: 'D', totalChargesCol: 'T', dateCols: 'E,O,P', notesCol: 'AA', claimNumberCol: 'C' }
        };
        let content = `
            <p class="text-sm mb-4">Generates the multi-tab daily action report and summary email text based on the latest logic.</p>
            <div class="grid grid-cols-2 gap-4 mb-4">
                <div><label class="block text-sm font-semibold">Today's Report:</label>${generateDatasetSelect('csr-today-ds')}</div>
                <div><label class="block text-sm font-semibold">Yesterday's (Optional):</label><select id="csr-yesterday-ds" class="w-full p-2 border rounded mt-1"></select></div>
            </div>
            <div class="mb-4">
                <label class="block text-sm font-semibold mb-2">Client Preset:</label>
                <select id="csr-client-preset" class="w-full p-2 border rounded"><option value="">-- Choose a Client --</option>${Object.keys(presets).map(p=>`<option value="${p}">${p.toUpperCase()}</option>`).join('')}</select>
            </div>
            <div id="csr-config-div" class="space-y-4 text-sm p-4 border rounded-md bg-gray-50 hidden">
                <div class="grid grid-cols-1 md:grid-cols-3 gap-x-6 gap-y-4">
                    <div><label id="csr-cleanAgeCol-label" class="font-medium">Clean Age:</label><input type="text" id="csr-cleanAgeCol" class="w-full p-1 border rounded uppercase"></div>
                    <div><label class="font-medium">Claim State:</label><input type="text" id="csr-claimStatusCol" class="w-full p-1 border rounded uppercase"></div>
                    <div><label class="font-medium">Claim Number:</label><input type="text" id="csr-claimNumberCol" class="w-full p-1 border rounded uppercase"></div>
                    <div><label class="font-medium">Payer:</label><input type="text" id="csr-payerCol" class="w-full p-1 border rounded uppercase"></div>
                    <div><label class="font-medium">DSNP Status:</label><input type="text" id="csr-dsnpCol" class="w-full p-1 border rounded uppercase"></div>
                    <div><label class="font-medium">Claim Type:</label><input type="text" id="csr-claimTypeCol" class="w-full p-1 border rounded uppercase"></div>
                    <div><label class="font-medium">Total Charges:</label><input type="text" id="csr-totalChargesCol" class="w-full p-1 border rounded uppercase"></div>
                    <div><label class="font-medium">W9/Notes Column:</label><input type="text" id="csr-notesCol" class="w-full p-1 border rounded uppercase"></div>
                    <div class="md:col-span-3"><label class="font-medium">Date Columns (comma-separated):</label><input type="text" id="csr-dateCols" class="w-full p-1 border rounded uppercase"></div>
                </div>
            </div>
        `;
        showModal('Daily Claim Status Report', content, runClaimStatusReport);
        document.getElementById('csr-yesterday-ds').innerHTML = `<option value="-1">-- None --</option>` + state.datasets.map((ds, i) => `<option value="${i}">${ds.name}</option>`).join('');
        const presetSelect = document.getElementById('csr-client-preset');
        const csrConfigDiv = document.getElementById('csr-config-div');
        presetSelect.onchange = () => {
            const client = presetSelect.value;
            if (client) {
                const p = presets[client];
                document.getElementById('csr-cleanAgeCol-label').textContent = p.label;
                const basePresetKeys = Object.keys(presets.solis); 
                basePresetKeys.forEach(key => {
                    if (key !== 'label') {
                        // ******** THE FIX IS HERE ********
                        // Correctly construct the element ID (e.g., 'csr-cleanAgeCol') from the preset key (e.g., 'cleanAgeCol')
                        const element = document.getElementById(`csr-${key}`);
                        if (element) {
                            element.value = p[key] || '';
                        }
                    }
                });
                csrConfigDiv.classList.remove('hidden');
            } else {
                csrConfigDiv.classList.add('hidden');
            }
        };
    });

    async function runClaimStatusReport() {
        showLoader(true, 'Generating Claim Status Report...');
        
        const colLetterToIndex = (letter) => {
            if (!letter || typeof letter !== 'string' || !/^[A-Z]+$/i.test(letter)) return -1;
            let col = 0; letter = letter.toUpperCase();
            for (let i = 0; i < letter.length; i++) col += (letter.charCodeAt(i) - 64) * Math.pow(26, letter.length - i - 1);
            return col - 1;
        };
        const parseCurrency = (value) => (value === null || value === undefined) ? NaN : parseFloat(String(value).replace(/[^0-9.-]/g, '')) || NaN;
        const truncateSheetName = (name) => name.replace(/[\\\/\?\*\[\]]/g, '').substring(0, 31);
        const getFormattedDate = () => {
            const d=new Date(), day=d.getDate(), month=["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"][d.getMonth()], year=d.getFullYear();
            const s = (day%10===1&&day!==11)?'st':((day%10===2&&day!==12)?'nd':((day%10===3&&day!==13)?'rd':'th'));
            return `${day}${s} ${month} ${year}`;
        }
        const extractDueDateFromNote = (noteText) => {
            const match = (noteText || '').match(/due\s*(by)?[:\s]*(\d{1,2}[\/-]\d{1,2}(?:[\/-]\d{2,4})?)/i);
            return match ? match[2] : '';
        }

        await new Promise(resolve => setTimeout(resolve, 50));

        try {
            if (!document.getElementById('csr-client-preset').value) throw new Error("Please choose a client preset.");
            const config = {};
            const ids = ['cleanAge', 'claimStatus', 'claimNumber', 'payer', 'dsnp', 'claimType', 'totalCharges', 'notes'];
            ids.forEach(id => config[`${id}Index`] = colLetterToIndex(document.getElementById(`csr-${id}Col`).value));
            config.dateIndices = document.getElementById('csr-dateCols').value.toUpperCase().trim().split(',').filter(c => c).map(c => colLetterToIndex(c.trim()));
            if (Object.values(config).some(val => val === -1 || (Array.isArray(val) && val.includes(-1)))) throw new Error("Invalid column letter entered.");
            const clientText = document.getElementById('csr-client-preset').options[document.getElementById('csr-client-preset').selectedIndex].text;
            
            const todayDS_idx = document.getElementById('csr-today-ds').value;
            const yesterdayDS_idx = document.getElementById('csr-yesterday-ds').value;
            const mainFile = state.datasets[todayDS_idx];
            const yesterdayFile = yesterdayDS_idx !== "-1" ? state.datasets[yesterdayDS_idx] : null;
            const hasYesterdayReport = !!yesterdayFile;

            const jsonToAOA = (ds) => [ds.headers, ...ds.data.map(row => ds.headers.map(h => row[h] ?? null))];
            const main_aoa = jsonToAOA(mainFile);
            if (main_aoa.length < 2) throw new Error("Main report is empty or has no data.");

            const getStatsForAOA = (aoa, config) => {
                const dayBuckets = { '0 - 20': 0, '21 - 29': 0, '30 - 59': 0, '60+': 0 };
                const stats = { 'PEND': { total: 0, ...dayBuckets }, 'ONHOLD': { total: 0, ...dayBuckets }, 'MANAGEMENTREVIEW': { total: 0, ...dayBuckets }, 'HC MGMT REV': { total: 0, ...dayBuckets }, 'W9_LETTER_NEEDED': { total: 0, ...dayBuckets }, 'W9_FOLLOW_UP': { total: 0, ...dayBuckets } };
                for(const row of aoa.slice(1)) {
                    if (row.every(cell => cell === null)) continue;
                    let claimState = String(row[config.claimStatusIndex] || '').trim().toUpperCase();
                    if (claimState === 'PREBATCH') continue;
                    const totalCharges = parseCurrency(row[config.totalChargesIndex]);
                    const claimType = String(row[config.claimTypeIndex] || '').trim().toUpperCase();
                    const cleanAge = parseInt(row[config.cleanAgeIndex], 10);
                    const daysValue = !isNaN(cleanAge) ? (cleanAge <= 20 ? '0 - 20' : (cleanAge <= 29 ? '21 - 29' : (cleanAge <= 59 ? '30 - 59' : '60+'))) : '';
                    let finalClaimState = claimState;
                    if (claimState === 'MANAGEMENTREVIEW' && !isNaN(totalCharges) && ((claimType.includes('PROFESSIONAL') && totalCharges > 3500) || (claimType.includes('INSTITUTIONAL') && totalCharges > 6500))) {
                        finalClaimState = 'HC MGMT REV';
                    }
                    if (stats[finalClaimState]) {
                        stats[finalClaimState].total++;
                        if (stats[finalClaimState][daysValue] !== undefined) stats[finalClaimState][daysValue]++;
                    }
                    const noteTextLower = String(row[config.notesIndex] || '').toLowerCase();
                    if (noteTextLower.includes('w9')) {
                        if (noteTextLower.includes('requested') || noteTextLower.includes('req') || noteTextLower.includes('due')) {
                           if (stats['W9_FOLLOW_UP'][daysValue] !== undefined) stats['W9_FOLLOW_UP'][daysValue]++;
                           stats['W9_FOLLOW_UP'].total++;
                        } else if (noteTextLower.includes('denied') || noteTextLower.includes('missing') || noteTextLower.includes('not on file') || noteTextLower.includes('not received')) {
                           if (stats['W9_LETTER_NEEDED'][daysValue] !== undefined) stats['W9_LETTER_NEEDED'][daysValue]++;
                           stats['W9_LETTER_NEEDED'].total++;
                        }
                    }
                }
                return stats;
            };

            const todayStats = getStatsForAOA(main_aoa, config);
            let yesterdayStats = null, yesterdayDataMap = new Map();

            if (hasYesterdayReport) {
                const yesterday_aoa = jsonToAOA(yesterdayFile);
                yesterdayStats = getStatsForAOA(yesterday_aoa, config);
                yesterday_aoa.slice(1).forEach(row => {
                    const claimNumber = row[config.claimNumberIndex];
                    if (claimNumber != null) {
                        yesterdayDataMap.set(String(claimNumber).trim(), {
                            state: String(row[config.claimStatusIndex] || '').trim().toUpperCase(),
                            type: String(row[config.claimTypeIndex] || '').trim().toUpperCase(),
                            charges: parseCurrency(row[config.totalChargesIndex])
                        });
                    }
                });
            }

            const masterSheetName = "All Processed Data", highDollarSheetName = "High Dollar - Pat & Shelley";
            const sheetsData = { [masterSheetName]: [], [highDollarSheetName]: [] };
            
            const headerRow = main_aoa[0], newHeader = [...headerRow];
            if (hasYesterdayReport) newHeader.splice(config.claimStatusIndex, 0, 'Yest. Claim State');
            let daysInsertIndex = config.cleanAgeIndex + 1;
            if (hasYesterdayReport && config.cleanAgeIndex >= config.claimStatusIndex) daysInsertIndex++; 
            newHeader.splice(daysInsertIndex, 0, 'Days');
            newHeader.push('Added (Owner)', 'Due Date');
            sheetsData[masterSheetName].push(newHeader);
            sheetsData[highDollarSheetName].push(newHeader);

            for (const originalRow of main_aoa.slice(1)) {
                if (originalRow.every(cell => cell === null)) continue;
                let claimState = String(originalRow[config.claimStatusIndex] || '').trim().toUpperCase();
                if (claimState === 'PREBATCH') continue;
                const newRow = [...originalRow];
                
                if (hasYesterdayReport) {
                    const claimNumber = String(originalRow[config.claimNumberIndex] || '').trim();
                    const yestData = yesterdayDataMap.get(claimNumber);
                    let yestStateForColumn = 'NEW';
                    if (yestData) {
                        yestStateForColumn = yestData.state;
                        if (yestData.state === 'MANAGEMENTREVIEW' && !isNaN(yestData.charges) && ((yestData.type.includes('PROFESSIONAL') && yestData.charges > 3500) || (yestData.type.includes('INSTITUTIONAL') && yestData.charges > 6500))) {
                            yestStateForColumn = 'HC MGMT REV';
                        }
                    }
                    newRow.splice(config.claimStatusIndex, 0, yestStateForColumn);
                }
                
                const claimType = String(originalRow[config.claimTypeIndex] || '').trim().toUpperCase();
                const totalCharges = parseCurrency(originalRow[config.totalChargesIndex]);
                const isHighCost = claimState === 'MANAGEMENTREVIEW' && !isNaN(totalCharges) && ((claimType.includes('PROFESSIONAL') && totalCharges > 3500) || (claimType.includes('INSTITUTIONAL') && totalCharges > 6500));
                const cleanAge = parseInt(originalRow[config.cleanAgeIndex], 10);
                const daysValue = !isNaN(cleanAge) ? (cleanAge <= 20 ? '0 - 20' : (cleanAge <= 29 ? '21 - 29' : (cleanAge <= 59 ? '30 - 59' : '60+'))) : '';
                let ownerValue = '';

                if (isHighCost) { 
                    ownerValue = 'Pat and Shelley';
                } 
                else if (['MANAGEMENTREVIEW', 'ONHOLD'].includes(claimState)) { ownerValue = 'Jessica'; } 
                else if (['PEND', 'APPROVED', 'DENY'].includes(claimState)) { ownerValue = 'Patrick'; } 
                else if (claimState === 'PR') { ownerValue = originalRow[config.payerIndex] || ''; }

                newRow.splice(daysInsertIndex, 0, daysValue);
                const noteText = String(originalRow[config.notesIndex] || '');
                newRow.push(ownerValue, extractDueDateFromNote(noteText));
                sheetsData[masterSheetName].push(newRow);

                if (isHighCost) {
                    sheetsData[highDollarSheetName].push(newRow);
                } else if (['PEND', 'ONHOLD', 'MANAGEMENTREVIEW', 'DENY'].includes(claimState) && ['21 - 29', '30 - 59', '60+'].includes(daysValue)) {
                    const dsnpRaw = String(originalRow[config.dsnpIndex] || '').toUpperCase();
                    let dsnpStatus = '';
                    if (dsnpRaw.includes('NON DSNP')) { dsnpStatus = 'NonDSNP'; } 
                    else if (dsnpRaw.includes('DSNP') || dsnpRaw === 'Y') { dsnpStatus = 'DSNP'; }

                    let statusTab = '', tabOwner = '';
                    if (claimState === 'MANAGEMENTREVIEW') { statusTab = 'MgmtRev'; tabOwner = 'Jessica'; }
                    else if (claimState === 'ONHOLD') { statusTab = 'OnHold'; tabOwner = 'Jessica'; }
                    else if (claimState === 'PEND') { statusTab = 'Pend'; tabOwner = 'Patrick'; }
                    else if (claimState === 'DENY') { statusTab = 'Deny'; tabOwner = 'Patrick'; }

                    if (dsnpStatus && tabOwner) {
                        const ageTab = { '21 - 29': '21-29d', '30 - 59': '30-59d', '60+': '60+d' }[daysValue];
                        const tabKey = truncateSheetName(`${dsnpStatus} ${statusTab} ${ageTab} - ${tabOwner}`);
                        if (!sheetsData[tabKey]) sheetsData[tabKey] = [newHeader];
                        sheetsData[tabKey].push(newRow);
                    }
                }
                if (config.notesIndex >= 0 && noteText.toLowerCase().includes('w9')) {
                    let w9SheetName = '';
                    const noteLower = noteText.toLowerCase();
                    if (noteLower.includes('requested') || noteLower.includes('req') || noteLower.includes('due')) w9SheetName = 'W9 Follow-Up - Pat';
                    else if (noteLower.includes('denied') || noteLower.includes('missing') || noteLower.includes('not on file') || noteLower.includes('not received')) w9SheetName = 'W9 Letter Needed - Jess';
                    else if (noteLower.includes('received') || noteLower.includes('reprocess') || noteLower.includes('rerun')) w9SheetName = 'W9 Received - Reprocess';
                    if (w9SheetName) {
                        if (!sheetsData[w9SheetName]) sheetsData[w9SheetName] = [newHeader];
                        sheetsData[w9SheetName].push(newRow);
                    }
                }
            }
            
            const newWorkbook = XLSX.utils.book_new();
            const sheetOrder = [masterSheetName, highDollarSheetName, ...Object.keys(sheetsData).filter(n => n.startsWith('W9 ')).sort(), ...Object.keys(sheetsData).filter(n => ![masterSheetName, highDollarSheetName].includes(n) && !n.startsWith('W9 ')).sort()];
            sheetOrder.forEach(sheetName => {
                if (sheetsData[sheetName] && sheetsData[sheetName].length > 1) { 
                    const ws = XLSX.utils.aoa_to_sheet(sheetsData[sheetName]);
                    ws['!autofilter'] = { ref: XLSX.utils.encode_range(XLSX.utils.decode_range(ws['!ref'])) };
                    XLSX.utils.book_append_sheet(newWorkbook, ws, sheetName);
                }
            });
            
            const ageBasisText = document.getElementById('csr-cleanAgeCol-label').textContent.replace(':', '');
            const combinedMgmtRevToday = { ...todayStats['MANAGEMENTREVIEW'] };
            const hcToday = todayStats['HC MGMT REV'];
            Object.keys(hcToday).forEach(key => combinedMgmtRevToday[key] = (combinedMgmtRevToday[key] || 0) + hcToday[key]);
            let combinedMgmtRevYesterday = null;
            if (yesterdayStats) {
                combinedMgmtRevYesterday = { ...yesterdayStats['MANAGEMENTREVIEW'] };
                const hcYesterday = yesterdayStats['HC MGMT REV'] || { total: 0, '0 - 20': 0, '21 - 29': 0, '30 - 59': 0, '60+': 0 };
                Object.keys(hcYesterday).forEach(key => combinedMgmtRevYesterday[key] = (combinedMgmtRevYesterday[key] || 0) + hcYesterday[key]);
            }
            const formatStatLine = (t, y) => yesterdayStats ? `${t} (${y})` : `${t}`;
            const createStatBlock = (title, today, yest) => {
                const y = yesterdayStats ? yest : null;
                return `Number of total claims ${title}: ${formatStatLine(today.total, y?.total ?? 0)}\n0 - 20 Days: ${formatStatLine(today['0 - 20'], y?.['0 - 20'] ?? 0)}\n21 - 29 Days: ${formatStatLine(today['21 - 29'], y?.['21 - 29'] ?? 0)}\n30 - 59 Days: ${formatStatLine(today['30 - 59'], y?.['30 - 59'] ?? 0)}\n60+ Days: ${formatStatLine(today['60+'], y?.['60+'] ?? 0)}`;
            };
            const yestCombined = yesterdayStats ? combinedMgmtRevYesterday : null;
            const yestHc = yesterdayStats ? yesterdayStats['HC MGMT REV'] : null;
            const mgmtReviewBlock = `Number of total claims Management Review: ${formatStatLine(combinedMgmtRevToday.total, yestCombined?.total ?? 0)}\nOf these, High Dollar Claims: ${formatStatLine(todayStats['HC MGMT REV'].total, yestHc?.total ?? 0)}\n0 - 20 Days: ${formatStatLine(combinedMgmtRevToday['0 - 20'], yestCombined?.['0 - 20'] ?? 0)}\n21 - 29 Days: ${formatStatLine(combinedMgmtRevToday['21 - 29'], yestCombined?.['21 - 29'] ?? 0)}\n30 - 59 Days: ${formatStatLine(combinedMgmtRevToday['30 - 59'], yestCombined?.['30 - 59'] ?? 0)}\n60+ Days: ${formatStatLine(combinedMgmtRevToday['60+'], yestCombined?.['60+'] ?? 0)}`;
            const intro = yesterdayStats ? `I have attached today's report for ${clientText}. Below are the highlights from this report, previous days are in (parenthesis):` : `I have attached today's report for ${clientText}. Below are the highlights from this report:`;
            const emailBody = `Hello Shelley, Jessica, and Pat,\n\n${intro}\n\n${createStatBlock('pending', todayStats['PEND'], yesterdayStats?.['PEND'])}\n\n${createStatBlock('On Hold', todayStats['ONHOLD'], yesterdayStats?.['ONHOLD'])}\n\n${mgmtReviewBlock}\n\n${createStatBlock('needing W9 Letter', todayStats['W9_LETTER_NEEDED'], yesterdayStats?.['W9_LETTER_NEEDED'])}\n\n${createStatBlock('for W9 Follow-Up', todayStats['W9_FOLLOW_UP'], yesterdayStats?.['W9_FOLLOW_UP'])}\n\nPlease let me know if you have any questions regarding this information.\n\nThis ${clientText} report is based on ${ageBasisText}.\n`;

            modalTitle.textContent = 'Report Generated!';
            modalBody.innerHTML = `<p class="mb-4">Your file is downloading automatically. You can copy the summary email text below.</p><textarea class="w-full h-64 p-2 border rounded font-mono text-sm">${emailBody.trim()}</textarea>`;
            modalConfirmBtn.textContent = 'Close';
            modalConfirmBtn.onclick = hideModal;
            
            XLSX.writeFile(newWorkbook, `${clientText} Daily Action Report for ${getFormattedDate()}.xlsx`);
            showLoader(false);

        } catch (error) {
            console.error("Report generation error:", error);
            alert(`Error: ${error.message}`);
            showLoader(false);
            if (configModal.style.display === 'flex') hideModal();
        }
    }
    
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
});
