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

    // --- INITIALIZATION ---
    initializeCoreListeners();
    initializeValidator();

    function initializeCoreListeners() {
        fileUploadInput.addEventListener('change', handleFileUpload);
        downloadBtn.addEventListener('click', handleDownload);
        document.body.addEventListener('click', handleDynamicClicks);
        document.body.addEventListener('change', handleDynamicChanges);
        document.body.addEventListener('input', handleDynamicInputs);
    }

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
                // We only support single-sheet files in the UI for simplicity, but workbook is stored
                const firstSheetName = workbook.SheetNames[0];
                const jsonData = XLSX.utils.sheet_to_json(workbook.Sheets[firstSheetName]);
                state.datasets.push({ name: file.name, data: jsonData, headers: Object.keys(jsonData[0] || {}), workbook: workbook });

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
            item.dataset.action = 'select-dataset';
            item.dataset.index = index;
            if (index === state.activeDatasetIndex) {
                item.classList.add('bg-indigo-700', 'text-white', 'font-bold');
            } else {
                item.classList.add('text-slate-300', 'hover:bg-slate-700');
            }
            item.textContent = ds.name;
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
        if (onConfirm) {
            newConfirmBtn.addEventListener('click', onConfirm);
        }
        modalConfirmBtn = newConfirmBtn;
    }

    function generateColumnCheckboxes(headers) { return headers.map(h => `<label class="flex items-center p-2 rounded hover:bg-slate-100"><input type="checkbox" class="h-4 w-4 rounded mr-2" data-column-name="${h}"><span class="text-sm">${h}</span></label>`).join(''); }
    function generateColumnSelect(headers, id) { return `<select id="${id}" class="w-full p-2 border rounded mt-1">${headers.map(h => `<option value="${h}">${h}</option>`).join('')}</select>`; }
    
    // --- STANDARD ACTIONS ---
    function addNewDataset(name, data, headers) {
        state.datasets.push({ name, data, headers }); // Note: new datasets don't have a workbook
        state.activeDatasetIndex = state.datasets.length - 1;
        updateUI();
    }
    
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
    // --- EVENT DELEGATION FOR ALL DYNAMIC CLICKS / CHANGES ---
    // ===================================================================================
    function handleDynamicClicks(event) {
        const target = event.target;
        const actionTarget = target.closest('[data-action]');
        if (!actionTarget) return;

        const { action, target: targetId, index, headerKey, ruleIndex } = actionTarget.dataset;

        switch(action) {
            case 'select-dataset':
                state.activeDatasetIndex = parseInt(index, 10);
                updateUI();
                break;
            case 'close-config-modal':
                closeModal('config-modal');
                break;
            case 'close-modal':
                closeModal(targetId);
                break;
            case 'add-new-category':
                addNewRuleCategory();
                break;
            case 'add-rule-to-category':
                addRuleToCategory();
                break;
            case 'delete-category':
                deleteRuleCategory();
                break;
            case 'download-pdf-dictionary':
                downloadFullPdfDictionary();
                break;
            case 'toggle-error-details':
                const details = actionTarget.closest('li').querySelector('.error-details');
                const isHidden = details.classList.contains('hidden');
                details.classList.toggle('hidden', !isHidden);
                actionTarget.textContent = isHidden ? '(hide)' : '(show)';
                break;
            case 'add-rule-to-inline-category':
                addRuleToInlineCategory(actionTarget);
                break;
            case 'remove-rule-pill':
                actionTarget.parentElement.remove();
                break;
            case 'delete-rule':
                deleteRule(headerKey, parseInt(ruleIndex, 10));
                break;
            case 'delete-column':
                deleteColumn(headerKey);
                break;
            case 'add-rule':
                addRule(actionTarget, headerKey);
                break;
            case 'select-category-to-edit':
                selectRuleCategoryToEdit(headerKey);
                break;
        }
    }

    function handleDynamicChanges(event) {
        const target = event.target;
        const { action, headerKey } = target.dataset;
        
        switch(action) {
            case 'update-description':
                updateDescription(headerKey, target.value);
                break;
            case 'apply-rule-category':
                applyRuleCategory(headerKey, target.value);
                break;
            case 'handle-category-choice':
                document.getElementById(`inline-editor-${target.dataset.index}`).classList.toggle('hidden', target.value !== '__CREATE_NEW__');
                break;
            case 'copy-settings':
                copySettingsFromExisting(target);
                break;
            case 'toggle-ignore':
                toggleIgnore(target.dataset.key);
                break;
        }
    }
    
    function handleDynamicInputs(event) {
        if (event.target.classList.contains('description-input')) {
            autoAdjustTextarea(event.target);
        }
    }
    
    // ===================================================================================
    // --- ROBUST VALIDATOR LOGIC ---
    // ===================================================================================
    function initializeValidator() {
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

    function closeModal(id) { document.getElementById(id).classList.add('hidden'); }

    function openDictionaryEditor() {
        const activeDS = getActiveDataset();
        const headers = activeDS ? activeDS.headers : [];
        document.getElementById('builder_dictionaryName').value = state.dictionary.dictionaryName;
        renderBuilderTable(headers);
        document.getElementById('dictionaryModal').classList.remove('hidden');
    }

    function openRuleCategoryEditor() {
        state.currentEditingCategory = null;
        document.getElementById('ruleCategoryEditorContainer').classList.add('hidden');
        renderRuleCategoryList();
        document.getElementById('ruleCategoryModal').classList.remove('hidden');
    }

    function downloadFullPdfDictionary() {
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

    function handleRestoreUpload(event) {
        const file = event.target.files[0];
        if (!file) return;
        const reader = new FileReader();
        reader.onload = (e) => {
            try {
                const backupData = JSON.parse(e.target.result);
                const dictionaryToRestore = backupData.dictionary || backupData;
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
        dataView.classList.remove('hidden');
        
        state.lastValidationResult = null; 
        state.ignoredErrors.clear();
        const activeDS = getActiveDataset();
        if (!activeDS) { alert('Please load a file first.'); return; }
        
        showLoader(true);
        
        setTimeout(() => {
            const fileHeaders = activeDS.headers;
            const dictHeaders = new Set(Object.values(state.dictionary.sheetsData).map(item => item["Column Name"].toUpperCase()));
            const missingHeaders = fileHeaders.filter(header => !dictHeaders.has(header.toUpperCase()));
            
            if (missingHeaders.length > 0) {
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
        let schemaHTML = `<div class="bg-white p-6 md:p-8 rounded-xl shadow-lg h-full flex flex-col"><div class="p-4 border-l-4 border-yellow-400 bg-yellow-50 rounded-r-lg"><h3 class="text-lg font-bold text-yellow-800">New Columns Detected</h3><p class="text-sm text-yellow-700 mt-2">Define rules for new columns below. Columns without a description will be excluded.</p></div><div class="mt-4 space-y-4 flex-grow overflow-y-auto">`;
        const categoryOptions = Object.keys(state.dictionary.ruleCategories || {}).sort().map(catName => `<option value="${catName}">${catName}</option>`).join('');
        missingHeaders.forEach((header, index) => { schemaHTML += `<div class="p-3 bg-slate-50 rounded-md" data-missing-header="${header}"><div class="grid grid-cols-1 md:grid-cols-2 gap-x-6 gap-y-4 items-start"><div><label class="block text-sm font-semibold text-slate-800">${header}</label><select data-action="handle-category-choice" data-index="${index}" class="category-select w-full mt-1 p-2 border rounded-md"><option value="">-- Select a Category --</option><option value="__CREATE_NEW__" class="font-bold text-blue-600">** Create New **</option>${categoryOptions}</select></div><div><label class="block text-sm font-medium text-gray-700">Description</label><textarea class="description-input w-full mt-1 p-2 border rounded-md" placeholder="Enter description..."></textarea><select data-action="copy-settings" class="w-full mt-2 p-1.5 border rounded-md bg-gray-50 text-sm"><option value="">-- Or, copy settings from existing --</option>${copyOptions}</select></div></div><div id="inline-editor-${index}" class="inline-editor hidden mt-4 p-4 border-t border-slate-200"><div class="mb-2"><label class="block text-sm font-medium">New Category Name:</label><input type="text" class="new-category-name w-full p-2 border rounded-md" placeholder="Enter Unique Name..."></div><div class="mb-2"><label class="block text-sm font-medium">Defined Rules:</label><div class="inline-rules-display flex flex-wrap gap-2 p-2 min-h-[40px] border rounded-md bg-white"></div></div><div><label class="block text-sm font-medium">Add New Rule:</label><div class="flex gap-2 items-center"><select class="inline-rule-type p-1 border rounded-md bg-white flex-1"><option value="REQUIRED">REQUIRED</option><option value="NOT REQUIRED">NOT REQUIRED</option><option value="ALLOWED_VALUES">ALLOWED_VALUES</option><option value="REGEX">REGEX</option><option value="VALID_DATE">VALID_DATE</option><option value="NO_FUTURE_DATE">NO_FUTURE_DATE</option></select><input type="text" placeholder="Value..." class="inline-rule-value p-1 border rounded-md flex-1"><button type="button" data-action="add-rule-to-inline-category" class="py-1 px-3 bg-blue-500 text-white rounded-md text-sm">+</button></div></div></div></div>`; });
        schemaHTML += `</div><div class="text-center mt-6 flex-shrink-0"><button id="workflow_proceedBtn" class="px-6 py-2 bg-green-600 hover:bg-green-700 text-white font-bold rounded-md">Apply & Continue</button></div></div>`;
        document.getElementById('validation-results-view').innerHTML = schemaHTML;
        dataView.classList.add('hidden');
        document.getElementById('validation-results-view').classList.remove('hidden');
        document.getElementById('workflow_proceedBtn').addEventListener('click', handleMappingProceed);
    }
    function copySettingsFromExisting(selectElement) { const selectedColumnName = selectElement.value; if (!selectedColumnName) return; const columnData = Object.values(state.dictionary.sheetsData).find(col => col['Column Name'] === selectedColumnName); if (columnData) { const container = selectElement.closest('[data-missing-header]'); const descriptionInput = container.querySelector('.description-input'); const categorySelect = container.querySelector('.category-select'); descriptionInput.value = columnData.description || ''; autoAdjustTextarea(descriptionInput); if (columnData.category && categorySelect.querySelector(`option[value="${columnData.category}"]`)) { categorySelect.value = columnData.category; const inlineEditor = container.querySelector('.inline-editor'); if (inlineEditor) { inlineEditor.classList.add('hidden'); } } } selectElement.value = ''; }
    function addRuleToInlineCategory(button) { const editor = button.closest('.inline-editor'); const type = editor.querySelector('.inline-rule-type').value; const valueInput = editor.querySelector('.inline-rule-value'); const value = valueInput.value.trim(); if (['ALLOWED_VALUES', 'REGEX'].includes(type) && !value) { alert("Value is required for this rule type."); return; } const display = editor.querySelector('.inline-rules-display'); const pill = document.createElement('span'); pill.className = 'rule-pill'; pill.dataset.type = type; pill.dataset.value = value; pill.innerHTML = `${type}: ${value || 'N/A'} <button type="button" data-action="remove-rule-pill" class="delete-btn ml-2">x</button>`; display.appendChild(pill); valueInput.value = ''; }
    function handleMappingProceed() { const assignments = []; const newCategories = {}; const allCategoryNames = new Set(Object.keys(state.dictionary.ruleCategories)); let hasError = false; const skippedColumns = []; document.querySelectorAll('[data-missing-header]').forEach(div => { const header = div.dataset.missingHeader; const categorySelect = div.querySelector('.category-select'); const descriptionInput = div.querySelector('.description-input'); if (categorySelect.value) { const description = descriptionInput.value.trim(); if (!description) { skippedColumns.push(header); return; } const assignment = { header, description }; if (categorySelect.value === '__CREATE_NEW__') { const editor = div.querySelector('.inline-editor'); const newCatName = editor.querySelector('.new-category-name').value.trim(); if (!newCatName) { alert(`Please enter a name for the new category for column "${header}".`); hasError = true; return; } if (allCategoryNames.has(newCatName)) { alert(`The category name "${newCatName}" already exists. Please choose a unique name.`); hasError = true; return; } assignment.category = newCatName; const rules = []; editor.querySelectorAll('.inline-rules-display .rule-pill').forEach(pill => { rules.push({ type: pill.dataset.type, value: pill.dataset.value, message: `Validation failed for rule ${pill.dataset.type}` }); }); newCategories[newCatName] = { rules }; allCategoryNames.add(newCatName); } else { assignment.category = categorySelect.value; } assignments.push(assignment); } else { skippedColumns.push(header); } }); if (hasError) return; const continueProcessing = () => { if (Object.keys(newCategories).length > 0) { state.dictionary.ruleCategories = { ...state.dictionary.ruleCategories, ...newCategories }; } state.pendingAssignments = assignments; state.lastValidationResult = { ...(state.lastValidationResult || {}), skippedColumns: skippedColumns }; applyPendingAssignments(); saveRobustDictionary(); commitMappingsAndValidate(); }; if (skippedColumns.length > 0) { if (confirm(`Some columns have no description and will be excluded. Continue?`)) { continueProcessing(); } } else { continueProcessing(); } }
    function applyPendingAssignments() { state.pendingAssignments.forEach(assignment => { const { header, category, description } = assignment; const headerKey = header.toUpperCase(); state.dictionary.sheetsData[headerKey] = { "Column Name": header, category: category, description: description, validation_rules: [] }; }); state.pendingAssignments = []; }
    function commitMappingsAndValidate() { document.getElementById('validation-results-view').classList.add('hidden'); showLoader(true); setTimeout(() => { runDataValidation(getActiveDataset(), state.dictionary); showLoader(false); }, 50); }
    function renderBuilderTable(headers) { const tbody = document.querySelector('#rulesTable tbody'); tbody.innerHTML = ''; const dictData = state.dictionary.sheetsData; const categoryOptions = Object.keys(state.dictionary.ruleCategories).sort().map(catName => `<option value="${catName}">${catName}</option>`).join(''); const allHeaders = [...new Set(headers.concat(Object.values(dictData).map(d => d["Column Name"])))].filter(Boolean).sort((a,b) => a.localeCompare(b)); allHeaders.forEach(header => { const headerKey = header.toUpperCase(); const columnData = Object.values(dictData).find(d => d["Column Name"].toUpperCase() === headerKey); if(!columnData) return; const row = tbody.insertRow(); let rulesToDisplay = []; if (columnData.category && state.dictionary.ruleCategories[columnData.category]) { rulesToDisplay = state.dictionary.ruleCategories[columnData.category].rules; } else if (columnData.validation_rules) { rulesToDisplay = columnData.validation_rules; } let rulesHTML = '<div class="flex flex-wrap gap-1">'; if(rulesToDisplay) { rulesToDisplay.forEach((rule, index) => { const deleteAction = columnData.category ? '' : `<button class="delete-btn ml-2" data-action="delete-rule" data-header-key="${headerKey}" data-rule-index="${index}">x</button>`; rulesHTML += `<span class="rule-pill">${rule.type}: ${rule.value || 'N/A'} ${deleteAction}</span>`; }); } rulesHTML += '</div>'; const addRuleHTML = `<div class="flex gap-2 items-center"><select class="p-1 border rounded-md bg-white w-1/3"><option value="REQUIRED">REQUIRED</option><option value="NOT REQUIRED">NOT REQUIRED</option><option value="ALLOWED_VALUES">ALLOWED_VALUES</option><option value="REGEX">REGEX</option><option value="VALID_DATE">VALID_DATE</option><option value="NO_FUTURE_DATE">NO_FUTURE_DATE</option></select><input type="text" placeholder="Value..." class="p-1 border rounded-md w-1/3"><button data-action="add-rule" data-header-key="${headerKey}" class="py-1 px-3 bg-blue-500 text-white rounded-md text-sm">+</button></div>`; const categoryHTML = `<td><select data-action="apply-rule-category" data-header-key="${headerKey}" class="p-1 border rounded-md w-full"><option value="">-- None --</option>${categoryOptions}</select></td>`; const descriptionHTML = `<td><textarea data-action="update-description" data-header-key="${headerKey}" class="description-input p-1 border rounded-md w-full">${columnData.description || ''}</textarea></td>`; row.innerHTML = `<td class="font-semibold align-middle">${header}<button class="delete-btn ml-4" data-action="delete-column" data-header-key="${headerKey}">Delete</button></td>${categoryHTML}${descriptionHTML}<td>${rulesHTML}</td><td>${addRuleHTML}</td>`; row.querySelector('select').value = columnData.category || ''; autoAdjustTextarea(row.querySelector('textarea')); }); }
    function applyRuleCategory(headerKey, categoryName) { const columnData = Object.values(state.dictionary.sheetsData).find(d=>d["Column Name"].toUpperCase() === headerKey); if (!columnData) return; if (categoryName === "") { columnData.validation_rules = []; columnData.category = ""; } else { const categoryTemplate = state.dictionary.ruleCategories[categoryName]; if (categoryTemplate && confirm(`This will replace existing rules for "${columnData['Column Name']}" with the "${categoryName}" category. Are you sure?`)) { columnData.validation_rules = []; columnData.category = categoryName; } } renderBuilderTable(Object.values(state.dictionary.sheetsData).map(r => r["Column Name"])); }
    function updateDescription(headerKey, newDescription) { const col = Object.values(state.dictionary.sheetsData).find(d=>d["Column Name"].toUpperCase() === headerKey); if (col) col.description = newDescription; }
    function addRule(buttonElement, headerKey) { const col = Object.values(state.dictionary.sheetsData).find(d=>d["Column Name"].toUpperCase() === headerKey); if(!col) return; const addRuleContainer = buttonElement.closest('div'); const select = addRuleContainer.querySelector('select'); const input = addRuleContainer.querySelector('input[type="text"]'); const ruleType = select.value, ruleValue = input.value.trim(); if (['ALLOWED_VALUES', 'REGEX'].includes(ruleType) && !ruleValue) { alert('Rule Value is required.'); return; } if(!col.validation_rules) col.validation_rules = []; col.validation_rules.push({ type: ruleType, value: ruleValue, message: `Validation failed for rule ${ruleType}` }); col.category = ''; renderBuilderTable(Object.values(state.dictionary.sheetsData).map(r => r["Column Name"])); }
    function deleteRule(headerKey, ruleIndex) { const col = Object.values(state.dictionary.sheetsData).find(d=>d["Column Name"].toUpperCase() === headerKey); if(col) { col.validation_rules.splice(ruleIndex, 1); col.category = ''; renderBuilderTable(Object.values(state.dictionary.sheetsData).map(r => r["Column Name"])); } }
    function deleteColumn(headerKey) { const colData = Object.values(state.dictionary.sheetsData).find(d=>d["Column Name"].toUpperCase() === headerKey); if(colData && confirm(`Delete the column "${colData["Column Name"]}"?`)){ const keyToDelete = Object.keys(state.dictionary.sheetsData).find(k => k.toUpperCase() === headerKey); if(keyToDelete) delete state.dictionary.sheetsData[keyToDelete]; renderBuilderTable(Object.values(state.dictionary.sheetsData).map(r => r["Column Name"])); } }
    function saveAndCloseBuilder() { state.dictionary.dictionaryName = document.getElementById('builder_dictionaryName').value; saveRobustDictionary(); closeModal('dictionaryModal'); }
    function handleCategoryEditorClose() { saveRobustDictionary(); closeModal('ruleCategoryModal'); if (state.pendingAssignments.length > 0) { applyPendingAssignments(); saveRobustDictionary(); commitMappingsAndValidate(); } }
    function renderRuleCategoryList() { const listEl = document.getElementById('ruleCategoryList'); listEl.innerHTML = ''; const sortedCategories = Object.keys(state.dictionary.ruleCategories).sort(); sortedCategories.forEach(catName => { const item = document.createElement('div'); item.className = 'category-list-item'; item.textContent = catName; item.dataset.action = 'select-category-to-edit'; item.dataset.headerKey = catName; if (catName === state.currentEditingCategory) item.classList.add('selected'); listEl.appendChild(item); }); }
    function selectRuleCategoryToEdit(categoryName) { state.currentEditingCategory = categoryName; renderRuleCategoryList(); const editorEl = document.getElementById('ruleCategoryEditorContainer'); editorEl.classList.remove('hidden'); document.getElementById('categoryEditorTitle').textContent = `Editing: ${categoryName}`; renderCategoryRulesDisplay(); }
    function renderCategoryRulesDisplay() { const displayEl = document.getElementById('categoryRulesDisplay'); displayEl.innerHTML = ''; const category = state.dictionary.ruleCategories[state.currentEditingCategory]; if (!category || !category.rules) return; category.rules.forEach((rule, index) => { const pill = document.createElement('span'); pill.className = 'rule-pill'; pill.innerHTML = `${rule.type}: ${rule.value || 'N/A'} <button class="delete-btn ml-2" data-action="remove-rule-pill" data-rule-index="${index}">x</button>`; displayEl.appendChild(pill); }); }
    function addNewRuleCategory() { const name = prompt("Enter a name for the new rule category:"); if (name && !state.dictionary.ruleCategories[name]) { state.dictionary.ruleCategories[name] = { rules: [] }; selectRuleCategoryToEdit(name); } else if (name) { alert("A category with this name already exists."); } }
    function addRuleToCategory() { const catName = state.currentEditingCategory; if (!catName) return; const type = document.getElementById('categoryRuleType').value; const valueInput = document.getElementById('categoryRuleValue'); const value = valueInput.value.trim(); if (['ALLOWED_VALUES', 'REGEX'].includes(type) && !value) { alert("Value is required for this rule type."); return; } state.dictionary.ruleCategories[catName].rules.push({ type, value, message: `Validation failed for rule ${type}` }); valueInput.value = ''; renderCategoryRulesDisplay(); }
    function deleteRuleCategory() { const catName = state.currentEditingCategory; if (catName && confirm(`Delete the "${catName}" category?`)) { delete state.dictionary.ruleCategories[catName]; state.currentEditingCategory = null; document.getElementById('ruleCategoryEditorContainer').classList.add('hidden'); renderRuleCategoryList(); } }
    function parseDate(value) { if (value === null || value === undefined || String(value).trim() === '') return null; if (typeof value === 'number' && value > 1 && value <= 2958465) { return new Date((value - 25569) * 86400 * 1000); } const date = new Date(String(value)); return isNaN(date.getTime()) ? null : date; }
    
    // The rest of the functions (runDataValidation, displayValidationReport, etc.) are the same as the previous correct version...
    // They are included here for completeness.

    function runDataValidation(activeDS, dictionary) {
        const { workbook, name: fileName } = activeDS;
        const analysisResults = {};
        const overallStats = { customIssueCount: 0, duplicateRowCount: 0, totalProcessedCells: 0, totalProcessedRows: 0, totalOriginalRows: 0 };
        const duplicateMethod = "thorough"; // This can be made a UI option later

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
                            const matchPercentage = (matchingCells / Math.max(rowA.length, 1)) * 100;
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
                        let message = rule.message || `Validation failed for rule ${rule.type}`;

                        switch (rule.type) {
                            case 'REQUIRED': isValid = !isCellEmpty; break;
                            case 'ALLOWED_VALUES_COMBINED': isValid = isCellEmpty || rule.value.has(String(cellValue).trim().toLowerCase()); break;
                            case 'REGEX': try { isValid = isCellEmpty || new RegExp(rule.value, 'i').test(String(cellValue)); } catch(e) { isValid = false; message="Invalid REGEX pattern in dictionary.";} break;
                            case 'VALID_DATE':
                                if (isCellEmpty) { isValid = true; break; }
                                isValid = parseDate(cellValue) !== null;
                                if(!isValid) message = "Value is not a recognizable date.";
                                break;
                            case 'NO_FUTURE_DATE':
                                if (isCellEmpty) { isValid = true; break; }
                                const parsedDate = parseDate(cellValue);
                                if (!parsedDate) {
                                    isValid = false; message = "Value is not a recognizable date.";
                                } else {
                                    const today = new Date();
                                    today.setHours(23, 59, 59, 999);
                                    if (parsedDate > today) {
                                        isValid = false; message = `Date ${parsedDate.toLocaleDateString()} is in the future.`;
                                    } else {
                                        isValid = true;
                                    }
                                }
                                break;
                        }
                        if (!isValid) {
                            currentSheetIssues.customValidation[header].push({ row: rowIndex + 2, value: cellValue, type: rule.type.replace('_COMBINED', ''), message });
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
        dataView.classList.add('hidden');
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
                html += `<li><div class="flex items-center"><label class="inline-flex items-center cursor-pointer"><input type="checkbox" data-action="toggle-ignore" data-key="${issueKey}" class="form-checkbox h-4 w-4 mr-2"><strong>Similar Rows Found:</strong> ${issues.duplicateRows.length} pairs.</label><button data-action="toggle-error-details" class="ml-2 text-blue-500 text-xs">(show)</button></div><div class="error-details hidden">${glimpseHTML}</div></li>`;
            }

            for (const col in issues.customValidation) {
                if (issues.customValidation[col].length > 0) {
                    hasIssues = true;
                    const issueKey = `val_${sheetName}_${col}`;
                    const colI = issues.customValidation[col];
                    html += `<li><div class="flex items-center"><label class="inline-flex items-center cursor-pointer"><input type="checkbox" data-action="toggle-ignore" data-key="${issueKey}" class="form-checkbox h-4 w-4 mr-2"><strong>Column "${col}":</strong> ${colI.length} issues.</label><button data-action="toggle-error-details" class="ml-2 text-blue-500 text-xs">(show)</button></div><div class="error-details hidden"><ul class="list-disc list-inside">${colI.slice(0, 50).map(i => `<li>Row ${i.row}: Failed '${i.type}' (Value: "${i.value}"). Message: ${i.message}</li>`).join('')}</ul></div></li>`;
                }
            }
            if (!hasIssues) html += `<li class="text-green-700 font-semibold">No issues found in this sheet.</li>`;
            html += `</ul></div>`;
        }
        html += `</div><div class="flex flex-col md:flex-row gap-4 mt-6 border-t pt-6"><button id="downloadReportBtn" class="flex-1 py-3 px-4 rounded-lg text-lg font-medium text-white bg-green-600 hover:bg-green-700">Download Scored Report</button><button id="downloadSubsetPdfBtn" class="flex-1 py-3 px-4 rounded-lg text-lg font-medium text-white bg-gray-600 hover:bg-gray-700">Download Subset PDF</button></div></div>`;
        view.innerHTML = html;
        document.getElementById('downloadReportBtn').addEventListener('click', downloadScoredReport);
        document.getElementById('downloadSubsetPdfBtn').addEventListener('click', downloadSubsetPdf);
        recalculateScores();
    }
    function autoAdjustTextarea(element) { element.style.height = 'auto'; element.style.height = (element.scrollHeight) + 'px'; }
    function toggleIgnore(key){if(state.ignoredErrors.has(key)){state.ignoredErrors.delete(key)}else{state.ignoredErrors.add(key)}document.querySelector(`input[data-key="${key}"]`).closest('li').classList.toggle('issue-ignored',state.ignoredErrors.has(key));recalculateScores()}
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
    
    function downloadSubsetPdf() {
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

});
