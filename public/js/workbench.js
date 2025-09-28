document.addEventListener('DOMContentLoaded', () => {
    // --- STATE MANAGEMENT ---
    const state = {
        datasets: [],
        activeDatasetIndex: 0,
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
    const modalConfirmBtn = document.getElementById('modal-confirm-btn');
    const modalCancelBtn = document.getElementById('modal-cancel-btn');
    const modalCloseBtn = document.getElementById('modal-close-btn');

    // --- INITIALIZATION ---
    fileUploadInput.addEventListener('change', handleFileUpload);
    downloadBtn.addEventListener('click', handleDownload);
    [modalCancelBtn, modalCloseBtn].forEach(btn => btn.addEventListener('click', hideModal));

    // --- FILE HANDLING ---
    async function handleFileUpload(event) {
        const files = Array.from(event.target.files);
        if (files.length === 0) return;

        showLoader(true, 'Reading files...');
        state.datasets = []; // Clear previous datasets

        for (const file of files) {
            try {
                const data = await readFile(file);
                const workbook = XLSX.read(data, { type: 'array', cellDates: true });
                
                // If only one sheet, add it directly.
                if (workbook.SheetNames.length === 1) {
                    const sheetName = workbook.SheetNames[0];
                    const jsonData = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName]);
                    state.datasets.push({
                        name: file.name,
                        data: jsonData,
                        headers: Object.keys(jsonData[0] || {}),
                    });
                } else {
                    // If multiple sheets, add each one as a separate dataset.
                    workbook.SheetNames.forEach(sheetName => {
                        const jsonData = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName]);
                        state.datasets.push({
                            name: `${file.name} - ${sheetName}`,
                            data: jsonData,
                            headers: Object.keys(jsonData[0] || {}),
                        });
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

    // --- UI RENDERING ---
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
            item.onclick = () => {
                state.activeDatasetIndex = index;
                updateUI();
            };
            loadedFilesList.appendChild(item);
        });
    }

    function renderActiveDataset() {
        const activeDataset = state.datasets[state.activeDatasetIndex];
        if (!activeDataset) return;

        tableTitle.textContent = activeDataset.name;
        renderDataTable(activeDataset.data, activeDataset.headers);
        statusBar.textContent = `Displaying ${activeDataset.data.length.toLocaleString()} rows and ${activeDataset.headers.length} columns. (Preview of first 200 rows)`;
    }

    function renderDataTable(data, headers) {
        const table = document.createElement('table');
        const thead = table.createTHead();
        const headerRow = thead.insertRow();
        headers.forEach(header => {
            const th = document.createElement('th');
            th.textContent = header;
            headerRow.appendChild(th);
        });

        const tbody = table.createTBody();
        const sampleData = data.slice(0, 200);
        sampleData.forEach(row => {
            const tr = tbody.insertRow();
            headers.forEach(header => {
                const td = tr.insertCell();
                td.textContent = row[header] ?? '';
            });
        });
        
        tableContainer.innerHTML = '';
        tableContainer.appendChild(table);
    }
    
    function showLoader(show, message = '') {
        if (show) {
            loaderOverlay.style.display = 'flex';
            if (message) {
                // You can add a message element to the overlay if desired
            }
        } else {
            loaderOverlay.style.display = 'none';
        }
    }

    // --- MODAL & CONFIGURATION ---
    function showModal(title, content, onConfirm) {
        modalTitle.textContent = title;
        modalBody.innerHTML = content;
        configModal.style.display = 'flex';
        // Clone and replace the button to remove old event listeners
        const newConfirmBtn = modalConfirmBtn.cloneNode(true);
        modalConfirmBtn.parentNode.replaceChild(newConfirmBtn, modalConfirmBtn);
        newConfirmBtn.addEventListener('click', onConfirm);
        // Re-assign the global reference
        window.modalConfirmBtn = newConfirmBtn;
    }

    function hideModal() {
        configModal.style.display = 'none';
    }

    function generateColumnCheckboxes(headers) {
        return headers.map(h => `
            <label class="flex items-center p-2 rounded hover:bg-gray-100">
                <input type="checkbox" class="h-4 w-4 rounded mr-2" data-column-name="${h}">
                <span class="text-sm">${h}</span>
            </label>
        `).join('');
    }

    function generateColumnSelect(headers, id) {
        return `<select id="${id}" class="w-full p-2 border rounded mt-1">${headers.map(h => `<option value="${h}">${h}</option>`).join('')}</select>`;
    }
    
    function generateDatasetSelect(id) {
        return `<select id="${id}" class="w-full p-2 border rounded mt-1">${state.datasets.map((ds, i) => `<option value="${i}">${ds.name}</option>`).join('')}</select>`;
    }

    // --- ACTION IMPLEMENTATIONS ---
    function getActiveDataset() {
        return state.datasets[state.activeDatasetIndex];
    }
    
    function addNewDataset(name, data, headers) {
        state.datasets.push({ name, data, headers });
        state.activeDatasetIndex = state.datasets.length - 1;
        updateUI();
    }
    
    // --- Attach Event Listeners to Action Buttons ---

    document.getElementById('action-trim-whitespace').addEventListener('click', () => {
        const headers = getActiveDataset().headers;
        const content = `
            <p class="text-sm mb-4">This will remove leading/trailing spaces from every cell in the selected column.</p>
            <label for="trim-column" class="block text-sm font-semibold">Select Column:</label>
            ${generateColumnSelect(headers, 'config-column')}
        `;
        showModal('Trim Whitespace', content, () => {
            const column = document.getElementById('config-column').value;
            showLoader(true);
            setTimeout(() => {
                const activeDataset = getActiveDataset();
                activeDataset.data.forEach(row => {
                    if (typeof row[column] === 'string') {
                        row[column] = row[column].trim();
                    }
                });
                renderActiveDataset();
                showLoader(false);
                hideModal();
            }, 50);
        });
    });

    document.getElementById('action-anonymize').addEventListener('click', () => {
        const headers = getActiveDataset().headers;
        const anonymizationTypes = [
            { value: 'NONE', label: 'Do Not Anonymize' },
            { value: 'FULL_NAME', label: 'Full Name' },
            { value: 'FIRST_NAME', label: 'First Name Only' },
            { value: 'LAST_NAME', label: 'Last Name Only' },
            { value: 'EMAIL', label: 'Email Address' },
            { value: 'PHONE', label: 'Phone Number' },
        ];
        const content = headers.map(h => `
            <div class="grid grid-cols-2 gap-4 items-center border-b pb-2 mb-2">
                <label class="font-semibold text-gray-700 truncate" title="${h}">${h}</label>
                <select data-header="${h}" class="column-mapper w-full p-2 border rounded-md">
                    ${anonymizationTypes.map(opt => `<option value="${opt.value}">${opt.label}</option>`).join('')}
                </select>
            </div>
        `).join('');
        showModal('Anonymize Personal Information (PII)', content, () => {
             const mappings = Array.from(document.querySelectorAll('.column-mapper'))
                .filter(select => select.value !== 'NONE')
                .map(select => ({ header: select.dataset.header, type: select.value }));
            
            if (mappings.length === 0) return alert('Please select at least one column to anonymize.');

            showLoader(true);
            setTimeout(() => {
                const fakeData = {
                    FIRST_NAME: ['Alex', 'Jordan', 'Casey', 'Taylor', 'Morgan', 'Skyler', 'Riley', 'Peyton', 'Jamie', 'Quinn'],
                    LAST_NAME: ['Smith', 'Johnson', 'Williams', 'Brown', 'Jones', 'Garcia', 'Miller', 'Davis', 'Rodriguez', 'Martinez'],
                    FULL_NAME: () => `${fakeData.FIRST_NAME[Math.floor(Math.random() * 10)]} ${fakeData.LAST_NAME[Math.floor(Math.random() * 10)]}`,
                    EMAIL: () => `user${Math.floor(1000 + Math.random() * 9000)}@example.com`,
                    PHONE: () => `(555) ${Math.floor(100 + Math.random() * 900)}-${Math.floor(1000 + Math.random() * 9000)}`
                };

                const activeDataset = getActiveDataset();
                const anonymizedData = activeDataset.data.map(row => {
                    const newRow = { ...row };
                    mappings.forEach(map => {
                        if (newRow[map.header] !== undefined) {
                            if (typeof fakeData[map.type] === 'function') {
                                newRow[map.header] = fakeData[map.type]();
                            } else {
                                newRow[map.header] = fakeData[map.type][Math.floor(Math.random() * 10)];
                            }
                        }
                    });
                    return newRow;
                });
                
                addNewDataset(`Anonymized - ${activeDataset.name}`, anonymizedData, activeDataset.headers);
                showLoader(false);
                hideModal();
            }, 50);
        });
    });

    document.getElementById('action-extract-columns').addEventListener('click', () => {
        const headers = getActiveDataset().headers;
        const content = `
            <p class="text-sm mb-4">Select the columns you want to keep. A new dataset will be created.</p>
            <div class="space-y-2">${generateColumnCheckboxes(headers)}</div>
        `;
        showModal('Extract Columns', content, () => {
            const selectedColumns = Array.from(modalBody.querySelectorAll('input:checked')).map(cb => cb.dataset.columnName);
            if (selectedColumns.length === 0) return alert('Please select at least one column.');
            
            showLoader(true);
            setTimeout(() => {
                const activeDataset = getActiveDataset();
                const newData = activeDataset.data.map(row => {
                    const newRow = {};
                    selectedColumns.forEach(col => {
                        newRow[col] = row[col];
                    });
                    return newRow;
                });
                addNewDataset(`Extracted - ${activeDataset.name}`, newData, selectedColumns);
                showLoader(false);
                hideModal();
            }, 50);
        });
    });

    document.getElementById('action-find-duplicates').addEventListener('click', () => {
        const headers = getActiveDataset().headers;
        const content = `
            <p class="text-sm mb-4">Select one or more columns to check for duplicate rows. A new dataset will be created with only the duplicate entries.</p>
            <div class="space-y-2">${generateColumnCheckboxes(headers)}</div>
        `;
        showModal('Find Duplicates', content, () => {
            const selectedColumns = Array.from(modalBody.querySelectorAll('input:checked')).map(cb => cb.dataset.columnName);
            if (selectedColumns.length === 0) return alert('Please select at least one column.');
            
            showLoader(true);
            setTimeout(() => {
                const activeDataset = getActiveDataset();
                const seen = new Map();
                const duplicates = [];
                activeDataset.data.forEach(row => {
                    const key = selectedColumns.map(col => row[col]).join('||');
                    if (seen.has(key)) {
                        // If it's the first time we've seen this duplicate, add the original row too
                        if (seen.get(key).first) {
                            duplicates.push(seen.get(key).row);
                            seen.get(key).first = false;
                        }
                        duplicates.push(row);
                    } else {
                        seen.set(key, { row: row, first: true });
                    }
                });
                addNewDataset(`Duplicates - ${activeDataset.name}`, duplicates, activeDataset.headers);
                showLoader(false);
                hideModal();
            }, 50);
        });
    });
    
    document.getElementById('action-merge-files').addEventListener('click', () => {
        if (state.datasets.length < 2) return alert("You need to upload at least two files to use the merge feature.");
        
        const content = `
            <div class="grid grid-cols-2 gap-4">
                <div>
                    <label class="block text-sm font-semibold">Left Table (Primary)</label>
                    ${generateDatasetSelect('config-ds1')}
                    <label class="block text-sm font-semibold mt-2">Key Column</label>
                    <select id="config-key1" class="w-full p-2 border rounded mt-1"></select>
                </div>
                <div>
                    <label class="block text-sm font-semibold">Right Table (to join)</label>
                    ${generateDatasetSelect('config-ds2')}
                    <label class="block text-sm font-semibold mt-2">Key Column</label>
                    <select id="config-key2" class="w-full p-2 border rounded mt-1"></select>
                </div>
            </div>
        `;

        function populateKeys() {
            const ds1_idx = document.getElementById('config-ds1').value;
            const ds2_idx = document.getElementById('config-ds2').value;
            document.getElementById('config-key1').innerHTML = state.datasets[ds1_idx].headers.map(h => `<option value="${h}">${h}</option>`).join('');
            document.getElementById('config-key2').innerHTML = state.datasets[ds2_idx].headers.map(h => `<option value="${h}">${h}</option>`).join('');
        }

        showModal('Merge Files (Left Join)', content, () => {
            const ds1_idx = document.getElementById('config-ds1').value;
            const ds2_idx = document.getElementById('config-ds2').value;
            const key1 = document.getElementById('config-key1').value;
            const key2 = document.getElementById('config-key2').value;

            showLoader(true);
            setTimeout(() => {
                const ds1 = state.datasets[ds1_idx];
                const ds2 = state.datasets[ds2_idx];

                const map2 = new Map(ds2.data.map(row => [row[key2], row]));
                const mergedData = ds1.data.map(row1 => {
                    const row2 = map2.get(row1[key1]);
                    return { ...row1, ...(row2 || {}) };
                });

                const newHeaders = [...ds1.headers, ...ds2.headers.filter(h => !ds1.headers.includes(h))];
                addNewDataset(`Merged - ${ds1.name} & ${ds2.name}`, mergedData, newHeaders);
                showLoader(false);
                hideModal();
            }, 50);
        });
        
        populateKeys();
        document.getElementById('config-ds1').addEventListener('change', populateKeys);
        document.getElementById('config-ds2').addEventListener('change', populateKeys);
    });
    
    document.getElementById('action-split-file').addEventListener('click', () => {
        const content = `
            <p class="text-sm mb-4">Split the active dataset into multiple CSV files contained in a ZIP archive.</p>
            <label for="config-rows" class="block text-sm font-semibold">Rows Per File</label>
            <input type="number" id="config-rows" value="100000" min="1" class="w-full p-2 border rounded mt-1">
        `;
        showModal('Split File by Row Count', content, () => {
            const rowsPerFile = parseInt(document.getElementById('config-rows').value, 10);
            if (isNaN(rowsPerFile) || rowsPerFile < 1) return alert("Invalid number of rows.");

            showLoader(true);
            setTimeout(() => {
                const activeDataset = getActiveDataset();
                const zip = new JSZip();
                let fileCount = 1;
                for (let i = 0; i < activeDataset.data.length; i += rowsPerFile) {
                    const chunk = activeDataset.data.slice(i, i + rowsPerFile);
                    const worksheet = XLSX.utils.json_to_sheet(chunk);
                    const csvContent = XLSX.utils.sheet_to_csv(worksheet);
                    zip.file(`split_${fileCount++}.csv`, csvContent);
                }
                zip.generateAsync({ type: 'blob' }).then(content => {
                    saveAs(content, `Split_${activeDataset.name}.zip`);
                    showLoader(false);
                    hideModal();
                });
            }, 50);
        });
    });

    // Dummy placeholders for actions not fully implemented in this round
    document.getElementById('action-stack-sheets').addEventListener('click', () => alert("Stack Sheets feature coming soon!"));
    document.getElementById('action-compare-sheets').addEventListener('click', () => alert("Compare Sheets feature coming soon!"));


    // --- DOWNLOADING ---
    function handleDownload() {
        const activeDataset = getActiveDataset();
        if (!activeDataset) return;
        
        showLoader(true);
        setTimeout(() => {
            const newWorksheet = XLSX.utils.json_to_sheet(activeDataset.data);
            const newWorkbook = XLSX.utils.book_new();
            XLSX.utils.book_append_sheet(newWorkbook, newWorksheet, "Result");
            
            const fileName = `Processed_${activeDataset.name.replace(/[^a-z0-9]/gi, '_').toLowerCase()}.xlsx`;
            XLSX.writeFile(newWorkbook, fileName);
            showLoader(false);
        }, 50);
    }
});
