document.addEventListener('DOMContentLoaded', () => {
    // --- STATE MANAGEMENT ---
    // A central object to hold the application's state.
    const state = {
        datasets: [], // Will hold all loaded file data { name, data, headers }
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
    const configPanelContent = document.getElementById('config-panel-content');
    const loaderOverlay = document.getElementById('loader-overlay');
    const downloadBtn = document.getElementById('download-btn');

    // --- INITIALIZATION ---
    fileUploadInput.addEventListener('change', handleFileUpload);
    downloadBtn.addEventListener('click', handleDownload);

    // --- FILE HANDLING ---
    async function handleFileUpload(event) {
        const files = event.target.files;
        if (files.length === 0) return;

        showLoader(true);
        state.datasets = []; // Clear previous datasets

        for (const file of files) {
            try {
                const data = await readFile(file);
                const workbook = XLSX.read(data, { type: 'array', cellDates: true });
                // For now, we only process the first sheet. Stacking comes later.
                const sheetName = workbook.SheetNames[0];
                const worksheet = workbook.Sheets[sheetName];
                const jsonData = XLSX.utils.sheet_to_json(worksheet);
                const headers = Object.keys(jsonData[0] || {});
                
                state.datasets.push({
                    name: file.name,
                    data: jsonData,
                    headers: headers,
                });
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
        } else {
            welcomeView.style.display = 'none';
            dataView.style.display = 'flex';
            actionsContainer.style.display = 'block';
            renderActiveDataset();
        }
    }

    function renderActiveDataset() {
        const activeDataset = state.datasets[state.activeDatasetIndex];
        if (!activeDataset) return;

        tableTitle.textContent = activeDataset.name;
        renderDataTable(activeDataset.data, activeDataset.headers);
        statusBar.textContent = `Displaying ${activeDataset.data.length.toLocaleString()} rows and ${activeDataset.headers.length} columns.`;
    }

    function renderDataTable(data, headers) {
        const table = document.createElement('table');
        
        // Headers
        const thead = table.createTHead();
        const headerRow = thead.insertRow();
        for (const header of headers) {
            const th = document.createElement('th');
            th.textContent = header;
            headerRow.appendChild(th);
        }

        // Body (sample of rows for performance)
        const tbody = table.createTBody();
        const sampleData = data.slice(0, 200); // Only render first 200 rows to keep UI snappy
        for (const row of sampleData) {
            const tr = tbody.insertRow();
            for (const header of headers) {
                const td = tr.insertCell();
                td.textContent = row[header] === null || row[header] === undefined ? '' : row[header];
            }
        }
        
        tableContainer.innerHTML = '';
        tableContainer.appendChild(table);
    }
    
    function showLoader(show) {
        loaderOverlay.style.display = show ? 'flex' : 'none';
    }

    // --- ACTION MODULES & EVENT LISTENERS ---

    // Action: Trim Whitespace
    document.getElementById('action-trim-whitespace').addEventListener('click', () => {
        configPanelContent.innerHTML = `
            <h3 class="font-semibold mb-2">Trim Whitespace</h3>
            <p class="text-sm mb-4">This will remove leading/trailing spaces from every cell in the selected column.</p>
            <label for="trim-column" class="block text-sm font-semibold">Select Column:</label>
            <select id="trim-column" class="w-full p-2 border rounded mt-1"></select>
            <button id="run-trim" class="w-full mt-4 bg-indigo-600 text-white font-semibold py-2 rounded hover:bg-indigo-700">Apply</button>
        `;
        
        const select = document.getElementById('trim-column');
        const activeDataset = state.datasets[state.activeDatasetIndex];
        select.innerHTML = activeDataset.headers.map(h => `<option value="${h}">${h}</option>`).join('');
        
        document.getElementById('run-trim').addEventListener('click', () => {
            const columnToTrim = select.value;
            showLoader(true);
            setTimeout(() => { // Use timeout to allow loader to show
                activeDataset.data.forEach(row => {
                    if (typeof row[columnToTrim] === 'string') {
                        row[columnToTrim] = row[columnToTrim].trim();
                    }
                });
                renderActiveDataset();
                showLoader(false);
                alert(`Whitespace trimmed in column: ${columnToTrim}`);
            }, 50);
        });
    });

    // Action: Find Duplicates
    document.getElementById('action-find-duplicates').addEventListener('click', () => {
        configPanelContent.innerHTML = `
            <h3 class="font-semibold mb-2">Find Duplicates</h3>
            <p class="text-sm mb-4">Select one or more columns to check for duplicate rows.</p>
            <div id="duplicate-columns" class="space-y-2"></div>
            <button id="run-find-duplicates" class="w-full mt-4 bg-indigo-600 text-white font-semibold py-2 rounded hover:bg-indigo-700">Find</button>
        `;
        
        const container = document.getElementById('duplicate-columns');
        const activeDataset = state.datasets[state.activeDatasetIndex];
        container.innerHTML = activeDataset.headers.map(h => `
            <label class="flex items-center p-2 rounded hover:bg-gray-100">
                <input type="checkbox" class="h-4 w-4 rounded mr-2" data-column-name="${h}">
                <span class="text-sm">${h}</span>
            </label>
        `).join('');
        
        document.getElementById('run-find-duplicates').addEventListener('click', () => {
            const selectedColumns = Array.from(container.querySelectorAll('input:checked')).map(cb => cb.dataset.columnName);
            if (selectedColumns.length === 0) {
                alert("Please select at least one column.");
                return;
            }
            showLoader(true);
            setTimeout(() => {
                const seen = new Set();
                const duplicates = [];
                activeDataset.data.forEach(row => {
                    const key = selectedColumns.map(col => row[col]).join('||');
                    if (seen.has(key)) {
                        duplicates.push(row);
                    } else {
                        seen.add(key);
                    }
                });
                
                // Create a new dataset view for the duplicates
                state.datasets.push({
                    name: `Duplicates from ${activeDataset.name}`,
                    data: duplicates,
                    headers: activeDataset.headers
                });
                state.activeDatasetIndex = state.datasets.length - 1;
                renderActiveDataset();
                showLoader(false);
                alert(`Found ${duplicates.length} duplicate rows.`);
            }, 50);
        });
    });

    // --- DOWNLOADING ---
    function handleDownload() {
        const activeDataset = state.datasets[state.activeDatasetIndex];
        if (!activeDataset) return;
        
        showLoader(true);
        setTimeout(() => {
            const newWorksheet = XLSX.utils.json_to_sheet(activeDataset.data);
            const newWorkbook = XLSX.utils.book_new();
            XLSX.utils.book_append_sheet(newWorkbook, newWorksheet, "Result");
            
            // Generate a safe filename
            const fileName = `Processed_${activeDataset.name.replace(/[^a-z0-9]/gi, '_').toLowerCase()}.xlsx`;
            XLSX.writeFile(newWorkbook, fileName);
            showLoader(false);
        }, 50);
    }
});
