// js/fileHandler.js

import { state } from './state.js';
import { showLoader } from './ui.js';

export async function handleFileUpload(event) {
    const files = Array.from(event.target.files);
    if (files.length === 0) return;
    
    showLoader(true);
    state.datasets = []; // Clear previous datasets
    
    for (const file of files) {
        try {
            const data = await readFile(file);
            const workbook = XLSX.read(data, { type: 'array', cellDates: true });

            if (workbook.SheetNames.length === 1) {
                const jsonData = XLSX.utils.sheet_to_json(workbook.Sheets[workbook.SheetNames[0]]);
                state.datasets.push({ 
                    name: file.name, 
                    data: jsonData, 
                    headers: Object.keys(jsonData[0] || {}) 
                });
            } else {
                 workbook.SheetNames.forEach(sheetName => {
                    const jsonData = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName]);
                    state.datasets.push({ 
                        name: `${file.name} - ${sheetName}`, 
                        data: jsonData, 
                        headers: Object.keys(jsonData[0] || {}) 
                    });
                });
            }
        } catch (error) {
            console.error("Error processing file:", file.name, error);
            alert(`Could not process file: ${file.name}`);
        }
    }
    
    state.activeDatasetIndex = 0;
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
