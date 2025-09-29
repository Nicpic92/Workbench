// js/fileHandler.js

import { state } from './state.js';
import { showLoader } from './ui.js';

export async function handleFileUpload(event) {
    const files = Array.from(event.target.files);
    if (files.length === 0) return;
    
    showLoader(true);
    state.datasets = []; // Clear previous datasets
    
    try {
        const allFileContents = await Promise.all(files.map(file => readFile(file)));

        allFileContents.forEach((fileContent, index) => {
            const file = files[index];
            try {
                const workbook = XLSX.read(fileContent, { type: 'array', cellDates: true });

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
                console.error("Error parsing workbook:", file.name, error);
                alert(`Could not parse the Excel file: ${file.name}`);
            }
        });
    } catch (error) {
        console.error("Error reading one or more files:", error);
        alert("There was an error reading one of the files. Please ensure it is not corrupted.");
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
