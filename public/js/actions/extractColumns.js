import { getActiveDataset, addNewDataset } from '../state.js';
import { showConfigModal, closeModal, showLoader, updateUI, generateColumnCheckboxes } from '../ui.js';

function extractColumnsAction() {
    const activeDS = getActiveDataset();
    if (!activeDS) return alert("Please load a file first.");

    const content = `<p class="text-sm mb-4">Select columns to keep.</p><div class="space-y-2 max-h-96 overflow-y-auto">${generateColumnCheckboxes(activeDS.headers)}</div>`;
    
    showConfigModal('Extract Columns', content, () => {
        const selected = Array.from(document.querySelectorAll('#config-modal input:checked')).map(cb => cb.dataset.columnName);
        if (selected.length === 0) return alert('Please select at least one column.');

        showLoader(true);
        setTimeout(() => {
            const newData = activeDS.data.map(row => selected.reduce((obj, key) => { if(row.hasOwnProperty(key)) obj[key] = row[key]; return obj; }, {}));
            addNewDataset(`Extracted - ${activeDS.name}`, newData, selected);
            updateUI();
            showLoader(false); 
            closeModal('config-modal');
        }, 50);
    });
}

export function initializeExtractColumnsAction() {
    document.getElementById('action-extract-columns').addEventListener('click', extractColumnsAction);
}
