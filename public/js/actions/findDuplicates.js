import { getActiveDataset, addNewDataset } from '../state.js';
import { showConfigModal, closeModal, showLoader, updateUI, generateColumnCheckboxes } from '../ui.js';

function findDuplicatesAction() {
    const activeDS = getActiveDataset();
    if (!activeDS) return alert("Please load a file first.");

    const content = `<p class="text-sm mb-4">Select columns to check for duplicates.</p><div class="space-y-2 max-h-96 overflow-y-auto">${generateColumnCheckboxes(activeDS.headers)}</div>`;
    
    showConfigModal('Find Duplicates', content, () => {
        const selected = Array.from(document.querySelectorAll('#config-modal input:checked')).map(cb => cb.dataset.columnName);
        if (selected.length === 0) return alert('Select at least one column.');

        showLoader(true);
        setTimeout(() => {
            const seen = new Map();
            const duplicates = [];
            activeDS.data.forEach(row => {
                const key = selected.map(col => row[col]).join('||');
                if (seen.has(key)) {
                    if (seen.get(key).first) { 
                        duplicates.push(seen.get(key).row); 
                        seen.get(key).first = false; 
                    }
                    duplicates.push(row);
                } else { 
                    seen.set(key, { row: row, first: true }); 
                }
            });

            if(duplicates.length > 0){
                addNewDataset(`Duplicates - ${activeDS.name}`, duplicates, activeDS.headers);
                updateUI();
            } else {
                alert("No duplicates found based on the selected columns.");
            }
            
            showLoader(false); 
            closeModal('config-modal');
        }, 50);
    });
}

export function initializeFindDuplicatesAction() {
    document.getElementById('action-find-duplicates').addEventListener('click', findDuplicatesAction);
}
