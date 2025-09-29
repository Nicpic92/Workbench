// js/actions/trimWhitespace.js

import { getActiveDataset } from '../state.js';
import { showConfigModal, closeModal, showLoader, generateColumnSelect, updateUI } from '../ui.js';

function trimWhitespaceAction() {
    const activeDS = getActiveDataset();
    if (!activeDS) return alert("Please load a file first.");

    const content = `<p class="text-sm mb-4">Select the column to trim.</p><label for="trim-column" class="block text-sm font-semibold">Column:</label>${generateColumnSelect(activeDS.headers, 'config-column')}`;
    
    showConfigModal('Trim Whitespace', content, () => {
        const column = document.getElementById('config-column').value;
        showLoader(true);
        setTimeout(() => {
            getActiveDataset().data.forEach(row => {
                if (typeof row[column] === 'string') {
                    row[column] = row[column].trim();
                }
            });
            updateUI(); // Re-render the table with the updated data
            showLoader(false); 
            closeModal('config-modal');
        }, 50);
    });
}

export function initializeTrimWhitespaceAction() {
    document.getElementById('action-trim-whitespace').addEventListener('click', trimWhitespaceAction);
}
