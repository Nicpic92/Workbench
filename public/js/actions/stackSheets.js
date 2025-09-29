// js/actions/stackSheets.js

import { state, addNewDataset } from '../state.js';
import { showConfigModal, closeModal, showLoader, updateUI } from '../ui.js';

function stackSheetsAction() {
    if (state.datasets.length < 2) return alert("Please load at least two files to stack.");

    const content = `<p class="text-sm mb-4">This will combine all ${state.datasets.length} currently loaded datasets into a single master sheet. Columns will be matched by header name.</p>`;
    
    showConfigModal('Stack All Sheets', content, () => {
        showLoader(true);
        setTimeout(() => {
            const allData = state.datasets.flatMap(ds => ds.data);
            const allHeaders = [...new Set(state.datasets.flatMap(ds => ds.headers))];
            
            addNewDataset(`Stacked - ${state.datasets.length} files`, allData, allHeaders);
            updateUI();
            showLoader(false); 
            closeModal('config-modal');
        }, 50);
    });
}

export function initializeStackSheetsAction() {
    document.getElementById('action-stack-sheets').addEventListener('click', stackSheetsAction);
}
