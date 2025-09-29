// js/actions/splitFile.js

import { getActiveDataset } from '../state.js';
import { showConfigModal, closeModal, showLoader } from '../ui.js';

function splitFileAction() {
    const activeDS = getActiveDataset();
    if (!activeDS) return alert("Please load a file first.");

    const content = `<p class="text-sm mb-4">Split the active dataset into multiple CSV files.</p><label for="config-rows" class="block text-sm font-semibold">Rows Per File</label><input type="number" id="config-rows" value="100000" min="1" class="w-full p-2 border rounded mt-1">`;
    
    showConfigModal('Split File by Row Count', content, () => {
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
            zip.generateAsync({ type: 'blob' }).then(content => {
                saveAs(content, `Split_${activeDS.name}.zip`);
                showLoader(false); 
                closeModal('config-modal');
            });
        }, 50);
    });
}

export function initializeSplitFileAction() {
    document.getElementById('action-split-file').addEventListener('click', splitFileAction);
}
