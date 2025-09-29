// js/actions/mergeFiles.js

import { state, getActiveDataset, addNewDataset } from '../state.js';
import { showConfigModal, closeModal, showLoader, updateUI, generateDatasetSelect } from '../ui.js';

function mergeFilesAction() {
    if (state.datasets.length < 2) return alert("Upload at least two files to merge.");

    const content = `<div class="grid grid-cols-2 gap-4"><div class="border-r pr-4"><label class="block text-sm font-semibold">Left Table (Primary)</label>${generateDatasetSelect('config-ds1')}<label class="block text-sm font-semibold mt-2">Key Column</label><select id="config-key1" class="w-full p-2 border rounded mt-1"></select></div><div><label class="block text-sm font-semibold">Right Table (to join)</label>${generateDatasetSelect('config-ds2')}<label class="block text-sm font-semibold mt-2">Key Column</label><select id="config-key2" class="w-full p-2 border rounded mt-1"></select></div></div>`;
    
    showConfigModal('Merge Files (Left Join)', content, () => {
        const ds1_idx = document.getElementById('config-ds1').value;
        const ds2_idx = document.getElementById('config-ds2').value;
        const key1 = document.getElementById('config-key1').value;
        const key2 = document.getElementById('config-key2').value;

        showLoader(true);
        setTimeout(() => {
            const ds1 = state.datasets[ds1_idx];
            const ds2 = state.datasets[ds2_idx];
            const map2 = new Map(ds2.data.map(row => [row[key2], row]));
            const mergedData = ds1.data.map(row1 => ({ ...row1, ...(map2.get(row1[key1]) || {}) }));
            const newHeaders = [...new Set([...ds1.headers, ...ds2.headers])];
            
            addNewDataset(`Merged - ${ds1.name} & ${ds2.name}`, mergedData, newHeaders);
            updateUI();
            showLoader(false); 
            closeModal('config-modal');
        }, 50);
    });

    const populateKeys = () => {
        ['1', '2'].forEach(n => {
            const ds_idx = document.getElementById(`config-ds${n}`).value;
            document.getElementById(`config-key${n}`).innerHTML = state.datasets[ds_idx].headers.map(h => `<option value="${h}">${h}</option>`).join('');
        });
    };
    populateKeys();
    document.getElementById('config-ds1').onchange = populateKeys;
    document.getElementById('config-ds2').onchange = populateKeys;
}

export function initializeMergeFilesAction() {
    document.getElementById('action-merge-files').addEventListener('click', mergeFilesAction);
}
