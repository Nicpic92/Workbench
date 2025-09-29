import { state, addNewDataset } from '../state.js';
import { showConfigModal, closeModal, showLoader, updateUI, generateDatasetSelect } from '../ui.js';

function compareSheetsAction() {
    if (state.datasets.length < 2) return alert("Upload at least two files to compare.");

    const content = `<div class="grid grid-cols-2 gap-4"><div class="border-r pr-4"><label class="block text-sm font-semibold">Original / Old File</label>${generateDatasetSelect('config-ds1')}<label class="block text-sm font-semibold mt-2">Unique ID Column</label><select id="config-key1" class="w-full p-2 border rounded mt-1"></select></div><div><label class="block text-sm font-semibold">New / Updated File</label>${generateDatasetSelect('config-ds2')}<label class="block text-sm font-semibold mt-2">Unique ID Column</label><select id="config-key2" class="w-full p-2 border rounded mt-1"></select></div></div>`;
    
    showConfigModal('Compare Sheets', content, () => {
        const ds1_idx = document.getElementById('config-ds1').value;
        const ds2_idx = document.getElementById('config-ds2').value;
        const key1 = document.getElementById('config-key1').value;
        const key2 = document.getElementById('config-key2').value;

        showLoader(true);
        setTimeout(() => {
            const ds1 = state.datasets[ds1_idx], ds2 = state.datasets[ds2_idx];
            const map1 = new Map(ds1.data.map(row => [row[key1], row]));
            const map2 = new Map(ds2.data.map(row => [row[key2], row]));
            const results = [];
            const allHeaders = [...new Set([...ds1.headers, ...ds2.headers])];
            
            map2.forEach((row2, key) => {
                const row1 = map1.get(key);
                if (!row1) { 
                    results.push({ Status: 'Added', ...row2 }); 
                } else {
                    let isModified = false;
                    for (const h of allHeaders) { if (String(row1[h] ?? '') !== String(row2[h] ?? '')) isModified = true; }
                    if (isModified) results.push({ Status: 'Modified', ...row2 });
                }
                map1.delete(key);
            });
            map1.forEach(row1 => results.push({ Status: 'Deleted', ...row1 }));
            
            addNewDataset(`Comparison - ${ds1.name} vs ${ds2.name}`, results, ['Status', ...allHeaders]);
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

export function initializeCompareSheetsAction() {
    document.getElementById('action-compare-sheets').addEventListener('click', compareSheetsAction);
}
