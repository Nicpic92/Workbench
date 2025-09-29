// js/actions/anonymize.js

import { getActiveDataset, addNewDataset } from '../state.js';
import { showConfigModal, closeModal, showLoader, updateUI } from '../ui.js';

function anonymizeAction() {
    const activeDS = getActiveDataset();
    if (!activeDS) return alert("Please load a file first.");

    const types = [{ v: 'NONE', l: 'Do Not Anonymize' }, { v: 'FULL_NAME', l: 'Full Name' }, { v: 'FIRST_NAME', l: 'First Name' }, { v: 'LAST_NAME', l: 'Last Name' }, { v: 'EMAIL', l: 'Email' }, { v: 'PHONE', l: 'Phone' }];
    const content = `<div class="space-y-3">${activeDS.headers.map(h => `<div class="grid grid-cols-2 gap-4 items-center"><label class="font-semibold truncate" title="${h}">${h}</label><select data-header="${h}" class="column-mapper w-full p-2 border rounded">${types.map(t => `<option value="${t.v}">${t.l}</option>`).join('')}</select></div>`).join('')}</div>`;
    
    showConfigModal('Anonymize Personal Information', content, () => {
        const mappings = Array.from(document.querySelectorAll('#config-modal .column-mapper')).filter(s => s.value !== 'NONE').map(s => ({ header: s.dataset.header, type: s.value }));
        if (mappings.length === 0) return alert('Please select at least one column to anonymize.');

        showLoader(true);
        setTimeout(() => {
            const fake = { FIRST: ['Alex', 'Jordan', 'Casey', 'Taylor'], LAST: ['Smith', 'Jones', 'Williams', 'Brown'], FULL: () => `${fake.FIRST[Math.floor(Math.random()*4)]} ${fake.LAST[Math.floor(Math.random()*4)]}`, EMAIL: () => `user${Math.floor(1000+Math.random()*9000)}@example.com`, PHONE: () => `(555) ${Math.floor(100+Math.random()*900)}-${Math.floor(1000+Math.random()*9000)}` };
            const anonData = activeDS.data.map(row => {
                const newRow = { ...row };
                mappings.forEach(m => {
                    if (newRow[m.header] !== undefined) newRow[m.header] = { FULL_NAME: fake.FULL(), FIRST_NAME: fake.FIRST[Math.floor(Math.random()*4)], LAST_NAME: fake.LAST[Math.floor(Math.random()*4)], EMAIL: fake.EMAIL(), PHONE: fake.PHONE() }[m.type];
                });
                return newRow;
            });
            addNewDataset(`Anonymized - ${activeDS.name}`, anonData, activeDS.headers);
            updateUI();
            showLoader(false); 
            closeModal('config-modal');
        }, 50);
    });
}

export function initializeAnonymizeAction() {
    document.getElementById('action-anonymize').addEventListener('click', anonymizeAction);
}
