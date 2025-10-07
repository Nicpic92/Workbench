import { state } from './state.js';
import { getNoteCategory } from './processing.js';
import { downloadPrebatchReport } from './generator.js';

// This module contains all functions that directly interact with the DOM (the HTML page).
// This keeps the logic for how the page looks and behaves separate from the data processing logic.

// --- Core UI Functions ---

export function displayStatus(message, type, showLoader = false) {
    const statusDiv = document.getElementById('status');
    statusDiv.textContent = message;
    statusDiv.style.color = type === 'error' ? 'red' : (type === 'success' ? 'green' : '#4f46e5');
    document.getElementById('loader').style.display = showLoader ? 'block' : 'none';
    document.getElementById('processBtn').disabled = !!showLoader;
}

export function resetUI() {
    ['review-container', 'final-downloads-container', 'movement-summary-container', 'approaching-critical-container', 'prebatch-container'].forEach(id => {
        const el = document.getElementById(id);
        if (el) el.classList.add('hidden');
    });
    document.getElementById('download-links-container').innerHTML = '';
    document.getElementById('copyEmailBtn').classList.add('hidden');
    document.getElementById('status').textContent = '';
    document.getElementById('assignment-editor-container').innerHTML = '';
    document.getElementById('approaching-critical-table-container').innerHTML = '';
}

export function getFormattedDate() {
    const d = new Date(), day = d.getDate(), month = d.toLocaleString('default', { month: 'short' }), year = d.getFullYear();
    const s = (day % 10 == 1 && day != 11) ? 'st' : ((day % 10 == 2 && day != 12) ? 'nd' : ((day % 10 == 3 && day != 13) ? 'rd' : 'th'));
    return `${day}${s} ${month} ${year}`;
}

// --- Review Step Display Functions ---

export function displayReviewStep() {
    if (state.prebatchClaims.length > 0) {
        document.getElementById('prebatch-summary').textContent = `Found ${state.prebatchClaims.length} claims in Prebatch status.`;
        document.getElementById('downloadPrebatchBtn').onclick = downloadPrebatchReport;
        document.getElementById('prebatch-container').classList.remove('hidden');
    } else {
        document.getElementById('prebatch-container').classList.add('hidden');
    }

    displayApproachingCriticalTable();
    displayAssignmentEditor();

    document.getElementById('review-container').classList.remove('hidden');
}

function displayApproachingCriticalTable() {
    const approachingCriticalClaims = state.processedClaimsList.filter(claim => claim.cleanAge === 27);
    const container = document.getElementById('approaching-critical-container');
    const tableContainer = document.getElementById('approaching-critical-table-container');

    if (approachingCriticalClaims.length > 0) {
        let tableHtml = `
            <div class="table-container border rounded-lg max-h-64 overflow-y-auto">
                <table class="min-w-full divide-y divide-gray-200">
                    <thead class="bg-gray-50 sticky top-0">
                        <tr>
                            <th class="px-4 py-2 text-left text-xs font-medium text-gray-500 uppercase">Claim #</th>
                            <th class="px-4 py-2 text-left text-xs font-medium text-gray-500 uppercase">Payer</th>
                            <th class="px-4 py-2 text-left text-xs font-medium text-gray-500 uppercase">Total Charges</th>
                            <th class="px-4 py-2 text-left text-xs font-medium text-gray-500 uppercase">Claim State</th>
                            <th class="px-4 py-2 text-left text-xs font-medium text-gray-500 uppercase">Current Owner</th>
                        </tr>
                    </thead>
                    <tbody class="bg-white divide-y divide-gray-200">`;

        for (const claim of approachingCriticalClaims) {
            const totalCharges = (parseFloat(String(claim.originalRow[claim.totalChargesIndex]).replace(/[^0-9.-]/g, '')) || 0)
                                 .toLocaleString('en-US', { style: 'currency', currency: 'USD' });
            
            tableHtml += `
                <tr>
                    <td class="px-4 py-2 whitespace-nowrap text-sm font-medium text-gray-900">${claim.claimNumber}</td>
                    <td class="px-4 py-2 whitespace-nowrap text-sm text-gray-700">${claim.originalRow[claim.payerIndex] || 'N/A'}</td>
                    <td class="px-4 py-2 whitespace-nowrap text-sm text-gray-700">${totalCharges}</td>
                    <td class="px-4 py-2 whitespace-nowrap text-sm text-gray-700">${claim.claimState}</td>
                    <td class="px-4 py-2 whitespace-nowrap text-sm font-semibold text-gray-800">${claim.finalOwner}</td>
                </tr>`;
        }

        tableHtml += `</tbody></table></div>`;
        tableContainer.innerHTML = tableHtml;
        container.classList.remove('hidden');
    } else {
        container.classList.add('hidden');
    }
}

function displayAssignmentEditor() {
    const assignmentContainer = document.getElementById('assignment-editor-container');
    assignmentContainer.innerHTML = `<h3 class="text-xl font-bold text-gray-900 mb-3">Assignment Editor</h3><p id="assignment-editor-description" class="text-gray-700 mb-4"></p>`;
    const categorizedNotes = {};
    let uniqueNoteCount = 0;
    
    state.processedClaimsList.forEach(claim => {
        const note = claim.noteText || "No Note";
        if (note !== "No Note") {
            const category = getNoteCategory(note);
            if (!categorizedNotes[category]) categorizedNotes[category] = {};
            if (!categorizedNotes[category][note]) {
                categorizedNotes[category][note] = { count: 0, defaultOwner: claim.defaultOwner };
                uniqueNoteCount++;
            }
            categorizedNotes[category][note].count++;
        }
    });

    if (uniqueNoteCount === 0) {
        document.getElementById('assignment-editor-description').textContent = "No notes were found in the report to assign.";
        assignmentContainer.innerHTML += `<div class="text-center py-4 text-gray-500">No notes found.</div>`;
    } else {
        document.getElementById('assignment-editor-description').textContent = `Found ${uniqueNoteCount} unique notes. The 'Default Assignment' now reflects yesterday's final assignment. Please review.`;
        Object.keys(categorizedNotes).sort().forEach(category => {
            const notes = categorizedNotes[category];
            const categoryId = category.replace(/\s|&/g, '-');
            let tableRowsHtml = '';
            Object.keys(notes).sort((a, b) => notes[b].count - notes[a].count).forEach(noteText => {
                const data = notes[noteText];
                tableRowsHtml += `<tr class="bg-white"><td class="px-6 py-4 text-sm text-gray-700 break-words">${noteText}</td><td class="px-6 py-4 text-sm text-gray-700">${data.count}</td><td class="px-6 py-4 text-sm font-medium text-gray-900">${data.defaultOwner || 'N/A'}</td><td class="px-6 py-4 text-sm text-gray-500"><select class="p-2 border rounded-md w-full assignment-override" data-note-text="${noteText.replace(/"/g, '&quot;')}" data-category="${categoryId}"><option value="">Keep Default</option><option value="Claims">Claims</option><option value="PV">PV</option></select></td></tr>`;
            });
            const categoryHtml = `<details class="bg-gray-50 border rounded-lg overflow-hidden" open><summary class="p-4 bg-gray-100 hover:bg-gray-200 flex justify-between items-center"><h4 class="text-lg font-bold text-gray-800">${category} (${Object.keys(notes).length} notes)</h4><div class="flex items-center space-x-2"><label class="text-sm font-medium">Assign All to:</label><select class="p-2 border rounded-md category-assign" data-category-target="${categoryId}"><option value="">-- Bulk Assign --</option><option value="Claims">Claims</option><option value="PV">PV</option></select></div></summary><div class="p-2"><div class="table-container border rounded-lg"><table class="min-w-full divide-y divide-gray-200 table-fixed"><thead class="bg-white sticky top-0"><tr><th scope="col" class="w-1/2 px-6 py-3 text-left text-xs font-medium text-gray-500 uppercase tracking-wider">Note Text</th><th scope="col" class="w-1/12 px-6 py-3 text-left text-xs font-medium text-gray-500 uppercase tracking-wider">Count</th><th scope="col" class="w-1/4 px-6 py-3 text-left text-xs font-medium text-gray-500 uppercase tracking-wider">Default Assignment</th><th scope="col" class="w-1/4 px-6 py-3 text-left text-xs font-medium text-gray-500 uppercase tracking-wider">New Assignment</th></tr></thead><tbody class="divide-y divide-gray-200">${tableRowsHtml}</tbody></table></div></div></details>`;
            assignmentContainer.innerHTML += categoryHtml;
        });

        // Add event listeners for bulk assignment dropdowns
        document.querySelectorAll('.category-assign').forEach(select => {
            select.addEventListener('change', (e) => {
                const targetCategory = e.target.dataset.categoryTarget;
                const newAssignment = e.target.value;
                if (!newAssignment) return;
                document.querySelectorAll(`.assignment-override[data-category="${targetCategory}"]`).forEach(noteSelect => { noteSelect.value = newAssignment; });
            });
        });
    }
}
