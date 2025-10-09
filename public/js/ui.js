import { state } from './state.js';
import { getNoteCategory } from './processing.js';
import { downloadPrebatchReport, generateAssignmentReport } from './generator.js';

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

export function displayWarning(message) {
    const warningContainer = document.getElementById('warning-container');
    const warningMessage = document.getElementById('warning-message');
    const editorDescription = document.getElementById('assignment-editor-description');

    if (warningContainer && warningMessage) {
        warningMessage.textContent = message;
        warningContainer.classList.remove('hidden');
    }
}

export function resetUI() {
    // Hide all major containers
    ['review-container', 'final-downloads-container', 'movement-summary-container', 'approaching-critical-container', 'prebatch-container', 'warning-container', 'assignment-upload-step'].forEach(id => {
        const el = document.getElementById(id);
        if (el) el.classList.add('hidden');
    });

    // Safely clear content of specific elements
    const downloadLinks = document.getElementById('download-links-container');
    if (downloadLinks) downloadLinks.innerHTML = '';

    const copyBtn = document.getElementById('copyEmailBtn');
    if (copyBtn) copyBtn.classList.add('hidden');

    const statusDiv = document.getElementById('status');
    if (statusDiv) statusDiv.textContent = '';
    
    const criticalTable = document.getElementById('approaching-critical-table-container');
    if (criticalTable) criticalTable.innerHTML = '';
    
    // Safely reset the assignment file input
    const assignmentFileInput = document.getElementById('assignmentFileInput');
    const assignmentFileName = document.getElementById('assignmentFileName');
    if (assignmentFileInput && assignmentFileName) {
        assignmentFileInput.value = '';
        assignmentFileName.textContent = 'No file selected.';
    }

    const generateBtn = document.getElementById('generateFinalReportsBtn');
    if (generateBtn) generateBtn.disabled = true;
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
    
    // Setup for the new assignment workflow
    document.getElementById('downloadAssignmentReportBtn').onclick = generateAssignmentReport;
    document.getElementById('assignment-upload-step').classList.remove('hidden');
    document.getElementById('generateFinalReportsBtn').disabled = true; // Disable until assignment file is uploaded

    document.getElementById('review-container').classList.remove('hidden');
}

function displayApproachingCriticalTable() {
    const approachingCriticalClaims = state.processedClaimsList.filter(claim => claim.cleanAge === 27);
    const container = document.getElementById('approaching-critical-container');
    const tableContainer = document.getElementById('approaching-critical-table-container');

    const countElement = document.getElementById('moving-to-critical-count');
    if (countElement) {
        countElement.textContent = approachingCriticalClaims.length.toLocaleString();
    }

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
                    <td class="px-4 py-2 whitespace-nowrap text-sm font-semibold text-gray-800">${claim.defaultOwner}</td>
                </tr>`;
        }

        tableHtml += `</tbody></table></div>`;
        tableContainer.innerHTML = tableHtml;
        container.classList.remove('hidden');
    } else {
        container.classList.add('hidden');
    }
}
