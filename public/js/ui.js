import { state } from './state.js';
import { getNoteCategory } from './processing.js';
import { downloadPrebatchReport, generateAssignmentReport } from './generator.js';

// This module contains all functions that directly interact with the DOM (the HTML page).
// This keeps the logic for how the page looks and behaves separate from the data processing logic.

// --- Core UI Functions ---

export function displayStatus(message, type, showLoader = false) {
    const statusDiv = document.getElementById('status');
    const loaderDiv = document.getElementById('loader');
    const processButton = document.getElementById('processBtn');

    if (statusDiv) {
        statusDiv.textContent = message;
        statusDiv.style.color = type === 'error' ? 'red' : (type === 'success' ? 'green' : '#4f46e5');
    }
    if (loaderDiv) loaderDiv.style.display = showLoader ? 'block' : 'none';
    if (processButton) processButton.disabled = !!showLoader;
}

export function displayWarning(message) {
    const warningContainer = document.getElementById('warning-container');
    const warningMessage = document.getElementById('warning-message');

    if (warningContainer && warningMessage) {
        warningMessage.textContent = message;
        warningContainer.classList.remove('hidden');
    }
}

export function resetUI() {
    // Configuration of elements to reset. This is a safer way to manage UI state.
    const elementsToReset = {
        'review-container': { action: 'hide' },
        'final-downloads-container': { action: 'hide' },
        'movement-summary-container': { action: 'hide' },
        'approaching-critical-container': { action: 'hide' },
        'prebatch-container': { action: 'hide' },
        'warning-container': { action: 'hide' },
        'assignment-upload-step': { action: 'hide' },
        'download-links-container': { action: 'clearHTML' },
        'approaching-critical-table-container': { action: 'clearHTML' },
        'status': { action: 'clearText' },
        'assignmentFileName': { action: 'setText', value: 'No file selected.' },
        'copyEmailBtn': { action: 'hide' },
        'generateFinalReportsBtn': { action: 'disable' },
        'assignmentFileInput': { action: 'resetValue' }
    };

    // Loop through the configuration and safely update the DOM
    for (const id in elementsToReset) {
        const element = document.getElementById(id);
        if (element) {
            const config = elementsToReset[id];
            switch (config.action) {
                case 'hide':
                    element.classList.add('hidden');
                    break;
                case 'clearHTML':
                    element.innerHTML = ''; // The operation that previously caused the error
                    break;
                case 'clearText':
                    element.textContent = '';
                    break;
                case 'setText':
                    element.textContent = config.value;
                    break;
                case 'disable':
                    element.disabled = true;
                    break;
                case 'resetValue':
                    element.value = '';
                    break;
            }
        } else {
            // This will log a warning in the browser's developer console if an element is missing, but it will NOT crash the app.
            console.warn(`resetUI failed to find element with ID: '${id}'. This may indicate an HTML/JS mismatch.`);
        }
    }
}


export function getFormattedDate() {
    const d = new Date(), day = d.getDate(), month = d.toLocaleString('default', { month: 'short' }), year = d.getFullYear();
    const s = (day % 10 == 1 && day != 11) ? 'st' : ((day % 10 == 2 && day != 12) ? 'nd' : ((day % 10 == 3 && day != 13) ? 'rd' : 'th'));
    return `${day}${s} ${month} ${year}`;
}

// --- Review Step Display Functions ---

export function displayReviewStep() {
    const prebatchContainer = document.getElementById('prebatch-container');
    if (prebatchContainer) {
        if (state.prebatchClaims.length > 0) {
            const summary = document.getElementById('prebatch-summary');
            const button = document.getElementById('downloadPrebatchBtn');
            if(summary) summary.textContent = `Found ${state.prebatchClaims.length} claims in Prebatch status.`;
            if(button) button.onclick = downloadPrebatchReport;
            prebatchContainer.classList.remove('hidden');
        } else {
            prebatchContainer.classList.add('hidden');
        }
    }

    displayApproachingCriticalTable();
    
    // Setup for the new assignment workflow
    const downloadBtn = document.getElementById('downloadAssignmentReportBtn');
    if (downloadBtn) downloadBtn.onclick = generateAssignmentReport;
    
    const uploadStep = document.getElementById('assignment-upload-step');
    if (uploadStep) uploadStep.classList.remove('hidden');
    
    const generateBtn = document.getElementById('generateFinalReportsBtn');
    if(generateBtn) generateBtn.disabled = true;

    const reviewContainer = document.getElementById('review-container');
    if (reviewContainer) reviewContainer.classList.remove('hidden');
}

function displayApproachingCriticalTable() {
    const approachingCriticalClaims = state.processedClaimsList.filter(claim => claim.cleanAge === 27);
    const container = document.getElementById('approaching-critical-container');
    const tableContainer = document.getElementById('approaching-critical-table-container');

    const countElement = document.getElementById('moving-to-critical-count');
    if (countElement) {
        countElement.textContent = approachingCriticalClaims.length.toLocaleString();
    }
    
    if (!container || !tableContainer) return; // Exit if essential elements are missing

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
