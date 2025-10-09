import { state, resetState } from './state.js';
import { clientPresets, gatherConfig } from './config.js';
import * as ui from './ui.js';
import { processYesterdayReport, processAndAssignClaims, calculateStats, analyzeClaimNotes } from './processing.js';
import { buildWorkbook, generatePdfReport } from './generator.js';

// This is the main entry point for the application. It orchestrates the other modules.
// It sets up event listeners and defines the core workflow when a user interacts with the tool.

// --- INITIALIZATION ---
document.addEventListener('DOMContentLoaded', initializeTool);

function initializeTool() {
    const clientSelect = document.getElementById('client-select');
    const fileInput = document.getElementById('fileInput');
    const yesterdayFileInput = document.getElementById('yesterdayFileInput');
    const assignmentFileInput = document.getElementById('assignmentFileInput');
    const processBtn = document.getElementById('processBtn');

    // --- Event Listeners ---
    clientSelect.addEventListener('change', handleClientSelection);
    fileInput.addEventListener('change', checkFiles);
    yesterdayFileInput.addEventListener('change', checkFiles);
    assignmentFileInput.addEventListener('change', handleAssignmentFileUpload);
    processBtn.addEventListener('click', performInitialProcessing);
    document.getElementById('generateFinalReportsBtn').addEventListener('click', generateFinalReports);
    document.getElementById('copyEmailBtn').addEventListener('click', copyEmailText);
}

// --- Event Handler Functions ---

function handleClientSelection() {
    const selectedClient = document.getElementById('client-select').value;
    const configContainer = document.getElementById('configuration-container');
    const uploadContainer = document.getElementById('upload-container');

    ui.resetUI();
    resetState(); // Reset state when client changes

    if (selectedClient && clientPresets[selectedClient]) {
        const preset = clientPresets[selectedClient];
        document.getElementById('cleanAgeCol-label').textContent = preset.label;
        Object.keys(preset).forEach(key => {
            if (key !== 'label') document.getElementById(key).value = preset[key];
        });
        configContainer.classList.remove('hidden');
        uploadContainer.classList.remove('hidden');
    } else {
        configContainer.classList.add('hidden');
        uploadContainer.classList.add('hidden');
    }
    checkFiles(); // Re-check file status
}

function checkFiles() {
    const fileInput = document.getElementById('fileInput');
    const yesterdayFileInput = document.getElementById('yesterdayFileInput');
    const processBtn = document.getElementById('processBtn');

    if(fileInput.files[0]) document.getElementById('fileName').textContent = `Selected: ${fileInput.files[0].name}`;
    if(yesterdayFileInput.files[0]) document.getElementById('yesterdayFileName').textContent = `Selected: ${yesterdayFileInput.files[0].name}`;

    processBtn.disabled = !fileInput.files[0];
}

async function performInitialProcessing() {
    const mainFile = document.getElementById('fileInput').files[0];
    const yesterdayFile = document.getElementById('yesterdayFileInput').files[0];
    const config = gatherConfig();

    if (!config) {
        ui.displayStatus('Configuration Error: Please check the entered column letters.', 'error');
        return;
    }

    ui.displayStatus('Processing... Please wait.', 'info', true);
    ui.resetUI();
    resetState();

    state.hasYesterdayFile = !!yesterdayFile;

    const readFileAsAOA = (file) => new Promise((resolve, reject) => {
        const reader = new FileReader();
        reader.onload = e => {
            try {
                const wb = XLSX.read(new Uint8Array(e.target.result), { type: 'array' });
                const sheetName = wb.SheetNames.includes('All Processed Data') ? 'All Processed Data' : wb.SheetNames[0];
                resolve(XLSX.utils.sheet_to_json(wb.Sheets[sheetName], { header: 1, defval: null }));
            } catch (err) { reject(err); }
        };
        reader.onerror = () => reject(new Error('File could not be read.'));
        reader.readAsArrayBuffer(file);
    });

    try {
        if (state.hasYesterdayFile) {
            const yesterday_aoa = await readFileAsAOA(yesterdayFile);
            const yestData = processYesterdayReport(yesterday_aoa);
            state.yesterdayStats = yestData.stats;
            state.yesterdayDataMap = yestData.dataMap;
        }

        const main_aoa = await readFileAsAOA(mainFile);
        state.mainReportHeader = main_aoa[0];
        state.prebatchClaims = main_aoa.slice(1).filter(row => String(row[config.claimStatusIndex] || '').toUpperCase().includes('PREBATCH'));

        state.processedClaimsList = processAndAssignClaims(main_aoa, config, state.yesterdayDataMap);

        const noteStats = { miscellaneous: 0, totalWithNotes: 0 };
        state.processedClaimsList.forEach(claim => {
            if (claim.noteText) {
                noteStats.totalWithNotes++;
                // **MODIFIED**: Use the new analysis function to check for "General Investigation"
                if (analyzeClaimNotes(claim.noteText).rootCause === 'General Investigation') {
                    noteStats.miscellaneous++;
                }
            }
        });

        if (noteStats.totalWithNotes > 10 && (noteStats.miscellaneous / noteStats.totalWithNotes > 0.9)) {
            const configuredNoteCol = document.getElementById('notesCol').value.toUpperCase();
            const warningMessage = `Warning: Over 90% of notes were categorized as 'General Investigation'. This often means the configured 'W9/Notes Column' (currently set to column '${configuredNoteCol}') is incorrect for this report. Please verify all column configurations above.`;
            ui.displayWarning(warningMessage);
        }

        // **MODIFIED**: Update header row for new report structure
        state.fileHeaderRow = [...state.mainReportHeader];
        if (state.hasYesterdayFile) {
            state.fileHeaderRow.splice(config.claimStatusIndex, 0, 'Yest. Claim State');
        }
        state.fileHeaderRow.push('Root Cause', 'Assigned To');


        state.processedClaimsList.forEach(claim => {
            const newRow = [...claim.originalRow];
            if (state.hasYesterdayFile) {
                newRow.splice(config.claimStatusIndex, 0, state.yesterdayDataMap.get(claim.claimNumber)?.state || 'NEW');
            }
            // **MODIFIED**: Push new data points to the end of the row
            newRow.push(claim.rootCause, claim.owner);

            claim.processedRow = newRow;
            claim.defaultOwner = claim.owner;
            claim.finalOwner = claim.owner; // Initialize finalOwner with the default
        });

        ui.displayStatus('Processing Complete. Please generate and complete the assignment report.', 'success');
        ui.displayReviewStep();

    } catch (error) {
        ui.displayStatus(`Error: ${error.message}`, 'error');
    }
}

async function handleAssignmentFileUpload(event) {
    const file = event.target.files[0];
    const fileNameDiv = document.getElementById('assignmentFileName');

    if (!file) {
        if(fileNameDiv) fileNameDiv.textContent = 'No file selected.';
        return;
    }

    if(fileNameDiv) fileNameDiv.textContent = `Selected: ${file.name}`;

    ui.displayStatus('Processing assignment file...', 'info', true);

    try {
        const reader = new FileReader();
        reader.onload = (e) => {
            const wb = XLSX.read(new Uint8Array(e.target.result), { type: 'array' });
            const ws = wb.Sheets[wb.SheetNames[0]];
            const jsonData = XLSX.utils.sheet_to_json(ws);

            // Assuming state.assignmentMap exists and is cleared in resetState
            state.assignmentMap.clear();
            let validAssignments = 0;

            jsonData.forEach(row => {
                const claimState = row['Claim State'];
                const noteText = row['Note / Edit Text'];
                const assignee = row['Assign To (Claims or PV)'];

                if (claimState && noteText && assignee) {
                    const assigneeUpper = assignee.toString().trim().toUpperCase();
                    if (assigneeUpper === 'CLAIMS' || assigneeUpper === 'PV') {
                        const key = `${claimState}||${noteText}`;
                        state.assignmentMap.set(key, assigneeUpper);
                        validAssignments++;
                    }
                }
            });

            if (validAssignments > 0) {
                ui.displayStatus(`Successfully loaded ${validAssignments} assignments. Ready to generate final reports.`, 'success');
                document.getElementById('generateFinalReportsBtn').disabled = false;
            } else {
                ui.displayStatus('Assignment file processed, but no valid assignments were found. Please check the file.', 'error');
                document.getElementById('generateFinalReportsBtn').disabled = true;
            }
        };
        reader.readAsArrayBuffer(file);
    } catch (error) {
        ui.displayStatus(`Error reading assignment file: ${error.message}`, 'error');
        document.getElementById('generateFinalReportsBtn').disabled = true;
    }
}


async function generateFinalReports() {
    ui.displayStatus('Applying assignments and generating final reports...', 'info', true);

    // Apply assignments from the map
    state.processedClaimsList.forEach(claim => {
        const noteText = claim.noteText || "No Note";
        const stateStr = claim.claimState || "UNKNOWN";
        const key = `${stateStr}||${noteText}`;
        const assignedOwner = state.assignmentMap.get(key);

        if (assignedOwner) {
            claim.finalOwner = assignedOwner;
            // The owner is the last element in the processedRow array
            claim.processedRow[claim.processedRow.length - 1] = assignedOwner;
        } else {
            // If no specific assignment, it keeps the default owner
            claim.finalOwner = claim.defaultOwner;
            claim.processedRow[claim.processedRow.length - 1] = claim.defaultOwner;
        }
    });

    if (state.hasYesterdayFile) {
        runDetailedCohortAnalysis();
    }

    const clientName = document.getElementById('client-select').options[document.getElementById('client-select').selectedIndex].text;
    const downloadsContainer = document.getElementById('download-links-container');
    downloadsContainer.innerHTML = '';

    // **MODIFIED**: Only a single overall report is generated now
    createDownloadLink(buildWorkbook(state.processedClaimsList, `${clientName} Daily Action Report`), `${clientName} Daily Action Report for ${ui.getFormattedDate()}.xlsx`, downloadsContainer);

    if (state.hasYesterdayFile) {
        const pdfButton = document.createElement('button');
        pdfButton.id = 'downloadPdfReportBtn';
        pdfButton.className = 'bg-red-600 text-white font-bold text-lg rounded-lg py-3 px-6 hover:bg-red-700';
        pdfButton.textContent = 'Download PDF Analysis';
        pdfButton.onclick = async () => {
            ui.displayStatus('Generating PDF report...', 'info', true);
            await generatePdfReport();
            ui.displayStatus('PDF Report generated!', 'success');
        };
        downloadsContainer.appendChild(pdfButton);
    }

    document.getElementById('final-downloads-container').classList.remove('hidden');
    document.getElementById('copyEmailBtn').classList.remove('hidden');
    ui.displayStatus('Excel reports generated successfully!', 'success');
}

function createDownloadLink(blob, fileName, container) {
    const link = document.createElement('a');
    link.href = URL.createObjectURL(blob);
    link.download = fileName;
    link.textContent = `Download ${fileName.split(' for')[0]}`;
    link.className = 'inline-block bg-indigo-600 text-white font-bold text-lg rounded-lg py-3 px-6 hover:bg-indigo-700';
    container.appendChild(link);
}

function copyEmailText() {
    const clientName = document.getElementById('client-select').options[document.getElementById('client-select').selectedIndex].text;
    let emailBody = `Hello Teams,\n\nAttached is today's Daily Action Report for ${clientName}.`;

    if (state.hasYesterdayFile) {
        const todayStats = calculateStats(state.processedClaimsList);
        const formatStatLine = (today, yesterday) => `${today} (Yest. ${yesterday ?? 0})`;

        const createStatBlock = (title, statusKey) => {
            const yestBlock = state.yesterdayStats?.[statusKey] || { total: 0, '28-30': 0, '21-27': 0, '31+': 0, '0-20': 0 };
            const todayBlock = todayStats?.[statusKey] || { total: 0, '28-30': 0, '21-27': 0, '31+': 0, '0-20': 0 };
            return `Number of total claims ${title}: ${formatStatLine(todayBlock.total, yestBlock.total)}\n` +
               `CRITICAL (28-30 Days): ${formatStatLine(todayBlock['28-30'], yestBlock['28-30'])}\n` +
               `PRIORITY (21-27 Days): ${formatStatLine(todayBlock['21-27'], yestBlock['21-27'])}\n` +
               `Backlog (31+ Days): ${formatStatLine(todayBlock['31+'], yestBlock['31+'])}\n` +
               `Queue (0-20 Days): ${formatStatLine(todayBlock['0-20'], yestBlock['0-20'])}`;
        };

        emailBody += `\n\nBelow are the detailed highlights from the report. For a full visual breakdown of claim movement, please see the attached 'Daily Claim-Flow Analysis' PDF.\n\n` +
              `${createStatBlock('pending', 'PEND')}\n\n` +
              `${createStatBlock('On Hold', 'ONHOLD')}\n\n` +
              `${createStatBlock('in Management Review', 'MANAGEMENTREVIEW')}`;
    }

    emailBody += `\n\nPlease let me know if you have any questions.`;

    navigator.clipboard.writeText(emailBody).then(() => {
        const btn = document.getElementById('copyEmailBtn');
        btn.textContent = 'Copied!';
        setTimeout(() => { btn.textContent = 'Copy Email Text'; }, 2000);
    }).catch(err => alert('Failed to copy text.'));
}

function runDetailedCohortAnalysis() {
    state.workflowMovement = { pvToClaims: 0, claimsToPv: 0, criticalToBacklog: 0, criticalWorked: 0 };
    state.detailedMovementStats = {};

    const todayClaimsMap = new Map(state.processedClaimsList.map(c => [c.claimNumber, c]));
    const config = gatherConfig();
    const prebatchClaimNumbers = new Set(state.prebatchClaims.map(row => String(row[config.claimNumberIndex] || '').trim()));

    const getBucketFromAge = (age) => {
        if (isNaN(age)) return 'UNKNOWN';
        if (age >= 28 && age <= 30) return 'Critical';
        if (age >= 31) return 'Backlog';
        if (age >= 21 && age <= 27) return 'Priority';
        return 'Queue';
    };

    const getStateKey = (stateStr) => (stateStr || '').includes('MANAGEMENT') ? 'MANAGEMENTREVIEW' : (stateStr || '');

    const yestCohorts = {};
    for (const [claimNumber, yestData] of state.yesterdayDataMap.entries()) {
        const yestState = getStateKey(yestData.state);
        const yestBucket = getBucketFromAge(yestData.cleanAge);
        if (!yestCohorts[yestState]) yestCohorts[yestState] = {};
        if (!yestCohorts[yestState][yestBucket]) yestCohorts[yestState][yestBucket] = [];
        yestCohorts[yestState][yestBucket].push(claimNumber);

        if(todayClaimsMap.has(claimNumber)){
            const todayOwner = todayClaimsMap.get(claimNumber).finalOwner;
            if(yestData.owner === 'PV' && todayOwner === 'Claims') state.workflowMovement.pvToClaims++;
            if(yestData.owner === 'Claims' && todayOwner === 'PV') state.workflowMovement.claimsToPv++;
        }
    }

    for(const yestState in yestCohorts){
        state.detailedMovementStats[yestState] = {};
        for(const yestBucket in yestCohorts[yestState]){
            const claimNumbersInCohort = yestCohorts[yestState][yestBucket];
            const breakdown = {
                totalYesterday: claimNumbersInCohort.length, movedToPrebatch: 0, resolvedOrRemoved: 0, movedTo: {}
            };

            for(const claimNumber of claimNumbersInCohort){
                if(prebatchClaimNumbers.has(claimNumber)){
                    breakdown.movedToPrebatch++;
                } else if (todayClaimsMap.has(claimNumber)){
                    const todayClaim = todayClaimsMap.get(claimNumber);
                    const todayState = getStateKey(todayClaim.claimState);
                    const todayBucket = getBucketFromAge(todayClaim.cleanAge);
                    const destinationKey = `${todayState}_${todayBucket}`;
                    breakdown.movedTo[destinationKey] = (breakdown.movedTo[destinationKey] || 0) + 1;
                } else {
                    breakdown.resolvedOrRemoved++;
                }
            }
            state.detailedMovementStats[yestState][yestBucket] = breakdown;
        }
    }

    if(state.detailedMovementStats.PEND && state.detailedMovementStats.PEND.Critical){
         const critPend = state.detailedMovementStats.PEND.Critical;
         state.workflowMovement.criticalWorked = (critPend.resolvedOrRemoved || 0) + (critPend.movedToPrebatch || 0);
         state.workflowMovement.criticalToBacklog = 0;
         for(const dest in critPend.movedTo){
             if(dest.includes('Backlog')) {
                 state.workflowMovement.criticalToBacklog += critPend.movedTo[dest];
             } else {
                const [, newBucket] = dest.split('_');
                if (newBucket !== 'Critical') {
                    state.workflowMovement.criticalWorked += critPend.movedTo[dest];
                }
             }
         }
    }

    document.getElementById('pv-to-claims-count').textContent = state.workflowMovement.pvToClaims.toLocaleString();
    document.getElementById('claims-to-pv-count').textContent = state.workflowMovement.claimsToPv.toLocaleString();
    document.getElementById('critical-to-backlog-count').textContent = state.workflowMovement.criticalToBacklog.toLocaleString();
    document.getElementById('critical-worked-count').textContent = state.workflowMovement.criticalWorked.toLocaleString();
    document.getElementById('movement-summary-container').classList.remove('hidden');
}
