// js/main.js

import { updateUI, handleDownload } from './ui.js';
import { handleFileUpload } from './fileHandler.js';

// Import initializers from every action
import { initializeAnonymizeAction } from './actions/anonymize.js';
import { initializeTrimWhitespaceAction } from './actions/trimWhitespace.js';
import { initializeExtractColumnsAction } from './actions/extractColumns.js';
import { initializeStackSheetsAction } from './actions/stackSheets.js';
import { initializeMergeFilesAction } from './actions/mergeFiles.js';
import { initializeFindDuplicatesAction } from './actions/findDuplicates.js';
import { initializeCompareSheetsAction } from './actions/compareSheets.js';
import { initializeSplitFileAction } from './actions/splitFile.js';

document.addEventListener('DOMContentLoaded', () => {
    // Initialize core functionality
    document.getElementById('file-upload').addEventListener('change', async (event) => {
        await handleFileUpload(event);
        updateUI(); // Update UI after files are loaded
    });
    document.getElementById('download-btn').addEventListener('click', handleDownload);

    // Initialize all the individual actions
    initializeTrimWhitespaceAction();
    initializeAnonymizeAction();
    initializeExtractColumnsAction();
    initializeStackSheetsAction();
    initializeMergeFilesAction();
    initializeFindDuplicatesAction();
    initializeCompareSheetsAction();
    initializeSplitFileAction();

    // Initial UI render
    updateUI();
});
