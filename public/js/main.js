// js/main.js

import { state } from './state.js';
import { updateUI } from './ui.js';
import { handleFileUpload } from './fileHandler.js';

// Import initializers from every action
import { initializeAnonymizeAction } from './actions/anonymize.js';
import { initializeTrimWhitespaceAction } from './actions/trimWhitespace.js';
// ... import all other action initializers

document.addEventListener('DOMContentLoaded', () => {
    // Initialize core functionality
    document.getElementById('file-upload').addEventListener('change', async (event) => {
        await handleFileUpload(event, state); // Pass state to the handler
        updateUI(); // Update UI after files are loaded
    });

    // Initialize all the individual actions
    initializeAnonymizeAction();
    initializeTrimWhitespaceAction();
    // ... initialize all other actions

    // Initial UI render
    updateUI();
});
