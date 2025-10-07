// This file holds configuration details, making them easy to update
// without searching through logic files.

// Presets for different clients, defining the column letters for key data points.
export const clientPresets = {
    solis: { label: 'Clean Age (Q):', cleanAgeCol: 'Q', claimStatusCol: 'I', payerCol: 'A', networkStatusCol: 'V', dsnpCol: 'Y', claimTypeCol: 'B', totalChargesCol: 'T', dateCols: 'E,O,P', notesCol: 'AA', claimNumberCol: 'C' },
    liberty: { label: 'Age (R):', cleanAgeCol: 'R', claimStatusCol: 'I', payerCol: 'A', networkStatusCol: 'V', dsnpCol: 'Y', claimTypeCol: 'B', totalChargesCol: 'T', dateCols: 'E,O,P', notesCol: 'AA', claimNumberCol: 'C' },
    
    // START: This is the section that has been updated
    secur: { 
        label: 'Clean Age (Q):', 
        cleanAgeCol: 'Q', 
        claimStatusCol: 'I', 
        payerCol: 'A', 
        networkStatusCol: 'V', // <-- This was changed from 'U' to 'V' to match your image
        dsnpCol: 'Y', 
        claimTypeCol: 'D', 
        totalChargesCol: 'T', 
        dateCols: 'E,O,P', 
        notesCol: 'AA', 
        claimNumberCol: 'C' 
    },
    // END: Update complete
    
    csh: { label: 'Age (R):', cleanAgeCol: 'R', claimStatusCol: 'I', payerCol: 'A', networkStatusCol: 'U', dsnpCol: 'Y', claimTypeCol: 'B', totalChargesCol: 'T', dateCols: 'E,O,P', notesCol: 'AA', claimNumberCol: 'C' }
};

// Gathers the column configurations from the UI input fields and converts
// Excel-style column letters (A, B, AA) into zero-based numeric indices.
export function gatherConfig() {
    try {
        const colLetterToIndex = (letter) => {
            if (!letter || !/^[A-Z]+$/i.test(letter)) return -1;
            let col = 0;
            letter = letter.toUpperCase();
            for (let i = 0; i < letter.length; i++) {
                col += (letter.charCodeAt(i) - 64) * Math.pow(26, letter.length - i - 1);
            }
            return col - 1;
        };
        const config = {};
        const ids = ['cleanAge', 'claimStatus', 'claimNumber', 'payer', 'networkStatus', 'dsnp', 'claimType', 'totalCharges', 'notes'];
        ids.forEach(id => {
            config[`${id}Index`] = colLetterToIndex(document.getElementById(`${id}Col`).value);
        });

        // Basic validation
        if (Object.values(config).some(val => val === -1)) {
            throw new Error("Invalid or empty column letter entered.");
        }
        return config;
    } catch (error) {
        // The error will be caught and displayed by the calling function in main.js
        console.error(`Configuration Error: ${error.message}`);
        return null;
    }
}
