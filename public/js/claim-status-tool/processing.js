// This module contains the core data processing and analysis logic.
// These functions are "pure" where possible, meaning they take data as input
// and return new data, without directly affecting the webpage. This makes
// them easier to test and reason about.

export function getHardcodedAssignment(noteText) { 
    // This function can be expanded with specific rules.
    // For now, it returns null to let the main logic handle assignment.
    return null; 
}
    
export function getNoteCategory(noteText) {
    const noteLower = noteText.toLowerCase();
    if (noteLower.includes('w9 req') || noteLower.includes('w9 requested') || noteLower.includes('w9 recvd') || noteLower.includes('w9 past due')) return 'W9 Form Management';
    if (noteLower.includes('rachael') || noteLower.includes('payer review') || noteLower.includes('hold for') || noteLower.includes('red tab') || noteLower.includes('move to pr')) return 'Manual Review (Rachael/Payer)';
    if (noteLower.startsWith('error -') || noteLower.startsWith('error-') || noteLower.includes('incorrectly') || noteLower.includes('missed to')) return 'Adjudication & Processing Errors';
    if (noteLower.includes('payment >') || noteLower.includes('net pay >') || noteLower.includes('>$10000') || noteLower.includes('exceeds total payment') || noteLower.includes('10k')) return 'High-Dollar Amount Review';
    if (noteLower.includes('contract') || noteLower.includes('provider not found') || noteLower.includes('no data for') || noteLower.includes('pay to name mismatch')) return 'Contract & Provider Data Issues';
    if (noteLower.includes('remap') || noteLower.includes('rerun') || noteLower.includes('reprocess') || noteLower.includes('pv updated')) return 'System Actions & Reprocessing';
    if (noteLower.includes('auth') || noteLower.includes('duplicate')) return 'Authorization & Duplicate Issues';
    return 'Miscellaneous';
}

export function processAndAssignClaims(aoa, config, yesterdayDataMap) {
    const claims = [];
    // Skip header row
    for (const row of aoa.slice(1)) {
        // Skip empty rows
        if (row.every(c => c === null)) continue;
        
        const claimState = String(row[config.claimStatusIndex] || '').trim().toUpperCase();
        // Skip prebatch claims from the main processing
        if (claimState.includes('PREBATCH')) continue;

        const noteText = String(row[config.notesIndex] || '');
        const cleanAge = parseInt(row[config.cleanAgeIndex], 10);
        let owner = null;
        
        // Use yesterday's owner if available
        const claimNumber = String(row[config.claimNumberIndex] || '').trim();
        if (claimNumber && yesterdayDataMap.has(claimNumber)) {
            owner = yesterdayDataMap.get(claimNumber).owner;
        }

        // If no owner from yesterday, determine one based on rules
        if (!owner) {
            owner = getHardcodedAssignment(noteText);
            if (owner === null) { // If no hardcoded rule matched
                const totalCharges = parseFloat(String(row[config.totalChargesIndex]).replace(/[^0-9.-]/g, '')) || 0;
                const claimType = String(row[config.claimTypeIndex] || '').toUpperCase();
                const isMgmtReview = claimState.includes('MANAGEMENT') && claimState.includes('REVIEW');
                const isHighCost = isMgmtReview && ((claimType.includes('PROFESSIONAL') && totalCharges > 3500) || (claimType.includes('INSTITUTIONAL') && totalCharges > 6500));
                
                if (isHighCost) owner = 'Claims';
                else if (isMgmtReview || claimState === 'ONHOLD') owner = 'PV';
                else if (['PEND', 'APPROVED', 'DENY'].includes(claimState)) owner = 'Claims';
                else if (claimState === 'PR') owner = row[config.payerIndex] || ''; // Assign to payer name
                else owner = 'PV'; // Default
            }
        }
        claims.push({ ...config, originalRow: row, noteText, cleanAge, claimState, owner, claimNumber });
    }
    return claims;
}

export function processYesterdayReport(aoa) {
    const headers = (aoa[0] || []).map(h => String(h || '').trim());
    const claimNumIndex = headers.indexOf('Claim Number');
    const ownerIndex = headers.indexOf('Added (Owner)');
    const stateIndex = headers.indexOf('Claim State');
    const cleanAgeHeaderNames = ['Clean Age', 'Age']; 
    let cleanAgeIndex = -1;

    for (const name of cleanAgeHeaderNames) {
        const idx = headers.findIndex(h => h.startsWith(name));
        if (idx !== -1) { 
            cleanAgeIndex = idx; 
            break; 
        }
    }
    
    if (claimNumIndex === -1 || ownerIndex === -1 || cleanAgeIndex === -1) {
        throw new Error("Yesterday's report must contain 'Claim Number', 'Added (Owner)', and a 'Clean Age'/'Age' column.");
    }

    const dataMap = new Map();
    const claimsList = [];
    for (const row of aoa.slice(1)) {
        if (row.every(c => c === null)) continue;
        const claimNumber = String(row[claimNumIndex] || '').trim();
        const owner = String(row[ownerIndex] || '').trim();
        const claimState = String(row[stateIndex] || '').trim().toUpperCase();
        const cleanAge = parseInt(row[cleanAgeIndex], 10);
        if (claimNumber && owner) {
            dataMap.set(claimNumber, { state: claimState, owner: owner, cleanAge: cleanAge });
        }
        claimsList.push({ claimState, cleanAge, owner });
    }
    return { stats: calculateStats(claimsList), dataMap };
}

export function calculateStats(claimsList) {
    const dayBuckets = { '28-30': 0, '21-27': 0, '31+': 0, '0-20': 0 };
    const stats = { 
        'PEND': { total: 0, ...dayBuckets }, 
        'ONHOLD': { total: 0, ...dayBuckets }, 
        'MANAGEMENTREVIEW': { total: 0, ...dayBuckets }, 
        'DENY': { total: 0, ...dayBuckets }, 
        'PR': { total: 0, ...dayBuckets }, 
        'APPROVED': { total: 0, ...dayBuckets} 
    };
    for(const claim of claimsList) {
        let finalClaimState = (claim.claimState || '').includes('MANAGEMENT') ? 'MANAGEMENTREVIEW' : (claim.claimState || '');
        if (stats[finalClaimState]) {
            stats[finalClaimState].total++;
            let daysValue = '';
            if (!isNaN(claim.cleanAge)) {
                if (claim.cleanAge >= 28 && claim.cleanAge <= 30) daysValue = '28-30';
                else if (claim.cleanAge >= 21 && claim.cleanAge <= 27) daysValue = '21-27';
                else if (claim.cleanAge >= 31) daysValue = '31+';
                else daysValue = '0-20';
                
                if (stats[finalClaimState][daysValue] !== undefined) {
                    stats[finalClaimState][daysValue]++;
                }
            }
        }
    }
    return stats;
}

export function calculateCycleTimeMetrics(processedClaimsList) {
    const metrics = {
        total_clean_nonpar: 0, met_goal_clean_nonpar: 0,
        total_other_nonpar: 0, met_goal_other_nonpar: 0,
        total_clean_par: 0,    met_goal_clean_par: 0,
        total_other_par: 0,    met_goal_other_par: 0,
    };
    const cleanStates = ['PEND', 'APPROVED', 'DENY', 'PR'];
    
    for (const claim of processedClaimsList) {
        const isNonPar = String(claim.originalRow[claim.networkStatusIndex] || '').toUpperCase().includes('OUT');
        const isClean = cleanStates.includes(claim.claimState) || claim.claimState.includes('MANAGEMENT') === false;

        if (isNonPar) {
            if (isClean) {
                metrics.total_clean_nonpar++;
                if (claim.cleanAge <= 30) metrics.met_goal_clean_nonpar++;
            } else {
                metrics.total_other_nonpar++;
                if (claim.cleanAge <= 60) metrics.met_goal_other_nonpar++;
            }
        } else { // Is Par
            if (isClean) {
                metrics.total_clean_par++;
                if (claim.cleanAge <= 30) metrics.met_goal_clean_par++;
            } else {
                metrics.total_other_par++;
                if (claim.cleanAge <= 60) metrics.met_goal_other_par++;
            }
        }
    }

    const calcRate = (met, total) => (total > 0 ? (met / total * 100) : 0).toFixed(2) + '%';

    return {
        cleanNonPar30: calcRate(metrics.met_goal_clean_nonpar, metrics.total_clean_nonpar),
        otherNonPar60: calcRate(metrics.met_goal_other_nonpar, metrics.total_other_nonpar),
        cleanPar30:    calcRate(metrics.met_goal_clean_par, metrics.total_clean_par),
        otherPar60:    calcRate(metrics.met_goal_other_par, metrics.total_other_par),
    };
}
