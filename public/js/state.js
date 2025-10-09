// This file manages the shared data and state for the tool.
// By centralizing the state, we ensure that all parts of the application
// are working with the same data.

export const state = {
    processedClaimsList: [],
    prebatchClaims: [],
    fileHeaderRow: [],
    mainReportHeader: [],
    yesterdayDataMap: new Map(),
    yesterdayStats: null,
    workflowMovement: { pvToClaims: 0, claimsToPv: 0, criticalToBacklog: 0, criticalWorked: 0 },
    prebatchMovementStats: {},
    detailedMovementStats: {},
    cycleTimeMetrics: {},
    assignmentMap: new Map(), // **THIS LINE WAS MISSING**
    hasYesterdayFile: false
};

// Resets the state to its initial values. This is crucial for ensuring
// that running the tool a second time doesn't use leftover data from the
// previous run.
export function resetState() {
    state.processedClaimsList = [];
    state.prebatchClaims = [];
    state.fileHeaderRow = [];
    state.mainReportHeader = [];
    state.yesterdayDataMap.clear();
    state.yesterdayStats = null;
    state.hasYesterdayFile = false;
    state.workflowMovement = { pvToClaims: 0, claimsToPv: 0, criticalToBacklog: 0, criticalWorked: 0 };
    state.prebatchMovementStats = {};
    state.detailedMovementStats = {};
    state.cycleTimeMetrics = {};
    state.assignmentMap.clear(); // Ensure the map is cleared on reset
}
