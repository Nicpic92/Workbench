// js/state.js

// The single source of truth for your application's data
export const state = {
    datasets: [],
    activeDatasetIndex: 0,
};

// A helper function to easily get the currently selected dataset
export function getActiveDataset() {
    return state.datasets[state.activeDatasetIndex];
}

// A function to add a new dataset to the state
export function addNewDataset(name, data, headers) {
    state.datasets.push({ name, data, headers });
    state.activeDatasetIndex = state.datasets.length - 1;
}
