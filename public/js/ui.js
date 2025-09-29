// js/ui.js

import { state, getActiveDataset } from './state.js';

// Get all DOM elements once
const welcomeView = document.getElementById('welcome-view');
const dataView = document.getElementById('data-view');
const actionsContainer = document.getElementById('actions-container');
const loadedFilesList = document.getElementById('loaded-files-list');
const tableTitle = document.getElementById('table-title');
const tableContainer = document.getElementById('data-table-container');
const statusBar = document.getElementById('status-bar');
const loaderOverlay = document.getElementById('loader-overlay');

// Modal Elements
const configModal = document.getElementById('config-modal');
const modalTitle = document.getElementById('modal-title');
const modalBody = document.getElementById('modal-body');
let modalConfirmBtn = document.getElementById('modal-confirm-btn');
const modalCancelBtn = document.getElementById('modal-cancel-btn');
const modalCloseBtn = document.getElementById('modal-close-btn');

// --- Core UI Functions ---

export function updateUI() {
    if (state.datasets.length === 0) {
        welcomeView.style.display = 'flex';
        dataView.classList.add('hidden');
        actionsContainer.style.display = 'none';
        loadedFilesList.innerHTML = '';
    } else {
        welcomeView.style.display = 'none';
        actionsContainer.style.display = 'block';
        dataView.classList.remove('hidden');
        renderLoadedFilesList();
        renderActiveDataset();
    }
}

function renderLoadedFilesList() {
    loadedFilesList.innerHTML = '';
    state.datasets.forEach((ds, index) => {
        const item = document.createElement('div');
        item.className = 'loaded-file-item';
        if (index === state.activeDatasetIndex) {
            item.classList.add('active');
        }
        item.textContent = ds.name;
        item.onclick = () => {
            state.activeDatasetIndex = index;
            updateUI();
        };
        loadedFilesList.appendChild(item);
    });
}

function renderActiveDataset() {
    const activeDataset = getActiveDataset();
    if (!activeDataset) return;
    tableTitle.textContent = activeDataset.name;
    renderDataTable(activeDataset.data, activeDataset.headers);
    statusBar.textContent = `Displaying ${activeDataset.data.length.toLocaleString()} rows and ${activeDataset.headers.length} columns. (Preview of first 200 rows)`;
}

function formatCellValue(value) {
    if (value === null || value === undefined) return '';
    if (value instanceof Date) {
        if (isNaN(value.getTime())) return String(value);
        const month = String(value.getUTCMonth() + 1).padStart(2, '0');
        const day = String(value.getUTCDate()).padStart(2, '0');
        const year = value.getUTCFullYear();
        return `${month}/${day}/${year}`;
    }
    return String(value);
}

function renderDataTable(data, headers) {
    const table = document.createElement('table');
    const thead = table.createTHead();
    const headerRow = thead.insertRow();
    headers.forEach(h => {
        const th = document.createElement('th');
        th.textContent = h;
        headerRow.appendChild(th);
    });
    const tbody = table.createTBody();
    data.slice(0, 200).forEach(row => {
        const tr = tbody.insertRow();
        headers.forEach(header => {
             const td = tr.insertCell();
             td.textContent = formatCellValue(row[header]);
        });
    });
    tableContainer.innerHTML = '';
    tableContainer.appendChild(table);
}

// --- Helper & Utility UI Functions ---

export function showLoader(show) {
    loaderOverlay.style.display = show ? 'flex' : 'none';
}

export function closeModal(id) {
    document.getElementById(id).classList.add('hidden');
}

export function showConfigModal(title, content, onConfirm) {
    modalTitle.textContent = title;
    modalBody.innerHTML = content;
    configModal.style.display = 'flex';
    // Clone and replace button to remove old listeners
    const newConfirmBtn = modalConfirmBtn.cloneNode(true);
    modalConfirmBtn.parentNode.replaceChild(newConfirmBtn, modalConfirmBtn);
    newConfirmBtn.addEventListener('click', onConfirm);
    modalConfirmBtn = newConfirmBtn;
    
    // Wire up close buttons for this specific modal instance
    modalCancelBtn.onclick = () => closeModal('config-modal');
    modalCloseBtn.onclick = () => closeModal('config-modal');
}

export function generateColumnCheckboxes(headers) {
    return headers.map(h => `<label class="flex items-center p-2 rounded hover:bg-slate-100"><input type="checkbox" class="h-4 w-4 rounded mr-2" data-column-name="${h}"><span class="text-sm">${h}</span></label>`).join('');
}

export function generateColumnSelect(headers, id) {
    return `<select id="${id}" class="w-full p-2 border rounded mt-1">${headers.map(h => `<option value="${h}">${h}</option>`).join('')}</select>`;
}

export function generateDatasetSelect(id) {
    return `<select id="${id}" class="w-full p-2 border rounded mt-1">${state.datasets.map((ds, i) => `<option value="${i}">${ds.name}</option>`).join('')}</select>`;
}

export function handleDownload() {
    const activeDataset = getActiveDataset();
    if (!activeDataset) return;
    showLoader(true);
    setTimeout(() => {
        const ws = XLSX.utils.json_to_sheet(activeDataset.data);
        const wb = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(wb, ws, "Result");
        XLSX.writeFile(wb, `Processed_${activeDataset.name.replace(/[^a-z0-9]/gi, '_').toLowerCase()}.xlsx`);
        showLoader(false);
    }, 50);
}
