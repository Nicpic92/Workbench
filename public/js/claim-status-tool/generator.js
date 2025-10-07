import { state } from './state.js';
import { getFormattedDate } from './ui.js';
import { gatherConfig } from './config.js'; // Needed for PDF generation
import { calculateStats } from './processing.js'; // Needed for email text

// This module is responsible for generating all final file outputs (Excel, PDF, etc.)

// --- Excel Report Generation ---

export function downloadPrebatchReport() {
    const clientName = document.getElementById('client-select').options[document.getElementById('client-select').selectedIndex].text;
    const ws = XLSX.utils.aoa_to_sheet([state.mainReportHeader, ...state.prebatchClaims]);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Prebatch Claims");
    XLSX.writeFile(wb, `${clientName} Prebatch Report for ${getFormattedDate()}.xlsx`);
}

export function buildWorkbook(claimsData, reportTitle, ownerFilter = null) {
    const masterSheetName = "All Processed Data", highDollarSheetName = "High Dollar";
    const sheetsData = { [masterSheetName]: [state.fileHeaderRow] };
    if (ownerFilter !== 'PV') sheetsData[highDollarSheetName] = [state.fileHeaderRow];
    
    const tabMetadata = {};
    const overallSummary = { par: { '28-29': 0, '21-27': 0, '30+': 0, '0-20': 0, total: 0 }, nonpar: { '28-29': 0, '21-27': 0, '30+': 0, '0-20': 0, total: 0 } };

    for (const claim of claimsData) {
        const networkType = String(claim.originalRow[claim.networkStatusIndex] || '').toUpperCase().includes('OUT') ? 'nonpar' : 'par';
        overallSummary[networkType].total++;
        if (!isNaN(claim.cleanAge)) {
            if (claim.cleanAge >= 28 && claim.cleanAge <= 29) overallSummary[networkType]['28-29']++;
            else if (claim.cleanAge >= 21 && claim.cleanAge <= 27) overallSummary[networkType]['21-27']++;
            else if (claim.cleanAge >= 30) overallSummary[networkType]['30+']++;
            else overallSummary[networkType]['0-20']++;
        }
        
        sheetsData[masterSheetName].push(claim.processedRow);
        
        const totalCharges = parseFloat(String(claim.originalRow[claim.totalChargesIndex]).replace(/[^0-9.-]/g, '')) || 0;
        const claimType = String(claim.originalRow[claim.claimTypeIndex] || '').toUpperCase();
        const isMgmtReview = claim.claimState.includes('MANAGEMENT') && claim.claimState.includes('REVIEW');
        const isHighCost = isMgmtReview && ((claimType.includes('PROFESSIONAL') && totalCharges > 3500) || (claimType.includes('INSTITUTIONAL') && totalCharges > 6500));
        
        if (isHighCost && ownerFilter !== 'PV') sheetsData[highDollarSheetName].push(claim.processedRow);
        
        const dsnpRaw = String(claim.originalRow[claim.dsnpIndex] || '').toUpperCase();
        const dsnpStatus = dsnpRaw.includes('NON DSNP') ? 'NonDSNP' : (dsnpRaw.includes('DSNP') || dsnpRaw === 'Y' ? 'DSNP' : '');
        
        let statusTab = '', tabOwner = '';
        if (isMgmtReview) { statusTab = 'MgmtRev'; tabOwner = 'PV'; }
        else if (claim.claimState === 'ONHOLD') { statusTab = 'OnHold'; tabOwner = 'PV'; }
        else if (claim.claimState === 'PEND') { statusTab = 'Pend'; tabOwner = 'Claims'; }
        else if (claim.claimState === 'DENY') { statusTab = 'Deny'; tabOwner = 'Claims'; }
        else if (claim.claimState === 'PR') { statusTab = 'PayerRev'; tabOwner = 'Claims'; }
        
        if (dsnpStatus && tabOwner && networkType && (ownerFilter === 'PV' || !isHighCost)) {
            let tabKey = '', priorityLevel = 0;
            if (claim.cleanAge >= 28 && claim.cleanAge <= 29) { priorityLevel = 1; tabKey = `CRITICAL (28-29d) ${networkType === 'par' ? 'Par' : 'NonPar'} ${statusTab} ${dsnpStatus}`; }
            else if (claim.cleanAge >= 21 && claim.cleanAge <= 27) { priorityLevel = 2; tabKey = `PRIORITY (21-27d) ${networkType === 'par' ? 'Par' : 'NonPar'} ${statusTab} ${dsnpStatus}`; }
            else if (claim.cleanAge >= 30) { priorityLevel = 3; tabKey = `Backlog (30+d) ${networkType === 'par' ? 'Par' : 'NonPar'} ${statusTab} ${dsnpStatus}`; }
            else { priorityLevel = 4; tabKey = `Queue (0-20d) ${networkType === 'par' ? 'Par' : 'NonPar'} ${statusTab} ${dsnpStatus}`; }
            
            if (!(ownerFilter && priorityLevel === 4)) {
                const truncatedTabKey = tabKey.replace(/[\\\/\?\*\[\]]/g, '').substring(0, 31);
                if (!sheetsData[truncatedTabKey]) { sheetsData[truncatedTabKey] = [state.fileHeaderRow]; tabMetadata[truncatedTabKey] = { owner: tabOwner, priority: priorityLevel }; }
                sheetsData[truncatedTabKey].push(claim.processedRow);
            }
        }

        const noteLower = claim.noteText.toLowerCase();
        if (noteLower.includes('w9')) {
            let w9SheetName = '', w9Owner = '';
            if (noteLower.includes('req')) { w9SheetName = 'W9 Follow-Up'; w9Owner = 'Claims'; }
            else if (noteLower.includes('denied') || noteLower.includes('missing')) { w9SheetName = 'W9 Letter Needed'; w9Owner = 'PV'; }
            else if (noteLower.includes('received') || noteLower.includes('reprocess')) { w9SheetName = 'W9 Received - Reprocess'; w9Owner = 'Claims'; }
            if (w9SheetName) {
                if (!sheetsData[w9SheetName]) { sheetsData[w9SheetName] = [state.fileHeaderRow]; tabMetadata[w9SheetName] = { owner: w9Owner, priority: 5 }; }
                sheetsData[w9SheetName].push(claim.processedRow);
            }
        }
    }

    const coverPageData = [[reportTitle], [`Date: ${getFormattedDate()}`], [], ["Overall Claim Summary"], ["Category", "28-29 Days (Critical)", "21-27 Days (Priority)", "30+ Days (Backlog)", "0-20 Days (Queue)", "Total Active Claims"], ["Par Claims", overallSummary.par['28-29'], overallSummary.par['21-27'], overallSummary.par['30+'], overallSummary.par['0-20'], overallSummary.par.total], ["Non-Par Claims", overallSummary.nonpar['28-29'], overallSummary.nonpar['21-27'], overallSummary.nonpar['30+'], overallSummary.nonpar['0-20'], overallSummary.nonpar.total], [], ["Core Strategy: Focus on claims nearing the 30-day threshold. Work tabs in priority order."], []];
    const allBreakoutTabs = Object.keys(tabMetadata);

    const addSectionToCover = (title, priority) => {
        const filteredTabs = allBreakoutTabs.filter(key => tabMetadata[key]?.priority === priority && (!ownerFilter || tabMetadata[key].owner === ownerFilter)).sort();
        if (filteredTabs.length > 0) {
            coverPageData.push([title], ["Tab Name", "Claim Count", "Assigned Owner"]);
            filteredTabs.forEach(key => {
                const totalCount = sheetsData[key].length - 1;
                let pvCount = 0, claimsCount = 0;
                for (const row of sheetsData[key].slice(1)) { if (row[state.fileHeaderRow.length - 2] === 'PV') pvCount++; else if (row[state.fileHeaderRow.length - 2] === 'Claims') claimsCount++; }
                coverPageData.push([key, totalCount, `PV (${pvCount}) Claims (${claimsCount})`]);
            });
            coverPageData.push([]);
        }
    };
    
    addSectionToCover("Priority 1: CRITICAL (28-29 days)", 1);
    addSectionToCover("Priority 2: PRIORITY (21-27 days)", 2);
    addSectionToCover("Priority 3: Backlog (30+ days)", 3);
    addSectionToCover("W9 and Other Tasks", 5);

    const wb = XLSX.utils.book_new();
    const coverWS = XLSX.utils.aoa_to_sheet(coverPageData);
    coverWS['!cols'] = [{ wch: 35 }, { wch: 20 }, { wch: 22 }];
    XLSX.utils.book_append_sheet(wb, coverWS, "Cover Page");
    
    let sheetOrder = ["Cover Page"];
    if (ownerFilter !== 'PV') sheetOrder.push(highDollarSheetName);
    sheetOrder.push(...allBreakoutTabs.sort((a,b) => (tabMetadata[a].priority - tabMetadata[b].priority) || a.localeCompare(b)), masterSheetName);
    
    sheetOrder.forEach(sheetName => { 
        if (sheetsData[sheetName] && sheetsData[sheetName].length > 1) { 
            const ws = XLSX.utils.aoa_to_sheet(sheetsData[sheetName]);
            ws['!autofilter'] = { ref: ws['!ref'] }; 
            XLSX.utils.book_append_sheet(wb, ws, sheetName); 
        } 
    });
    
    return new Blob([XLSX.write(wb, { bookType: 'xlsx', type: 'array' })], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
}


// --- PDF Report Generation ---

export async function generatePdfReport() {
    const { jsPDF } = window.jspdf;
    const doc = new jsPDF({ orientation: 'p', unit: 'mm', format: 'a4' });
    
    await createTitlePage(doc);
    doc.addPage();
    await createChartsPage(doc);
    doc.addPage();
    await createDetailedTablesPage(doc);

    const pageCount = doc.internal.getNumberOfPages();
    for (let i = 1; i <= pageCount; i++) {
        doc.setPage(i);
        doc.setFontSize(8);
        doc.setTextColor(150);
        doc.text('Spreadsheet Simplicity - Confidential & Proprietary', 14, 290);
        doc.text(`Page ${i} of ${pageCount}`, 190, 290);
    }

    const clientName = document.getElementById('client-select').value;
    doc.save(`${clientName.toUpperCase()}_Daily_Claim-Flow_Analysis_${getFormattedDate()}.pdf`);
}

function formatCurrency(value) {
    if (value === null || isNaN(value)) return '$0';
    if (value < 1000) return `$${value.toFixed(0)}`;
    if (value < 1000000) return `$${(value / 1000).toFixed(1)}K`;
    return `$${(value / 1000000).toFixed(1)}M`;
}

async function createTitlePage(doc) {
    const clientName = document.getElementById('client-select').options[document.getElementById('client-select').selectedIndex].text;
    let currentY = 20;

    doc.setFontSize(22);
    doc.setFont('helvetica', 'bold');
    doc.text('Daily Claim-Flow Analysis', 14, currentY);
    currentY += 8;
    
    doc.setFontSize(14);
    doc.setFont('helvetica', 'normal');
    doc.text(`Prepared for: ${clientName}`, 14, currentY);
    currentY += 6;
    doc.text(`Date: ${getFormattedDate()}`, 14, currentY);
    currentY += 12;

    doc.setFontSize(16);
    doc.setFont('helvetica', 'bold');
    doc.text('Introduction', 14, currentY);
    doc.setLineWidth(0.5);
    doc.line(14, currentY + 2, 48, currentY + 2);
    currentY += 8;

    const introText = "This report provides a daily analysis of claim inventory movement and highlights key operational metrics. The following pages offer a visual and statistical breakdown of claim cohorts to support workflow management and strategic oversight.";
    doc.setFontSize(11);
    doc.setFont('helvetica', 'normal');
    const summaryText = doc.splitTextToSize(introText, 182);
    doc.text(summaryText, 14, currentY);
    currentY += (summaryText.length * 5) + 4;
    
    doc.setFontSize(16);
    doc.setFont('helvetica', 'bold');
    doc.text('Summary of Findings', 14, currentY);
    doc.setLineWidth(0.5);
    doc.line(14, currentY + 2, 65, currentY + 2);
    currentY += 8;

    const config = gatherConfig();
    
    let totalClaimsProcessed = 0;
    for (const stateName in state.detailedMovementStats) {
        for (const bucket in state.detailedMovementStats[stateName]) {
            const cohort = state.detailedMovementStats[stateName][bucket];
            for (const destination in cohort.movedTo) {
                if (destination.startsWith('APPROVED') || destination.startsWith('DENY')) {
                    totalClaimsProcessed += cohort.movedTo[destination];
                }
            }
        }
    }

    let valueRecoveredFromDenials = 0;
    const todayClaimsMap = new Map(state.processedClaimsList.map(c => [c.claimNumber, c]));
    for (const [claimNumber, yestData] of state.yesterdayDataMap.entries()) {
        if (yestData.state === 'DENY') {
            const todayClaim = todayClaimsMap.get(claimNumber);
            if (todayClaim && (todayClaim.claimState === 'APPROVED' || todayClaim.claimState.includes('PREBATCH'))) {
                 const charges = parseFloat(String(todayClaim.originalRow[config.totalChargesIndex]).replace(/[^0-9.-]/g, '')) || 0;
                 valueRecoveredFromDenials += charges;
            }
        }
    }

    const totalCriticalYesterday = state.detailedMovementStats.PEND?.Critical?.totalYesterday || 0;
    const criticalSuccessRate = totalCriticalYesterday > 0 ? ((state.workflowMovement.criticalWorked / totalCriticalYesterday) * 100).toFixed(1) + '%' : 'N/A';

    const todayStats = calculateStats(state.processedClaimsList);
    const yestBacklog = state.yesterdayStats?.PEND?.['30+'] ?? 0;
    const todayBacklog = todayStats?.PEND?.['30+'] ?? 0;
    const backlogChange = todayBacklog - yestBacklog;
    let backlogText;
    if (backlogChange > 0) backlogText = `increased by ${backlogChange} claims.`;
    else if (backlogChange < 0) backlogText = `decreased by ${Math.abs(backlogChange)} claims.`;
    else backlogText = "remained stable.";

    const summaryPoints = [
        `A total of ${totalClaimsProcessed.toLocaleString()} claims reached a final adjudication status (Approved or Denied).`,
        `The team achieved a ${criticalSuccessRate} success rate in resolving claims from yesterday's critical aging bucket.`,
        `Approximately ${formatCurrency(valueRecoveredFromDenials)} in revenue was recovered from overturned denials.`,
        `The high-priority Backlog (30+ Days) inventory has ${backlogText}`
    ];

    doc.setFontSize(11);
    doc.setFont('helvetica', 'normal');
    summaryPoints.forEach(point => {
        doc.text("•", 16, currentY);
        const pointText = doc.splitTextToSize(point, 172);
        doc.text(pointText, 22, currentY);
        currentY += (pointText.length * 5) + 2;
    });

    let valueInPend = 0, deniedDollars = 0;
    let pendAgeSum = 0, pendAgeCount = 0, oonCount = 0;
    
    state.processedClaimsList.forEach(claim => {
        const charges = parseFloat(String(claim.originalRow[config.totalChargesIndex]).replace(/[^0-9.-]/g, '')) || 0;
        if(claim.claimState === 'PEND') {
            valueInPend += charges;
            if (!isNaN(claim.cleanAge)) { pendAgeSum += claim.cleanAge; pendAgeCount++; }
        }
        if(claim.claimState === 'DENY') deniedDollars += charges;
        if(String(claim.originalRow[config.networkStatusIndex] || '').toUpperCase().includes('OUT')) oonCount++;
    });

    const avgPendAge = pendAgeCount > 0 ? (pendAgeSum / pendAgeCount).toFixed(1) + ' Days' : 'N/A';
    const oonPercent = state.processedClaimsList.length > 0 ? ((oonCount / state.processedClaimsList.length) * 100).toFixed(1) + '%' : 'N/A';
    const deniedYesterdayCount = state.yesterdayStats?.DENY?.total || 0;
    let overturnedCount = 0;
    if (state.detailedMovementStats.DENY) {
        for (const bucket in state.detailedMovementStats.DENY) {
            const cohort = state.detailedMovementStats.DENY[bucket];
            overturnedCount += cohort.movedToPrebatch;
            for (const dest in cohort.movedTo) { if (dest.startsWith('APPROVED')) { overturnedCount += cohort.movedTo[dest]; } }
        }
    }
    const denialOverturnRate = deniedYesterdayCount > 0 ? ((overturnedCount / deniedYesterdayCount) * 100).toFixed(1) + '%' : 'N/A';
    
    const kpis = [
        { label: 'Total Claims Processed', value: totalClaimsProcessed.toLocaleString() },
        { label: 'Value in PEND Inventory', value: formatCurrency(valueInPend) },
        { label: 'Denied Dollars at Risk', value: formatCurrency(deniedDollars) },
        { label: 'Average Clean Age (PEND)', value: avgPendAge },
        { label: 'Denial Overturn Rate', value: denialOverturnRate },
        { label: 'Out-of-Network %', value: oonPercent }
    ];

    const kpiStartY = currentY + 4;
    doc.setFontSize(16);
    doc.setFont('helvetica', 'bold');
    doc.text('Key Performance Indicators', 14, kpiStartY);
    doc.line(14, kpiStartY + 2, 70, kpiStartY + 2);
    
    let currentX = 14;
    let kpiY = kpiStartY + 10;
    const rectWidth = 58;
    const rectHeight = 30;
    
    kpis.forEach((kpi, index) => {
        if (index === 3) {
            kpiY += rectHeight + 5;
            currentX = 14;
        }
        doc.setFillColor(240, 240, 255);
        doc.roundedRect(currentX, kpiY, rectWidth, rectHeight, 3, 3, 'F');
        doc.setFontSize(18);
        doc.setFont('helvetica', 'bold');
        doc.text(kpi.value, currentX + rectWidth/2, kpiY + 15, { align: 'center'});
        doc.setFontSize(9);
        doc.setFont('helvetica', 'normal');
        const label = doc.splitTextToSize(kpi.label, rectWidth - 5);
        doc.text(label, currentX + rectWidth/2, kpiY + 23, { align: 'center'});
        currentX += rectWidth + 5;
    });
    
    let finalY = kpiY + rectHeight + 5;

    if (Object.keys(state.cycleTimeMetrics).length > 0) {
         doc.setFontSize(16);
         doc.setFont('helvetica', 'bold');
         doc.text('Claims Cycle Time Performance', 14, finalY);
         doc.line(14, finalY + 2, 85, finalY + 2);

         const tableBody = [
             [`95% of Clean/Non-Par Claims in 30 Days`, state.cycleTimeMetrics.cleanNonPar30],
             [`All Other Non-Par Claims in 60 Days`, state.cycleTimeMetrics.otherNonPar60],
             [`95% of Clean/Par Claims in 30 Days`, state.cycleTimeMetrics.cleanPar30],
             [`All Other Par Claims in 60 Days`, state.cycleTimeMetrics.otherPar60],
         ];

         doc.autoTable({
             startY: finalY + 5,
             head: [['Performance Measure', 'Current Performance']],
             body: tableBody,
             theme: 'grid',
             headStyles: { fillColor: [41, 128, 186] },
         });
         finalY = doc.autoTable.previous.finalY;
    }
}

async function createChartsPage(doc) {
    doc.setFontSize(18);
    doc.setFont('helvetica', 'bold');
    doc.text("Graphical Analysis - The Big Picture", 14, 20);

    const renderChartToImage = async (canvasId, config, width, height) => {
        const container = document.getElementById('chart-render-container');
        const canvas = document.createElement('canvas');
        canvas.id = canvasId;
        canvas.width = width;
        canvas.height = height;
        container.appendChild(canvas);
        
        return new Promise((resolve) => {
            const chartInstance = new Chart(canvas, {
                ...config,
                options: {
                    ...config.options,
                    animation: {
                        onComplete: () => {
                            const imgData = chartInstance.canvas.toDataURL('image/png');
                            chartInstance.destroy();
                            container.innerHTML = '';
                            resolve(imgData);
                        }
                    }
                }
            });
        });
    };

    const flowData = {};
    const destinations = new Set(['Prebatch', 'Resolved/Removed']);
    for(const stateName in state.detailedMovementStats) {
        for(const bucket in state.detailedMovementStats[stateName]) {
            const cohort = state.detailedMovementStats[stateName][bucket];
            const cohortName = `${stateName} - ${bucket}`;
            flowData[cohortName] = { total: cohort.totalYesterday, 'Prebatch': cohort.movedToPrebatch, 'Resolved/Removed': cohort.resolvedOrRemoved };
            for(const dest in cohort.movedTo) { destinations.add(dest); flowData[cohortName][dest] = cohort.movedTo[dest]; }
        }
    }
    const sortedDestinations = [...destinations].sort();
    const flowChartConfig = {
        type: 'bar',
        data: {
            labels: Object.keys(flowData),
            datasets: sortedDestinations.map((dest, i) => ({
                label: dest.replace(/_/g, ' - '),
                data: Object.values(flowData).map(d => d[dest] || 0),
                backgroundColor: `hsl(${i * 360 / sortedDestinations.length}, 70%, 50%)`
            }))
        },
        options: { indexAxis: 'y', scales: { x: { stacked: true }, y: { stacked: true } }, responsive: false, plugins: { title: { display: true, text: "Yesterday's Claim Flow by Destination" }, legend: { display: false } } }
    };
    const flowChartImg = await renderChartToImage('flowChart', flowChartConfig, 1000, 800);
    doc.addImage(flowChartImg, 'PNG', 14, 25, 180, 100);

    const pendData = { Queue: 0, Priority: 0, Critical: 0, Backlog: 0 };
    let totalPend = 0;
    if (state.detailedMovementStats.PEND) {
        for (const bucket in state.detailedMovementStats.PEND) {
            pendData[bucket] = state.detailedMovementStats.PEND[bucket].totalYesterday;
            totalPend += pendData[bucket];
        }
    }
    const pendChartConfig = {
        type: 'doughnut',
        data: { labels: Object.keys(pendData), datasets: [{ data: Object.values(pendData), backgroundColor: ['#36A2EB', '#FFCE56', '#FF6384', '#4BC0C0'] }] },
        options: { responsive: false, plugins: { title: { display: true, text: `Composition of Yesterday's PEND Inventory (Total: ${totalPend})` } } }
    };
    const pendChartImg = await renderChartToImage('pendChart', pendChartConfig, 400, 400);
    doc.addImage(pendChartImg, 'PNG', 14, 135, 80, 80);

    const critOutcomes = { 'Moved to Prebatch': 0, 'Approved': 0, 'Remained Critical': 0, 'Aged to Backlog': 0, 'Denied': 0 };
    if (state.detailedMovementStats.PEND && state.detailedMovementStats.PEND.Critical) {
        const crit = state.detailedMovementStats.PEND.Critical;
        critOutcomes['Moved to Prebatch'] = crit.movedToPrebatch;
        critOutcomes['Aged to Backlog'] = state.workflowMovement.criticalToBacklog;
        for(const [dest, count] of Object.entries(crit.movedTo)) {
            const [newState, newBucket] = dest.split('_');
            if(newState === 'APPROVED') critOutcomes['Approved'] += count;
            else if(newState === 'DENY') critOutcomes['Denied'] += count;
            else if(newBucket === 'Critical') critOutcomes['Remained Critical'] += count;
        }
    }
    
    const totalCritOutcomes = Object.values(critOutcomes).reduce((a, b) => a + b, 0);
    if (totalCritOutcomes > 0) {
        const critChartConfig = {
            type: 'bar',
            data: { labels: Object.keys(critOutcomes), datasets: [{ label: 'Number of Claims', data: Object.values(critOutcomes), backgroundColor: '#FF6384' }] },
            options: { responsive: false, plugins: { title: { display: true, text: "Outcomes of Yesterday's Critical Claims" }, legend: { display: false } } }
        };
        const critChartImg = await renderChartToImage('critChart', critChartConfig, 600, 400);
        doc.addImage(critChartImg, 'PNG', 105, 135, 90, 60);
    } else {
        doc.setFontSize(10);
        doc.setFont('helvetica', 'italic');
        doc.setTextColor(128);
        doc.text("No Critical Claims to Analyze for this Period.", 150, 165, { align: 'center' });
    }
}

async function createDetailedTablesPage(doc) {
    doc.setFontSize(18);
    doc.setFont('helvetica', 'bold');
    doc.text("Detailed Cohort Movement Analysis", 14, 20);
    
    let finalY = 25;

    const sortedStates = Object.keys(state.detailedMovementStats).sort();
    for(const stateName of sortedStates) {
        const sortedBuckets = Object.keys(state.detailedMovementStats[stateName]).sort((a,b) => {
            const order = { 'Critical': 1, 'Priority': 2, 'Backlog': 3, 'Queue': 4 };
            return (order[a] || 99) - (order[b] || 99);
        });

        for(const bucket of sortedBuckets) {
            const cohortData = state.detailedMovementStats[stateName][bucket];
            if (cohortData.totalYesterday === 0) continue;

            const tableBody = [];
            if (cohortData.resolvedOrRemoved > 0) tableBody.push(['Resolved/Removed from report', cohortData.resolvedOrRemoved, `${(cohortData.resolvedOrRemoved / cohortData.totalYesterday * 100).toFixed(1)}%`, 'Positive']);
            if (cohortData.movedToPrebatch > 0) tableBody.push(['Moved to Prebatch', cohortData.movedToPrebatch, `${(cohortData.movedToPrebatch / cohortData.totalYesterday * 100).toFixed(1)}%`, 'Positive']);
            
            const sortedDestinations = Object.entries(cohortData.movedTo).sort((a,b) => b[1] - a[1]);
            sortedDestinations.forEach(([dest, count]) => {
                 const [newState, newBucket] = dest.split('_');
                 let impact = 'Neutral';
                 if (newBucket === 'Backlog' && bucket !== 'Backlog') impact = 'Negative';
                 if (newState.startsWith('APPROVED')) impact = 'Positive';
                 tableBody.push([`Moved to: ${newState} - ${newBucket}`, count, `${(count/cohortData.totalYesterday*100).toFixed(1)}%`, impact]);
            });

            const tableHeight = (tableBody.length + 1) * 10 + 20;
            if (finalY + tableHeight > 280) {
                doc.addPage();
                finalY = 20;
            }

            doc.autoTable({
                startY: finalY + 5,
                head: [['Destination', 'Count', 'Percentage', 'Impact']],
                body: tableBody,
                didDrawPage: (tableData) => {
                    doc.setFontSize(12);
                    doc.setFont('helvetica', 'bold');
                    doc.text(`From: ${stateName} - ${bucket} (Yesterday's Total: ${cohortData.totalYesterday})`, tableData.settings.margin.left, finalY + 2);
                },
                willDrawCell: (data) => {
                    if (data.column.dataKey === 'Impact' && data.cell.section === 'body') {
                        if (data.cell.text[0] === 'Positive') doc.setTextColor(0, 128, 0);
                        else if (data.cell.text[0] === 'Negative') doc.setTextColor(255, 0, 0);
                    }
                },
                didDrawCell: () => doc.setTextColor(0, 0, 0),
                margin: { top: finalY + 10 }
            });

            finalY = doc.autoTable.previous.finalY + 5;
        }
    }
}
