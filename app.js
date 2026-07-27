/*
===========================================
 Telemedicine Hub Performance Logic
 Hub North 24 Parganas
===========================================
*/

const hubState = {
    masterDoctors: [],
    uploadedSummary: [],
    section: "1" // 1 = Block Wise, 2 = Hospital Wise
};

document.addEventListener('DOMContentLoaded', () => {
    initHubDashboard();
});

function initHubDashboard() {
    // ডিফল্ট হিসেবে আজকের তারিখ সেট করা
    const today = new Date().toISOString().split('T')[0];
    document.getElementById('reportDateInput').value = today;
    
    // রিপোর্ট ডেট আপডেট লিসেনার
    document.getElementById('reportDateInput').addEventListener('change', updateReportDateLabel);
    updateReportDateLabel();
    
    setupEventListeners();
    fetchHubMasterData();
}

function updateReportDateLabel() {
    const inputDate = document.getElementById('reportDateInput').value;
    if(inputDate) {
        const [y, m, d] = inputDate.split('-');
        document.getElementById('reportDate').textContent = `${d}/${m}/${y}`;
    }
}

function getFormattedDate() {
    const inputDate = document.getElementById('reportDateInput').value;
    const [y, m, d] = inputDate.split('-');
    return `${d}/${m}/${y}`; // 25/07/2026 ফরম্যাট
}

function setReportDate() {
    const today = new Date();
    const formattedDate = today.toLocaleDateString('en-IN', { 
        day: '2-digit', month: 'short', year: 'numeric' 
    });
    if (document.getElementById('reportDate')) {
        document.getElementById('reportDate').textContent = formattedDate;
    }
}

function getFilenameDate() {
    return new Date().toLocaleDateString('en-IN', { 
        day: '2-digit', month: 'short', year: 'numeric' 
    }).replace(/ /g, '_');
}

/**
 * Fetch Doctor Master Data from Google Apps Script
 */
const CONFIG = {

    // Hub Master Google Sheet API Web App URL
    HUB_MASTER_URL: "https://script.google.com/macros/s/AKfycbyo7mEdHSLbb6WyszbTUHe4NrRFgqkoJJCI8LNnEAxZvkPOgFgXScXxORBrNPJ2qjd0Pg/exec"
};
async function fetchHubMasterData() {
    showLoading(true);
    const masterUrl = CONFIG.HUB_MASTER_URL || CONFIG.API_URL;

    try {
        const res = await fetch(masterUrl, { method: 'GET', redirect: 'follow' });
        if (res.ok) {
            const data = await res.json();
            hubState.masterDoctors = parseMasterArray(data);
            
            if (hubState.uploadedSummary.length > 0) {
                processAndRenderHubReport();
            }
        }
    } catch (err) {
        console.error('Master Data Fetch Error:', err);
    } finally {
        showLoading(false);
    }
}

/**
 * Master Array Parser
 */
function parseMasterArray(rawList) {
    if (!Array.isArray(rawList)) return [];
    return rawList.map(item => {
        const keys = Object.keys(item);
        
        const typeKey = keys.find(k => /type|hf/i.test(k.trim())) || keys[0];
        const nameKey = keys.find(k => /block|hospital|facility|location|name/i.test(k.trim())) || keys[1];
        const docKey  = keys.find(k => /doctor|user|doc/i.test(k.trim())) || keys[2];

        const rawType = String(item[typeKey] || '').trim();
        const rawName = String(item[nameKey] || '').trim();
        const rawDoc  = String(item[docKey] || '').trim();

        return {
            hfType: rawType,
            locationName: rawName,
            doctorName: rawDoc,
            cleanDocName: normalizeDoctorName(rawDoc)
        };
    }).filter(d => d.doctorName !== '');
}

/**
 * Doctor Name Normalization
 */
function normalizeDoctorName(str) {
    if (!str) return '';
    return str.toString()
        .toUpperCase()
        .replace(/\bDR\b/g, '')     // Removes 'DR'
        .replace(/\bMR\b/g, '')     // Removes 'MR'
        .replace(/\./g, '')         // Removes full stops
        .replace(/\(.*\)/g, '')     // Removes brackets
        .replace(/[^A-Z]/g, '')     // Keeps ONLY letters A-Z
        .trim();
}

/**
 * Doctor Match Score
 */
function getMatchScore(masterClean, excelClean) {
    if (!masterClean || !excelClean) return 0;
    if (masterClean === excelClean) return 100;
    if (masterClean.includes(excelClean) || excelClean.includes(masterClean)) {
        return Math.min(masterClean.length, excelClean.length);
    }
    return 0;
}

function setupEventListeners() {
    document.getElementById('generateReport')?.addEventListener('click', handleFileUpload);
    document.getElementById('sectionSelect')?.addEventListener('change', (e) => {
        hubState.section = e.target.value;
        if (hubState.uploadedSummary.length > 0) {
            processAndRenderHubReport();
        }
    });

    document.getElementById('btnExportExcel')?.addEventListener('click', exportHubExcel);
    document.getElementById('btnExportPDF')?.addEventListener('click', downloadHubPDF);
    document.getElementById('btnExportJPG')?.addEventListener('click', downloadHubJPG);
    document.getElementById('btnShareWhatsApp')?.addEventListener('click', shareHubWhatsApp);
}

function handleFileUpload() {
    const fileInput = document.getElementById('excelFile');
    const file = fileInput?.files[0];

    if (!file) {
        alert('অনুগ্রহ করে ConsultationReceiveSummary ফাইলটি নির্বাচন করুন।');
        return;
    }

    showLoading(true);

    const reader = new FileReader();
    reader.onload = function (e) {
        try {
            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, { type: 'array' });
            const sheet = workbook.Sheets[workbook.SheetNames[0]];
            const uploadedRows = XLSX.utils.sheet_to_json(sheet, { defval: "" });

            hubState.uploadedSummary = uploadedRows;
            processAndRenderHubReport();
        } catch (err) {
            console.error(err);
            alert('Excel ফাইল প্রসেস করতে সমস্যা হয়েছে।');
        } finally {
            showLoading(false);
        }
    };

    reader.readAsArrayBuffer(file);
}

/**
 * Processing Logic
 */
function processAndRenderHubReport() {
    const uploadedList = [];

    // Filter Column B for "HUB NORTH TWENTY FOUR PRGS"
    hubState.uploadedSummary.forEach(row => {
        const keys = Object.keys(row);
        
        const facilityCol  = keys.find(k => /health facility|facility/i.test(k.trim())) || keys[1]; 
        const docCol       = keys.find(k => /user|doctor|doc/i.test(k.trim())) || keys[3];            
        const completedCol = keys.find(k => /completed|consultation/i.test(k.trim()));
        const inProcessCol = keys.find(k => /process|in process/i.test(k.trim()));

        const facilityName = String(row[facilityCol] || '').trim().toUpperCase();

        if (facilityName === "HUB NORTH TWENTY FOUR PRGS") {
            const rawDoc = String(row[docCol] || '').trim();
            const cleanName = normalizeDoctorName(rawDoc);
            const completed = completedCol ? (parseInt(row[completedCol], 10) || 0) : 0;
            const inProcess = inProcessCol ? (parseInt(row[inProcessCol], 10) || 0) : 0;

            if (cleanName && completed > 0) {
                uploadedList.push({
                    rawDoc,
                    rawFacility: row[facilityCol],
                    cleanName,
                    completed,
                    inProcess,
                    matched: false
                });
            }
        }
    });

    const targetSection = hubState.section; // "1" = Block, "2" = Hospital/DHQ
    let matchedDoctors = [];
    let unmatchedDoctors = [];

    // Filter master list for target section (Block or Hospital)
    const sectionMasters = hubState.masterDoctors.filter(doc => {
        const isBlock = /^block$/i.test(doc.hfType);
        return targetSection === "1" ? isBlock : !isBlock;
    });

    if (sectionMasters.length > 0) {
        const masterResults = sectionMasters.map(m => ({
            ...m,
            completed: 0,
            inProcess: 0
        }));

        uploadedList.forEach(uploaded => {
            let bestIndex = -1;
            let bestScore = 0;

            masterResults.forEach((master, idx) => {
                const score = getMatchScore(master.cleanDocName, uploaded.cleanName);
                if (score > bestScore) {
                    bestScore = score;
                    bestIndex = idx;
                }
            });

            if (bestIndex !== -1 && bestScore > 0) {
                masterResults[bestIndex].completed += uploaded.completed;
                masterResults[bestIndex].inProcess += uploaded.inProcess;
                uploaded.matched = true;
            }
        });

        // Filter matched doctors who performed consultations for Main Doctor Table
        matchedDoctors = masterResults.filter(m => m.completed > 0);

        // Identify Unmatched Doctors
        uploadedList.forEach(uploaded => {
            if (!uploaded.matched) {
                const isPresentInAny = hubState.masterDoctors.some(m => getMatchScore(m.cleanDocName, uploaded.cleanName) > 0);
                if (!isPresentInAny) {
                    unmatchedDoctors.push(uploaded);
                }
            }
        });

    } else {
        uploadedList.forEach(item => {
            matchedDoctors.push({
                hfType: targetSection === "1" ? "Block" : "Hospital",
                locationName: item.rawFacility || "North 24 Parganas",
                doctorName: item.rawDoc,
                completed: item.completed,
                inProcess: item.inProcess
            });
        });
    }

    matchedDoctors.sort((a, b) => a.locationName.localeCompare(b.locationName));

    renderUnmatchedDoctorsTableTop(unmatchedDoctors);
    
    // Render Doctor Wise Main Table
    renderHubTable(matchedDoctors);

    // Render Location Summary Table below Main Table (Includes Zero consultation blocks/hospitals)
    renderLocationSummaryTable(sectionMasters, matchedDoctors);
    updateDashboardSummary(uploadedList, hubState.masterDoctors);
}

function renderUnmatchedDoctorsTableTop(unmatchedList) {
    let unmatchedContainer = document.getElementById('unmatchedContainerTop');
    const mainReportCard = document.getElementById('hubReportCard');
    
    if (!unmatchedContainer) {
        unmatchedContainer = document.createElement('div');
        unmatchedContainer.id = 'unmatchedContainerTop';
        unmatchedContainer.className = 'mb-4 no-print';
        mainReportCard.before(unmatchedContainer);
    }

    if (!unmatchedList || unmatchedList.length === 0) {
        unmatchedContainer.innerHTML = '';
        return;
    }

    let rowsHtml = unmatchedList.map((doc, idx) => `
        <tr class="table-warning">
            <td class="text-center border-dark">${idx + 1}</td>
            <td class="border-dark fw-bold text-danger">${doc.rawDoc}</td>
            <td class="border-dark">${doc.rawFacility}</td>
            <td class="text-center border-dark fw-bold">${doc.completed}</td>
            <td class="text-center border-dark">${doc.inProcess}</td>
        </tr>
    `).join('');

    unmatchedContainer.innerHTML = `
        <div class="card border-danger shadow-sm">
            <div class="card-header bg-danger text-white fw-bold">
                <i class="fa-solid fa-triangle-exclamation me-2"></i> Warning: Doctors Not Found in Master Google Sheet (${unmatchedList.length})
            </div>
            <div class="card-body py-2">
                <p class="text-muted small mb-2">HUB NORTH TWENTY FOUR PRGS-এর নিচের ডাক্তারদের নাম মাস্টার শিটে পাওয়া যায়নি। অনুগ্রহ করে তাদের নাম গুগুল শিটে যোগ করুন:</p>
                <div class="table-responsive">
                    <table class="table table-bordered border-dark table-sm mb-0">
                        <thead class="table-dark text-center">
                            <tr>
                                <th>#</th>
                                <th>Doctor Name (Uploaded Excel)</th>
                                <th>Facility Name</th>
                                <th>Completed</th>
                                <th>In Process</th>
                            </tr>
                        </thead>
                        <tbody>
                            ${rowsHtml}
                        </tbody>
                    </table>
                </div>
            </div>
        </div>
    `;
}

/**
 * Render Main Doctor Wise Report Table
 */
function renderHubTable(doctorList) {
    const tbody = document.getElementById('hubReportBody');
    const titleHeader = document.getElementById('reportTitleHeader');
    const colHeaderName = document.getElementById('colHeaderName');

    if (!tbody) return;
    tbody.innerHTML = '';

    const isBlock = hubState.section === "1";
    titleHeader.textContent = isBlock 
        ? "Block Wise Doctor Tele-Consultation Report" 
        : "Hospital Wise Doctor Tele-Consultation Report";
    colHeaderName.textContent = isBlock ? "Name of the Block" : "Name of the Hospital";

    if (doctorList.length === 0) {
        tbody.innerHTML = `<tr><td colspan="5" class="text-center py-4 text-danger fw-bold">কোনো সক্রিয় ডাক্তার তথ্য পাওয়া যায়নি।</td></tr>`;
    } else {
        let globalSlNo = 1;
        let grandActiveDocs = 0;
        let grandConsultations = 0;
        let grandInProcess = 0;

        const grouped = {};
        doctorList.forEach(doc => {
            if (!grouped[doc.locationName]) grouped[doc.locationName] = [];
            grouped[doc.locationName].push(doc);
        });

        Object.keys(grouped).sort().forEach(locName => {
            const docs = grouped[locName];

            docs.forEach(doc => {
                grandConsultations += doc.completed;
                grandInProcess += doc.inProcess;

                const tr = document.createElement('tr');
                
                // Complete Consultation Row Color Code
                const rowColorClass = getConsultationColorClass(doc.completed);
                tr.className = rowColorClass;

                // In Process Cell Specific Color Logic (Light Red when >= 1)
                const inProcessClass = doc.inProcess >= 1 ? 'cell-in-process-red' : '';

                tr.innerHTML = `
                    <td class="text-center border-dark fw-bold">${globalSlNo++}</td>
                    <td class="border-dark fw-bold">${doc.locationName}</td>
                    <td class="border-dark fw-bold">${doc.doctorName}</td>
                    <td class="text-center border-dark fw-bold">${doc.completed.toLocaleString()}</td>
                    <td class="text-center border-dark ${inProcessClass}">${doc.inProcess.toLocaleString()}</td>
                `;
                tbody.appendChild(tr);
            });

            grandActiveDocs += docs.length;
        });

        // Grand Total Row - এখানে সাইজ ঠিক করে দেওয়া হলো
const grandTr = document.createElement('tr');
grandTr.className = "bg-grandtotal border-dark"; 

const grandInProcessClass = grandInProcess >= 1 ? 'cell-in-process-red' : '';

grandTr.innerHTML = `
    <td colspan="2" class="text-center border-dark text-uppercase">
        <strong style="font-size: 10px;">GRAND TOTAL \n (Active Doctors: ${grandActiveDocs})</strong>
    </td>
    <td class="border-dark text-center fw-bold" style="font-size: 14px;">TOTAL</td>
    <td class="text-center border-dark fw-extrabold text-success" style="font-size: 18px;">${grandConsultations.toLocaleString()}</td>
    <td class="text-center border-dark fw-bold ${grandInProcessClass}" style="font-size: 16px;">${grandInProcess.toLocaleString()}</td>
`;
tbody.appendChild(grandTr);
    } 
}

/**
 * Render Location Wise Summary Table Below Main Report (Block / Hospital Wise)
 * Colorful Main Header & Sub-Header
 */
function renderLocationSummaryTable(sectionMasters, activeDoctorList) {
    const mainReportCard = document.getElementById('hubReportCard');
    let summaryContainer = document.getElementById('locationSummaryContainer');

    if (!summaryContainer) {
        summaryContainer = document.createElement('div');
        summaryContainer.id = 'locationSummaryContainer';
        summaryContainer.className = 'card shadow mb-3 border-dark mt-3';
        mainReportCard.after(summaryContainer);
    }

    const isBlock = hubState.section === "1";
    const locTitle = isBlock ? "Block Wise Performance Summary" : "Hospital Wise Performance Summary";
    const locColHeader = isBlock ? "Name Of Block" : "Name of Hospital";

    // Build Master Location Map with Active Doctor Counts and Total Consultations
    const locationMap = {};

    // 1. Collect all unique locations from Master Sheet
    sectionMasters.forEach(m => {
        if (m.locationName && !locationMap[m.locationName]) {
            locationMap[m.locationName] = {
                activeDoctors: 0,
                totalConsultation: 0
            };
        }
    });

    // 2. Map Active Consultations & Doctor Counts
    activeDoctorList.forEach(doc => {
        if (!locationMap[doc.locationName]) {
            locationMap[doc.locationName] = { activeDoctors: 0, totalConsultation: 0 };
        }
        locationMap[doc.locationName].activeDoctors += 1;
        locationMap[doc.locationName].totalConsultation += doc.completed;
    });

    const sortedLocations = Object.keys(locationMap).sort();

    let slNo = 1;
    let grandActiveDocs = 0;
    let grandConsultations = 0;

    // renderLocationSummaryTable ফাংশনের ভেতরে এটি পরিবর্তন করুন:

let rowsHtml = sortedLocations.map(locName => {
    const data = locationMap[locName];
    grandActiveDocs += data.activeDoctors;
    grandConsultations += data.totalConsultation;

    // এখানে লজিকটি ঠিক করুন:
    // জিরো কনসালটেশন হলে সরাসরি "bg-deep-red" ক্লাস পাবে
    const rowBg = data.totalConsultation === 0 ? 'bg-deep-red' : getConsultationColorClass(data.totalConsultation);

    return `
        <tr class="${rowBg}">
            <td class="text-center border-dark fw-bold">${slNo++}</td>
            <td class="border-dark fw-bold">${locName}</td>
            <td class="text-center border-dark fw-bold ${data.activeDoctors === 0 ? 'text-danger' : 'text-primary'}">${data.activeDoctors}</td>
            <td class="text-center border-dark fw-bold ${data.totalConsultation === 0 ? 'text-white' : ''}">${data.totalConsultation.toLocaleString()}</td>
        </tr>
    `;
}).join('');

    summaryContainer.innerHTML = `
        <div class="card-header summary-header-bg fw-bold h6 mb-0 text-center py-2">
            <i class="fa-solid fa-chart-pie me-2"></i>${locTitle}
        </div>
        <div class="table-responsive">
            <table class="table table-bordered border-dark align-middle mb-0 custom-black-table">
                <thead class="summary-subhead-blue text-center">
                    <tr>
                        <th style="width: 10%;">Sl No.</th>
                        <th>${locColHeader}</th>
                        <th style="width: 25%;">No of Active Doctor</th>
                        <th style="width: 25%;">Total Consultation</th>
                    </tr>
                </thead>
                <tbody>
                    ${rowsHtml.length > 0 ? rowsHtml : '<tr><td colspan="4" class="text-center py-3 text-muted">কোনো ডাটা পাওয়া যায়নি।</td></tr>'}
                    <tr class="bg-grandtotal border-dark fs-6">
                        <td colspan="2" class="text-center border-dark text-uppercase">
                            <strong>GRAND TOTAL</strong>
                        </td>
                        <td class="text-center border-dark fw-extrabold text-primary h6 mb-0">${grandActiveDocs}</td>
                        <td class="text-center border-dark fw-extrabold text-success h5 mb-0">${grandConsultations.toLocaleString()}</td>
                    </tr>
                </tbody>
            </table>
        </div>
    `;
}

function getConsultationColorClass(count) {
    if (count === 0) return "bg-deep-black"; // 
    if (count >= 1 && count <= 59) return "bg-light-red";
    if (count >= 60 && count <= 99) return "bg-light-orange";
    if (count >= 100) return "bg-light-green";
    return "";
}

/**
 * Downloads & Exports
 */
function downloadHubPDF() {
    showLoading(true);

    const dateStr = getFormattedDate().replace(/\//g, '_'); 
    const sectionName = hubState.section === "1" ? "Block_Wise" : "Hospital_Wise";
    const fileName = `${dateStr}_N24PGS_Telemedicine_Hub_${sectionName}_Report.pdf`;

    const printElement = document.querySelector('.print-container');
    const clone = printElement.cloneNode(true);

    clone.querySelectorAll('.no-print, .btn, #excelFile, #sectionSelect, #loading').forEach(el => el.remove());

    // ১. এখানে আমরা রো কাউন্ট করব
    const mainTableBody = clone.querySelector("#hubReportBody");
    const rowCount = mainTableBody ? mainTableBody.querySelectorAll('tr').length : 20; // ডিফল্ট ২০ ধরা হলো

    // ২. রো অনুযায়ী ফন্ট সাইজ এবং প্যাডিং সেট করা
    // অনেক রো থাকলে ফন্ট ছোট হবে, কম থাকলে বড় হবে
    let fontSize = '12px';
    let padding = '6px';
    let lineHeight = '1.2';

    if (rowCount > 30) {
        fontSize = '10px'; // অনেক ডাক্তার থাকলে ছোট ফন্ট
        padding = '2px';
    } else if (rowCount > 20) {
        fontSize = '12px';
        padding = '4px';
    } else {
        fontSize = '14px'; // কম ডাক্তার থাকলে বড় ফন্ট
        padding = '8px';
    }

    // হেডার এবং অন্যান্য সেটিং
    const header = clone.querySelector("#governmentHeader");
    const headerClone = header.cloneNode(true);
    headerClone.style.marginTop = "20px";
    
    const summaryCard = clone.querySelector("#locationSummaryContainer");
    if (summaryCard) {
        summaryCard.prepend(headerClone);
    }

    // ৩. ডাইনামিক স্টাইল ইনজেকশন
    const pdfStyle = document.createElement('style');
    pdfStyle.innerHTML = `
        .print-container { width: 100% !important; padding: 10px !important; }
        table { table-layout: fixed !important; width: 100% !important; }
        td, th { 
            font-size: ${fontSize} !important; 
            padding: ${padding} !important; 
            line-height: ${lineHeight} !important;
            word-wrap: break-word !important; 
        }
        tr { page-break-inside: avoid !important; }
        #locationSummaryContainer { page-break-before: always !important; }
        h3 { font-size: 20px !important; }
        h5 { font-size: 14px !important; }
        .bg-grandtotal { background-color: #e9ecef !important; }
    `;
    clone.appendChild(pdfStyle);

    document.body.appendChild(clone);

    const opt = {
        margin: [0.2, 0.2, 0.2, 0.2],
        filename: fileName,
        image: { type: 'jpeg', quality: 0.98 },
        html2canvas: { scale: 2, useCORS: true, letterRendering: true },
        jsPDF: { unit: 'in', format: 'a4', orientation: 'portrait' },
        pagebreak: { mode: 'css' }
    };

    html2pdf().set(opt).from(clone).save().then(() => {
        document.body.removeChild(clone);
        showLoading(false);
    }).catch(err => {
        console.error("PDF Error:", err);
        if (clone.parentNode) document.body.removeChild(clone);
        showLoading(false);
    });
}

function downloadHubJPG() {
    showLoading(true);

    // ১. সিলেক্ট করা তারিখ ব্যবহার করে ফাইলের নাম তৈরি
    const dateStr = getFormattedDate().replace(/\//g, '_'); 
    const sectionName = hubState.section === "1" ? "Block_Wise" : "Hospital_Wise";
    const fileName = `${dateStr}_N24PGS_Telemedicine_Hub_${sectionName}_Report.jpg`;

    // ২. পুরো .print-container ডিভটিকে ধরলাম (এতে হেডার এবং তারিখসহ সব থাকবে)
    const printElement = document.querySelector('.print-container');
    const container = printElement.cloneNode(true);

    // ৩. অপ্রয়োজনীয় বাটন এবং এলিমেন্ট মুছে ফেলা
    container.querySelectorAll('.no-print, .btn, #excelFile, #sectionSelect, #loading').forEach(el => el.remove());

    // ৪. ব্যাকগ্রাউন্ড সাদা করা এবং প্যাডিং ঠিক রাখা
    container.style.padding = '20px';
    container.style.backgroundColor = '#ffffff';
    document.body.appendChild(container);

    // ৫. ইমেজ জেনারেট করা
    html2canvas(container, { scale: 2, useCORS: true }).then(canvas => {
        const link = document.createElement('a');
        link.download = fileName; // সিলেক্ট করা তারিখ অনুযায়ী নাম
        link.href = canvas.toDataURL('image/jpeg', 0.95);
        link.click();

        document.body.removeChild(container);
        showLoading(false);
    }).catch(err => {
        console.error("JPG Error:", err);
        if (container.parentNode) document.body.removeChild(container);
        showLoading(false);
        alert("JPG জেনারেট করতে সমস্যা হয়েছে।");
    });
}

function exportHubExcel() {
    const table1 = document.querySelector("#hubReportCard table");
    const table2 = document.querySelector("#locationSummaryContainer table");
    if (!table1) return;

    const dateStr = getFilenameDate();
    const sectionName = hubState.section === "1" ? "Block_Wise" : "Hospital_Wise";
    const fileName = `${dateStr}_N24PGS_Telemedicine_Hub_${sectionName}_Report.xlsx`;

    const wb = XLSX.utils.book_new();
    const sheet1 = XLSX.utils.table_to_sheet(table1);
    XLSX.utils.book_append_sheet(wb, sheet1, "Doctor Wise Performance");

    if (table2) {
        const sheet2 = XLSX.utils.table_to_sheet(table2);
        XLSX.utils.book_append_sheet(wb, sheet2, "Performance Summary");
    }

    XLSX.writeFile(wb, fileName);
}

function shareHubWhatsApp() {
    const reportDate = document.getElementById('reportDate')?.textContent || '';
    const sectionName = hubState.section === "1" ? "Block Wise Doctor" : "Hospital Wise Doctor";
    
    const message = `*Telemedicine Hub Performance Report (${sectionName})*%0A` +
                    `*Hub:* North 24 Parganas%0A` +
                    `*Date:* ${reportDate}%0A%0A` +
                    `_Generated from North 24 Parganas Hub Performance Dashboard._`;
                    
    window.open(`https://api.whatsapp.com/send?text=${message}`, '_blank');
}

function showLoading(isLoading) {
    const loadingElem = document.getElementById('loading');
    if (loadingElem) {
        loadingElem.classList.toggle('d-none', !isLoading);
    }
}

/**
 * Update Hub Total Summary Stats - Master Data Matching Logic
 */
function updateDashboardSummary(uploadedList, masterDoctors) {
    const statsContainer = document.getElementById('hubSummaryStats');
    
    // ১. শুধুমাত্র মাস্টার শিটে থাকা ডাক্তারদের লিস্ট তৈরি
    let matchedInMaster = new Set();
    let totalCons = 0;

    // ২. এক্সেল ডাটাকে মাস্টার ডাটার সাথে ম্যাচ করা
    uploadedList.forEach(uploaded => {
        // মাস্টার ডাটার সাথে ম্যাচ করছি (normalizeDoctorName ব্যবহার করে)
        const isMatched = masterDoctors.some(m => getMatchScore(m.cleanDocName, uploaded.cleanName) > 0);
        
        if (isMatched) {
            matchedInMaster.add(uploaded.cleanName); // শুধু মাস্টার শিটে থাকা ডাক্তারদের নাম যোগ করছি
            totalCons += uploaded.completed;         // শুধু তাদের কনসালটেশন যোগ করছি
        }
    });

    // ৩. সামারি কার্ড রেন্ডার
    statsContainer.innerHTML = `
        <div class="card border-primary shadow-sm mt-3">
            <div class="card-body py-2 text-center bg-light">
                <h6 class="text-primary fw-bold mb-0">
                    <i class="fa-solid fa-chart-line me-2"></i> MASTER SYNCED PERFORMANCE (Total): 
                    <span class="text-dark">
                        Active Doctors: <b>${matchedInMaster.size}</b> | 
                        Total Consultations: <b>${totalCons.toLocaleString()}</b>
                    </span>
                </h6>
            </div>
        </div>
    `;
}