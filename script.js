// Menambahkan event listener ke tombol proses utama
document.getElementById('process-btn').addEventListener('click', handleFiles);
// Menambahkan event listener ke input file untuk membuat form dinamis
document.getElementById('excel-file').addEventListener('change', generateDynamicInputs);

// Mengatur tanggal default saat halaman pertama kali dimuat
document.addEventListener("DOMContentLoaded", function() {
    const today = new Date();
    const yyyy = today.getFullYear();
    const mm = String(today.getMonth() + 1).padStart(2, '0');
    const dd = String(today.getDate()).padStart(2, '0');
    const todayString = `${yyyy}-${mm}-${dd}`;
    document.getElementById('createdDate').value = todayString;
    document.getElementById('startDate').value = todayString;
    document.getElementById('endDate').value = todayString;
});

// Fungsi untuk membuat input Test Case ID secara dinamis
function generateDynamicInputs() {
    const fileInput = document.getElementById('excel-file');
    const files = fileInput.files;
    const container = document.getElementById('test-case-id-inputs');
    container.innerHTML = ''; 

    if (files.length > 0) {
        const title = document.createElement('h5');
        title.className = 'col-12';
        title.textContent = 'Pengaturan per File';
        container.appendChild(title);
    }

    for (let i = 0; i < files.length; i++) {
        const file = files[i];
        const idInputId = `testCaseId-${i}`;
        const descInputId = `encDescription-${i}`;

        const inputGroupHtml = `
            <div class="col-md-6 mb-3">
                <label for="${idInputId}" class="form-label">
                    <small>ID Awal untuk File: <strong>${file.name}</strong></small>
                </label>
                <input type="text" class="form-control" id="${idInputId}" placeholder="Contoh: FM-IT-RAC-INI-2507-05739">
            </div>
            <div class="col-md-6 mb-3">
                <label for="${descInputId}" class="form-label">
                    <small>ENC Description untuk: <strong>${file.name}</strong></small>
                </label>
                <input type="text" class="form-control" id="${descInputId}" value="-">
            </div>
        `;
        container.innerHTML += inputGroupHtml;
    }
}

// Fungsi untuk menampilkan notifikasi
function showNotification(message, type = 'danger') {
    const notificationArea = document.getElementById('notification-area');
    notificationArea.innerHTML = `<div class="alert alert-${type} alert-dismissible fade show" role="alert">
        ${message}
        <button type="button" class="btn-close" data-bs-dismiss="alert" aria-label="Close"></button>
    </div>`;
}

// Helper untuk format tanggal
function formatDate(date) {
    if (!date) return '';
    const parts = date.split('-');
    const year = parseInt(parts[0], 10);
    const month = parseInt(parts[1], 10) - 1;
    const day = parseInt(parts[2], 10);
    const d = new Date(year, month, day);
    const months = ["Jan", "Feb", "Mar", "Apr", "Mei", "Jun", "Jul", "Agu", "Sep", "Okt", "Nov", "Des"];
    return `${d.getDate()}-${months[d.getMonth()]}-${d.getFullYear()}`;
}

// Helper untuk membersihkan nama sheet
function sanitizeSheetName(fileName) {
    return fileName.replace(/[:\\/?*[\]]/g, "").substring(0, 31);
}

// Fungsi utama untuk menangani file
async function handleFiles() {
    const fileInput = document.getElementById('excel-file');
    const processBtn = document.getElementById('process-btn');
    const btnText = document.getElementById('btn-text');
    const spinner = document.getElementById('spinner');
    const files = fileInput.files;

    if (files.length === 0) {
        showNotification('⚠️ Silakan pilih satu atau beberapa file Excel terlebih dahulu.');
        return;
    }

    processBtn.disabled = true;
    btnText.textContent = `Memproses ${files.length} file...`;
    spinner.classList.remove('d-none');

    const outputWorkbook = XLSX.utils.book_new();

    try {
        const settings = getSettingsFromForm();
        const globalSuffixCounter = { current: 1 };

        for (let i = 0; i < files.length; i++) {
            const file = files[i];
            
            try {
                const baseIdForFile = document.getElementById(`testCaseId-${i}`).value;
                const encDescriptionForFile = document.getElementById(`encDescription-${i}`).value;

                const data = await file.arrayBuffer();
                const workbook = XLSX.read(data);
                const sourceSheet = workbook.Sheets[workbook.SheetNames[0]];
                const sourceData = XLSX.utils.sheet_to_json(sourceSheet, { header: 1 });
                
                if (sourceData.length < 2) {
                    console.warn(`File "${file.name}" dilewati karena tidak berisi data.`);
                    continue;
                }
                
                const processedData = processSingleFile(sourceData, settings, baseIdForFile, encDescriptionForFile, globalSuffixCounter);
                
                const newWorksheet = XLSX.utils.aoa_to_sheet(processedData);
                applyStyling(newWorksheet, processedData);
                const sheetName = sanitizeSheetName(file.name);
                XLSX.utils.book_append_sheet(outputWorkbook, newWorksheet, sheetName);

            } catch (error) {
                console.error(`Error pada file "${file.name}":`, error);
                const errorMessage = `
                    Terjadi kesalahan saat memproses file: <strong>${file.name}</strong>.<br>
                    Pesan Error: ${error.message}<br>
                    <em>Silakan periksa file tersebut atau hubungi developer.</em>`;
                showNotification(errorMessage, 'danger');
                
                processBtn.disabled = false;
                btnText.textContent = 'Proses Data & Unduh Hasil';
                spinner.classList.add('d-none');
                return; 
            }
        }
        
        if (outputWorkbook.SheetNames.length === 0) {
             showNotification('Tidak ada file yang berhasil diproses. Periksa kembali semua file Anda.', 'warning');
             return;
        }

        const today = new Date();
        const fileName = `Hasil_Gabungan_${today.getFullYear()}${String(today.getMonth()+1).padStart(2,'0')}${String(today.getDate()).padStart(2,'0')}.xlsx`;
        
        XLSX.writeFile(outputWorkbook, fileName, { cellStyles: true });
        
        showNotification(`✅ Proses Selesai! File <strong>${fileName}</strong> dengan ${outputWorkbook.SheetNames.length} sheet telah berhasil diunduh.`, 'success');

    } catch (generalError) {
        console.error("General Error:", generalError);
        showNotification('Terjadi kesalahan umum. Silakan coba lagi.', 'danger');
    } finally {
        processBtn.disabled = false;
        btnText.textContent = 'Proses Data & Unduh Hasil';
        spinner.classList.add('d-none');
    }
}

// Fungsi untuk mengambil pengaturan dari form
function getSettingsFromForm() {
    return {
        createdDate: formatDate(document.getElementById('createdDate').value),
        startDate: formatDate(document.getElementById('startDate').value),
        endDate: formatDate(document.getElementById('endDate').value),
        testMethod: "Positive",
        testCaseStatus: "Tested",
        executionStatus: "Passed",
        testCaseType: "Functional",
        testPriority: "Medium"
    };
}

// Logika inti pemrosesan
function processSingleFile(sourceData, settings, baseId, encDescription, globalSuffixCounter) {
    const processedData = [];
    const header = [
        "No.", "ENC Ref. No.", "ENC Description", "Test Case ID", "Module", "Path",
        "Test Case Type", "Test Priority", "Test Case Name (ID)", "Test Case Description (ID)",
        "Test Prerequisite", "Test Method", "Tools", "Created Date", "Created by",
        "Actual Test Start Date", "Actual Test End Date", "Tested by", "Test Case Status",
        "Test Execution Status", "Notes"
    ];
    processedData.push(header);

    let nomorUrut = 1;
    let prerequisiteSteps = [];
    let pathForCurrentGroup = '';
    const numSourceCols = sourceData[0]?.length || 0;

    for (let i = 1; i < sourceData.length; i++) {
        const row = sourceData[i];
        if (row && row[1] && String(row[1]).trim() !== "") {
            if (processedData.length > 1) {
                const lastOutputRow = processedData[processedData.length - 1];
                lastOutputRow[5] = pathForCurrentGroup;
                lastOutputRow[10] = prerequisiteSteps.join('\n');
            }
            prerequisiteSteps = [];
            pathForCurrentGroup = '';

            const newRow = [];
            newRow[0] = nomorUrut;
            newRow[1] = baseId;
            newRow[2] = encDescription;
            newRow[3] = `${baseId}-${globalSuffixCounter.current}`;
            newRow[4] = numSourceCols >= 39 ? row[38] : '';
            newRow[5] = '';
            newRow[6] = settings.testCaseType;
            newRow[7] = settings.testPriority;
            newRow[8] = numSourceCols >= 2 ? row[1] : '';
            newRow[9] = numSourceCols >= 3 ? row[2] : '';
            newRow[10] = '';
            newRow[11] = settings.testMethod;
            newRow[12] = '-';
            newRow[13] = settings.createdDate;
            newRow[14] = numSourceCols >= 38 ? row[37] : '';
            newRow[15] = settings.startDate;
            newRow[16] = settings.endDate;
            newRow[17] = numSourceCols >= 38 ? row[37] : '';
            newRow[18] = settings.testCaseStatus;
            newRow[19] = settings.executionStatus;
            newRow[20] = numSourceCols >= 33 ? row[32] : '';
            processedData.push(newRow);
            
            nomorUrut++;
            globalSuffixCounter.current++;
        }
        
        const stepText = (numSourceCols >= 14 && row && row[13]) ? String(row[13]).trim() : '';
        if (stepText !== "") {
            prerequisiteSteps.push(`${prerequisiteSteps.length + 1}. ${stepText}`);
            const lowerStepText = stepText.toLowerCase();
            const searchText = "akses menu ";
            if (lowerStepText.includes(searchText)) {
                pathForCurrentGroup = stepText.substring(lowerStepText.indexOf(searchText) + searchText.length);
            }
        }
    }
    
    if (processedData.length > 1) {
        const lastOutputRow = processedData[processedData.length - 1];
        lastOutputRow[5] = pathForCurrentGroup;
        lastOutputRow[10] = prerequisiteSteps.join('\n');
    }
    
    return processedData;
}

// Fungsi styling
function applyStyling(worksheet, data) {
    const header = data[0];
    const wscols = header.map(h => ({ wch: (h?.length || 5) + 15 }));
    worksheet['!cols'] = wscols;

    for (let R = 0; R < data.length; ++R) {
        for (let C = 0; C < data[R].length; ++C) {
            const cell_address = XLSX.utils.encode_cell({ r: R, c: C });
            const cell = worksheet[cell_address];
            if (!cell) continue;

            cell.s = cell.s || {};
            cell.s.font = cell.s.font || {};
            cell.s.alignment = cell.s.alignment || {};
            
            cell.s.font.name = 'Tahoma';
            cell.s.font.sz = 10;
            cell.s.alignment.vertical = 'top';
            
            if (R === 0) {
                cell.s.font.bold = true;
            }

            if (C === 5 || C === 10) {
                cell.s.alignment.wrapText = true;
            }
        }
    }
}