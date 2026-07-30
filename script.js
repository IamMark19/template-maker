const fileInput = document.getElementById('fileInput');
const tableContainer = document.getElementById('tableContainer');
const exportBtn = document.getElementById('exportBtn');
const exportTechBtn = document.getElementById('exportTechBtn');
const clearBtn = document.getElementById('clearBtn');

let rawData = [];
let groupedRows = [];

fileInput.addEventListener('change', handleFile, false);
exportBtn.addEventListener('click', () => handleExport(false), false);
exportTechBtn.addEventListener('click', () => handleExport(true), false);
clearBtn.addEventListener('click', clearPreview, false);

function isHyperlinkObj(v) { 
    return v && typeof v === "object" && !!v.hyperlink; 
}

function parseHyperlinkFormula(f) {
    if (!f) return null;
    const m = /HYPERLINK\(\s*"([^"]+)"\s*,\s*"([^"]*)"\s*\)/i.exec(f);
    if (m) return { hyperlink: m[1], text: m[2] || m[1] };
    return null;
}

// Utility function to format date as M/D/YY (e.g., 7/30/26)
function getFormattedTodayDate() {
    const today = new Date();
    const month = today.getMonth() + 1;
    const day = today.getDate();
    const year = String(today.getFullYear()).slice(-2);
    return `${month}/${day}/${year}`;
}

function handleFile(e) {
    const f = e.target.files[0];
    if (!f) return;
    const reader = new FileReader();

    reader.onload = function(ev) {
        try {
            const data = new Uint8Array(ev.target.result);
            const wb = XLSX.read(data, { type: 'array' });
            const sheet = wb.Sheets[wb.SheetNames[0]];
            if (!sheet || !sheet['!ref']) return;

            const range = XLSX.utils.decode_range(sheet['!ref']);
            let headerRowIndex = -1;

            // 1. Locate the header row containing "Department"
            for (let r = range.s.r; r <= range.e.r; r++) {
                for (let c = range.s.c; c <= range.e.c; c++) {
                    const cell = sheet[XLSX.utils.encode_cell({ r, c })];
                    const val = cell ? String(XLSX.utils.format_cell(cell)).trim().toLowerCase() : "";
                    if (val === "department") {
                        headerRowIndex = r;
                        break;
                    }
                }
                if (headerRowIndex !== -1) break;
            }

            if (headerRowIndex === -1) headerRowIndex = range.s.r;

            // 2. Extract column header titles
            const headersRaw = [];
            for (let c = range.s.c; c <= range.e.c; c++) {
                const cell = sheet[XLSX.utils.encode_cell({ r: headerRowIndex, c })];
                headersRaw.push(cell ? XLSX.utils.format_cell(cell).trim() : `Column${c + 1}`);
            }

            // 3. Process data rows AS-IS for preview
            const jsonData = [];
            for (let r = headerRowIndex + 1; r <= range.e.r; r++) {
                const rowObj = {};
                let rowHasValue = false;

                for (let c = range.s.c; c <= range.e.c; c++) {
                    let colName = headersRaw[c - range.s.c];

                    // Standardize header names
                    if (colName === "Change ID Caused By Request") colName = "Change ID";
                    if (colName === "Change Title Caused By Request" || colName === "Subject") colName = "Description";

                    const cell = sheet[XLSX.utils.encode_cell({ r, c })];
                    if (!cell) { rowObj[colName] = ""; continue; }

                    let text = XLSX.utils.format_cell(cell);
                    let linkObj = (cell.l && cell.l.Target) 
                        ? { text: text || cell.l.Target, hyperlink: cell.l.Target } 
                        : parseHyperlinkFormula(cell.f);

                    if (linkObj) {
                        rowObj[colName] = linkObj;
                        if (String(linkObj.text || "").trim()) rowHasValue = true;
                    } else {
                        rowObj[colName] = text;
                        if (String(text).trim()) rowHasValue = true;
                    }
                }

                if (rowHasValue) {
                    jsonData.push(rowObj);
                }
            }

            rawData = jsonData;
            groupedRows = groupByDepartment(jsonData);

            // Preview UI default layout
            const previewHeaders = [
                'Change Type', 'RequestID', 'Change ID', 
                'Description', 'Requester', 'UAT Owner', 
                'UAT Date', 'Request Status'
            ];

            renderPreview(previewHeaders, groupedRows);

            exportBtn.disabled = false;
            exportTechBtn.disabled = false;
            clearBtn.disabled = false;
        } catch (err) {
            console.error(err);
            alert('Failed to parse Excel file.');
        }
    };
    reader.readAsArrayBuffer(f);
}

function groupByDepartment(rows) {
    const groups = [];
    let currentDept = null;
    let currentRows = [];

    for (const r of rows) {
        const deptVal = r["Department"];
        const dept = isHyperlinkObj(deptVal) ? (deptVal.text || "") : String(deptVal || "");

        if (!dept.trim()) {
            if (currentDept) currentRows.push(r);
            continue;
        }

        if (!currentDept || dept !== currentDept) {
            if (currentDept) groups.push({ department: currentDept, rows: currentRows });
            currentDept = dept;
            currentRows = [r];
        } else {
            currentRows.push(r);
        }
    }
    if (currentDept) groups.push({ department: currentDept, rows: currentRows });
    return groups;
}

function renderPreview(headers, groups) {
    tableContainer.innerHTML = "";

    const table = document.createElement('table');
    let thead = `<thead><tr>${headers.map(h => `<th>${h}</th>`).join('')}</tr></thead>`;
    let tbody = "<tbody>";

    groups.forEach(group => {
        // Department Group Banner
        tbody += `<tr><td colspan="${headers.length}" style="background:#5451e0; color:white; font-weight:bold; text-align:left; padding:8px 12px;">${group.department}</td></tr>`;
        
        group.rows.forEach(row => {
            tbody += "<tr>";
            headers.forEach(h => {
                const val = row[h];
                if (isHyperlinkObj(val)) {
                    tbody += `<td><a href="${val.hyperlink}" target="_blank" style="color:#0000ff; text-decoration:underline;">${val.text || val.hyperlink}</a></td>`;
                } else {
                    tbody += `<td>${val || ""}</td>`;
                }
            });
            tbody += "</tr>";
        });
    });

    table.innerHTML = thead + tbody + "</tbody>";
    tableContainer.appendChild(table);
}

async function handleExport(includeTechnician) {
    if (!rawData.length) return;
    const workbook = new ExcelJS.Workbook();
    const sheet = workbook.addWorksheet('Announcement');

    const purpleHeaderFill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'FF5451E0' }
    };

    const whiteBoldFont = {
        name: 'Segoe UI',
        bold: true,
        color: { argb: 'FFFFFFFF' },
        size: 10
    };

    const lightBorder = {
        top: { style: 'thin', color: { argb: 'FFE2E8F0' } },
        bottom: { style: 'thin', color: { argb: 'FFE2E8F0' } },
        left: { style: 'thin', color: { argb: 'FFE2E8F0' } },
        right: { style: 'thin', color: { argb: 'FFE2E8F0' } }
    };

    const exportDateStr = getFormattedTodayDate();

    if (includeTechnician) {
        // ==========================================
        // 1. EXPORT WITH TECHNICIAN (FLAT TABLE)
        // ==========================================
        const exportHeaders = [
            'Department', 'Change Type', 'RequestID', 'Change ID', 
            'Description', 'Technician', 'Requester', 'UAT Owner', 
            'UAT Date', 'Request Status'
        ];

        // Header Row
        const hRow = sheet.addRow(exportHeaders);
        hRow.height = 24;
        hRow.eachCell(c => {
            c.font = whiteBoldFont;
            c.fill = purpleHeaderFill;
            c.alignment = { vertical: 'middle', horizontal: 'left' };
        });

        // Add Data Rows directly (flat layout)
        rawData.forEach(row => {
            const exportRow = { ...row };

            // Rules:
            const reqVal = exportRow['Requester'];
            exportRow['UAT Owner'] = isHyperlinkObj(reqVal) ? (reqVal.text || reqVal.hyperlink) : reqVal;
            exportRow['UAT Date'] = exportDateStr;

            const rowValues = exportHeaders.map(h => {
                const val = exportRow[h];
                if (isHyperlinkObj(val)) return val.text || val.hyperlink;
                return val || "";
            });

            const r = sheet.addRow(rowValues);
            r.height = 20;

            r.eachCell({ includeEmpty: true }, (cell, colIdx) => {
                const colKey = exportHeaders[colIdx - 1];
                const rawVal = exportRow[colKey];

                cell.font = { name: 'Segoe UI', size: 10 };
                cell.alignment = { vertical: 'middle', horizontal: 'left' };
                cell.border = lightBorder;

                if (isHyperlinkObj(rawVal)) {
                    cell.value = { text: rawVal.text || rawVal.hyperlink, hyperlink: rawVal.hyperlink };
                    cell.font = { name: 'Segoe UI', size: 10, color: { argb: "FF0000FF" }, underline: true };
                }
            });
        });

        // Set column widths
        const customWidths = {
            'Department': 25,
            'Change Type': 15,
            'RequestID': 16,
            'Change ID': 14,
            'Description': 65,
            'Technician': 22,
            'Requester': 22,
            'UAT Owner': 22,
            'UAT Date': 14,
            'Request Status': 16
        };

        sheet.columns.forEach((col, idx) => {
            const headerName = exportHeaders[idx];
            col.width = customWidths[headerName] || 20;
        });

    } else {
        // ==========================================
        // 2. STANDARD EXPORT (GROUPED BANNERS)
        // ==========================================
        const exportHeaders = [
            'Change Type', 'RequestID', 'Change ID', 
            'Description', 'Requester', 'UAT Owner', 
            'UAT Date', 'Request Status'
        ];

        // Main Header Row
        const hRow = sheet.addRow(exportHeaders);
        hRow.height = 24;
        hRow.eachCell(c => {
            c.font = whiteBoldFont;
            c.fill = purpleHeaderFill;
            c.alignment = { vertical: 'middle', horizontal: 'left' };
        });

        // Department Banner & Rows
        groupedRows.forEach(group => {
            const dRow = sheet.addRow([group.department]);
            dRow.height = 20;
            sheet.mergeCells(dRow.number, 1, dRow.number, exportHeaders.length);
            
            const deptCell = dRow.getCell(1);
            deptCell.font = whiteBoldFont;
            deptCell.fill = purpleHeaderFill;
            deptCell.alignment = { vertical: 'middle', horizontal: 'left' };

            group.rows.forEach(row => {
                const exportRow = { ...row };

                const reqVal = exportRow['Requester'];
                exportRow['UAT Owner'] = isHyperlinkObj(reqVal) ? (reqVal.text || reqVal.hyperlink) : reqVal;
                exportRow['UAT Date'] = exportDateStr;

                const rowValues = exportHeaders.map(h => {
                    const val = exportRow[h];
                    if (isHyperlinkObj(val)) return val.text || val.hyperlink;
                    return val || "";
                });

                const r = sheet.addRow(rowValues);
                r.height = 20;

                r.eachCell({ includeEmpty: true }, (cell, colIdx) => {
                    const colKey = exportHeaders[colIdx - 1];
                    const rawVal = exportRow[colKey];

                    cell.font = { name: 'Segoe UI', size: 10 };
                    cell.alignment = { vertical: 'middle', horizontal: 'left' };
                    cell.border = lightBorder;

                    if (isHyperlinkObj(rawVal)) {
                        cell.value = { text: rawVal.text || rawVal.hyperlink, hyperlink: rawVal.hyperlink };
                        cell.font = { name: 'Segoe UI', size: 10, color: { argb: "FF0000FF" }, underline: true };
                    }
                });
            });
        });

        // Set column widths
        const customWidths = {
            'Change Type': 15,
            'RequestID': 16,
            'Change ID': 14,
            'Description': 65,
            'Requester': 22,
            'UAT Owner': 22,
            'UAT Date': 14,
            'Request Status': 16
        };

        sheet.columns.forEach((col, idx) => {
            const headerName = exportHeaders[idx];
            col.width = customWidths[headerName] || 20;
        });
    }

    const buffer = await workbook.xlsx.writeBuffer();
    const fileName = includeTechnician ? 'Release_Template_with_Tech.xlsx' : 'Release_Template_detailed.xlsx';
    saveAs(new Blob([buffer]), fileName);
}

function clearPreview() {
    rawData = []; 
    groupedRows = [];
    fileInput.value = "";
    tableContainer.innerHTML = '<div class="placeholder"><p>hehehehe</p></div>';
    exportBtn.disabled = true;
    exportTechBtn.disabled = true;
    clearBtn.disabled = true;
}