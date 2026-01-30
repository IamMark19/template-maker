const fileInput = document.getElementById('fileInput');
const tableContainer = document.getElementById('tableContainer');
const exportBtn = document.getElementById('exportBtn');
const clearBtn = document.getElementById('clearBtn');

let rawData = [];
let headerRow = null;
let groupedRows = [];

fileInput.addEventListener('change', handleFile, false);
exportBtn.addEventListener('click', handleExport, false);
clearBtn.addEventListener('click', clearPreview, false);

function isHyperlinkObj(v) { return v && typeof v === "object" && !!v.hyperlink; }

function isEmptyRow(r) {
    if (!r) return true;
    return Object.values(r).every(v => {
        if (isHyperlinkObj(v)) return String(v.text || "").trim() === "";
        return v === null || v === undefined || String(v).trim() === "";
    });
}

function uniqueHeaders(headers) {
    const seen = new Map();
    return headers.map((h, idx) => {
        let key = (h == null ? "" : String(h)).trim();
        if (!key) key = `Column${idx + 1}`;
        const count = (seen.get(key) || 0) + 1;
        seen.set(key, count);
        return count === 1 ? key : `${key}_${count}`;
    });
}

function parseHyperlinkFormula(f) {
    if (!f) return null;
    const m = /HYPERLINK\(\s*"([^"]+)"\s*,\s*"([^"]*)"\s*\)/i.exec(f);
    if (m) return { hyperlink: m[1], text: m[2] || m[1] };
    return null;
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
            let headerRowIndex = range.s.r;
            let bestFilled = -1;
            let deptRow = -1;

            for (let r = range.s.r; r <= Math.min(range.e.r, range.s.r + 20); r++) {
                let filled = 0;
                let hasDept = false;
                for (let c = range.s.c; c <= range.e.c; c++) {
                    const cell = sheet[XLSX.utils.encode_cell({ r, c })];
                    const text = cell ? XLSX.utils.format_cell(cell) : "";
                    if (String(text).trim()) filled++;
                    if (String(text).trim().toLowerCase() === "department") hasDept = true;
                }
                if (hasDept) { deptRow = r; break; }
                if (filled > bestFilled) { bestFilled = filled; headerRowIndex = r; }
            }
            if (deptRow >= 0) headerRowIndex = deptRow;

            const headersRaw = [];
            for (let c = range.s.c; c <= range.e.c; c++) {
                const cell = sheet[XLSX.utils.encode_cell({ r: headerRowIndex, c })];
                headersRaw.push(cell ? XLSX.utils.format_cell(cell) : `Column${c - range.s.c + 1}`);
            }
            const headers = uniqueHeaders(headersRaw);

            const jsonData = [];
            for (let r = headerRowIndex + 1; r <= range.e.r; r++) {
                const rowObj = {};
                let rowHasValue = false;
                for (let c = range.s.c; c <= range.e.c; c++) {
                    const h = headers[c - range.s.c];
                    const cell = sheet[XLSX.utils.encode_cell({ r, c })];
                    if (!cell) { rowObj[h] = ""; continue; }

                    let text = XLSX.utils.format_cell(cell);
                    let linkObj = (cell.l && cell.l.Target) ? { text: text || cell.l.Target, hyperlink: cell.l.Target } : parseHyperlinkFormula(cell.f);

                    if (linkObj) {
                        rowObj[h] = linkObj;
                        if (String(linkObj.text || "").trim()) rowHasValue = true;
                    } else {
                        rowObj[h] = text;
                        if (String(text).trim()) rowHasValue = true;
                    }
                }
                if (rowHasValue) jsonData.push(rowObj);
            }

            rawData = jsonData;
            headerRow = headers;
            groupedRows = groupByDepartment(jsonData);
            renderPreview(headerRow, groupedRows);

            exportBtn.disabled = false;
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
        if (isEmptyRow(r)) continue;
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
    
    const copyBtn = document.createElement('button');
    copyBtn.innerHTML = "<span>📋</span> Copy to HTML Table";
    copyBtn.className = "btn btn-glow primary";
    copyBtn.style.marginBottom = "20px";
    copyBtn.onclick = copyTableAsHTML;
    tableContainer.appendChild(copyBtn);

    const table = document.createElement('table');
    const displayHeaders = headers.length > 1 ? headers.slice(1) : [headers[0]];
    
    let thead = `<thead><tr>${displayHeaders.map(h => `<th>${h}</th>`).join('')}</tr></thead>`;
    let tbody = "<tbody>";

    groups.forEach(group => {
        tbody += `<tr><td colspan="${displayHeaders.length}" style="background:#4f46e5; color:white; font-weight:bold; text-align:center;">${group.department}</td></tr>`;
        group.rows.forEach(row => {
            tbody += "<tr>";
            displayHeaders.forEach(h => {
                const val = row[h];
                if (isHyperlinkObj(val)) {
                    tbody += `<td><a href="${val.hyperlink}" target="_blank" style="color:#4f46e5; font-weight:600;">${val.text || val.hyperlink}</a></td>`;
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

function copyTableAsHTML() {
    const table = tableContainer.querySelector('table');
    if (!table) return;
    const wrapper = `<div style="max-width:100%; overflow-x:auto;">${table.outerHTML}</div>`;
    navigator.clipboard.writeText(wrapper).then(() => {
        const gifAlert = document.getElementById('gifAlert');
        gifAlert.style.display = 'flex';
        setTimeout(() => { gifAlert.style.display = 'none'; }, 2500);
    });
}

async function handleExport() {
    if (!groupedRows.length) return;
    const workbook = new ExcelJS.Workbook();
    const sheet = workbook.addWorksheet('Announcement');
    const displayHeaders = headerRow.length > 1 ? headerRow.slice(1) : [headerRow[0]];

    const hRow = sheet.addRow(displayHeaders);
    hRow.eachCell(c => {
        c.font = { bold: true, color: { argb: 'FFFFFFFF' } };
        c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: '4f46e5' } };
    });

    groupedRows.forEach(group => {
        const dRow = sheet.addRow([group.department]);
        sheet.mergeCells(dRow.number, 1, dRow.number, displayHeaders.length);
        dRow.getCell(1).font = { bold: true, color: { argb: 'FFFFFFFF' } };
        dRow.getCell(1).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: '6366f1' } };

        group.rows.forEach(row => {
            const dataRow = displayHeaders.map(h => isHyperlinkObj(row[h]) ? (row[h].text || row[h].hyperlink) : (row[h] || ""));
            const r = sheet.addRow(dataRow);
            r.eachCell((cell, colIdx) => {
                const v = row[displayHeaders[colIdx - 1]];
                if (isHyperlinkObj(v)) {
                    cell.value = { text: v.text || v.hyperlink, hyperlink: v.hyperlink };
                    cell.font = { color: { argb: "FF0000FF" }, underline: true };
                }
            });
        });
    });

    const buffer = await workbook.xlsx.writeBuffer();
    saveAs(new Blob([buffer]), 'Release_Template_detailed.xlsx');
}

function clearPreview() {
    rawData = []; groupedRows = [];
    fileInput.value = "";
    tableContainer.innerHTML = '<div class="placeholder"><div class="placeholder-icon">📊</div><p>Upload an Excel file to see the magic happen</p></div>';
    exportBtn.disabled = true;
    clearBtn.disabled = true;
}