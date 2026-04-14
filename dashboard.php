<?php 
session_start(); 
include "db.php";
if(!isset($_SESSION['user'])){ header("Location: login.php"); exit; }
?>
<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>Finance Procurement Report System </title>

  <link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/handsontable/dist/handsontable.full.min.css">
  <script src="https://cdn.jsdelivr.net/npm/handsontable/dist/handsontable.full.min.js"></script>
  
  <script src="https://cdnjs.cloudflare.com/ajax/libs/exceljs/4.3.0/exceljs.min.js"></script>
  <script src="https://cdnjs.cloudflare.com/ajax/libs/FileSaver.js/2.0.5/FileSaver.min.js"></script>

  <style>
    :root {
      --primary-green: #1b5e20;
      --secondary-green: #2e7d32;
      --accent-yellow: #fbc02d;
      --danger-red: #d32f2f;
      --info-blue: #0277bd;
      --dark-grey: #263238;
      --light-grey: #f8f9fa;
      --border-radius: 12px;
      --shadow-sm: 0 2px 4px rgba(0,0,0,0.05);
      --shadow-md: 0 4px 12px rgba(0,0,0,0.1);
    }
      
    body { margin:0; font-family: 'Segoe UI', Roboto, sans-serif; background:#f4f7f6; color: #333; overflow-x: hidden; }
    
    .header {
      background: linear-gradient(135deg, var(--primary-green) 0%, var(--secondary-green) 100%);
      color: #ffffff; padding: 10px 20px; display: flex; flex-wrap: wrap; justify-content: space-between;
      align-items: center; box-shadow: 0 4px 20px rgba(0,0,0,0.15); border-top: 4px solid var(--accent-yellow); gap: 10px;
      align-items: center; box-shadow: 0 4px 20px rgba(0,0,0,0.15); border-bottom: 4px solid var(--accent-yellow); gap: 10px;
    }
    .header .title-container { display: flex; align-items: center; gap: 15px; flex: 1; min-width: 250px; justify-content: center; }
    .header-logo { height: clamp(35px, 5vw, 55px); width: auto; filter: drop-shadow(0 2px 4px rgba(0,0,0,0.2)); }
    .header span { font-size: clamp(0.8rem, 2vw, 1.1rem); font-weight: 700; text-transform: uppercase; letter-spacing: 1px; text-align: center; line-height: 1.2; }
    
    .logout { 
      background: rgba(255,255,255,0.1); padding: 8px 18px; color: white; text-decoration: none; 
      border-radius: 6px; font-size: 13px; font-weight: bold; transition: 0.3s; border: 1px solid rgba(255,255,255,0.3);
    }
    .logout:hover { background: var(--danger-red); border-color: transparent; }
    
    .nav-bar { display:flex; gap:10px; background: var(--dark-grey); padding:10px 15px; overflow-x: auto; white-space: nowrap; }
    .nav-bar button { 
        background: transparent; color: #adb5bd; border:none; padding:10px 15px; 
        border-radius: 6px; cursor:pointer; font-weight: 600; transition: 0.3s; font-size: 14px;
    }
    .nav-bar button.active { background: var(--primary-green); color: white; }
    
    .toolbar { 
      background: white; padding: 15px; display: flex; flex-wrap: wrap; gap: 10px; 
      margin: 15px; border-radius: var(--border-radius); box-shadow: var(--shadow-sm); align-items: center; 
    }
    .toolbar button { 
      background: #eceff1; color: #455a64; border: 1px solid #cfd8dc; padding: 8px 14px; border-radius: 6px; 
      cursor: pointer; font-size: 13px; font-weight: 600; flex-grow: 1; max-width: 200px;
    }
    .separator { border-left: 1px solid #cfd8dc; margin: 0 5px; height: 25px; display: none; }
    @media (min-width: 768px) { .separator { display: block; } .toolbar button { flex-grow: 0; } }
    
    #spreadsheet { height: 70vh; margin:0 15px 20px; border-radius: var(--border-radius); border:1px solid #ddd; overflow:hidden; background: white; box-shadow: var(--shadow-md); }

    /* Custom Styles for Handsontable Cells */
    .boldText { font-weight: 600; }
    .wrapText { white-space: pre-wrap !important; word-wrap: break-word; }
    .cellBorder { border: 1px solid #ccc !important; }
    .noBorder { border: none !important; }
    .redText { color: #d32f2f !important; }
    .htCenter { text-align: center !important; }

    /* NEW: Style for the Header Row (Row 10) in the System */
    .headerRow {
      background-color: #1b5e20 !important; /* Matches --primary-green */
      color: #ffffff !important;           /* White text */
      font-weight: 700 !important;
      text-align: center !important;
      vertical-align: middle !important;
    }

    #filesPage { display:none; padding:20px; max-width: 1300px; margin: 0 auto; }
    .page-header { display: flex; flex-wrap: wrap; justify-content: space-between; align-items: flex-end; margin-bottom: 20px; gap: 10px; border-bottom: 2px solid #e0e0e0; padding-bottom: 10px; }
    
    .stats-container { display: grid; grid-template-columns: repeat(auto-fit, minmax(250px, 1fr)); gap: 15px; margin-bottom: 30px; }
    .stat-card { 
        background: white; padding: 30px; border-radius: var(--border-radius); 
        box-shadow: var(--shadow-sm); position: relative; overflow: hidden; border: 1px solid #eee;
    }
    .stat-card p { font-size: 28px; margin: 5px 0 0; }

    .file-table-container { 
        background: white; border-radius: var(--border-radius); 
        box-shadow: var(--shadow-md); overflow-x: auto; border: 1px solid #eee;
    }
    #filesPage table { border-collapse:collapse; width:100%; min-width: 600px; }
    #filesPage th, #filesPage td { padding: 12px; text-align: left; border-bottom: 1px solid #eee; }

    #saveModal { 
      display:none; position:fixed; top:50%; left:50%; transform:translate(-50%,-50%); 
      background:white; padding:25px; border-radius:15px; width: 90%; max-width: 400px;
      box-shadow: 0 25px 60px rgba(0,0,0,0.3); z-index:1000;
    }
    .modal-overlay { display:none; position:fixed; top:0; left:0; width:100%; height:100%; background:rgba(0,0,0,0.5); z-index:999; backdrop-filter: blur(3px); }

    .badge { padding: 4px 8px; border-radius: 4px; font-size: 11px; font-weight: bold; }
    .status-done { background: #e8f5e9; color: #2e7d32; }
    .status-outgoing { background: #e3f2fd; color: #0277bd; }

    #activeFileNameDisplay { 
      margin: 10px 15px; padding: 12px; background: white; 
      border-radius: 8px; display: flex; flex-wrap: wrap; align-items: center; gap: 15px; 
      font-size: 14px; box-shadow: var(--shadow-sm); border-left: 5px solid var(--primary-green);
    }
  </style>
</head>

<body>

<div class="header">
    <div class="title-container">
        <img src="Logo1.png" class="header-logo" alt="Logo">
        <span>16th Installation Management Battalion Philippine Army<br><small style="font-weight:400; opacity:0.9;">Finance Procurement Management System</small></span>
        <img src="Logo.png" class="header-logo" alt="Logo">
    </div>
    <a href="logout.php" class="logout">Sign Out</a>
</div>

<div class="nav-bar">
  <button id="navSheet" class="active" onclick="showPage('sheetPage')">📊 Dashboard</button>
  <button onclick="resetAndNew()">📄 New Project</button>
  <button id="navFiles" onclick="showPage('filesPage')">📂 Saved Projects</button>
</div>

<div id="sheetPage">
  <div id="activeFileNameDisplay">
      <span id="currentFileNameTxt" style="font-weight:700; color:var(--dark-grey);">New Project</span>
      <div class="status-update-box" id="statusBox" style="display:none; align-items:center; gap:10px; border-left: 1px solid #eee; padding-left:15px;">
          <span style="color:#94a3b8; font-weight:600; font-size:11px; text-transform:uppercase;">Status</span>
          <select id="editFileStatus" style="padding:5px; border-radius:5px;">
              <option value="Outgoing">Outgoing</option>
              <option value="Done">Done</option>
          </select>
      </div>
  </div>

  <div class="toolbar">
    <button onclick="addRow()">➕ Add Row</button>
    <button onclick="addCol()">➕ Add Column</button>
    <div class="separator"></div>
    <button onmousedown="event.preventDefault(); applyMerge()">Merge Cells</button>
    <button onmousedown="event.preventDefault(); applyUnmerge()">Unmerge</button>
    <div class="separator"></div>
    <button onclick="handleSaveAction()" id="saveBtn" style="background:var(--primary-green); color:white; border:none; min-width:140px;">💾 Save Changes</button>
  </div>
  <div id="spreadsheet"></div>
</div>

<div id="filesPage">
  <div class="page-header">
    <h2 class="page-title">File <span>Records</span></h2>
  </div>
  
  <div class="stats-container" style="display: grid; grid-template-columns: repeat(3, 1fr); gap: 10px; margin-bottom: 10px;">
    <div class="stat-card" style="padding: 10px; background: white; border-radius: 6px; border: 1px solid #eee; box-shadow: var(--shadow-sm); border-top: 3px solid var(--dark-grey);">
        <h4 style="margin: 0; font-size: 11px; text-transform: uppercase; color: #666;">Total</h4>
        <p id="statTotal" style="margin: 5px 0 0; font-size: 20px; font-weight: 800; color: var(--dark-grey);">0</p>
    </div>
    <div class="stat-card" style="padding: 10px; background: white; border-radius: 6px; border: 1px solid #eee; box-shadow: var(--shadow-sm); border-top: 3px solid var(--info-blue);">
        <h4 style="margin: 0; font-size: 11px; text-transform: uppercase; color: #666;">On-Going</h4>
        <p id="statOutgoing" style="margin: 5px 0 0; font-size: 20px; font-weight: 800; color: var(--info-blue);">0</p>
    </div>
    <div class="stat-card" style="padding: 10px; background: white; border-radius: 6px; border: 1px solid #eee; box-shadow: var(--shadow-sm); border-top: 3px solid var(--secondary-green);">
        <h4 style="margin: 0; font-size: 11px; text-transform: uppercase; color: #666;">Completed</h4>
        <p id="statDone" style="margin: 5px 0 0; font-size: 20px; font-weight: 800; color: var(--secondary-green);">0</p>
    </div>
  </div>

  <div class="file-table-container">
      <table>
        <thead>
            <tr>
                <th width="20%">Modified</th>
                <th width="15%">Status</th>
                <th>File Name</th>
                <th width="150px">Actions</th>
            </tr>
        </thead>
        <tbody id="filesTable"></tbody>
      </table>
  </div>
</div>

<div class="modal-overlay" id="overlay" onclick="closeModal()"></div>
<div id="saveModal">
  <h3 style="margin-top:0;">Archive Project</h3>
  <div style="margin-bottom:15px;">
    <label style="font-size:12px; font-weight:bold; display:block; margin-bottom:5px;">Project Title</label>
    <input type="text" id="modalFileName" placeholder="Enter file name..." style="width:100%; padding:10px; border:1px solid #ddd; border-radius:8px; box-sizing:border-box;">
  </div>
  <div style="margin-bottom:20px;">
    <label style="font-size:12px; font-weight:bold; display:block; margin-bottom:5px;">Status</label>
    <label><input type="radio" name="status" value="Outgoing" checked> Outgoing</label>
    <label style="margin-left:15px;"><input type="radio" name="status" value="Done"> Completed</label>
  </div>
  <div style="display:flex; gap:10px;">
    <button onclick="confirmSaveNew()" style="flex:2; padding:12px; background:var(--primary-green); color:white; border:none; border-radius:8px; cursor:pointer; font-weight:bold;">Create Entry</button>
    <button onclick="closeModal()" style="flex:1; padding:12px; background:#eee; border:none; border-radius:8px; cursor:pointer;">Cancel</button>
  </div>
</div>

<script>
// --- HANDSONTABLE CONFIG ---
const container = document.getElementById('spreadsheet');
const leftLogo = "Logo1.png"; 
const rightLogo = "Logo.png";
const LOGO_WIDTH = 90;
let currentOpenFileIndex = null; 

function getHeaderOffset(totalCols) {
    const headerWidth = 6; 
    return Math.max(0, Math.floor((totalCols - headerWidth) / 2));
}

function logoRenderer(instance, td, row, col, prop, value, cellProperties) {
    Handsontable.renderers.TextRenderer.apply(this, arguments);
    if (value && (value.includes('.png') || value.includes('.jpg'))) {
        td.innerHTML = `<img src="${value}" style="width:70px; height:70px; object-fit:contain; display:block; margin:auto;">`;
    }
    td.style.background = '#fff';
    td.style.verticalAlign = 'middle';
    td.style.textAlign = 'center';
    td.style.border = 'none'; 
}

function getDynamicConfig(totalCols) {
    const start = getHeaderOffset(totalCols);
    const leftLogoCol = start;
    const centerStart = start + 1;
    const rightLogoCol = start + 5;
    const textSpan = 4;
    let merges = [
        {row: 0, col: leftLogoCol, rowspan: 3, colspan: 1}, 
        {row: 0, col: rightLogoCol, rowspan: 3, colspan: 1}
    ];
    [0, 1, 2, 3, 5, 6, 7].forEach(r => {
        merges.push({row: r, col: centerStart, rowspan: 1, colspan: textSpan});
    });
    return { leftLogoCol, rightLogoCol, centerStart, centerEnd: start + 4, merges };
}

function refreshHeaderData() {
    const totalCols = hot.countCols();
    const config = getDynamicConfig(totalCols);
    let currentData = hot.getData();

    for(let r=0; r<=9; r++) {
        for(let c=0; c<totalCols; c++) { currentData[r][c] = ''; }
    }

    currentData[0][config.leftLogoCol] = leftLogo;
    currentData[0][config.centerStart] = 'HEADQUARTERS';
    currentData[0][config.rightLogoCol] = rightLogo;
    currentData[1][config.centerStart] = '16th INSTALLATION MANAGEMENT BATTALION';
    currentData[2][config.centerStart] = 'INSTALLATION MANAGEMENT COMMAND, PHILIPPINE ARMY';
    currentData[3][config.centerStart] = 'Camp Riego de Dios, Tanza, Cavite';
    currentData[5][config.centerStart] = 'Procurement Status Report';
    currentData[6][config.centerStart] = 'as of: 20 December 2025';
    currentData[7][config.centerStart] = 'ON GOING PROJECTS 2025';

    const columnHeaders = ["L/I", "FUND CATEGORY", "ASA Nr", "URP Nr", "ORS Nr", "Contract/PO Nr", "Name of Projects", "Supplier/ Contractor", "ABC (PhP)", "Bid Price (PhP)", "Contract Cost (PhP)", "Date Awarded (NOA)", "Date of NTP", "Date of PO", "Date of PO signature @ Enduser", "Date of PO received @ G10, RCPA", "Date of PO received @ 14FAU", "Date received @ Enduser (10days validity)", "Date of NOD received @ G10, RCPA", "Date of NOD received @ 14FPAO", "Date of Delivery", "Date of DV signature @ Enduser", "Date of DV received @ G10, RCPA", "Date of DV received @ 14FAU", "STATUS", "Remarks"];

    if (totalCols < columnHeaders.length) {
        hot.alter('insert_col_end', columnHeaders.length - totalCols);
    }

    columnHeaders.forEach((headerText, index) => {
        currentData[9][index] = headerText;
    });

    hot.updateSettings({
        data: currentData,
        mergeCells: config.merges,
        colWidths: (index) => (index === config.leftLogoCol || index === config.rightLogoCol) ? LOGO_WIDTH : 130
    });
    hot.render();
}

let hot = new Handsontable(container, {
    data: Array.from({length:40},()=>Array(26).fill('')), 
    rowHeaders: true,
    colHeaders: true,
    contextMenu: true,
    manualColumnResize: true,
    manualRowResize: true,
    mergeCells: true, 
    licenseKey: 'non-commercial-and-evaluation',
    autoRowSize: true,
    autoColumnSize: true,
    rowHeights: (row) => {
        if (row === 0) return 80;
        if (row >= 1 && row <= 8) return 25;
        if (row === 9) return 75;
        return undefined;
    },
    cells: function(row, col) {
        let cellProp = {};
        const config = getDynamicConfig(this.instance.countCols());
        cellProp.className = 'boldText wrapText'; 
        
        if (row <= 8) { cellProp.className += ' noBorder'; }

        // --- UPDATED: Applying the dark green color to Row 10 (index 9) ---
        if (row === 9) { 
            cellProp.className = 'headerRow cellBorder'; 
        }

        if ((row === 0 && col === config.leftLogoCol) || (row === 0 && col === config.rightLogoCol)) {
            cellProp.renderer = logoRenderer;
        }
        if (row <= 7 && col >= config.centerStart && col <= config.centerEnd) {
            cellProp.className += ' htCenter';
            if (row === 7) cellProp.className += ' redText';
        }
        if (row >= 10) { cellProp.className += ' cellBorder'; }
        return cellProp;
    }
});

refreshHeaderData();

// --- APP LOGIC ---
function resetAndNew() {
    if(confirm("Discard current changes and start a new project?")) {
        currentOpenFileIndex = null;
        document.getElementById('currentFileNameTxt').innerText = "New Project";
        document.getElementById('saveBtn').innerText = "💾 Save Project";
        document.getElementById('statusBox').style.display = 'none';
        hot.loadData(Array.from({length:40},()=>Array(26).fill('')));
        refreshHeaderData();
        showPage('sheetPage');
    }
}

function handleSaveAction() {
    if (currentOpenFileIndex !== null) updateExistingFile();
    else openSaveModal();
}

function updateExistingFile() {
    let files = JSON.parse(localStorage.getItem("files") || "[]");
    let f = files[currentOpenFileIndex];
    f.data = hot.getData();
    f.status = document.getElementById('editFileStatus').value;
    f.meta = getSheetMeta();
    f.merges = hot.getPlugin('mergeCells').mergedCellsCollection.mergedCells;
    f.date = new Date().toLocaleString();
    localStorage.setItem("files", JSON.stringify(files));
    alert("Archive Updated!");
}

function confirmSaveNew(){
    let name = document.getElementById("modalFileName").value.trim();
    let status = document.querySelector('input[name="status"]:checked').value;
    if(!name) return alert("Please enter project name");

    let files = JSON.parse(localStorage.getItem("files") || "[]");
    files.push({ 
        name, status, 
        date: new Date().toLocaleString(), 
        data: hot.getData(),
        meta: getSheetMeta(),
        merges: hot.getPlugin('mergeCells').mergedCellsCollection.mergedCells 
    });

    localStorage.setItem("files", JSON.stringify(files));
    currentOpenFileIndex = files.length - 1;
    document.getElementById('currentFileNameTxt').innerText = name;
    document.getElementById('editFileStatus').value = status;
    document.getElementById('statusBox').style.display = 'flex';
    document.getElementById('saveBtn').innerText = "💾 Update";
    closeModal();
    alert("Project Saved!");
}

function getSheetMeta() {
    let meta = [];
    for(let r=0; r<hot.countRows(); r++){
        for(let c=0; c<hot.countCols(); c++){
            let cls = hot.getCellMeta(r,c).className;
            if(cls) meta.push({r, c, cls});
        }
    }
    return meta;
}

function loadFiles(){
    let t = document.getElementById("filesTable");
    let files = JSON.parse(localStorage.getItem("files") || "[]");
    
    document.getElementById('statTotal').innerText = files.length;
    document.getElementById('statOutgoing').innerText = files.filter(f=>f.status==='Outgoing').length;
    document.getElementById('statDone').innerText = files.filter(f=>f.status==='Done').length;

    t.innerHTML = files.length ? "" : "<tr><td colspan='4' style='text-align:center; padding:50px;'>No projects found.</td></tr>";
    
    files.forEach((f, i) => {
        let statusCls = f.status === 'Done' ? 'status-done' : 'status-outgoing';
        t.innerHTML += `<tr>
            <td>${f.date}</td>
            <td><span class="badge ${statusCls}">${f.status}</span></td>
            <td><strong>${f.name}</strong></td>
            <td style="text-align: right; padding: 8px; white-space: nowrap;">
    <div style="display: flex; gap: 4px; justify-content: flex-end;">
        <button onclick="loadFile(${i})" title="Open Project" style="cursor: pointer; background: #e3f2fd; color: #1976d2; border: 1px solid #bbdefb; padding: 4px 8px; border-radius: 4px; font-size: 12px; display: flex; align-items: center; gap: 4px;">
            📂 <span style="font-weight: 600;">Open</span>
        </button>
        <button onclick="downloadExcel(${i})" title="Export to Excel" style="cursor: pointer; background: #e8f5e9; color: #2e7d32; border: 1px solid #c8e6c9; padding: 4px 8px; border-radius: 4px; font-size: 12px; display: flex; align-items: center; gap: 4px;">
            📊 <span style="font-weight: 600;">Export</span>
        </button>
        <button onclick="deleteFile(${i})" title="Delete Archive" style="cursor: pointer; background: #ffebee; color: #c62828; border: 1px solid #ffcdd2; padding: 4px 8px; border-radius: 4px; font-size: 12px; display: flex; align-items: center; gap: 4px;">
            🗑️ <span style="font-weight: 600;">Delete</span>
        </button>
    </div>
</td>
        </tr>`;
    });
}

function loadFile(i){
    let files = JSON.parse(localStorage.getItem("files") || "[]");
    let f = files[i];
    currentOpenFileIndex = i;
    document.getElementById('currentFileNameTxt').innerText = f.name;
    document.getElementById('editFileStatus').value = f.status;
    document.getElementById('statusBox').style.display = 'flex';
    document.getElementById('saveBtn').innerText = "💾 Update";
    hot.loadData(f.data);
    if(f.meta) f.meta.forEach(m => hot.setCellMeta(m.r, m.c, 'className', m.cls));
    hot.updateSettings({ mergeCells: f.merges || [] });
    hot.render();
    showPage('sheetPage');
}

function showPage(p){
    document.getElementById('sheetPage').style.display = p==='sheetPage'?'block':'none';
    document.getElementById('filesPage').style.display = p==='filesPage'?'block':'none';
    document.getElementById('navSheet').className = p==='sheetPage'?'active':'';
    document.getElementById('navFiles').className = p==='filesPage'?'active':'';
    if(p==='filesPage') loadFiles();
    setTimeout(() => hot.render(), 100); 
}

function openSaveModal(){ document.getElementById('saveModal').style.display='block'; document.getElementById('overlay').style.display='block'; }
function closeModal(){ document.getElementById('saveModal').style.display='none'; document.getElementById('overlay').style.display='none'; }

function deleteFile(i) {
    if(confirm("Delete this archive?")) {
        let files = JSON.parse(localStorage.getItem("files") || "[]");
        files.splice(i, 1);
        localStorage.setItem("files", JSON.stringify(files));
        loadFiles();
    }
}

// --- UPDATED EXCEL DOWNLOAD FUNCTION ---
async function downloadExcel(i) {
    let files = JSON.parse(localStorage.getItem("files") || "[]");
    let f = files[i];
    
    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet('Finance Report');

    async function addImageToSheet(url, worksheet, row, col) {
        try {
            const response = await fetch(url);
            const blob = await response.blob();
            const arrayBuffer = await blob.arrayBuffer();
            const imageId = workbook.addImage({
                buffer: arrayBuffer,
                extension: 'png',
            });
            worksheet.addImage(imageId, {
                tl: { col: col, row: row },
                ext: { width: 80, height: 80 },
                editAs: 'oneCell'
            });
        } catch (e) { console.error("Logo fetch error:", e); }
    }

    worksheet.columns = f.data[0].map(() => ({ width: 18 }));

    f.data.forEach((rowData, rowIndex) => {
        const cleanRowData = rowData.map(val => (val === "Logo1.png" || val === "Logo.png") ? "" : val);
        const row = worksheet.addRow(cleanRowData);
        row.height = rowIndex === 0 ? 65 : (rowIndex === 9 ? 60 : 20);
        
        cleanRowData.forEach((cellValue, colIndex) => {
            const cell = row.getCell(colIndex + 1);
            cell.alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };

            if (rowIndex <= 8) {
                cell.font = { bold: true, name: 'Arial', size: 10 };
            }
            if (rowIndex === 9) {
                cell.font = { bold: true, color: { argb: 'FFFFFFFF' } };
                cell.fill = {
                    type: 'pattern',
                    pattern: 'solid',
                    fgColor: { argb: 'FF1B5E20' }
                };
                cell.border = { top: {style:'thin'}, left: {style:'thin'}, bottom: {style:'thin'}, right: {style:'thin'} };
            }
            if (rowIndex === 7) {
                cell.font = { bold: true, color: { argb: 'FFFF0000' }, size: 12 };
            }
            if (rowIndex >= 10) {
                cell.border = { top: {style:'thin'}, left: {style:'thin'}, bottom: {style:'thin'}, right: {style:'thin'} };
            }
        });
    });

    if (f.merges) {
        f.merges.forEach(m => {
            try {
                worksheet.mergeCells(m.row + 1, m.col + 1, m.row + m.rowspan, m.col + m.colspan);
            } catch(e) { }
        });
    }

    const config = getDynamicConfig(f.data[0].length);
    await addImageToSheet(leftLogo, worksheet, 0, config.leftLogoCol);
    await addImageToSheet(rightLogo, worksheet, 0, config.rightLogoCol);

    const buffer = await workbook.xlsx.writeBuffer();
    saveAs(new Blob([buffer]), `${f.name}.xlsx`);
}

function applyMerge() { 
    let sel = hot.getSelected();
    if(sel) sel.forEach(s => hot.getPlugin('mergeCells').merge(Math.min(s[0],s[2]), Math.min(s[1],s[3]), Math.max(s[0],s[2]), Math.max(s[1],s[3])));
}
function applyUnmerge() {
    let sel = hot.getSelected();
    if(sel) sel.forEach(s => hot.getPlugin('mergeCells').unmerge(Math.min(s[0],s[2]), Math.min(s[1],s[3]), Math.max(s[0],s[2]), Math.max(s[1],s[3])));
}
function addCol() { hot.alter('insert_col_end'); refreshHeaderData(); }
function addRow() { hot.alter('insert_row_below'); }

window.addEventListener('resize', () => hot.render());
</script>
</body>
</html>