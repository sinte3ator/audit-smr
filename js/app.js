'use strict';

/* ============================== СОСТОЯНИЕ ============================== */
const S = {
  wb: null, fileName: '', originalBuffer: null, saveHandle: null,
  currentSheet: '', sheetData: {}, sheetColors: {},
  uniqueSections: [], columns: [], works: [],
  myCells: {}, foremanCells: {}, cellColors: {},
  selSection: '', selColIndex: -1, selWorkIdx: 0,
  filterSection: '', filterColIndex: -1, filterWorkIdxs: new Set(),
  hideAccepted: false, darkMode: false,
  pendingChanges: {}, originalCells: {},
  isCellLocked: false, isMobile: window.innerWidth < 1024,
  undoStack: [], lastChangeTime: 0, summaryFilter: null,
  cachedSheets: new Set(), autoSaveInterval: 30000,
  isLocalStorageAvailable: false, demoMode: false
};
const MS = { floors: new Set(), works: new Set() };
const MAX_UNDO_STEPS = 30;
let pendingAction = null;
let deferredPrompt = null;
let swRegistration = null;
let refreshing = false;

/* ============================== УТИЛИТЫ ============================== */
function checkLocalStorage() {
  try { const t = '__storage_test__'; localStorage.setItem(t, t); localStorage.removeItem(t); S.isLocalStorageAvailable = true; return true; }
  catch (e) { S.isLocalStorageAvailable = false; return false; }
}
function isMobile() { return window.innerWidth < 1024; }
function vibrate(ms = 30) { if (isMobile() && navigator.vibrate) navigator.vibrate(ms); }
function showToast(message, type = 'success') {
  const toast = document.createElement('div');
  toast.className = `toast toast-${type}`;
  toast.textContent = message;
  document.body.appendChild(toast);
  setTimeout(() => { toast.style.opacity = '0'; setTimeout(() => toast.remove(), 300); }, 3000);
}
function esc(s) { const d = document.createElement('div'); d.textContent = s == null ? '' : String(s); return d.innerHTML; }
function escAttr(s) { return String(s == null ? '' : s).replace(/'/g, "\\'").replace(/"/g, '&quot;'); }
function hasAnyContent(v) { if (v === null || v === undefined) return false; return String(v).trim().length > 0; }
function parseNum(v) {
  if (v === null || v === undefined || v === '') return null;
  if (typeof v === 'number') return isFinite(v) ? v : null;
  const s = String(v).replace(',', '.').replace('%', '').replace(/\s/g, '');
  if (s === '—' || s === '-') return null;
  const n = parseFloat(s);
  return isNaN(n) ? null : n;
}
function formatVal(v) { const n = parseNum(v); return n === null ? '—' : Math.round(n) + '%'; }
function clampPct(n) { return Math.max(0, Math.min(100, Math.round(n))); }

const INDEXED_COLORS = ['#000000','#FFFFFF','#FF0000','#00FF00','#0000FF','#FFFF00','#FF00FF','#00FFFF','#000000','#FFFFFF','#FF0000','#00FF00','#0000FF','#FFFF00','#FF00FF','#00FFFF','#800000','#008000','#000080','#808000','#800080','#008080','#C0C0C0','#808080','#9999FF','#993366','#FFFFCC','#CCFFFF','#660066','#FF8080','#0066CC','#CCCCFF','#000080','#FF00FF','#FFFF00','#00FFFF','#800080','#800000','#008080','#0000FF','#00CCFF','#CCFFFF','#CCFFCC','#FFFF99','#99CCFF','#FF99CC','#CC99FF','#FFCC99','#3366FF','#33CCCC','#99CC00','#FFCC00','#FF9900','#FF6600','#666699','#969696','#003366','#339966','#003300','#333300','#993300','#993366','#333399','#333333'];
function applyTint(hex, tint) {
  if (!hex || tint == null || tint === 0) return hex;
  hex = hex.replace('#', ''); if (hex.length === 8) hex = hex.slice(2);
  let r = parseInt(hex.substring(0, 2), 16), g = parseInt(hex.substring(2, 4), 16), b = parseInt(hex.substring(4, 6), 16);
  if (tint < 0) { r = Math.round(r * (1 + tint)); g = Math.round(g * (1 + tint)); b = Math.round(b * (1 + tint)); }
  else { r = Math.round(r * (1 - tint) + 255 * tint); g = Math.round(g * (1 - tint) + 255 * tint); b = Math.round(b * (1 - tint) + 255 * tint); }
  return '#' + [r, g, b].map(x => Math.max(0, Math.min(255, x)).toString(16).padStart(2, '0')).join('');
}
function isLightColor(hex) {
  if (!hex) return true; hex = hex.replace('#', '');
  const r = parseInt(hex.substring(0, 2), 16), g = parseInt(hex.substring(2, 4), 16), b = parseInt(hex.substring(4, 6), 16);
  return (0.299 * r + 0.587 * g + 0.114 * b) / 255 > 0.55;
}
function getTotalPendingChanges() {
  let total = 0;
  for (const k in S.pendingChanges) if (S.pendingChanges[k] && S.pendingChanges[k].size) total += S.pendingChanges[k].size;
  return total;
}
function getSheetPendingChanges(name) { if (!S.pendingChanges[name]) S.pendingChanges[name] = new Map(); return S.pendingChanges[name]; }
function getSheetOriginals(name) { if (!S.originalCells[name]) S.originalCells[name] = new Map(); return S.originalCells[name]; }
function captureOriginal(sheetName, cellRef) {
  const originals = getSheetOriginals(sheetName);
  if (!originals.has(cellRef)) {
    const ws = S.wb && S.wb.Sheets[sheetName];
    const cell = ws ? ws[cellRef] : undefined;
    originals.set(cellRef, cell ? Object.assign({}, cell) : null);
  }
}
function restoreCellSnapshot(ws, cellRef, snapshot) {
  if (!ws) return;
  if (snapshot === null || snapshot === undefined) delete ws[cellRef];
  else ws[cellRef] = Object.assign({}, snapshot);
}

/* ============================== МОДАЛКИ ============================== */
function showAppModal(title, message, buttons) {
  document.getElementById('appModalTitle').textContent = title;
  document.getElementById('appModalMessage').innerHTML = message;
  const box = document.getElementById('appModalButtons');
  box.innerHTML = '';
  buttons.forEach(b => {
    const btn = document.createElement('button');
    btn.className = 'w-full px-4 py-3 font-semibold rounded-xl ' + (b.cls || 'bg-surface-200 dark:bg-surface-700 text-surface-900 dark:text-white');
    btn.textContent = b.label;
    btn.onclick = () => { closeAppModal(); if (b.fn) b.fn(); };
    box.appendChild(btn);
  });
  document.getElementById('appModal').classList.add('open');
}
function closeAppModal() { document.getElementById('appModal').classList.remove('open'); }
function showWarningModal(action) { pendingAction = action; document.getElementById('warningModal').classList.add('open'); }
function closeWarningModal() { document.getElementById('warningModal').classList.remove('open'); pendingAction = null; }
async function saveAndExit() {
  const action = pendingAction;
  closeWarningModal();
  const ok = await exportToFile();
  if (ok && action) action();
}
function exitWithoutSaving() {
  const action = pendingAction;
  revertCurrentSheetChanges();
  closeWarningModal();
  if (action) action();
}
function revertCurrentSheetChanges() {
  const sheetName = S.currentSheet;
  const changes = getSheetPendingChanges(sheetName);
  const originals = getSheetOriginals(sheetName);
  const ws = S.wb && S.wb.Sheets[sheetName];
  changes.forEach((val, cellRef) => {
    const snap = originals.has(cellRef) ? originals.get(cellRef) : null;
    restoreCellSnapshot(ws, cellRef, snap);
    const m = cellRef.match(/^([A-Z]+)(\d+)$/);
    if (m && ws) {
      const colIdx = colNameToIndex(m[1]) - 1;
      const rowNum = parseInt(m[2], 10) - 1;
      const workIdx = S.works.findIndex(w => w.myRow === rowNum);
      if (workIdx >= 0) S.myCells[workIdx + '_' + colIdx] = snap && snap.v != null ? snap.v : null;
    }
  });
  changes.clear(); originals.clear(); S.cachedSheets.delete(sheetName);
  S.undoStack = S.undoStack.filter(a => (a.sheetName || (a.changes && a.changes[0] && a.changes[0].sheetName)) !== sheetName);
  updateSelectionPanel(); renderMatrix(); updateSummary(); rebuildSheetSelect(); updateUndoButton();
}

/* ============================== ЧТЕНИЕ EXCEL / ЦВЕТА ============================== */
function colNameToIndex(colName) { let i = 0; for (let k = 0; k < colName.length; k++) i = i * 26 + (colName.charCodeAt(k) - 64); return i; }

async function getSheetFilePath(zip, sheetName) {
  try {
    const wbFile = zip.file('xl/workbook.xml');
    if (!wbFile) return 'xl/worksheets/sheet1.xml';
    const parser = new DOMParser();
    const wbXml = parser.parseFromString(await wbFile.async('text'), 'text/xml');
    const sheets = wbXml.getElementsByTagName('sheet');
    let rid = null;
    for (let i = 0; i < sheets.length; i++) {
      if (sheets[i].getAttribute('name') === sheetName) { rid = sheets[i].getAttribute('r:id') || sheets[i].getAttribute('id'); break; }
    }
    if (!rid) return 'xl/worksheets/sheet1.xml';
    const relsFile = zip.file('xl/_rels/workbook.xml.rels');
    if (!relsFile) return 'xl/worksheets/sheet1.xml';
    const relsXml = parser.parseFromString(await relsFile.async('text'), 'text/xml');
    const rels = relsXml.getElementsByTagName('Relationship');
    for (let i = 0; i < rels.length; i++) {
      if (rels[i].getAttribute('Id') === rid) {
        let target = rels[i].getAttribute('Target') || '';
        if (target.charAt(0) === '/') return target.slice(1);
        if (target.indexOf('xl/') === 0) return target;
        return 'xl/' + target.replace(/^\.\//, '');
      }
    }
  } catch (e) { console.warn('resolve sheet path', e); }
  return 'xl/worksheets/sheet1.xml';
}

async function extractExcelColorsForSheet(arrayBuffer, sheetName) {
  const colorMap = {};
  try {
    const zip = await JSZip.loadAsync(arrayBuffer);
    const parser = new DOMParser();
    const sheetPath = await getSheetFilePath(zip, sheetName);
    const sheetFile = zip.file(sheetPath);
    if (!sheetFile) return colorMap;
    const themeColors = ['#FFFFFF','#000000','#E7E6E6','#44546A','#4472C4','#ED7D31','#A5A5A5','#FFC000','#5B9BD5','#70AD47'];
    const themeFile = zip.file('xl/theme/theme1.xml');
    if (themeFile) {
      const themeXml = parser.parseFromString(await themeFile.async('text'), 'text/xml');
      const clrScheme = themeXml.getElementsByTagName('a:clrScheme')[0];
      if (clrScheme) {
        ['a:lt1','a:dk1','a:lt2','a:dk2','a:accent1','a:accent2','a:accent3','a:accent4','a:accent5','a:accent6'].forEach((tag, idx) => {
          const el = clrScheme.getElementsByTagName(tag)[0];
          if (el) {
            const srgb = el.getElementsByTagName('a:srgbClr')[0], sys = el.getElementsByTagName('a:sysClr')[0];
            if (srgb && srgb.getAttribute('val')) themeColors[idx] = '#' + srgb.getAttribute('val');
            else if (sys && sys.getAttribute('lastClr')) themeColors[idx] = '#' + sys.getAttribute('lastClr');
          }
        });
      }
    }
    const styleFillHex = [];
    const stylesFile = zip.file('xl/styles.xml');
    if (stylesFile) {
      const stylesXml = parser.parseFromString(await stylesFile.async('text'), 'text/xml');
      const fills = Array.from(stylesXml.querySelectorAll('fills > fill')).map(fill => {
        const pf = fill.querySelector('patternFill'); if (!pf) return null;
        const pType = pf.getAttribute('patternType'); if (!pType || pType === 'none') return null;
        const fg = pf.querySelector('fgColor'); if (!fg) return null;
        let hex = null;
        if (fg.getAttribute('rgb')) { const rgb = fg.getAttribute('rgb'); hex = '#' + (rgb.length === 8 ? rgb.slice(2) : rgb); }
        else if (fg.getAttribute('theme') !== null) { const t = parseInt(fg.getAttribute('theme'), 10); if (themeColors[t]) hex = themeColors[t]; }
        else if (fg.getAttribute('indexed') !== null) { const i = parseInt(fg.getAttribute('indexed'), 10); if (INDEXED_COLORS[i]) hex = INDEXED_COLORS[i]; }
        if (hex && fg.getAttribute('tint') !== null) hex = applyTint(hex, parseFloat(fg.getAttribute('tint')));
        return hex;
      });
      Array.from(stylesXml.querySelectorAll('cellXfs > xf')).forEach((xf, idx) => {
        const fillId = parseInt(xf.getAttribute('fillId') || '0', 10);
        styleFillHex[idx] = fills[fillId] || null;
      });
    }
    const sheetXml = parser.parseFromString(await sheetFile.async('text'), 'text/xml');
    const cells = sheetXml.getElementsByTagName('c');
    for (let i = 0; i < cells.length; i++) {
      const c = cells[i], sAttr = c.getAttribute('s'), rAttr = c.getAttribute('r');
      if (sAttr !== null && rAttr) {
        const hex = styleFillHex[parseInt(sAttr, 10)];
        if (hex && hex.toUpperCase() !== '#FFFFFF' && hex.toUpperCase() !== '#000000') colorMap[rAttr] = hex;
      }
    }
  } catch (e) { console.warn('XML styles parse error for ' + sheetName, e); }
  return colorMap;
}

async function patchExcelFile(arrayBuffer, changes, activeSheetName) {
  const zip = await JSZip.loadAsync(arrayBuffer);
  const parser = new DOMParser(), serializer = new XMLSerializer();
  const sheetPath = await getSheetFilePath(zip, activeSheetName);
  const sheetFile = zip.file(sheetPath);
  if (!sheetFile) throw new Error('Sheet XML not found: ' + sheetPath);
  const sheetXml = parser.parseFromString(await sheetFile.async('text'), 'text/xml');
  const sheetData = sheetXml.getElementsByTagName('sheetData')[0];
  if (!sheetData) throw new Error('sheetData not found');
  const ns = sheetXml.documentElement.namespaceURI || 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
  const createEl = tag => sheetXml.createElementNS(ns, tag);
  const rowMap = new Map();
  const rows = sheetData.getElementsByTagName('row');
  for (let i = 0; i < rows.length; i++) rowMap.set(parseInt(rows[i].getAttribute('r'), 10), rows[i]);
  for (const [cellRef, value] of changes) {
    const match = cellRef.match(/^([A-Z]+)(\d+)$/); if (!match) continue;
    const newColIdx = colNameToIndex(match[1]);
    const rowNumber = parseInt(match[2], 10);
    let row = rowMap.get(rowNumber);
    if (!row) {
      row = createEl('row'); row.setAttribute('r', String(rowNumber));
      let inserted = false;
      for (let i = 0; i < sheetData.children.length; i++) {
        const exR = parseInt(sheetData.children[i].getAttribute('r'), 10);
        if (rowNumber < exR) { sheetData.insertBefore(row, sheetData.children[i]); inserted = true; break; }
      }
      if (!inserted) sheetData.appendChild(row);
      rowMap.set(rowNumber, row);
    }
    const cells = row.getElementsByTagName('c');
    let cell = null;
    for (let i = 0; i < cells.length; i++) if (cells[i].getAttribute('r') === cellRef) { cell = cells[i]; break; }
    if (!cell) {
      cell = createEl('c'); cell.setAttribute('r', cellRef);
      let inserted = false;
      for (let i = 0; i < cells.length; i++) {
        const m = (cells[i].getAttribute('r') || '').match(/^([A-Z]+)/);
        if (m && newColIdx < colNameToIndex(m[1])) { row.insertBefore(cell, cells[i]); inserted = true; break; }
      }
      if (!inserted) row.appendChild(cell);
    }
    cell.removeAttribute('t');
    const oldF = cell.getElementsByTagNameNS(ns, 'f')[0]; if (oldF) oldF.remove();
    const oldIs = cell.getElementsByTagNameNS(ns, 'is')[0]; if (oldIs) oldIs.remove();
    let vElem = cell.getElementsByTagNameNS(ns, 'v')[0];
    if (!vElem) { vElem = createEl('v'); cell.appendChild(vElem); }
    vElem.textContent = String(value);
  }
  zip.file(sheetPath, serializer.serializeToString(sheetXml));
  // Удаляем устаревшую цепочку вычислений, чтобы Excel не ругался на изменённые формулы
  if (zip.file('xl/calcChain.xml')) {
    zip.remove('xl/calcChain.xml');
    const ctFile = zip.file('[Content_Types].xml');
    if (ctFile) {
      const ctXml = parser.parseFromString(await ctFile.async('text'), 'text/xml');
      Array.from(ctXml.getElementsByTagName('Override')).forEach(ov => {
        if (ov.getAttribute('PartName') === '/xl/calcChain.xml') ov.parentNode.removeChild(ov);
      });
      zip.file('[Content_Types].xml', serializer.serializeToString(ctXml));
    }
    const relsFile = zip.file('xl/_rels/workbook.xml.rels');
    if (relsFile) {
      const relsXml = parser.parseFromString(await relsFile.async('text'), 'text/xml');
      Array.from(relsXml.getElementsByTagName('Relationship')).forEach(rel => {
        if ((rel.getAttribute('Target') || '') === 'calcChain.xml') rel.parentNode.removeChild(rel);
      });
      zip.file('xl/_rels/workbook.xml.rels', serializer.serializeToString(relsXml));
    }
  }
  return await zip.generateAsync({ type: 'arraybuffer', compression: 'DEFLATE' });
}

function readExcelFile(file, handle) {
  if (typeof XLSX === 'undefined' || typeof JSZip === 'undefined') {
    showError('Библиотеки не загружены. Выполните node tools/fetch-libs.mjs и обновите страницу.');
    return;
  }
  if (file.size > 50 * 1024 * 1024) { showError('Файл слишком большой (максимум 50 МБ)'); return; }
  const reader = new FileReader();
  reader.onload = async function (e) {
    try {
      const buffer = e.target.result;
      S.originalBuffer = buffer;
      S.wb = XLSX.read(new Uint8Array(buffer), { type: 'array', cellDates: true });
      if (!S.wb.SheetNames || !S.wb.SheetNames.length) { showError('Файл не содержит листов'); return; }
      S.fileName = file.name; S.saveHandle = handle || null; S.demoMode = false;
      S.pendingChanges = {}; S.originalCells = {}; S.undoStack = [];
      S.currentSheet = S.wb.SheetNames[0];
      S.sheetColors = {};
      for (const name of S.wb.SheetNames) S.sheetColors[name] = await extractExcelColorsForSheet(buffer, name);
      S.cellColors = S.sheetColors[S.currentSheet] || {};
      rebuildSheetSelect(); parseSheet(); checkDraft();
      idbPut('lastFile', { name: S.fileName, buffer: buffer, timestamp: Date.now() }).catch(() => {});
    } catch (err) {
      console.error('File read error:', err);
      showToast('Ошибка чтения: ' + err.message, 'error');
      showError('Ошибка чтения файла: ' + err.message);
    }
  };
  reader.onerror = () => { showToast('Не удалось прочитать файл', 'error'); showError('Не удалось прочитать файл'); };
  reader.readAsArrayBuffer(file);
}

/* ============================== ЛИСТ / ПАРСИНГ ============================== */
function rebuildSheetSelect() {
  const select = document.getElementById('filterSheet'); if (!select || !S.wb) return;
  select.innerHTML = S.wb.SheetNames.map(name => {
    const changes = S.pendingChanges[name] ? S.pendingChanges[name].size : 0;
    const label = changes > 0 ? `${name} (${changes})` : name;
    return `<option value="${esc(name)}" ${name === S.currentSheet ? 'selected' : ''}>${esc(label)}</option>`;
  }).join('');
  select.onchange = e => { if (e.target.value !== S.currentSheet) switchSheet(e.target.value); };
}
function switchSheet(name) { doSwitchSheet(name); }
function doSwitchSheet(newSheetName) {
  S.currentSheet = newSheetName;
  S.cellColors = S.sheetColors[newSheetName] || {};
  S.filterSection = ''; S.filterColIndex = -1; S.filterWorkIdxs.clear();
  S.selSection = ''; S.selColIndex = -1; S.selWorkIdx = 0;
  parseSheet(); applyPendingToCurrentSheet(); renderMatrix(); rebuildSheetSelect();
}
function applyPendingToCurrentSheet() {
  const changes = getSheetPendingChanges(S.currentSheet);
  if (!changes.size) return;
  const ws = S.wb.Sheets[S.currentSheet];
  S.works.forEach((work, workIdx) => {
    S.columns.forEach(col => {
      const cellRef = XLSX.utils.encode_cell({ r: work.myRow, c: col.colIndex });
      if (changes.has(cellRef)) {
        const val = changes.get(cellRef);
        S.myCells[workIdx + '_' + col.colIndex] = val;
        if (ws[cellRef]) { ws[cellRef].v = val; ws[cellRef].t = 'n'; delete ws[cellRef].w; }
        else ws[cellRef] = { v: val, t: 'n' };
      }
    });
  });
}
function parseSheet() {
  const ws = S.wb.Sheets[S.currentSheet];
  if (!ws || !ws['!ref']) { showError('Лист пуст или структура не распознана'); return; }
  const range = XLSX.utils.decode_range(ws['!ref']);
  S.columns = []; S.uniqueSections = [];
  let lastSection = '';
  for (let c = 3; c <= range.e.c; c++) {
    const secCell = ws[XLSX.utils.encode_cell({ r: 0, c })];
    const secVal = secCell && secCell.v != null ? String(secCell.v).trim() : '';
    if (secVal) lastSection = secVal;
    const flCell = ws[XLSX.utils.encode_cell({ r: 2, c })];
    const flVal = flCell && flCell.v != null ? String(flCell.v).trim() : '';
    if (!flVal || flVal.toLowerCase() === 'статус') continue;
    if (!lastSection) continue;
    if (!S.uniqueSections.includes(lastSection)) S.uniqueSections.push(lastSection);
    S.columns.push({ colIndex: c, section: lastSection, floor: flVal });
  }
  if (!S.columns.length) { showError('Не удалось распознать секции и этажи'); return; }
  S.works = []; S.myCells = {}; S.foremanCells = {};
  for (let r = 3; r <= range.e.r; r++) {
    const cellA = ws[XLSX.utils.encode_cell({ r: r, c: 0 })];
    const valA = cellA && cellA.v != null ? String(cellA.v).trim().toUpperCase() : '';
    if (valA.includes('ПРОРАБ') && r > 3) {
      const workRef = XLSX.utils.encode_cell({ r: r - 1, c: 2 });
      const workNameCell = ws[workRef];
      const workName = workNameCell && workNameCell.v != null ? String(workNameCell.v).trim() : ('Работа #' + (S.works.length + 1));
      const workIdx = S.works.length;
      S.works.push({ name: workName, myRow: r - 1, foremanRow: r, rowColor: S.cellColors[workRef] || null });
      for (let i = 0; i < S.columns.length; i++) {
        const c = S.columns[i].colIndex, key = workIdx + '_' + c;
        const myCell = ws[XLSX.utils.encode_cell({ r: r - 1, c })];
        S.myCells[key] = myCell && myCell.v != null ? myCell.v : null;
        const frCell = ws[XLSX.utils.encode_cell({ r: r, c })];
        S.foremanCells[key] = frCell && frCell.v != null ? frCell.v : null;
      }
    }
  }
  if (!S.works.length) { showError('Не найдены строки с отметкой "ПРОРАБ"'); return; }
  S.filterSection = S.uniqueSections[0];
  const firstSecCols = S.columns.filter(col => col.section === S.filterSection);
  S.selSection = S.filterSection;
  S.selColIndex = firstSecCols.length ? firstSecCols[0].colIndex : -1;
  S.selWorkIdx = 0;
  showApp(); initFilters(); updateSelectionPanel(); renderMatrix(); rebuildMultiselects(); updateSummary();
}
function showApp() {
  document.getElementById('uploadScreen').classList.add('hidden');
  document.getElementById('appScreen').classList.remove('hidden');
  document.getElementById('fileName').textContent = S.fileName;
  document.getElementById('tablePlaceholder').classList.add('hidden');
  document.getElementById('tableContainer').classList.remove('hidden');
  document.getElementById('summaryBar').classList.remove('hidden');
}
function toggleFilters() {
  const content = document.getElementById('filtersContent'), arrow = document.getElementById('filtersArrow');
  content.classList.toggle('open');
  arrow.style.transform = content.classList.contains('open') ? 'rotate(180deg)' : '';
}
function initFilters() {
  const fSec = document.getElementById('filterSection');
  fSec.innerHTML = S.uniqueSections.map(s => `<option value="${esc(s)}">${esc(s)}</option>`).join('');
  fSec.value = S.filterSection;
  rebuildFloorSelects(); rebuildWorksDropdown();
  fSec.onchange = e => {
    S.filterSection = e.target.value; S.filterColIndex = -1; S.selSection = S.filterSection;
    const secCols = S.columns.filter(c => c.section === S.filterSection);
    if (secCols.length) S.selColIndex = secCols[0].colIndex;
    rebuildFloorSelects(); rebuildWorksDropdown(); updateSelectionPanel(); renderMatrix(); rebuildMultiselects(); updateSummary();
  };
  document.getElementById('filterFloor').onchange = e => {
    S.filterColIndex = parseInt(e.target.value, 10);
    if (S.filterColIndex !== -1) { S.selColIndex = S.filterColIndex; updateSelectionPanel(); }
    renderMatrix();
  };
  document.getElementById('hideAccepted').onchange = e => {
    S.hideAccepted = e.target.checked; renderMatrix(); rebuildMultiselects(); rebuildWorksDropdown(); updateSummary();
  };
}
function rebuildFloorSelects() {
  const fFl = document.getElementById('filterFloor');
  const secCols = S.columns.filter(col => col.section === S.filterSection);
  fFl.innerHTML = `<option value="-1">Все этажи (${secCols.length})</option>` + secCols.map(col => `<option value="${col.colIndex}">${esc(col.floor)}</option>`).join('');
  fFl.value = String(S.filterColIndex);
}
function getVisibleWorks() {
  let visibleWorks = S.works.map((w, idx) => Object.assign({}, w, { idx }));
  const secCols = S.columns.filter(col => col.section === S.filterSection);
  if (S.hideAccepted) visibleWorks = visibleWorks.filter(w => !shouldHideRow(w, secCols));
  if (S.summaryFilter === 'done') visibleWorks = visibleWorks.filter(w => isWorkDone(w, secCols));
  else if (S.summaryFilter === 'mismatch') visibleWorks = visibleWorks.filter(w => countMismatches(w, secCols) > 0);
  return visibleWorks;
}
function rebuildWorksDropdown() {
  const container = document.getElementById('filterWorkOptions');
  const visibleWorks = getVisibleWorks();
  container.innerHTML = visibleWorks.length ? visibleWorks.map(w => {
    const checked = S.filterWorkIdxs.has(w.idx);
    return `<label class="multiselect-option" data-idx="${w.idx}" data-name="${esc(w.name).toLowerCase()}">
      <input type="checkbox" value="${w.idx}" ${checked ? 'checked' : ''} onchange="handleFilterWorkChange(this)">
      <span class="multiselect-option-label">${esc(w.name)}</span></label>`;
  }).join('') : '<div class="px-3 py-2 text-xs text-surface-400 italic">Нет работ</div>';
  updateWorksDropdownLabel();
}
function handleFilterWorkChange(cb) {
  const idx = parseInt(cb.value, 10);
  if (cb.checked) S.filterWorkIdxs.add(idx); else S.filterWorkIdxs.delete(idx);
  renderMatrix(); updateSummary(); updateWorksDropdownLabel();
}
function filterWorkAll(selectAll, e) {
  e.stopPropagation();
  const visibleWorks = getVisibleWorks();
  document.querySelectorAll('#filterWorkOptions input[type="checkbox"]').forEach(cb => {
    const idx = parseInt(cb.value, 10);
    const isVisible = visibleWorks.some(w => w.idx === idx);
    if (selectAll && isVisible) { cb.checked = true; S.filterWorkIdxs.add(idx); }
    else if (!selectAll) { cb.checked = false; S.filterWorkIdxs.delete(idx); }
  });
  renderMatrix(); updateSummary(); updateWorksDropdownLabel();
}
function updateWorksDropdownLabel() {
  const label = document.querySelector('#filterWorkMs .multiselect-label');
  const count = S.filterWorkIdxs.size;
  if (count === 0) label.textContent = 'Все работы';
  else if (count === 1) label.textContent = S.works[Array.from(S.filterWorkIdxs)[0]].name;
  else label.textContent = `Выбрано: ${count}`;
}
function filterWorksDropdown(query) {
  const q = query.trim().toLowerCase();
  let visibleCount = 0;
  document.querySelectorAll('#filterWorkOptions .multiselect-option').forEach(opt => {
    const name = opt.getAttribute('data-name') || '';
    if (!q || name.includes(q)) { opt.classList.remove('hidden'); visibleCount++; }
    else opt.classList.add('hidden');
  });
  document.getElementById('filterWorkEmpty').classList.toggle('hidden', visibleCount > 0);
}
function shouldHideRow(work, visibleCols) {
  if (!S.hideAccepted || !visibleCols.length) return false;
  let allForemanEmpty = true, allMyEmpty = true, allMatch = true, hasAnyForemanData = false;
  for (const col of visibleCols) {
    const key = work.idx + '_' + col.colIndex;
    const myVal = S.myCells[key], frVal = S.foremanCells[key];
    const frHas = hasAnyContent(frVal), myHas = hasAnyContent(myVal);
    if (frHas) {
      allForemanEmpty = false;
      const myNum = parseNum(myVal), frNum = parseNum(frVal);
      if (frNum !== null) { hasAnyForemanData = true; if (myNum === null || myNum !== frNum) allMatch = false; }
    }
    if (myHas) allMyEmpty = false;
  }
  if (allForemanEmpty && allMyEmpty) return true;
  if (allMatch && hasAnyForemanData) return true;
  return false;
}
function countMismatches(work, visibleCols) {
  let count = 0;
  for (const col of visibleCols) {
    const key = work.idx + '_' + col.colIndex;
    const myNum = parseNum(S.myCells[key]), frNum = parseNum(S.foremanCells[key]);
    if (frNum !== null) {
      if (myNum === null) count++;
      else { const diff = Math.abs(myNum - frNum); if ((frNum === 100 && myNum < 100) || diff > 15) count++; }
    }
  }
  return count;
}

/* ============================== МАТРИЦА ============================== */
function renderMatrix() {
  const secName = S.filterSection;
  let visibleCols = S.columns.filter(col => col.section === secName);
  if (S.filterColIndex !== -1) visibleCols = visibleCols.filter(col => col.colIndex === S.filterColIndex);
  let visibleWorks = S.works.map((w, idx) => Object.assign({}, w, { idx }));
  if (S.filterWorkIdxs.size) visibleWorks = visibleWorks.filter(w => S.filterWorkIdxs.has(w.idx));
  if (S.summaryFilter === 'done') {
    const secCols = S.columns.filter(col => col.section === secName);
    visibleWorks = visibleWorks.filter(w => isWorkDone(w, secCols));
  } else if (S.summaryFilter === 'mismatch') visibleWorks = visibleWorks.filter(w => countMismatches(w, visibleCols) > 0);

  const thead = document.getElementById('matrixHead'), tbody = document.getElementById('matrixBody');
  let theadHTML = '<tr class="bg-surface-100 dark:bg-surface-700">';
  theadHTML += '<th class="sticky-col px-2 md:px-3 py-2.5 text-xs font-semibold text-surface-600 dark:text-surface-300 border-b border-r border-surface-200 dark:border-surface-600 bg-surface-100 dark:bg-surface-700 text-left min-w-[160px] md:min-w-[240px] max-w-[200px] md:max-w-[320px]">Вид работы</th>';
  for (let i = 0; i < visibleCols.length; i++) {
    theadHTML += `<th class="px-1 md:px-2 py-2.5 text-xs font-bold text-surface-700 dark:text-surface-200 border-b border-r border-surface-200 dark:border-surface-600 bg-surface-100 dark:bg-surface-700 whitespace-nowrap min-w-[60px] md:min-w-[68px] text-center">${esc(visibleCols[i].floor)}</th>`;
  }
  theadHTML += '</tr>';
  thead.innerHTML = theadHTML;

  let tbodyHTML = '', shownRowsCount = 0;
  for (let wi = 0; wi < visibleWorks.length; wi++) {
    const work = visibleWorks[wi];
    if (shouldHideRow(work, visibleCols)) continue;
    const mismatchCount = countMismatches(work, visibleCols);
    const isRowSelected = visibleCols.some(col => S.selColIndex === col.colIndex && S.selWorkIdx === work.idx);
    const rowHighlightClass = isRowSelected ? 'bg-brand-100/50 dark:bg-brand-900/30' : '';
    let rowCellsHTML = '';
    for (let ci = 0; ci < visibleCols.length; ci++) {
      const col = visibleCols[ci], key = work.idx + '_' + col.colIndex;
      const myVal = S.myCells[key], frVal = S.foremanCells[key];
      const myNum = parseNum(myVal), frValNum = parseNum(frVal);
      const frHasContent = hasAnyContent(frVal);
      const myRef = XLSX.utils.encode_cell({ r: work.myRow, c: col.colIndex });
      const frRef = XLSX.utils.encode_cell({ r: work.foremanRow, c: col.colIndex });
      const excelHex = S.cellColors[myRef] || S.cellColors[frRef] || null;
      const isSelected = S.selColIndex === col.colIndex && S.selWorkIdx === work.idx;
      if (!frHasContent) {
        let lockedBg = 'bg-surface-100/40 dark:bg-surface-800/20 cell-locked';
        if (isSelected) lockedBg += ' ring-2 ring-brand-600 ring-inset';
        rowCellsHTML += `<td class="px-1 py-1 border-r border-surface-200/70 dark:border-surface-700/50 ${lockedBg} cursor-not-allowed transition-all" onclick="selectLockedCell('${escAttr(col.section)}', ${col.colIndex}, ${work.idx})">`;
        rowCellsHTML += `<div class="cell-content opacity-50"><span class="text-surface-300 dark:text-surface-600 text-xs">—</span></div></td>`;
      } else {
        const myStr = myVal !== null && myVal !== '' ? formatVal(myVal) : '—';
        const frStr = formatVal(frVal);
        let inlineStyle = '', cellBg = 'bg-white dark:bg-surface-800/40 hover:bg-surface-50 dark:hover:bg-surface-700/50';
        let myColor = 'text-surface-800 dark:text-surface-200';
        let frBadge = 'bg-surface-200/80 dark:bg-surface-700 text-surface-700 dark:text-surface-300';
        if (excelHex) {
          inlineStyle = `style="background-color: ${excelHex};"`;
          cellBg = 'hover:brightness-95 cell-with-color';
          const light = isLightColor(excelHex);
          myColor = light ? 'text-slate-900' : 'text-white';
          frBadge = light ? 'bg-black/10 text-slate-900' : 'bg-white/20 text-white';
        } else if (myNum === null) myColor = 'text-surface-400 dark:text-surface-500';
        if (isSelected) cellBg += ' ring-2 ring-brand-600 ring-inset';
        rowCellsHTML += `<td ${inlineStyle} class="px-1 py-1 border-r border-surface-200/70 dark:border-surface-700/50 ${cellBg} cursor-pointer transition-all" onclick="selectCell('${escAttr(col.section)}', ${col.colIndex}, ${work.idx})">`;
        rowCellsHTML += `<div class="cell-content"><span class="cell-my ${myColor}">${myStr}</span><span class="cell-foreman ${frBadge}">Пр: ${frStr}</span></div></td>`;
      }
    }
    shownRowsCount++;
    const rowNameStyle = work.rowColor ? `style="background-color: ${work.rowColor}; color: ${isLightColor(work.rowColor) ? '#0f172a' : '#ffffff'};"` : '';
    const rowNameClass = work.rowColor ? '' : 'text-surface-800 dark:text-surface-200 bg-white dark:bg-surface-800';
    const mismatchIcon = mismatchCount > 0 ? `<span class="inline-block ml-1 text-amber-500" title="${mismatchCount} расхождений">⚠ ${mismatchCount}</span>` : '';
    const workCellHighlight = isRowSelected && !work.rowColor ? 'bg-brand-100/50 dark:bg-brand-900/30' : '';
    tbodyHTML += `<tr class="border-b border-surface-200/70 dark:border-surface-700/60 table-row-hover ${rowHighlightClass}">`;
    tbodyHTML += `<td ${rowNameStyle} class="sticky-col px-2 md:px-3 py-2 text-xs md:text-sm font-medium border-r border-surface-200 dark:border-surface-700 ${rowNameClass} ${workCellHighlight} leading-snug">${esc(work.name)}${mismatchIcon}</td>`;
    tbodyHTML += rowCellsHTML + '</tr>';
  }
  if (!shownRowsCount) tbodyHTML = `<tr><td colspan="${visibleCols.length + 1}" class="p-8 text-center text-sm text-surface-400">Нет работ, подходящих под фильтры</td></tr>`;
  tbody.innerHTML = tbodyHTML;
}
function isWorkDone(work, visibleCols) {
  let allMatch = true, hasAnyForemanData = false;
  for (const col of visibleCols) {
    const key = work.idx + '_' + col.colIndex;
    const frVal = S.foremanCells[key];
    if (hasAnyContent(frVal)) {
      const myNum = parseNum(S.myCells[key]), frNum = parseNum(frVal);
      if (frNum !== null) { hasAnyForemanData = true; if (myNum !== frNum) { allMatch = false; break; } }
    }
  }
  return allMatch && hasAnyForemanData;
}

/* ============================== ВЫБОР / ПАНЕЛЬ ============================== */
function selectCell(secName, colIndex, workIdx) {
  S.selSection = secName; S.selColIndex = colIndex; S.selWorkIdx = workIdx;
  updateSelectionPanel(); renderMatrix();
  if (isMobile()) openBottomSheet();
}
function selectLockedCell(secName, colIndex, workIdx) {
  S.selSection = secName; S.selColIndex = colIndex; S.selWorkIdx = workIdx;
  updateSelectionPanel(); renderMatrix();
  showToast('Нет отметки прораба — ввод невозможен', 'warning');
}
function updateSelectionPanel() {
  const colObj = S.columns.find(c => c.colIndex === S.selColIndex);
  const workObj = S.works[S.selWorkIdx];
  const set = (id, val) => { const el = document.getElementById(id); if (el) el.textContent = val; };
  set('selSection', S.selSection || '—');
  set('selFloor', colObj ? colObj.floor : '—');
  set('selWork', workObj ? workObj.name : '—');
  const myValueDisplay = document.getElementById('myValueDisplay');
  const foremanValueDisplay = document.getElementById('foremanValueDisplay');
  const lockedWarning = document.getElementById('lockedWarning');
  const acceptBtn = document.getElementById('acceptBtn');
  const submitBtn = document.getElementById('submitBtn');
  const manualInput = document.getElementById('manualValue');
  if (colObj && workObj) {
    const key = S.selWorkIdx + '_' + S.selColIndex;
    const myVal = S.myCells[key], frVal = S.foremanCells[key];
    const frHasContent = hasAnyContent(frVal);
    if (myValueDisplay) myValueDisplay.textContent = myVal !== null && myVal !== '' ? formatVal(myVal) : '—';
    if (foremanValueDisplay) foremanValueDisplay.textContent = frHasContent ? formatVal(frVal) : '—';
    S.isCellLocked = !frHasContent;
    if (lockedWarning) lockedWarning.classList.toggle('hidden', !S.isCellLocked);
    [acceptBtn, submitBtn].forEach(b => { if (b) { b.disabled = S.isCellLocked; b.classList.toggle('opacity-50', S.isCellLocked); b.classList.toggle('cursor-not-allowed', S.isCellLocked); } });
    if (manualInput) { manualInput.disabled = S.isCellLocked; manualInput.classList.toggle('opacity-50', S.isCellLocked); }
  } else {
    if (myValueDisplay) myValueDisplay.textContent = '—';
    if (foremanValueDisplay) foremanValueDisplay.textContent = '—';
    S.isCellLocked = true;
    if (lockedWarning) lockedWarning.classList.remove('hidden');
    [acceptBtn, submitBtn].forEach(b => { if (b) { b.disabled = true; b.classList.add('opacity-50', 'cursor-not-allowed'); } });
    if (manualInput) { manualInput.disabled = true; manualInput.classList.add('opacity-50', 'cursor-not-allowed'); }
  }
  if (isMobile() && document.getElementById('bottomSheet').classList.contains('open')) renderBottomSheetContent();
  updateUndoButton();
}
function updateUndoButton() { document.getElementById('undoBtn').disabled = S.undoStack.length === 0; }
function openBottomSheet() {
  renderBottomSheetContent();
  document.getElementById('bottomSheetBackdrop').classList.add('open');
  document.getElementById('bottomSheet').classList.add('open');
}
function closeBottomSheet() {
  document.getElementById('bottomSheetBackdrop').classList.remove('open');
  document.getElementById('bottomSheet').classList.remove('open');
}
function renderBottomSheetContent() {
  const colObj = S.columns.find(c => c.colIndex === S.selColIndex);
  const workObj = S.works[S.selWorkIdx];
  const key = S.selWorkIdx + '_' + S.selColIndex;
  const myVal = S.myCells[key], frVal = S.foremanCells[key];
  const frHasContent = hasAnyContent(frVal), frNum = parseNum(frVal);
  document.getElementById('bottomSheetContent').innerHTML = `
    <div class="flex items-center justify-between mb-4">
      <h2 class="text-lg font-bold text-surface-900 dark:text-white">Ввод факта</h2>
      <button onclick="closeBottomSheet()" class="p-2 rounded-lg hover:bg-surface-100 dark:hover:bg-surface-700">
        <svg class="w-5 h-5" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" d="M6 18L18 6M6 6l12 12"/></svg>
      </button>
    </div>
    <div class="p-3 bg-surface-50 dark:bg-surface-700/50 rounded-2xl border border-surface-200 dark:border-surface-600 mb-4 card-shadow">
      <div class="text-xs mb-2">
        <div><span class="text-surface-500 dark:text-surface-400">Секция:</span> <span class="font-semibold text-surface-900 dark:text-white">${esc(S.selSection || '—')}</span></div>
        <div><span class="text-surface-500 dark:text-surface-400">Этаж:</span> <span class="font-semibold text-surface-900 dark:text-white">${colObj ? esc(colObj.floor) : '—'}</span></div>
        <div class="mt-1"><span class="text-surface-500 dark:text-surface-400">Работа:</span> <div class="font-medium text-surface-900 dark:text-white leading-snug">${workObj ? esc(workObj.name) : '—'}</div></div>
      </div>
      <div class="grid grid-cols-2 gap-2">
        <div><div class="text-xs text-surface-500 dark:text-surface-400 mb-1">Мой факт</div>
          <div class="px-3 py-2 text-lg font-bold font-mono bg-white dark:bg-surface-800 border border-surface-200 dark:border-surface-600 rounded-lg text-center">${myVal !== null && myVal !== '' ? formatVal(myVal) : '—'}</div></div>
        <div><div class="text-xs text-surface-500 dark:text-surface-400 mb-1">Прораб</div>
          <div class="px-3 py-2 text-lg font-bold font-mono bg-white dark:bg-surface-800 border border-surface-200 dark:border-surface-600 rounded-lg text-center">${frHasContent ? formatVal(frVal) : '—'}</div></div>
      </div>
    </div>
    ${!frHasContent ? `
    <div class="mb-4 p-3 bg-amber-50 dark:bg-amber-900/20 border border-amber-200 dark:border-amber-800 rounded-xl text-sm text-amber-800 dark:text-amber-300 flex items-start gap-2">
      <svg class="w-5 h-5 flex-shrink-0 mt-0.5" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" d="M12 9v3.75m-9.303 3.376c-.866 1.5.216 3.374 1.948 3.374h14.71c1.732 0 2.814-1.874 1.948-3.374L13.949 3.378c-.866-1.5-3.032-1.5-3.898 0L2.697 16.126zM12 15.75h.007v.008H12v-.008z"/></svg>
      <span>Нет отметки прораба — ввод невозможен</span>
    </div>` : `
    <button onclick="acceptForemanValue(); closeBottomSheet();" class="w-full flex items-center justify-center gap-2 px-4 py-4 btn-gradient text-white text-base font-bold rounded-xl shadow-sm mb-3">
      <svg class="w-5 h-5" fill="none" stroke="currentColor" stroke-width="2.5" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" d="M4.5 12.75l6 6 9-13.5"/></svg>
      Факт верный (${frNum !== null ? frNum + '%' : '—'})
    </button>
    <div class="mb-3">
      <label class="block text-sm font-medium text-surface-500 dark:text-surface-400 mb-2">Ввести значение (0–100)</label>
      <div class="flex gap-2">
        <input type="number" id="mobileManualValue" min="0" max="100" step="1" placeholder="0 — 100" inputmode="numeric" class="flex-1 px-4 py-3 text-base border border-surface-300 dark:border-surface-600 rounded-xl bg-surface-50 dark:bg-surface-700 text-surface-900 dark:text-white placeholder-surface-400 focus:ring-2 focus:ring-brand-500 outline-none">
        <button onclick="submitMobileManualValue()" class="px-5 py-3 btn-brand text-white text-base font-bold rounded-xl whitespace-nowrap">OK</button>
      </div>
    </div>`}`;
}

/* ============================== ЗАПИСЬ ЗНАЧЕНИЙ ============================== */
function submitMobileManualValue() {
  const input = document.getElementById('mobileManualValue'); if (!input) return;
  const val = parseFloat(input.value);
  if (isNaN(val) || val < 0 || val > 100) { showToast('Введите число от 0 до 100', 'error'); return; }
  writeValueToCell(clampPct(val)); vibrate(30);
  showToast('Записано: ' + clampPct(val) + '%', 'success');
  renderMatrix(); renderBottomSheetContent();
}
function submitManualValue() {
  if (S.isCellLocked) { showFormStatus('Нет отметки прораба', 'error'); return; }
  if (S.selColIndex === -1 || !S.works[S.selWorkIdx]) { showFormStatus('Сначала выберите ячейку', 'error'); return; }
  const input = document.getElementById('manualValue');
  const val = parseFloat(input.value);
  if (isNaN(val) || val < 0 || val > 100) { showFormStatus('Введите число от 0 до 100', 'error'); return; }
  writeValueToCell(clampPct(val)); input.value = '';
  showFormStatus('Записано: ' + clampPct(val) + '%', 'success'); renderMatrix();
}
function acceptForemanValue() {
  if (S.isCellLocked) { showFormStatus('Нет отметки прораба', 'error'); return; }
  if (S.selColIndex === -1 || !S.works[S.selWorkIdx]) { showFormStatus('Сначала выберите ячейку', 'error'); return; }
  const frNum = parseNum(S.foremanCells[S.selWorkIdx + '_' + S.selColIndex]);
  if (frNum === null) { showFormStatus('У прораба нет значения', 'error'); return; }
  writeValueToCell(clampPct(frNum));
  showFormStatus('Подтверждено: ' + formatVal(frNum), 'success'); vibrate(30); renderMatrix();
}
function writeValueToCell(val) {
  const workObj = S.works[S.selWorkIdx], colIndex = S.selColIndex;
  const cellRef = XLSX.utils.encode_cell({ r: workObj.myRow, c: colIndex });
  const oldValue = S.myCells[S.selWorkIdx + '_' + colIndex];
  captureOriginal(S.currentSheet, cellRef);
  const originalSnapshot = getSheetOriginals(S.currentSheet).get(cellRef);
  S.undoStack.push({ type: 'single', sheetName: S.currentSheet, cellRef, workIdx: S.selWorkIdx, colIndex, oldValue, originalSnapshot, timestamp: Date.now() });
  if (S.undoStack.length > MAX_UNDO_STEPS) S.undoStack.shift();
  getSheetPendingChanges(S.currentSheet).set(cellRef, val);
  const ws = S.wb.Sheets[S.currentSheet];
  if (ws[cellRef]) { ws[cellRef].v = val; ws[cellRef].t = 'n'; delete ws[cellRef].w; }
  else ws[cellRef] = { v: val, t: 'n' };
  S.myCells[S.selWorkIdx + '_' + colIndex] = val;
  updateSelectionPanel(); rebuildSheetSelect(); updateSummary(); scheduleAutoSave();
}
function applyValueToCell(workIdx, colIndex, value) {
  const work = S.works[workIdx];
  const cellRef = XLSX.utils.encode_cell({ r: work.myRow, c: colIndex });
  const oldValue = S.myCells[workIdx + '_' + colIndex];
  captureOriginal(S.currentSheet, cellRef);
  const originalSnapshot = getSheetOriginals(S.currentSheet).get(cellRef);
  getSheetPendingChanges(S.currentSheet).set(cellRef, value);
  const ws = S.wb.Sheets[S.currentSheet];
  if (ws[cellRef]) { ws[cellRef].v = value; ws[cellRef].t = 'n'; delete ws[cellRef].w; }
  else ws[cellRef] = { v: value, t: 'n' };
  S.myCells[workIdx + '_' + colIndex] = value;
  return { cellRef, workIdx, colIndex, oldValue, originalSnapshot, sheetName: S.currentSheet };
}
function applyForemanByFloors() {
  if (!MS.floors.size) { showToast('Выберите хотя бы один этаж', 'warning'); return; }
  const targetCols = S.columns.filter(col => col.section === S.filterSection && MS.floors.has(col.colIndex));
  const changes = []; let applied = 0, skipped = 0;
  S.works.forEach((work, workIdx) => targetCols.forEach(col => {
    const frNum = parseNum(S.foremanCells[workIdx + '_' + col.colIndex]);
    if (frNum === null) { skipped++; return; }
    changes.push(applyValueToCell(workIdx, col.colIndex, clampPct(frNum))); applied++;
  }));
  if (changes.length) { S.undoStack.push({ type: 'mass', changes, timestamp: Date.now() }); if (S.undoStack.length > MAX_UNDO_STEPS) S.undoStack.shift(); }
  MS.floors.clear();
  document.querySelectorAll('#msFloorsOptions input[type="checkbox"]').forEach(cb => cb.checked = false);
  updateMultiselectLabel('msFloors'); updateSelectionPanel(); renderMatrix(); updateSummary(); rebuildSheetSelect();
  vibrate([30, 50, 30]);
  showToast(`Применено: ${applied}${skipped ? `, пропущено: ${skipped}` : ''}`, applied ? 'success' : 'warning');
  scheduleAutoSave();
}
function applyForemanByWorks() {
  if (!MS.works.size) { showToast('Выберите хотя бы одну работу', 'warning'); return; }
  const secCols = S.columns.filter(col => col.section === S.filterSection);
  const changes = []; let applied = 0, skipped = 0;
  MS.works.forEach(workIdx => {
    if (!S.works[workIdx]) return;
    secCols.forEach(col => {
      const frNum = parseNum(S.foremanCells[workIdx + '_' + col.colIndex]);
      if (frNum === null) { skipped++; return; }
      changes.push(applyValueToCell(workIdx, col.colIndex, clampPct(frNum))); applied++;
    });
  });
  if (changes.length) { S.undoStack.push({ type: 'mass', changes, timestamp: Date.now() }); if (S.undoStack.length > MAX_UNDO_STEPS) S.undoStack.shift(); }
  MS.works.clear();
  document.querySelectorAll('#msWorksOptions input[type="checkbox"]').forEach(cb => cb.checked = false);
  updateMultiselectLabel('msWorks'); updateSelectionPanel(); renderMatrix(); updateSummary(); rebuildSheetSelect();
  vibrate([30, 50, 30]);
  showToast(`Применено: ${applied}${skipped ? `, пропущено: ${skipped}` : ''}`, applied ? 'success' : 'warning');
  scheduleAutoSave();
}

/* ============================== МУЛЬТИВЫБОР ============================== */
function toggleMultiselect(id, e) {
  e.stopPropagation();
  document.querySelectorAll('.multiselect.open').forEach(ms => { if (ms.id !== id) ms.classList.remove('open'); });
  document.getElementById(id).classList.toggle('open');
}
document.addEventListener('click', e => { if (!e.target.closest('.multiselect')) document.querySelectorAll('.multiselect.open').forEach(ms => ms.classList.remove('open')); });
function multiselectAll(id, selectAll, e) {
  e.stopPropagation();
  const container = document.getElementById(id + 'Options'); if (!container) return;
  const baseSet = id.includes('Floors') ? MS.floors : MS.works;
  container.querySelectorAll('input[type="checkbox"]').forEach(cb => {
    cb.checked = selectAll;
    const val = parseInt(cb.value, 10);
    if (selectAll) baseSet.add(val); else baseSet.delete(val);
  });
  updateMultiselectLabel(id);
}
function updateMultiselectLabel(id) {
  const set = id.includes('Floors') ? MS.floors : MS.works;
  const label = document.querySelector(`#${id} .multiselect-label`); if (!label) return;
  label.textContent = `${id.includes('Floors') ? 'Выбрать этажи' : 'Выбрать работы'} (${set.size})`;
}
function handleMsFloorChange(cb) { const v = parseInt(cb.value, 10); if (cb.checked) MS.floors.add(v); else MS.floors.delete(v); updateMultiselectLabel('msFloors'); }
function handleMsWorkChange(cb) { const v = parseInt(cb.value, 10); if (cb.checked) MS.works.add(v); else MS.works.delete(v); updateMultiselectLabel('msWorks'); }
function filterMsWorks(id, query) {
  const q = query.trim().toLowerCase();
  document.querySelectorAll(`#${id}Options .multiselect-option`).forEach(opt => {
    const name = opt.getAttribute('data-name') || '';
    opt.classList.toggle('hidden', !(!q || name.includes(q)));
  });
}
function rebuildMultiselects() { rebuildMultiselectsFor('msFloors', 'msWorks'); }
function rebuildMultiselectsFor(floorsId, worksId) {
  const floorsContainer = document.getElementById(floorsId + 'Options');
  if (floorsContainer) {
    const secCols = S.columns.filter(col => col.section === S.filterSection);
    MS.floors.clear();
    floorsContainer.innerHTML = secCols.length ? secCols.map(col =>
      `<label class="multiselect-option"><input type="checkbox" value="${col.colIndex}" onchange="handleMsFloorChange(this)"><span class="multiselect-option-label">${esc(col.floor)}</span></label>`).join('')
      : '<div class="px-3 py-2 text-xs text-surface-400">Нет этажей</div>';
    updateMultiselectLabel(floorsId);
  }
  const worksContainer = document.getElementById(worksId + 'Options');
  if (worksContainer) {
    const secCols = S.columns.filter(col => col.section === S.filterSection);
    let visibleWorks = S.works.map((w, idx) => Object.assign({}, w, { idx }));
    if (S.filterWorkIdxs.size) visibleWorks = visibleWorks.filter(w => S.filterWorkIdxs.has(w.idx));
    if (S.hideAccepted) visibleWorks = visibleWorks.filter(w => !shouldHideRow(w, secCols));
    MS.works.clear();
    worksContainer.innerHTML = visibleWorks.map(w =>
      `<label class="multiselect-option" data-name="${esc(w.name).toLowerCase()}"><input type="checkbox" value="${w.idx}" onchange="handleMsWorkChange(this)"><span class="multiselect-option-label">${esc(w.name)}</span></label>`).join('');
    updateMultiselectLabel(worksId);
  }
}

/* ============================== UNDO ============================== */
function undoLastChange() {
  if (!S.undoStack.length) return;
  const action = S.undoStack.pop();
  const list = action.type === 'single' ? [action] : action.changes;
  const touched = new Set();
  list.forEach(ch => {
    const sheetName = ch.sheetName || S.currentSheet;
    const ws = S.wb.Sheets[sheetName];
    const originals = getSheetOriginals(sheetName);
    const snap = ch.originalSnapshot !== undefined ? ch.originalSnapshot : (originals.has(ch.cellRef) ? originals.get(ch.cellRef) : null);
    restoreCellSnapshot(ws, ch.cellRef, snap);
    getSheetPendingChanges(sheetName).delete(ch.cellRef);
    if (sheetName === S.currentSheet) S.myCells[ch.workIdx + '_' + ch.colIndex] = ch.oldValue;
    touched.add(sheetName);
  });
  touched.forEach(name => { if (getSheetPendingChanges(name).size === 0) S.cachedSheets.delete(name); });
  updateSelectionPanel(); renderMatrix(); updateSummary(); rebuildSheetSelect(); updateUndoButton();
  showToast('Изменение отменено', 'success'); vibrate([30, 50, 30]); scheduleAutoSave();
}

/* ============================== СОХРАНЕНИЕ / ЭКСПОРТ ============================== */
async function buildPatchedBuffer() {
  let buffer = S.originalBuffer;
  const total = getTotalPendingChanges();
  if (total > 0 && buffer) {
    for (const sheetName in S.pendingChanges) {
      const changes = S.pendingChanges[sheetName];
      if (changes && changes.size) buffer = await patchExcelFile(buffer, changes, sheetName);
    }
  } else if (!buffer && S.wb) {
    buffer = XLSX.write(S.wb, { bookType: 'xlsx', type: 'array' });
  }
  return buffer;
}
function clearAfterPersist(buffer) {
  S.originalBuffer = buffer;
  for (const k in S.pendingChanges) S.pendingChanges[k].clear();
  S.originalCells = {}; S.undoStack = []; S.cachedSheets.clear();
  clearDraft(); rebuildSheetSelect(); updateUndoButton(); updateSummary();
  idbPut('lastFile', { name: S.fileName, buffer: buffer, timestamp: Date.now() }).catch(() => {});
}
function reloadFromBuffer(buffer) {
  S.wb = XLSX.read(new Uint8Array(buffer), { type: 'array', cellDates: true });
  if (!S.wb.SheetNames.includes(S.currentSheet)) S.currentSheet = S.wb.SheetNames[0];
  parseSheet(); rebuildSheetSelect();
}
async function saveToFile() {
  if (typeof XLSX === 'undefined') { showToast('Библиотека XLSX не загружена', 'error'); return false; }
  const total = getTotalPendingChanges();
  if (!total && !S.originalBuffer && !S.wb) { showToast('Нет данных для сохранения', 'warning'); return false; }
  if (!total) { showToast('Нет изменений', 'warning'); return true; }
  try {
    const buffer = await buildPatchedBuffer();
    if (!buffer) { showToast('Нет данных для сохранения', 'warning'); return false; }
    if (S.saveHandle) {
      const writable = await S.saveHandle.createWritable();
      await writable.write(buffer); await writable.close();
      clearAfterPersist(buffer); reloadFromBuffer(buffer);
      showToast('Файл сохранён ✓', 'success'); vibrate([30, 50, 30]); return true;
    }
    downloadBuffer(buffer, S.fileName || 'audit.xlsx');
    clearAfterPersist(buffer); reloadFromBuffer(buffer);
    showToast('Файл скачан (браузер не умеет перезаписывать исходный файл)', 'success');
    vibrate([30, 50, 30]); return true;
  } catch (err) {
    console.error('Save error:', err);
    showToast('Ошибка сохранения: ' + err.message, 'error'); return false;
  }
}
async function exportToFile() {
  if (typeof XLSX === 'undefined') { showToast('Библиотека XLSX не загружена', 'error'); return false; }
  if (!getTotalPendingChanges() && !S.originalBuffer && !S.wb) { showToast('Нет данных для экспорта', 'warning'); return false; }
  try {
    const buffer = await buildPatchedBuffer();
    if (!buffer) { showToast('Нет данных для экспорта', 'warning'); return false; }
    const now = new Date();
    const p = n => String(n).padStart(2, '0');
    const dateStr = `${p(now.getDate())}-${p(now.getMonth() + 1)}-${String(now.getFullYear()).slice(2)}_${p(now.getHours())}-${p(now.getMinutes())}-${p(now.getSeconds())}`;
    downloadBuffer(buffer, `audit_${dateStr}.xlsx`);
    clearAfterPersist(buffer); reloadFromBuffer(buffer);
    showToast('Файл экспортирован ✓', 'success'); vibrate([30, 50, 30]); return true;
  } catch (err) {
    console.error('Export error:', err);
    showToast('Ошибка экспорта: ' + err.message, 'error'); return false;
  }
}
function downloadBuffer(buffer, filename) {
  const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
  const url = URL.createObjectURL(blob);
  const link = document.createElement('a');
  link.href = url; link.download = filename;
  document.body.appendChild(link); link.click(); document.body.removeChild(link);
  setTimeout(() => URL.revokeObjectURL(url), 100);
}
async function shareFile() {
  if (!S.originalBuffer && !S.wb) { showToast('Нет файла', 'error'); return; }
  try {
    const buffer = await buildPatchedBuffer();
    const now = new Date(); const p = n => String(n).padStart(2, '0');
    const dateStr = `${p(now.getDate())}-${p(now.getMonth() + 1)}-${String(now.getFullYear()).slice(2)}_${p(now.getHours())}-${p(now.getMinutes())}-${p(now.getSeconds())}`;
    const filename = `audit_${dateStr}.xlsx`;
    const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    const file = new File([blob], filename, { type: blob.type });
    if (navigator.canShare && navigator.canShare({ files: [file] })) {
      try {
        await navigator.share({ files: [file], title: 'Аудит СМР', text: 'Отчёт по факту выполнения работ' });
        showToast('Файл отправлен ✓', 'success');
      } catch (err) { if (err.name !== 'AbortError') { console.error(err); showToast('Ошибка шеринга', 'error'); } }
    } else { downloadBuffer(buffer, filename); showToast('Файл скачан ✓', 'success'); }
  } catch (err) { showToast('Ошибка: ' + err.message, 'error'); }
}

/* ============================== ЧЕРНОВИК / АВТОСОХРАНЕНИЕ ============================== */
function scheduleAutoSave() { S.lastChangeTime = Date.now(); }
function saveDraft() {
  if (!S.isLocalStorageAvailable || !S.fileName) return;
  try {
    const draftKey = `draft_${S.fileName}`;
    const allChanges = {}; let total = 0;
    for (const sheetName in S.pendingChanges) {
      if (S.pendingChanges[sheetName].size) { allChanges[sheetName] = Array.from(S.pendingChanges[sheetName].entries()); total += S.pendingChanges[sheetName].size; }
    }
    if (!total) { localStorage.removeItem(draftKey); return; }
    localStorage.setItem(draftKey, JSON.stringify({ allChanges, total, timestamp: Date.now() }));
  } catch (e) { console.warn('save draft', e); }
}
function checkDraft() {
  if (!S.isLocalStorageAvailable || !S.fileName) return;
  try {
    const draftKey = `draft_${S.fileName}`;
    const draftStr = localStorage.getItem(draftKey);
    if (!draftStr) return;
    const draft = JSON.parse(draftStr);
    if (!draft.allChanges || !draft.total) { localStorage.removeItem(draftKey); return; }
    const d = new Date(draft.timestamp); const p = n => String(n).padStart(2, '0');
    const timeStr = `${p(d.getDate())}.${p(d.getMonth() + 1)} ${p(d.getHours())}:${p(d.getMinutes())}`;
    showAppModal('Найден черновик',
      `Несохранённые правки (${draft.total} ячеек) от ${timeStr}. Восстановить?`,
      [
        { label: 'Восстановить', cls: 'btn-gradient text-white', fn: () => {
            for (const sheetName in draft.allChanges) { S.pendingChanges[sheetName] = new Map(draft.allChanges[sheetName]); S.cachedSheets.add(sheetName); }
            applyPendingToCurrentSheet(); renderMatrix(); updateUndoButton(); rebuildSheetSelect(); updateSummary();
            showToast('Черновик восстановлен', 'success');
          } },
        { label: 'Удалить черновик', cls: 'bg-red-600 hover:bg-red-700 text-white', fn: () => localStorage.removeItem(draftKey) },
        { label: 'Позже', fn: () => {} }
      ]);
  } catch (e) { console.warn('check draft', e); }
}
function clearDraft() { if (S.isLocalStorageAvailable && S.fileName) { try { localStorage.removeItem(`draft_${S.fileName}`); } catch (e) {} } }

/* ============================== SUMMARY / ФИЛЬТРЫ ============================== */
function updateSummary() {
  const secCols = S.columns.filter(col => col.section === S.filterSection);
  let total = 0, done = 0, mismatch = 0;
  S.works.forEach((w, idx) => {
    const work = Object.assign({}, w, { idx });
    if (shouldHideRow(work, secCols)) return;
    total++;
    if (isWorkDone(work, secCols)) done++;
    if (countMismatches(work, secCols) > 0) mismatch++;
  });
  document.getElementById('summaryTotal').textContent = `📊 ${total} работ`;
  document.getElementById('summaryDone').textContent = `✓ ${done}`;
  document.getElementById('summaryMismatch').textContent = `⚠ ${mismatch}`;
}
function filterBySummary(type) {
  S.summaryFilter = S.summaryFilter === type ? null : type;
  renderMatrix(); rebuildWorksDropdown();
}

/* ============================== ФАЙЛОВЫЕ СОБЫТИЯ ============================== */
function handleDragOver(e) { e.preventDefault(); e.stopPropagation(); document.getElementById('dropZone').classList.add('drop-active'); }
function handleDragLeave(e) { e.preventDefault(); e.stopPropagation(); document.getElementById('dropZone').classList.remove('drop-active'); }
function handleDrop(e) {
  e.preventDefault(); e.stopPropagation();
  document.getElementById('dropZone').classList.remove('drop-active');
  if (e.dataTransfer.files.length) processFile(e.dataTransfer.files[0], null);
}
function handleFileSelect(e) { if (e.target.files.length) processFile(e.target.files[0], null); e.target.value = ''; }
async function openFilePicker() {
  const proceed = async () => {
    if (window.showOpenFilePicker) {
      try {
        const [handle] = await window.showOpenFilePicker({
          multiple: false,
          types: [{ description: 'Excel', accept: { 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet': ['.xlsx'] } }]
        });
        const file = await handle.getFile();
        processFile(file, handle);
      } catch (err) { if (err.name !== 'AbortError') document.getElementById('fileInputSecondary').click(); }
    } else {
      const el = document.getElementById('fileInputSecondary') || document.getElementById('fileInput');
      el.click();
    }
  };
  if (getTotalPendingChanges() > 0) showWarningModal(proceed); else proceed();
}
function processFile(file, handle) {
  const ext = file.name.split('.').pop().toLowerCase();
  if (ext !== 'xlsx') { showError('Поддерживаются только .xlsx'); return; }
  readExcelFile(file, handle);
}
function showFormStatus(msg, type) {
  const el = document.getElementById('formStatus'); if (!el) return;
  el.textContent = msg;
  el.className = 'mt-3 p-3 rounded-xl text-sm font-medium transition-all ' +
    (type === 'success' ? 'bg-green-50 dark:bg-green-900/20 text-green-700 dark:text-green-400 border border-green-200 dark:border-green-800'
      : 'bg-red-50 dark:bg-red-900/20 text-red-700 dark:text-red-400 border border-red-200 dark:border-red-800');
  el.classList.remove('hidden');
  setTimeout(() => el.classList.add('hidden'), 3000);
}
function showError(msg) {
  const el = document.getElementById('uploadError');
  el.textContent = msg; el.classList.remove('hidden');
  setTimeout(() => el.classList.add('hidden'), 6000);
}
function toggleTheme() {
  S.darkMode = !S.darkMode;
  document.documentElement.classList.toggle('dark', S.darkMode);
  if (S.isLocalStorageAvailable) { try { localStorage.setItem('ctrl-dark', S.darkMode ? '1' : '0'); } catch (e) {} }
}

/* ============================== ДЕМО-ПРОЕКТ ============================== */
function buildDemo() {
  if (typeof XLSX === 'undefined') { showToast('Библиотека XLSX не загружена. Выполните node tools/fetch-libs.mjs', 'error'); return; }
  const sections = ['Секция 1', 'Секция 2', 'Секция 3'];
  const floorsPerSection = 5;
  const works = ['Монолитные работы', 'Кладка перегородок', 'Штукатурка стен', 'Стяжка пола', 'Электромонтаж', 'Сантехника (ВК)', 'Окна ПВХ', 'Чистовая отделка'];
  const aoa = [];
  const row0 = ['', '', '']; const row1 = ['', '', '']; const row2 = ['', '', 'Этаж'];
  let col = 3;
  sections.forEach(sec => { for (let f = 1; f <= floorsPerSection; f++) { row0[col] = f === 1 ? sec : ''; row1[col] = ''; row2[col] = String(f); col++; } });
  aoa.push(row0, row1, row2);
  const totalCols = col;
  works.forEach((w, wi) => {
    const myRow = ['', '', w]; const frRow = ['ПРОРАБ', '', ''];
    for (let c = 3; c < totalCols; c++) {
      const floor = ((c - 3) % floorsPerSection) + 1;
      const base = 100 - (floor - 1) * 18 - wi * 6;
      const frVal = Math.max(0, Math.min(100, base + ((c * 7 + wi * 13) % 15) - 7));
      frRow[c] = frVal <= 5 ? null : frVal;
      myRow[c] = null;
    }
    aoa.push(myRow, frRow);
  });
  const ws = XLSX.utils.aoa_to_sheet(aoa);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, 'Корпус 1');
  const buffer = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
  S.originalBuffer = buffer;
  S.wb = XLSX.read(new Uint8Array(buffer), { type: 'array', cellDates: true });
  S.fileName = 'Демо — ЖК Северный парк.xlsx';
  S.saveHandle = null; S.demoMode = true;
  S.pendingChanges = {}; S.originalCells = {}; S.undoStack = [];
  S.currentSheet = S.wb.SheetNames[0];
  S.sheetColors = { [S.currentSheet]: {} };
  S.cellColors = {};
  rebuildSheetSelect(); parseSheet();
  showToast('Демо-проект загружен. Данные вымышленные.', 'success');
}

/* ============================== INDEXEDDB ============================== */
const IDB_NAME = 'audit-smr-store', IDB_VERSION = 1, IDB_STORE = 'kv';
function openIDB() {
  return new Promise((resolve, reject) => {
    if (!window.indexedDB) { reject(new Error('IndexedDB недоступна')); return; }
    const req = indexedDB.open(IDB_NAME, IDB_VERSION);
    req.onupgradeneeded = () => { const db = req.result; if (!db.objectStoreNames.contains(IDB_STORE)) db.createObjectStore(IDB_STORE); };
    req.onsuccess = () => resolve(req.result);
    req.onerror = () => reject(req.error);
  });
}
async function idbPut(key, value) { const db = await openIDB(); return new Promise((res, rej) => { const tx = db.transaction(IDB_STORE, 'readwrite'); tx.objectStore(IDB_STORE).put(value, key); tx.oncomplete = () => { db.close(); res(); }; tx.onerror = () => { db.close(); rej(tx.error); }; }); }
async function idbGet(key) { const db = await openIDB(); return new Promise((res, rej) => { const tx = db.transaction(IDB_STORE, 'readonly'); const rq = tx.objectStore(IDB_STORE).get(key); rq.onsuccess = () => { db.close(); res(rq.result); }; rq.onerror = () => { db.close(); rej(rq.error); }; }); }
async function idbDel(key) { const db = await openIDB(); return new Promise((res, rej) => { const tx = db.transaction(IDB_STORE, 'readwrite'); tx.objectStore(IDB_STORE).delete(key); tx.oncomplete = () => { db.close(); res(); }; tx.onerror = () => { db.close(); rej(tx.error); }; }); }
async function tryRestoreLastFile() {
  let rec; try { rec = await idbGet('lastFile'); } catch (e) { return; }
  if (!rec || !rec.buffer) return;
  const d = new Date(rec.timestamp); const p = n => String(n).padStart(2, '0');
  const timeStr = `${p(d.getDate())}.${p(d.getMonth() + 1)} ${p(d.getHours())}:${p(d.getMinutes())}`;
  showAppModal('Продолжить работу?',
    `Последний файл: <b>${esc(rec.name)}</b><br>Сохранён: ${timeStr}.<br>Открыть его заново без выбора файла?`,
    [
      { label: 'Открыть файл', cls: 'btn-gradient text-white', fn: () => {
          const file = new File([rec.buffer], rec.name, { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
          readExcelFile(file, null);
        } },
      { label: 'Не сейчас', fn: () => {} },
      { label: 'Забыть файл', cls: 'bg-red-600 hover:bg-red-700 text-white', fn: () => idbDel('lastFile').catch(() => {}) }
    ]);
}

/* ============================== PWA ============================== */
function initPWA() {
  makeAppleTouchIcon();
  updateNetworkStatus();
  window.addEventListener('online', updateNetworkStatus);
  window.addEventListener('offline', updateNetworkStatus);
  window.addEventListener('beforeinstallprompt', e => {
    e.preventDefault(); deferredPrompt = e;
    const btn = document.getElementById('installBtn');
    btn.classList.remove('hidden'); btn.classList.add('inline-flex');
  });
  window.addEventListener('appinstalled', () => {
    deferredPrompt = null;
    const btn = document.getElementById('installBtn');
    btn.classList.add('hidden'); btn.classList.remove('inline-flex');
    showToast('Приложение установлено ✓', 'success');
  });
  if ('serviceWorker' in navigator && location.protocol !== 'file:') {
    navigator.serviceWorker.register('sw.js').then(reg => {
      swRegistration = reg;
      if (reg.waiting) showUpdateButton();
      reg.addEventListener('updatefound', () => {
        const nw = reg.installing;
        if (nw) nw.addEventListener('statechange', () => { if (nw.state === 'installed' && navigator.serviceWorker.controller) showUpdateButton(); });
      });
    }).catch(err => console.warn('SW registration failed', err));
    navigator.serviceWorker.addEventListener('controllerchange', () => { if (!refreshing) { refreshing = true; location.reload(); } });
  }
}
function showUpdateButton() {
  const btn = document.getElementById('updateBtn');
  btn.classList.remove('hidden'); btn.classList.add('inline-flex');
  showToast('Доступно обновление приложения', 'success');
}
function applyUpdate() {
  if (swRegistration && swRegistration.waiting) swRegistration.waiting.postMessage({ type: 'SKIP_WAITING' });
  else location.reload();
}
async function promptInstall() {
  if (deferredPrompt) {
    deferredPrompt.prompt();
    const choice = await deferredPrompt.userChoice;
    deferredPrompt = null;
    const btn = document.getElementById('installBtn');
    btn.classList.add('hidden'); btn.classList.remove('inline-flex');
    if (choice.outcome === 'accepted') showToast('Установка началась…', 'success');
  } else {
    showAppModal('Установка приложения',
      'Откройте меню браузера и выберите <b>«Установить приложение»</b> / <b>«Добавить на главный экран»</b>.<br><br>На iPhone/iPad: Safari → кнопка «Поделиться» → «На экран „Домой“».',
      [{ label: 'Понятно', cls: 'btn-gradient text-white' }]);
  }
}
function updateNetworkStatus() {
  const online = navigator.onLine;
  const chip = document.getElementById('netStatus'), dot = document.getElementById('netDot'), text = document.getElementById('netText');
  if (!chip) return;
  text.textContent = online ? 'Онлайн' : 'Офлайн';
  dot.className = 'w-1.5 h-1.5 rounded-full ' + (online ? 'bg-green-500' : 'bg-amber-500');
  chip.className = 'hidden sm:inline-flex items-center gap-1.5 px-2 py-1 rounded-full text-[11px] font-medium ' +
    (online ? 'bg-green-50 dark:bg-green-900/30 text-green-700 dark:text-green-400' : 'bg-amber-50 dark:bg-amber-900/30 text-amber-700 dark:text-amber-400');
}
function makeAppleTouchIcon() {
  try {
    const size = 180, c = document.createElement('canvas'); c.width = c.height = size;
    const ctx = c.getContext('2d');
    const g = ctx.createLinearGradient(0, 0, size, size);
    g.addColorStop(0, '#b83a9e'); g.addColorStop(1, '#5f1152');
    ctx.fillStyle = g; ctx.fillRect(0, 0, size, size);
    ctx.fillStyle = '#ffffff';
    ctx.font = 'bold ' + Math.round(size * 0.3) + 'px Arial, sans-serif';
    ctx.textAlign = 'center'; ctx.textBaseline = 'middle';
    ctx.fillText('СМР', size / 2, size / 2 + size * 0.02);
    const link = document.createElement('link');
    link.rel = 'apple-touch-icon'; link.href = c.toDataURL('image/png');
    document.head.appendChild(link);
  } catch (e) {}
}

/* ============================== СТАРТ ============================== */
(function init() {
  checkLocalStorage();
  try { if (S.isLocalStorageAvailable && localStorage.getItem('ctrl-dark') === '1') { S.darkMode = true; document.documentElement.classList.add('dark'); } } catch (e) {}
  setTimeout(() => { const sp = document.getElementById('splashScreen'); if (sp) sp.classList.add('hidden'); }, 1200);

  const mv = document.getElementById('manualValue');
  if (mv) mv.addEventListener('keydown', e => { if (e.key === 'Enter') submitManualValue(); });

  document.addEventListener('keydown', e => {
    if (e.key === 'Escape') { closeBottomSheet(); closeWarningModal(); closeAppModal(); }
    if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 's') { e.preventDefault(); if (!document.getElementById('appScreen').classList.contains('hidden')) saveToFile(); }
    if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 'z') {
      const tag = (document.activeElement && document.activeElement.tagName) || '';
      if (tag !== 'INPUT' && tag !== 'TEXTAREA') { e.preventDefault(); undoLastChange(); }
    }
  });

  window.addEventListener('pagehide', () => { if (getTotalPendingChanges() > 0) saveDraft(); });
  document.addEventListener('visibilitychange', () => { if (document.visibilityState === 'hidden' && getTotalPendingChanges() > 0) saveDraft(); });
  window.addEventListener('resize', () => { S.isMobile = window.innerWidth < 1024; });

  setInterval(() => { if (getTotalPendingChanges() > 0) saveDraft(); }, S.autoSaveInterval);

  initPWA();

  if (typeof XLSX === 'undefined' || typeof JSZip === 'undefined') {
    showError('Библиотеки не загружены. Выполните "node tools/fetch-libs.mjs" в папке проекта и обновите страницу.');
  }

  const params = new URLSearchParams(location.search);
  if (params.get('demo') === '1') buildDemo();
  else tryRestoreLastFile();
})();