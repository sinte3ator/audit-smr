'use strict';

/* ================= СОСТОЯНИЕ ================= */
const S = {
  wb: null,
  fileName: '',
  originalBuffer: null,
  saveHandle: null,
  currentSheet: '',
  sheetData: {},
  sheetColors: {},
  uniqueSections: [],
  columns: [],
  works: [],
  myCells: {},
  foremanCells: {},
  cellColors: {},
  selSection: '',
  selColIndex: -1,
  selWorkIdx: 0,
  filterSection: '',
  filterColIndex: -1,
  filterWorkIdxs: new Set(),
  hideAccepted: false,
  darkMode: false,
  pendingChanges: {},
  originalCells: {},
  isCellLocked: false,
  isMobile: window.innerWidth < 1024,
  undoStack: [],
  lastChangeTime: 0,
  summaryFilter: null,
  cachedSheets: new Set(),
  autoSaveInterval: 30000,
  isLocalStorageAvailable: false,
  demoMode: false
};

const MS = { floors: new Set(), works: new Set() };
const MAX_UNDO_STEPS = 30;
let pendingAction = null;

/* ================= УТИЛИТЫ ================= */
function checkLocalStorage() {
  try {
    const t = '__t';
    localStorage.setItem(t, t);
    localStorage.removeItem(t);
    S.isLocalStorageAvailable = true;
  } catch (e) {
    S.isLocalStorageAvailable = false;
  }
}

function isMobile() {
  return window.innerWidth < 1024;
}

function vibrate(ms = 30) {
  if (isMobile() && navigator.vibrate) navigator.vibrate(ms);
}

function showToast(message, type = 'success', action = null) {
  const toast = document.createElement('div');
  toast.className = 'toast toast-' + type;

  const span = document.createElement('span');
  span.textContent = message;
  toast.appendChild(span);

  const kill = () => {
    toast.style.opacity = '0';
    setTimeout(() => toast.remove(), 300);
  };

  if (action) {
    const b = document.createElement('button');
    b.className = 'toast-action';
    b.textContent = action.label;
    b.onclick = () => {
      action.fn();
      kill();
    };
    toast.appendChild(b);
  }

  document.body.appendChild(toast);
  setTimeout(kill, action ? 6500 : 3000);
}

function esc(s) {
  const d = document.createElement('div');
  d.textContent = s == null ? '' : String(s);
  return d.innerHTML;
}

function hasAnyContent(v) {
  return v != null && String(v).trim().length > 0;
}

function parseNum(v) {
  if (v == null || v === '') return null;
  if (typeof v === 'number') return isFinite(v) ? v : null;

  const s = String(v).replace(',', '.').replace('%', '').replace(/\s/g, '');
  if (s === '—' || s === '-') return null;

  const n = parseFloat(s);
  return isNaN(n) ? null : n;
}

function formatVal(v) {
  const n = parseNum(v);
  return n === null ? '—' : Math.round(n) + '%';
}

function clampPct(n) {
  return Math.max(0, Math.min(100, Math.round(n)));
}

function getTotalPendingChanges() {
  let t = 0;
  for (const k in S.pendingChanges) {
    if (S.pendingChanges[k] && S.pendingChanges[k].size) t += S.pendingChanges[k].size;
  }
  return t;
}

function getSheetPendingChanges(n) {
  if (!S.pendingChanges[n]) S.pendingChanges[n] = new Map();
  return S.pendingChanges[n];
}

function getSheetOriginals(n) {
  if (!S.originalCells[n]) S.originalCells[n] = new Map();
  return S.originalCells[n];
}

function captureOriginal(sheet, ref) {
  const o = getSheetOriginals(sheet);
  if (!o.has(ref)) {
    const ws = S.wb && S.wb.Sheets[sheet];
    const c = ws ? ws[ref] : undefined;
    o.set(ref, c ? Object.assign({}, c) : null);
  }
}

function restoreCellSnapshot(ws, ref, snap) {
  if (!ws) return;
  if (snap == null) delete ws[ref];
  else ws[ref] = Object.assign({}, snap);
}

/* ================= МОДАЛКИ ================= */
function showAppModal(title, message, buttons) {
  document.getElementById('appModalTitle').textContent = title;
  document.getElementById('appModalMessage').innerHTML = message;

  const box = document.getElementById('appModalButtons');
  box.innerHTML = '';

  buttons.forEach(b => {
    const btn = document.createElement('button');
    btn.className = 'w-full px-4 py-3 font-semibold rounded-xl ' +
      (b.cls || 'bg-surface-200 dark:bg-surface-700 text-surface-900 dark:text-white');
    btn.textContent = b.label;
    btn.onclick = () => {
      closeAppModal();
      b.fn && b.fn();
    };
    box.appendChild(btn);
  });

  document.getElementById('appModal').classList.add('open');
}

function closeAppModal() {
  document.getElementById('appModal').classList.remove('open');
}

function showWarningModal(action) {
  pendingAction = action;
  document.getElementById('warningModal').classList.add('open');
}

function closeWarningModal() {
  document.getElementById('warningModal').classList.remove('open');
  pendingAction = null;
}

async function saveAndExit() {
  const a = pendingAction;
  closeWarningModal();
  const ok = await saveToFile();
  if (ok && a) a();
}

function exitWithoutSaving() {
  const a = pendingAction;
  revertCurrentSheetChanges();
  closeWarningModal();
  if (a) a();
}

function revertCurrentSheetChanges() {
  const sheet = S.currentSheet;
  const changes = getSheetPendingChanges(sheet);
  const originals = getSheetOriginals(sheet);
  const ws = S.wb && S.wb.Sheets[sheet];

  changes.forEach((val, ref) => {
    const snap = originals.has(ref) ? originals.get(ref) : null;
    restoreCellSnapshot(ws, ref, snap);

    const m = ref.match(/^([A-Z]+)(\d+)$/);
    if (m && ws) {
      const c = colNameToIndex(m[1]) - 1;
      const r = parseInt(m[2], 10) - 1;
      const wi = S.works.findIndex(w => w.myRow === r);
      if (wi >= 0) S.myCells[wi + '_' + c] = snap && snap.v != null ? snap.v : null;
    }
  });

  changes.clear();
  originals.clear();
  S.cachedSheets.delete(sheet);

  S.undoStack = S.undoStack.filter(a =>
    (a.sheetName || (a.changes && a.changes[0] && a.changes[0].sheetName)) !== sheet
  );

  refreshUI();
}

/* ================= EXCEL ================= */
const INDEXED_COLORS = [
  '#000000', '#FFFFFF', '#FF0000', '#00FF00', '#0000FF', '#FFFF00', '#FF00FF', '#00FFFF',
  '#000000', '#FFFFFF', '#FF0000', '#00FF00', '#0000FF', '#FFFF00', '#FF00FF', '#00FFFF',
  '#800000', '#008000', '#000080', '#808000', '#800080', '#008080', '#C0C0C0', '#808080',
  '#9999FF', '#993366', '#FFFFCC', '#CCFFFF', '#660066', '#FF8080', '#0066CC', '#CCCCFF',
  '#000080', '#FF00FF', '#FFFF00', '#00FFFF', '#800080', '#800000', '#008080', '#0000FF',
  '#00CCFF', '#CCFFFF', '#CCFFCC', '#FFFF99', '#99CCFF', '#FF99CC', '#CC99FF', '#FFCC99',
  '#3366FF', '#33CCCC', '#99CC00', '#FFCC00', '#FF9900', '#FF6600', '#666699', '#969696',
  '#003366', '#339966', '#003300', '#333300', '#993300', '#993366', '#333399', '#333333'
];

function applyTint(hex, tint) {
  if (!hex || tint == null || tint === 0) return hex;

  hex = hex.replace('#', '');
  if (hex.length === 8) hex = hex.slice(2);

  let r = parseInt(hex.substring(0, 2), 16);
  let g = parseInt(hex.substring(2, 4), 16);
  let b = parseInt(hex.substring(4, 6), 16);

  if (tint < 0) {
    r = Math.round(r * (1 + tint));
    g = Math.round(g * (1 + tint));
    b = Math.round(b * (1 + tint));
  } else {
    r = Math.round(r * (1 - tint) + 255 * tint);
    g = Math.round(g * (1 - tint) + 255 * tint);
    b = Math.round(b * (1 - tint) + 255 * tint);
  }

  return '#' + [r, g, b]
    .map(x => Math.max(0, Math.min(255, x)).toString(16).padStart(2, '0'))
    .join('');
}

function isLightColor(hex) {
  if (!hex) return true;

  hex = hex.replace('#', '');
  const r = parseInt(hex.substring(0, 2), 16);
  const g = parseInt(hex.substring(2, 4), 16);
  const b = parseInt(hex.substring(4, 6), 16);

  return (0.299 * r + 0.587 * g + 0.114 * b) / 255 > 0.55;
}

function colNameToIndex(c) {
  let i = 0;
  for (let k = 0; k < c.length; k++) {
    i = i * 26 + (c.charCodeAt(k) - 64);
  }
  return i;
}

async function getSheetFilePath(zip, sheetName) {
  try {
    const wbFile = zip.file('xl/workbook.xml');
    if (!wbFile) return 'xl/worksheets/sheet1.xml';

    const parser = new DOMParser();
    const wbXml = parser.parseFromString(await wbFile.async('text'), 'text/xml');
    const sheets = wbXml.getElementsByTagName('sheet');
    let rid = null;

    for (let i = 0; i < sheets.length; i++) {
      if (sheets[i].getAttribute('name') === sheetName) {
        rid = sheets[i].getAttribute('r:id') || sheets[i].getAttribute('id');
        break;
      }
    }

    if (!rid) return 'xl/worksheets/sheet1.xml';

    const relsFile = zip.file('xl/_rels/workbook.xml.rels');
    if (!relsFile) return 'xl/worksheets/sheet1.xml';

    const relsXml = parser.parseFromString(await relsFile.async('text'), 'text/xml');
    const rels = relsXml.getElementsByTagName('Relationship');

    for (let i = 0; i < rels.length; i++) {
      if (rels[i].getAttribute('Id') === rid) {
        let t = rels[i].getAttribute('Target') || '';
        if (t.charAt(0) === '/') return t.slice(1);
        if (t.indexOf('xl/') === 0) return t;
        return 'xl/' + t.replace(/^\.\//, '');
      }
    }
  } catch (e) {
    console.warn('sheet path', e);
  }

  return 'xl/worksheets/sheet1.xml';
}

async function extractExcelColorsForSheet(buffer, sheetName) {
  const colorMap = {};

  try {
    const zip = await JSZip.loadAsync(buffer);
    const parser = new DOMParser();

    const sheetPath = await getSheetFilePath(zip, sheetName);
    const sheetFile = zip.file(sheetPath);
    if (!sheetFile) return colorMap;

    const themeColors = [
      '#FFFFFF', '#000000', '#E7E6E6', '#44546A', '#4472C4',
      '#ED7D31', '#A5A5A5', '#FFC000', '#5B9BD5', '#70AD47'
    ];

    const themeFile = zip.file('xl/theme/theme1.xml');
    if (themeFile) {
      const x = parser.parseFromString(await themeFile.async('text'), 'text/xml');
      const cs = x.getElementsByTagName('a:clrScheme')[0];

      if (cs) {
        ['a:lt1', 'a:dk1', 'a:lt2', 'a:dk2', 'a:accent1', 'a:accent2', 'a:accent3', 'a:accent4', 'a:accent5', 'a:accent6']
          .forEach((tag, idx) => {
            const el = cs.getElementsByTagName(tag)[0];
            if (el) {
              const s = el.getElementsByTagName('a:srgbClr')[0];
              const y = el.getElementsByTagName('a:sysClr')[0];

              if (s && s.getAttribute('val')) {
                themeColors[idx] = '#' + s.getAttribute('val');
              } else if (y && y.getAttribute('lastClr')) {
                themeColors[idx] = '#' + y.getAttribute('lastClr');
              }
            }
          });
      }
    }

    const styleFillHex = [];
    const stylesFile = zip.file('xl/styles.xml');

    if (stylesFile) {
      const x = parser.parseFromString(await stylesFile.async('text'), 'text/xml');

      const fills = Array.from(x.querySelectorAll('fills > fill')).map(fill => {
        const pf = fill.querySelector('patternFill');
        if (!pf) return null;

        const pt = pf.getAttribute('patternType');
        if (!pt || pt === 'none') return null;

        const fg = pf.querySelector('fgColor');
        if (!fg) return null;

        let hex = null;

        if (fg.getAttribute('rgb')) {
          const rgb = fg.getAttribute('rgb');
          hex = '#' + (rgb.length === 8 ? rgb.slice(2) : rgb);
        } else if (fg.getAttribute('theme') !== null) {
          const t = parseInt(fg.getAttribute('theme'), 10);
          if (themeColors[t]) hex = themeColors[t];
        } else if (fg.getAttribute('indexed') !== null) {
          const i = parseInt(fg.getAttribute('indexed'), 10);
          if (INDEXED_COLORS[i]) hex = INDEXED_COLORS[i];
        }

        if (hex && fg.getAttribute('tint') !== null) {
          hex = applyTint(hex, parseFloat(fg.getAttribute('tint')));
        }

        return hex;
      });

      Array.from(x.querySelectorAll('cellXfs > xf')).forEach((xf, idx) => {
        styleFillHex[idx] = fills[parseInt(xf.getAttribute('fillId') || '0', 10)] || null;
      });
    }

    const sx = parser.parseFromString(await sheetFile.async('text'), 'text/xml');
    const cells = sx.getElementsByTagName('c');

    for (let i = 0; i < cells.length; i++) {
      const c = cells[i];
      const s = c.getAttribute('s');
      const r = c.getAttribute('r');

      if (s !== null && r) {
        const hex = styleFillHex[parseInt(s, 10)];
        if (hex && hex.toUpperCase() !== '#FFFFFF' && hex.toUpperCase() !== '#000000') {
          colorMap[r] = hex;
        }
      }
    }
  } catch (e) {
    console.warn('colors ' + sheetName, e);
  }

  return colorMap;
}

async function patchExcelFile(buffer, changes, sheetName) {
  const zip = await JSZip.loadAsync(buffer);
  const parser = new DOMParser();
  const serializer = new XMLSerializer();

  const sheetPath = await getSheetFilePath(zip, sheetName);
  const sheetFile = zip.file(sheetPath);

  if (!sheetFile) throw new Error('Лист не найден: ' + sheetPath);

  const xml = parser.parseFromString(await sheetFile.async('text'), 'text/xml');
  const sheetData = xml.getElementsByTagName('sheetData')[0];

  if (!sheetData) throw new Error('sheetData не найден');

  const ns = xml.documentElement.namespaceURI || 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
  const createEl = t => xml.createElementNS(ns, t);

  const rowMap = new Map();
  const rows = sheetData.getElementsByTagName('row');

  for (let i = 0; i < rows.length; i++) {
    rowMap.set(parseInt(rows[i].getAttribute('r'), 10), rows[i]);
  }

  for (const [ref, value] of changes) {
    const m = ref.match(/^([A-Z]+)(\d+)$/);
    if (!m) continue;

    const newCol = colNameToIndex(m[1]);
    const rowNum = parseInt(m[2], 10);

    let row = rowMap.get(rowNum);

    if (!row) {
      row = createEl('row');
      row.setAttribute('r', String(rowNum));

      let ins = false;
      for (let i = 0; i < sheetData.children.length; i++) {
        if (rowNum < parseInt(sheetData.children[i].getAttribute('r'), 10)) {
          sheetData.insertBefore(row, sheetData.children[i]);
          ins = true;
          break;
        }
      }

      if (!ins) sheetData.appendChild(row);
      rowMap.set(rowNum, row);
    }

    const cells = row.getElementsByTagName('c');
    let cell = null;

    for (let i = 0; i < cells.length; i++) {
      if (cells[i].getAttribute('r') === ref) {
        cell = cells[i];
        break;
      }
    }

    if (!cell) {
      cell = createEl('c');
      cell.setAttribute('r', ref);

      let ins = false;
      for (let i = 0; i < cells.length; i++) {
        const mm = (cells[i].getAttribute('r') || '').match(/^([A-Z]+)/);
        if (mm && newCol < colNameToIndex(mm[1])) {
          row.insertBefore(cell, cells[i]);
          ins = true;
          break;
        }
      }

      if (!ins) row.appendChild(cell);
    }

    cell.removeAttribute('t');

    const f = cell.getElementsByTagNameNS(ns, 'f')[0];
    if (f) f.remove();

    const is = cell.getElementsByTagNameNS(ns, 'is')[0];
    if (is) is.remove();

    let v = cell.getElementsByTagNameNS(ns, 'v')[0];
    if (!v) {
      v = createEl('v');
      cell.appendChild(v);
    }

    v.textContent = String(value);
  }

  zip.file(sheetPath, serializer.serializeToString(xml));

  if (zip.file('xl/calcChain.xml')) {
    zip.remove('xl/calcChain.xml');

    const ct = zip.file('[Content_Types].xml');
    if (ct) {
      const x = parser.parseFromString(await ct.async('text'), 'text/xml');
      Array.from(x.getElementsByTagName('Override')).forEach(o => {
        if (o.getAttribute('PartName') === '/xl/calcChain.xml') {
          o.parentNode.removeChild(o);
        }
      });
      zip.file('[Content_Types].xml', serializer.serializeToString(x));
    }

    const rf = zip.file('xl/_rels/workbook.xml.rels');
    if (rf) {
      const x = parser.parseFromString(await rf.async('text'), 'text/xml');
      Array.from(x.getElementsByTagName('Relationship')).forEach(r => {
        if ((r.getAttribute('Target') || '') === 'calcChain.xml') {
          r.parentNode.removeChild(r);
        }
      });
      zip.file('xl/_rels/workbook.xml.rels', serializer.serializeToString(x));
    }
  }

  return await zip.generateAsync({ type: 'arraybuffer', compression: 'DEFLATE' });
}

function readExcelFile(file, handle) {
  if (typeof XLSX === 'undefined' || typeof JSZip === 'undefined') {
    showError('Библиотеки не загружены. Запустите START.bat и обновите страницу.');
    return;
  }

  if (file.size > 50 * 1024 * 1024) {
    showError('Файл слишком большой (максимум 50 МБ)');
    return;
  }

  const reader = new FileReader();

  reader.onload = async e => {
    try {
      const buffer = e.target.result;
      S.originalBuffer = buffer;

      S.wb = XLSX.read(new Uint8Array(buffer), { type: 'array', cellDates: true });

      if (!S.wb.SheetNames || !S.wb.SheetNames.length) {
        showError('Файл не содержит листов');
        return;
      }

      S.fileName = file.name;
      S.saveHandle = handle || null;
      S.demoMode = false;

      S.pendingChanges = {};
      S.originalCells = {};
      S.undoStack = [];
      S.summaryFilter = null;

      S.currentSheet = S.wb.SheetNames[0];
      S.sheetColors = {};

      for (const n of S.wb.SheetNames) {
        S.sheetColors[n] = await extractExcelColorsForSheet(buffer, n);
      }

      S.cellColors = S.sheetColors[S.currentSheet] || {};

      rebuildSheetSelect();
      parseSheet();
      checkDraft();

      idbPut('lastFile', { name: S.fileName, buffer, timestamp: Date.now() }).catch(() => {});

      closeDrawer();
    } catch (err) {
      console.error(err);
      showToast('Ошибка чтения: ' + err.message, 'error');
      showError('Ошибка чтения файла: ' + err.message);
    }
  };

  reader.onerror = () => showError('Не удалось прочитать файл');
  reader.readAsArrayBuffer(file);
}

/* ================= ЛИСТ / ПАРСИНГ ================= */
function rebuildSheetSelect() {
  const sel = document.getElementById('filterSheet');
  if (!sel || !S.wb) return;

  sel.innerHTML = S.wb.SheetNames.map(n => {
    const c = S.pendingChanges[n] ? S.pendingChanges[n].size : 0;
    const label = c > 0 ? `${n} (${c})` : n;
    return `<option value="${esc(n)}" ${n === S.currentSheet ? 'selected' : ''}>${esc(label)}</option>`;
  }).join('');

  sel.onchange = e => {
    if (e.target.value !== S.currentSheet) doSwitchSheet(e.target.value);
  };
}

function doSwitchSheet(name) {
  S.currentSheet = name;
  S.cellColors = S.sheetColors[name] || {};

  S.filterSection = '';
  S.filterColIndex = -1;
  S.filterWorkIdxs.clear();
  S.summaryFilter = null;

  S.selSection = '';
  S.selColIndex = -1;
  S.selWorkIdx = 0;

  parseSheet();
  applyPendingToCurrentSheet();
  renderMatrix();
  rebuildSheetSelect();
  updateSummary();
}

function applyPendingToCurrentSheet() {
  const changes = getSheetPendingChanges(S.currentSheet);
  if (!changes.size) return;

  const ws = S.wb.Sheets[S.currentSheet];

  S.works.forEach((w, wi) => S.columns.forEach(col => {
    const ref = XLSX.utils.encode_cell({ r: w.myRow, c: col.colIndex });

    if (changes.has(ref)) {
      const v = changes.get(ref);
      S.myCells[wi + '_' + col.colIndex] = v;

      if (ws[ref]) {
        ws[ref].v = v;
        ws[ref].t = 'n';
        delete ws[ref].w;
      } else {
        ws[ref] = { v, t: 'n' };
      }
    }
  }));
}

function parseSheet() {
  const ws = S.wb.Sheets[S.currentSheet];

  if (!ws || !ws['!ref']) {
    showError('Лист пуст или структура не распознана');
    return;
  }

  const range = XLSX.utils.decode_range(ws['!ref']);

  S.columns = [];
  S.uniqueSections = [];
  let lastSection = '';

  for (let c = 3; c <= range.e.c; c++) {
    const sc = ws[XLSX.utils.encode_cell({ r: 0, c })];
    const sv = sc && sc.v != null ? String(sc.v).trim() : '';
    if (sv) lastSection = sv;

    const fc = ws[XLSX.utils.encode_cell({ r: 2, c })];
    const fv = fc && fc.v != null ? String(fc.v).trim() : '';

    if (!fv || fv.toLowerCase() === 'статус' || !lastSection) continue;

    if (!S.uniqueSections.includes(lastSection)) S.uniqueSections.push(lastSection);

    S.columns.push({
      colIndex: c,
      section: lastSection,
      floor: fv
    });
  }

  if (!S.columns.length) {
    showError('Не удалось распознать секции и этажи');
    return;
  }

  S.works = [];
  S.myCells = {};
  S.foremanCells = {};

  for (let r = 3; r <= range.e.r; r++) {
    const a = ws[XLSX.utils.encode_cell({ r, c: 0 })];
    const av = a && a.v != null ? String(a.v).trim().toUpperCase() : '';

    if (av.includes('ПРОРАБ') && r > 3) {
      const workRef = XLSX.utils.encode_cell({ r: r - 1, c: 2 });
      const wc = ws[workRef];

      const name = wc && wc.v != null ? String(wc.v).trim() : ('Работа #' + (S.works.length + 1));
      const wi = S.works.length;

      S.works.push({
        name,
        myRow: r - 1,
        foremanRow: r,
        rowColor: S.cellColors[workRef] || null
      });

      S.columns.forEach(col => {
        const key = wi + '_' + col.colIndex;

        const mc = ws[XLSX.utils.encode_cell({ r: r - 1, c: col.colIndex })];
        S.myCells[key] = mc && mc.v != null ? mc.v : null;

        const fc = ws[XLSX.utils.encode_cell({ r, c: col.colIndex })];
        S.foremanCells[key] = fc && fc.v != null ? fc.v : null;
      });
    }
  }

  if (!S.works.length) {
    showError('Не найдены строки с отметкой "ПРОРАБ"');
    return;
  }

  S.filterSection = S.uniqueSections[0];

  const first = S.columns.filter(c => c.section === S.filterSection);

  S.selSection = S.filterSection;
  S.selColIndex = first.length ? first[0].colIndex : -1;
  S.selWorkIdx = 0;

  showApp();
  initFilters();
  updateSelectionPanel();
  renderMatrix();
  rebuildMultiselects();
  renderDrawerMass();
  updateSummary();
}

function showApp() {
  document.getElementById('uploadScreen').classList.add('hidden');
  document.getElementById('appScreen').classList.remove('hidden');
  document.getElementById('fileName').textContent = S.fileName;
  document.getElementById('tablePlaceholder').classList.add('hidden');
  document.getElementById('tableContainer').classList.remove('hidden');
}

function toggleFilters() {
  const c = document.getElementById('filtersContent');
  const a = document.getElementById('filtersArrow');

  c.classList.toggle('open');
  a.style.transform = c.classList.contains('open') ? 'rotate(180deg)' : '';
}

function initFilters() {
  const fSec = document.getElementById('filterSection');

  fSec.innerHTML = S.uniqueSections.map(s => `<option value="${esc(s)}">${esc(s)}</option>`).join('');
  fSec.value = S.filterSection;

  rebuildFloorSelects();
  rebuildWorksDropdown();

  fSec.onchange = e => {
    S.filterSection = e.target.value;
    S.filterColIndex = -1;
    S.selSection = S.filterSection;
    S.summaryFilter = null;

    const sc = S.columns.filter(c => c.section === S.filterSection);
    if (sc.length) S.selColIndex = sc[0].colIndex;

    rebuildFloorSelects();
    rebuildWorksDropdown();
    updateSelectionPanel();
    renderMatrix();
    rebuildMultiselects();
    renderDrawerMass();
    updateSummary();
  };

  document.getElementById('filterFloor').onchange = e => {
    S.filterColIndex = parseInt(e.target.value, 10);
    if (S.filterColIndex !== -1) {
      S.selColIndex = S.filterColIndex;
      updateSelectionPanel();
    }
    renderMatrix();
    updateSummary();
  };

  document.getElementById('hideAccepted').onchange = e => {
    S.hideAccepted = e.target.checked;
    renderMatrix();
    rebuildMultiselects();
    rebuildWorksDropdown();
    renderDrawerMass();
    updateSummary();
  };
}

function rebuildFloorSelects() {
  const f = document.getElementById('filterFloor');
  const sc = S.columns.filter(c => c.section === S.filterSection);

  f.innerHTML = `<option value="-1">Все этажи (${sc.length})</option>` +
    sc.map(c => `<option value="${c.colIndex}">${esc(c.floor)}</option>`).join('');

  f.value = String(S.filterColIndex);
}

function getVisibleWorks() {
  let vw = S.works.map((w, idx) => Object.assign({}, w, { idx }));

  const sc = S.columns.filter(c => c.section === S.filterSection);

  if (S.hideAccepted) {
    vw = vw.filter(w => !shouldHideRow(w, sc));
  }

  return vw;
}

function rebuildWorksDropdown() {
  const cont = document.getElementById('filterWorkOptions');
  const vw = getVisibleWorks();

  cont.innerHTML = vw.length
    ? vw.map(w => `
      <label class="multiselect-option" data-idx="${w.idx}" data-name="${esc(w.name).toLowerCase()}">
        <input type="checkbox" value="${w.idx}" ${S.filterWorkIdxs.has(w.idx) ? 'checked' : ''} onchange="handleFilterWorkChange(this)">
        <span class="multiselect-option-label">${esc(w.name)}</span>
      </label>
    `).join('')
    : '<div class="px-3 py-2 text-xs text-surface-400 italic">Нет работ</div>';

  updateWorksDropdownLabel();
}

function handleFilterWorkChange(cb) {
  const i = parseInt(cb.value, 10);
  cb.checked ? S.filterWorkIdxs.add(i) : S.filterWorkIdxs.delete(i);

  renderMatrix();
  updateSummary();
  updateWorksDropdownLabel();
}

function filterWorkAll(all, e) {
  e.stopPropagation();

  document.querySelectorAll('#filterWorkOptions input[type="checkbox"]').forEach(cb => {
    const i = parseInt(cb.value, 10);
    cb.checked = all;
    all ? S.filterWorkIdxs.add(i) : S.filterWorkIdxs.delete(i);
  });

  renderMatrix();
  updateSummary();
  updateWorksDropdownLabel();
}

function updateWorksDropdownLabel() {
  const l = document.querySelector('#filterWorkMs .multiselect-label');
  const n = S.filterWorkIdxs.size;

  l.textContent = n === 0
    ? 'Все работы'
    : n === 1
      ? ((S.works[[...S.filterWorkIdxs][0]] || {}).name || '1 работа')
      : `Выбрано: ${n}`;
}

function filterWorksDropdown(q) {
  q = q.trim().toLowerCase();
  let vis = 0;

  document.querySelectorAll('#filterWorkOptions .multiselect-option').forEach(o => {
    const ok = !q || (o.getAttribute('data-name') || '').includes(q);
    o.classList.toggle('hidden', !ok);
    if (ok) vis++;
  });

  document.getElementById('filterWorkEmpty').classList.toggle('hidden', vis > 0);
}

function shouldHideRow(work, cols) {
  if (!S.hideAccepted || !cols.length) return false;

  let allFrEmpty = true;
  let allMyEmpty = true;
  let allMatch = true;
  let hasFr = false;

  for (const col of cols) {
    const k = work.idx + '_' + col.colIndex;
    const my = S.myCells[k];
    const fr = S.foremanCells[k];

    if (hasAnyContent(fr)) {
      allFrEmpty = false;

      const mn = parseNum(my);
      const fn = parseNum(fr);

      if (fn !== null) {
        hasFr = true;
        if (mn === null || mn !== fn) allMatch = false;
      }
    }

    if (hasAnyContent(my)) allMyEmpty = false;
  }

  return (allFrEmpty && allMyEmpty) || (allMatch && hasFr);
}

function countMismatches(work, cols) {
  let n = 0;

  for (const col of cols) {
    const k = work.idx + '_' + col.colIndex;

    const mn = parseNum(S.myCells[k]);
    const fn = parseNum(S.foremanCells[k]);

    if (fn !== null) {
      if (mn === null) {
        n++;
      } else if ((fn === 100 && mn < 100) || Math.abs(mn - fn) > 15) {
        n++;
      }
    }
  }

  return n;
}

function isWorkDone(work, cols) {
  let ok = true;
  let has = false;

  for (const col of cols) {
    const k = work.idx + '_' + col.colIndex;
    const fr = S.foremanCells[k];

    if (hasAnyContent(fr)) {
      const mn = parseNum(S.myCells[k]);
      const fn = parseNum(fr);

      if (fn !== null) {
        has = true;
        if (mn !== fn) {
          ok = false;
          break;
        }
      }
    }
  }

  return ok && has;
}

/* ================= ШТОРКА ВЫБОРА РАБОТ (телефон) ================= */
function openWorkSheet() {
  document.getElementById('workSheetSearch').value = '';
  renderWorkSheet();
  document.getElementById('workSheetBackdrop').classList.add('open');
  document.getElementById('workSheet').classList.add('open');
}

function closeWorkSheet() {
  document.getElementById('workSheetBackdrop').classList.remove('open');
  document.getElementById('workSheet').classList.remove('open');
}

function renderWorkSheet() {
  const vw = getVisibleWorks();
  const box = document.getElementById('workSheetOptions');

  box.innerHTML = vw.length
    ? vw.map(w => `
      <label class="multiselect-option" data-name="${esc(w.name).toLowerCase()}">
        <input type="checkbox" value="${w.idx}" ${S.filterWorkIdxs.has(w.idx) ? 'checked' : ''} onchange="handleWorkSheetChange(this)">
        <span class="multiselect-option-label">${esc(w.name)}</span>
      </label>
    `).join('')
    : '<div class="p-3 text-sm text-surface-400 italic">Нет работ</div>';

  document.getElementById('workSheetCount').textContent = S.filterWorkIdxs.size ? `(${S.filterWorkIdxs.size})` : '';
}

function handleWorkSheetChange(cb) {
  const i = parseInt(cb.value, 10);
  cb.checked ? S.filterWorkIdxs.add(i) : S.filterWorkIdxs.delete(i);

  document.getElementById('workSheetCount').textContent = S.filterWorkIdxs.size ? `(${S.filterWorkIdxs.size})` : '';

  renderMatrix();
  updateSummary();
  updateWorksDropdownLabel();
  rebuildWorksDropdown();
  renderDrawerMass();
}

function workSheetAll(all) {
  document.querySelectorAll('#workSheetOptions input[type="checkbox"]').forEach(cb => {
    if (cb.closest('.multiselect-option').classList.contains('hidden')) return;

    const i = parseInt(cb.value, 10);
    cb.checked = all;
    all ? S.filterWorkIdxs.add(i) : S.filterWorkIdxs.delete(i);
  });

  handleWorkSheetRefresh();
}

function handleWorkSheetRefresh() {
  document.getElementById('workSheetCount').textContent = S.filterWorkIdxs.size ? `(${S.filterWorkIdxs.size})` : '';

  renderMatrix();
  updateSummary();
  updateWorksDropdownLabel();
  rebuildWorksDropdown();
  renderDrawerMass();
}

function filterWorkSheet(q) {
  q = q.trim().toLowerCase();

  document.querySelectorAll('#workSheetOptions .multiselect-option').forEach(o => {
    o.classList.toggle('hidden', !!q && !(o.getAttribute('data-name') || '').includes(q));
  });
}

/* ================= МАТРИЦА ================= */
function renderMatrix() {
  const sec = S.filterSection;

  let cols = S.columns.filter(c => c.section === sec);
  if (S.filterColIndex !== -1) cols = cols.filter(c => c.colIndex === S.filterColIndex);

  let vw = S.works.map((w, idx) => Object.assign({}, w, { idx }));

  if (S.filterWorkIdxs.size) vw = vw.filter(w => S.filterWorkIdxs.has(w.idx));
  if (S.summaryFilter === 'mismatch') vw = vw.filter(w => countMismatches(w, cols) > 0);

  const thead = document.getElementById('matrixHead');
  const tbody = document.getElementById('matrixBody');

  let th = '<tr class="bg-surface-100 dark:bg-surface-700"><th class="sticky-col px-2 md:px-3 py-2.5 text-xs font-semibold text-surface-600 dark:text-surface-300 border-b border-r border-surface-200 dark:border-surface-600 bg-surface-100 dark:bg-surface-700 text-left min-w-[160px] md:min-w-[240px] max-w-[200px] md:max-w-[320px]">Вид работы</th>';

  cols.forEach(c => {
    th += `<th class="px-1 md:px-2 py-2.5 text-xs font-bold text-surface-700 dark:text-surface-200 border-b border-r border-surface-200 dark:border-surface-600 bg-surface-100 dark:bg-surface-700 whitespace-nowrap min-w-[60px] md:min-w-[68px] text-center">${esc(c.floor)}</th>`;
  });

  thead.innerHTML = th + '</tr>';

  let html = '';
  let shown = 0;

  for (const work of vw) {
    if (shouldHideRow(work, cols)) continue;

    const mism = countMismatches(work, cols);
    const rowSel = cols.some(c => S.selColIndex === c.colIndex && S.selWorkIdx === work.idx);

    let cells = '';

    for (const col of cols) {
      const k = work.idx + '_' + col.colIndex;
      const my = S.myCells[k];
      const fr = S.foremanCells[k];

      const mn = parseNum(my);
      const frHas = hasAnyContent(fr);

      const myRef = XLSX.utils.encode_cell({ r: work.myRow, c: col.colIndex });
      const frRef = XLSX.utils.encode_cell({ r: work.foremanRow, c: col.colIndex });

      const hex = S.cellColors[myRef] || S.cellColors[frRef] || null;
      const isSel = S.selColIndex === col.colIndex && S.selWorkIdx === work.idx;

      const data = `data-work="${work.idx}" data-col="${col.colIndex}" data-sec="${esc(col.section)}"${frHas ? '' : ' data-locked="1"'}`;

      if (!frHas) {
        cells += `<td ${data} class="px-1 py-1 border-r border-surface-200/70 dark:border-surface-700/50 bg-surface-100/40 dark:bg-surface-800/20 cell-locked cursor-not-allowed ${isSel ? 'ring-2 ring-brand-600 ring-inset' : ''}"><div class="cell-content opacity-50"><span class="text-surface-300 dark:text-surface-600 text-xs">—</span></div></td>`;
      } else {
        const myStr = (my != null && my !== '') ? formatVal(my) : '—';
        const frStr = formatVal(fr);

        let style = '';
        let bg = 'bg-white dark:bg-surface-800/40';
        let myColor = 'text-surface-800 dark:text-surface-200';
        let frBadge = 'bg-surface-200/80 dark:bg-surface-700 text-surface-700 dark:text-surface-300';

        if (hex) {
          style = `style="background-color:${hex};"`;
          bg = 'cell-with-color';

          const light = isLightColor(hex);
          myColor = light ? 'text-slate-900' : 'text-white';
          frBadge = light ? 'bg-black/10 text-slate-900' : 'bg-white/20 text-white';
        } else if (mn === null) {
          myColor = 'text-surface-400 dark:text-surface-500';
        }

        if (isSel) bg += ' ring-2 ring-brand-600 ring-inset';

        cells += `<td ${data} ${style} class="px-1 py-1 border-r border-surface-200/70 dark:border-surface-700/50 ${bg} cursor-pointer"><div class="cell-content"><span class="cell-my ${myColor}">${myStr}</span><span class="cell-foreman ${frBadge}">Пр: ${frStr}</span></div></td>`;
      }
    }

    shown++;

    let stickyStyle = '';
    let stickyClass = '';

    if (work.rowColor) {
      stickyStyle = `style="background-color:${work.rowColor};color:${isLightColor(work.rowColor) ? '#0f172a' : '#ffffff'};"`;
      stickyClass = '';
    } else if (rowSel) {
      stickyClass = 'bg-brand-200 dark:bg-brand-800 text-surface-900 dark:text-surface-100';
    } else {
      stickyClass = 'bg-white dark:bg-surface-800 text-surface-800 dark:text-surface-200';
    }

    const mismBadge = mism > 0
      ? `<span class="inline-block ml-1 text-amber-500" title="${mism} расхождений">⚠ ${mism}</span>`
      : '';

    html += `<tr class="border-b border-surface-200/70 dark:border-surface-700/60 table-row-hover ${rowSel ? 'bg-brand-100/50 dark:bg-brand-900/30' : ''}"><td ${stickyStyle} class="sticky-col px-2 md:px-3 py-2 text-xs md:text-sm font-medium border-r border-surface-200 dark:border-surface-700 ${stickyClass} leading-snug">${esc(work.name)}${mismBadge}</td>${cells}</tr>`;
  }

  tbody.innerHTML = shown
    ? html
    : `<tr><td colspan="${cols.length + 1}" class="p-8 text-center text-sm text-surface-400">Нет работ, подходящих под фильтры</td></tr>`;
}

/* Взаимодействие с матрицей: клик + долгое нажатие */
function initMatrixInteractions() {
  const tbody = document.getElementById('matrixBody');

  let lp = null;
  let suppressUntil = 0;

  tbody.addEventListener('click', e => {
    if (Date.now() < suppressUntil) return;

    const td = e.target.closest('td[data-work]');
    if (!td) return;

    const w = +td.dataset.work;
    const c = +td.dataset.col;
    const s = td.dataset.sec;

    td.dataset.locked ? selectLockedCell(s, c, w) : selectCell(s, c, w);
  });

  tbody.addEventListener('pointerdown', e => {
    if (e.pointerType === 'mouse') return;

    const td = e.target.closest('td[data-work]:not([data-locked])');
    if (!td) return;

    td.classList.add('cell-pressing');

    lp = {
      x: e.clientX,
      y: e.clientY,
      td,
      timer: setTimeout(() => {
        lp = null;
        td.classList.remove('cell-pressing');
        suppressUntil = Date.now() + 600;

        const w = +td.dataset.work;
        const c = +td.dataset.col;
        const s = td.dataset.sec;

        S.selSection = s;
        S.selColIndex = c;
        S.selWorkIdx = w;

        updateSelectionPanel();
        acceptForemanValue();
        renderMatrix();

        showToast('Подтверждено: ' + formatVal(S.myCells[w + '_' + c]), 'success', {
          label: 'Отменить',
          fn: undoLastChange
        });

        vibrate(40);
      }, 550)
    };
  });

  const cancel = () => {
    if (lp) {
      clearTimeout(lp.timer);
      lp.td.classList.remove('cell-pressing');
      lp = null;
    }
  };

  tbody.addEventListener('pointerup', cancel);
  tbody.addEventListener('pointercancel', cancel);

  tbody.addEventListener('pointermove', e => {
    if (lp && (Math.abs(e.clientX - lp.x) > 12 || Math.abs(e.clientY - lp.y) > 12)) cancel();
  });

  tbody.addEventListener('contextmenu', e => {
    if (e.target.closest('td[data-work]')) e.preventDefault();
  });
}

/* ================= ВЫБОР / ПАНЕЛЬ ФАКТА ================= */
function selectCell(sec, col, wi) {
  S.selSection = sec;
  S.selColIndex = col;
  S.selWorkIdx = wi;

  updateSelectionPanel();
  renderMatrix();

  if (isMobile()) openBottomSheet();
}

function selectLockedCell(sec, col, wi) {
  S.selSection = sec;
  S.selColIndex = col;
  S.selWorkIdx = wi;

  updateSelectionPanel();
  renderMatrix();

  showToast('Нет отметки прораба — ввод невозможен', 'warning');
}

function updateSelectionPanel() {
  const col = S.columns.find(c => c.colIndex === S.selColIndex);
  const work = S.works[S.selWorkIdx];

  const set = (id, v) => {
    const el = document.getElementById(id);
    if (el) el.textContent = v;
  };

  set('selSection', S.selSection || '—');
  set('selFloor', col ? col.floor : '—');
  set('selWork', work ? work.name : '—');

  const myD = document.getElementById('myValueDisplay');
  const frD = document.getElementById('foremanValueDisplay');
  const lock = document.getElementById('lockedWarning');

  const acc = document.getElementById('acceptBtn');
  const sub = document.getElementById('submitBtn');
  const inp = document.getElementById('manualValue');

  if (col && work) {
    const k = S.selWorkIdx + '_' + S.selColIndex;
    const my = S.myCells[k];
    const fr = S.foremanCells[k];
    const frHas = hasAnyContent(fr);

    if (myD) myD.textContent = (my != null && my !== '') ? formatVal(my) : '—';
    if (frD) frD.textContent = frHas ? formatVal(fr) : '—';

    S.isCellLocked = !frHas;
  } else {
    if (myD) myD.textContent = '—';
    if (frD) frD.textContent = '—';

    S.isCellLocked = true;
  }

  if (lock) lock.classList.toggle('hidden', !S.isCellLocked);

  [acc, sub].forEach(b => {
    if (b) {
      b.disabled = S.isCellLocked;
      b.classList.toggle('opacity-50', S.isCellLocked);
    }
  });

  if (inp) {
    inp.disabled = S.isCellLocked;
    inp.classList.toggle('opacity-50', S.isCellLocked);
  }

  if (isMobile() && document.getElementById('bottomSheet').classList.contains('open')) {
    renderBottomSheetContent();
  }

  updateUndoButton();
}

function updateUndoButton() {
  const b = document.getElementById('undoBtn');
  const ready = S.undoStack.length > 0;

  b.disabled = !ready;
  b.classList.toggle('undo-ready', ready);

  b.title = ready ? 'Отменить последнее действие (Ctrl+Z)' : 'Нечего отменять';
}

function updateSaveIndicator() {
  const n = getTotalPendingChanges();

  const badge = document.getElementById('saveBadge');
  const btn = document.getElementById('saveBtn');

  if (badge) {
    badge.textContent = n > 99 ? '99+' : String(n);
    badge.classList.toggle('hidden', n === 0);
  }

  if (btn) {
    btn.title = n > 0 ? `Сохранить — ${n} несохранённых правок (Ctrl+S)` : 'Сохранить (Ctrl+S)';
  }
}

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
  const col = S.columns.find(c => c.colIndex === S.selColIndex);
  const work = S.works[S.selWorkIdx];

  const k = S.selWorkIdx + '_' + S.selColIndex;
  const my = S.myCells[k];
  const fr = S.foremanCells[k];

  const frHas = hasAnyContent(fr);
  const frNum = parseNum(fr);

  const lockedHtml = !frHas
    ? `<div class="mb-4 p-3 bg-amber-50 dark:bg-amber-900/20 border border-amber-200 dark:border-amber-800 rounded-xl text-sm text-amber-800 dark:text-amber-300">🔒 Нет отметки прораба — ввод невозможен</div>`
    : `
      <button onclick="acceptForemanValue(); closeBottomSheet();" class="w-full flex items-center justify-center gap-2 px-4 py-4 btn-gradient text-white text-base font-bold rounded-xl shadow-sm mb-3">
        <svg class="w-5 h-5" fill="none" stroke="currentColor" stroke-width="2.5" viewBox="0 0 24 24">
          <path stroke-linecap="round" stroke-linejoin="round" d="M4.5 12.75l6 6 9-13.5"/>
        </svg>
        Факт верный (${frNum !== null ? frNum + '%' : '—'})
      </button>

      <div class="mb-1">
        <label class="block text-sm font-medium text-surface-500 dark:text-surface-400 mb-2">Ввести значение (0–100)</label>
        <div class="flex gap-2">
          <input type="number" id="mobileManualValue" min="0" max="100" step="1" placeholder="0 — 100" inputmode="numeric" class="flex-1 px-4 py-3 text-base border border-surface-300 dark:border-surface-600 rounded-xl bg-surface-50 dark:bg-surface-700 text-surface-900 dark:text-white focus:ring-2 focus:ring-brand-500 outline-none">
          <button onclick="submitMobileManualValue()" class="px-5 py-3 btn-brand text-white text-base font-bold rounded-xl">OK</button>
        </div>
      </div>
    `;

  document.getElementById('bottomSheetContent').innerHTML = `
    <div class="flex items-center justify-between mb-4">
      <h2 class="text-lg font-bold text-surface-900 dark:text-white">Ввод факта</h2>
      <button onclick="closeBottomSheet()" class="hdr-btn">
        <svg class="w-5 h-5" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24">
          <path stroke-linecap="round" d="M6 18L18 6M6 6l12 12"/>
        </svg>
      </button>
    </div>

    <div class="p-3 bg-surface-50 dark:bg-surface-700/50 rounded-2xl border border-surface-200 dark:border-surface-600 mb-4 card-shadow">
      <div class="text-xs mb-2">
        <div>
          <span class="text-surface-500 dark:text-surface-400">Секция:</span>
          <span class="font-semibold text-surface-900 dark:text-white">${esc(S.selSection || '—')}</span>
        </div>
        <div>
          <span class="text-surface-500 dark:text-surface-400">Этаж:</span>
          <span class="font-semibold text-surface-900 dark:text-white">${col ? esc(col.floor) : '—'}</span>
        </div>
        <div class="mt-1">
          <span class="text-surface-500 dark:text-surface-400">Работа:</span>
          <div class="font-medium text-surface-900 dark:text-white leading-snug">${work ? esc(work.name) : '—'}</div>
        </div>
      </div>

      <div class="grid grid-cols-2 gap-2">
        <div>
          <div class="text-xs text-surface-500 dark:text-surface-400 mb-1">Мой факт</div>
          <div class="px-3 py-2 text-lg font-bold font-mono bg-white dark:bg-surface-800 border border-surface-200 dark:border-surface-600 rounded-lg text-center">
            ${(my != null && my !== '') ? formatVal(my) : '—'}
          </div>
        </div>
        <div>
          <div class="text-xs text-surface-500 dark:text-surface-400 mb-1">Прораб</div>
          <div class="px-3 py-2 text-lg font-bold font-mono bg-white dark:bg-surface-800 border border-surface-200 dark:border-surface-600 rounded-lg text-center">
            ${frHas ? formatVal(fr) : '—'}
          </div>
        </div>
      </div>
    </div>

    ${lockedHtml}
  `;

  const mi = document.getElementById('mobileManualValue');
  if (mi) {
    mi.addEventListener('keydown', e => {
      if (e.key === 'Enter') submitMobileManualValue();
    });
  }
}

function submitMobileManualValue() {
  const inp = document.getElementById('mobileManualValue');
  if (!inp) return;

  const v = parseFloat(inp.value);

  if (isNaN(v) || v < 0 || v > 100) {
    showToast('Введите число от 0 до 100', 'error');
    return;
  }

  writeValueToCell(clampPct(v));
  vibrate(30);
  showToast('Записано: ' + clampPct(v) + '%', 'success');

  renderMatrix();
  renderBottomSheetContent();
}

function submitManualValue() {
  if (S.isCellLocked) {
    showFormStatus('Нет отметки прораба', 'error');
    return;
  }

  if (S.selColIndex === -1 || !S.works[S.selWorkIdx]) {
    showFormStatus('Сначала выберите ячейку', 'error');
    return;
  }

  const inp = document.getElementById('manualValue');
  const v = parseFloat(inp.value);

  if (isNaN(v) || v < 0 || v > 100) {
    showFormStatus('Введите число от 0 до 100', 'error');
    return;
  }

  writeValueToCell(clampPct(v));
  inp.value = '';

  showFormStatus('Записано: ' + clampPct(v) + '%', 'success');
  renderMatrix();
}

function acceptForemanValue() {
  if (S.isCellLocked) {
    showFormStatus('Нет отметки прораба', 'error');
    return;
  }

  if (S.selColIndex === -1 || !S.works[S.selWorkIdx]) {
    showFormStatus('Сначала выберите ячейку', 'error');
    return;
  }

  const fr = parseNum(S.foremanCells[S.selWorkIdx + '_' + S.selColIndex]);

  if (fr === null) {
    showFormStatus('У прораба нет значения', 'error');
    return;
  }

  writeValueToCell(clampPct(fr));
  showFormStatus('Подтверждено: ' + formatVal(fr), 'success');
  vibrate(30);
  renderMatrix();
}

/* ================= ЗАПИСЬ ================= */
function writeValueToCell(val) {
  const work = S.works[S.selWorkIdx];
  const col = S.selColIndex;

  const ref = XLSX.utils.encode_cell({ r: work.myRow, c: col });
  const old = S.myCells[S.selWorkIdx + '_' + col];

  captureOriginal(S.currentSheet, ref);

  const snap = getSheetOriginals(S.currentSheet).get(ref);

  const entry = {
    type: 'single',
    sheetName: S.currentSheet,
    cellRef: ref,
    workIdx: S.selWorkIdx,
    colIndex: col,
    oldValue: old,
    originalSnapshot: snap,
    timestamp: Date.now()
  };

  S.undoStack.push(entry);
  if (S.undoStack.length > MAX_UNDO_STEPS) S.undoStack.shift();

  getSheetPendingChanges(S.currentSheet).set(ref, val);

  const ws = S.wb.Sheets[S.currentSheet];

  if (ws[ref]) {
    ws[ref].v = val;
    ws[ref].t = 'n';
    delete ws[ref].w;
  } else {
    ws[ref] = { v: val, t: 'n' };
  }

  S.myCells[S.selWorkIdx + '_' + col] = val;

  updateSelectionPanel();
  rebuildSheetSelect();
  updateSummary();
  scheduleAutoSave();
}

function applyValueToCell(wi, col, val) {
  const work = S.works[wi];
  const ref = XLSX.utils.encode_cell({ r: work.myRow, c: col });

  const old = S.myCells[wi + '_' + col];

  captureOriginal(S.currentSheet, ref);

  const snap = getSheetOriginals(S.currentSheet).get(ref);

  getSheetPendingChanges(S.currentSheet).set(ref, val);

  const ws = S.wb.Sheets[S.currentSheet];

  if (ws[ref]) {
    ws[ref].v = val;
    ws[ref].t = 'n';
    delete ws[ref].w;
  } else {
    ws[ref] = { v: val, t: 'n' };
  }

  S.myCells[wi + '_' + col] = val;

  return {
    cellRef: ref,
    workIdx: wi,
    colIndex: col,
    oldValue: old,
    originalSnapshot: snap,
    sheetName: S.currentSheet
  };
}

function applyForemanByFloors() {
  if (!MS.floors.size) {
    showToast('Выберите хотя бы один этаж', 'warning');
    return;
  }

  const cols = S.columns.filter(c => c.section === S.filterSection && MS.floors.has(c.colIndex));

  const changes = [];
  let applied = 0;
  let skipped = 0;

  S.works.forEach((w, wi) => cols.forEach(col => {
    const fr = parseNum(S.foremanCells[wi + '_' + col.colIndex]);

    if (fr === null) {
      skipped++;
      return;
    }

    changes.push(applyValueToCell(wi, col.colIndex, clampPct(fr)));
    applied++;
  }));

  finishMass(changes, applied, skipped, 'floors');
}

function applyForemanByWorks() {
  if (!MS.works.size) {
    showToast('Выберите хотя бы одну работу', 'warning');
    return;
  }

  const cols = S.columns.filter(c => c.section === S.filterSection);

  const changes = [];
  let applied = 0;
  let skipped = 0;

  MS.works.forEach(wi => {
    if (!S.works[wi]) return;

    cols.forEach(col => {
      const fr = parseNum(S.foremanCells[wi + '_' + col.colIndex]);

      if (fr === null) {
        skipped++;
        return;
      }

      changes.push(applyValueToCell(wi, col.colIndex, clampPct(fr)));
      applied++;
    });
  });

  finishMass(changes, applied, skipped, 'works');
}

function finishMass(changes, applied, skipped, kind) {
  let entry = null;

  if (changes.length) {
    entry = { type: 'mass', changes, timestamp: Date.now() };
    S.undoStack.push(entry);
    if (S.undoStack.length > MAX_UNDO_STEPS) S.undoStack.shift();
  }

  MS[kind].clear();

  document.querySelectorAll(kind === 'floors' ? '#msFloorsOptions input' : '#msWorksOptions input')
    .forEach(cb => cb.checked = false);

  updateMultiselectLabel(kind === 'floors' ? 'msFloors' : 'msWorks');

  refreshUI();
  scheduleAutoSave();
  vibrate([30, 50, 30]);

  showToast(
    `Применено: ${applied}${skipped ? `, пропущено: ${skipped}` : ''}`,
    applied ? 'success' : 'warning',
    entry ? { label: 'Отменить', fn: () => undoFromToast(entry) } : null
  );
}

/* ================= МУЛЬТИВЫБОР (десктоп) ================= */
function toggleMultiselect(id, e) {
  e.stopPropagation();

  if (id === 'filterWorkMs' && isMobile()) {
    openWorkSheet();
    return;
  }

  document.querySelectorAll('.multiselect.open').forEach(ms => {
    if (ms.id !== id) ms.classList.remove('open');
  });

  document.getElementById(id).classList.toggle('open');
}

document.addEventListener('click', e => {
  if (!e.target.closest('.multiselect')) {
    document.querySelectorAll('.multiselect.open').forEach(ms => ms.classList.remove('open'));
  }
});

function multiselectAll(id, all, e) {
  e.stopPropagation();

  const cont = document.getElementById(id + 'Options');
  if (!cont) return;

  const set = id.includes('Floors') ? MS.floors : MS.works;

  cont.querySelectorAll('input[type="checkbox"]').forEach(cb => {
    cb.checked = all;
    const v = parseInt(cb.value, 10);
    all ? set.add(v) : set.delete(v);
  });

  updateMultiselectLabel(id);
  renderDrawerMass();
}

function updateMultiselectLabel(id) {
  const set = id.includes('Floors') ? MS.floors : MS.works;
  const l = document.querySelector(`#${id} .multiselect-label`);

  if (!l) return;

  l.textContent = `${id.includes('Floors') ? 'Выбрать этажи' : 'Выбрать работы'} (${set.size})`;
}

function handleMsFloorChange(cb) {
  const v = parseInt(cb.value, 10);
  cb.checked ? MS.floors.add(v) : MS.floors.delete(v);

  updateMultiselectLabel('msFloors');
  updateDrawerMassMeta();
}

function handleMsWorkChange(cb) {
  const v = parseInt(cb.value, 10);
  cb.checked ? MS.works.add(v) : MS.works.delete(v);

  updateMultiselectLabel('msWorks');
  updateDrawerMassMeta();
}

function filterMsWorks(id, q) {
  q = q.trim().toLowerCase();

  document.querySelectorAll(`#${id}Options .multiselect-option`).forEach(o => {
    o.classList.toggle('hidden', !!q && !(o.getAttribute('data-name') || '').includes(q));
  });
}

function rebuildMultiselects() {
  const fc = document.getElementById('msFloorsOptions');

  if (fc) {
    const sc = S.columns.filter(c => c.section === S.filterSection);
    MS.floors.clear();

    fc.innerHTML = sc.length
      ? sc.map(c => `
        <label class="multiselect-option">
          <input type="checkbox" value="${c.colIndex}" onchange="handleMsFloorChange(this)">
          <span class="multiselect-option-label">${esc(c.floor)}</span>
        </label>
      `).join('')
      : '<div class="px-3 py-2 text-xs text-surface-400">Нет этажей</div>';

    updateMultiselectLabel('msFloors');
  }

  const wc = document.getElementById('msWorksOptions');

  if (wc) {
    let vw = getVisibleWorks();
    if (S.filterWorkIdxs.size) vw = vw.filter(w => S.filterWorkIdxs.has(w.idx));

    MS.works.clear();

    wc.innerHTML = vw.map(w => `
      <label class="multiselect-option" data-name="${esc(w.name).toLowerCase()}">
        <input type="checkbox" value="${w.idx}" onchange="handleMsWorkChange(this)">
        <span class="multiselect-option-label">${esc(w.name)}</span>
      </label>
    `).join('');

    updateMultiselectLabel('msWorks');
  }
}

/* ================= БОКОВАЯ ПАНЕЛЬ ================= */
function openDrawer() {
  renderDrawerMass();
  document.getElementById('drawerBackdrop').classList.add('open');
  document.getElementById('sideDrawer').classList.add('open');
}

function closeDrawer() {
  document.getElementById('drawerBackdrop').classList.remove('open');
  document.getElementById('sideDrawer').classList.remove('open');
}

function drawerWorksSource() {
  let vw = getVisibleWorks();
  if (S.filterWorkIdxs.size) vw = vw.filter(w => S.filterWorkIdxs.has(w.idx));
  return vw;
}

function renderDrawerMass() {
  const sc = S.columns.filter(c => c.section === S.filterSection);

  const fl = document.getElementById('drawerFloorsList');

  if (fl) {
    fl.innerHTML = sc.length
      ? sc.map(c => `
        <label class="mass-item">
          <input type="checkbox" value="${c.colIndex}" ${MS.floors.has(c.colIndex) ? 'checked' : ''} onchange="handleDrawerFloorChange(this)">
          <span>${esc(c.floor)}</span>
        </label>
      `).join('')
      : '<div class="p-3 text-xs text-surface-400">Нет этажей</div>';
  }

  const wl = document.getElementById('drawerWorksList');

  if (wl) {
    wl.innerHTML = drawerWorksSource().map(w => `
      <label class="mass-item" data-name="${esc(w.name).toLowerCase()}">
        <input type="checkbox" value="${w.idx}" ${MS.works.has(w.idx) ? 'checked' : ''} onchange="handleDrawerWorkChange(this)">
        <span>${esc(w.name)}</span>
      </label>
    `).join('') || '<div class="p-3 text-xs text-surface-400">Нет работ</div>';
  }

  updateDrawerMassMeta();
}

function handleDrawerFloorChange(cb) {
  const v = parseInt(cb.value, 10);
  cb.checked ? MS.floors.add(v) : MS.floors.delete(v);

  updateMultiselectLabel('msFloors');
  updateDrawerMassMeta();
}

function handleDrawerWorkChange(cb) {
  const v = parseInt(cb.value, 10);
  cb.checked ? MS.works.add(v) : MS.works.delete(v);

  updateMultiselectLabel('msWorks');
  updateDrawerMassMeta();
}

function drawerAll(kind, all) {
  if (kind === 'floors') {
    S.columns.filter(c => c.section === S.filterSection)
      .forEach(c => all ? MS.floors.add(c.colIndex) : MS.floors.delete(c.colIndex));

    updateMultiselectLabel('msFloors');
  } else {
    drawerWorksSource().forEach(w => all ? MS.works.add(w.idx) : MS.works.delete(w.idx));

    updateMultiselectLabel('msWorks');
  }

  renderDrawerMass();
}

function filterDrawerWorks(q) {
  q = q.trim().toLowerCase();

  document.querySelectorAll('#drawerWorksList .mass-item').forEach(o => {
    o.classList.toggle('hidden', !!q && !(o.getAttribute('data-name') || '').includes(q));
  });
}

function massEligible(kind){
  const cols = S.columns.filter(c => c.section === S.filterSection);
  let n = 0;

  if (kind === 'floors') {
    const tc = cols.filter(c => MS.floors.has(c.colIndex));

    S.works.forEach((w, wi) => tc.forEach(c => {
      const fr = parseNum(S.foremanCells[wi + '_' + c.colIndex]);
      const my = parseNum(S.myCells[wi + '_' + c.colIndex]);

      if (fr !== null && my !== fr) n++;
    }));
  } else {
    MS.works.forEach(wi => {
      if (!S.works[wi]) return;

      cols.forEach(c => {
        const fr = parseNum(S.foremanCells[wi + '_' + c.colIndex]);
        const my = parseNum(S.myCells[wi + '_' + c.colIndex]);

        if (fr !== null && my !== fr) n++;
      });
    });
  }

  return n;
}

function updateDrawerMassMeta() {
  const ef = massEligible('floors');
  const ew = massEligible('works');

  const bf = document.getElementById('drawerMassFloorsBtn');
  const bw = document.getElementById('drawerMassWorksBtn');

  const hf = document.getElementById('drawerFloorsHint');
  const hw = document.getElementById('drawerWorksHint');

  if (bf) bf.disabled = !MS.floors.size || !ef;
  if (bw) bw.disabled = !MS.works.size || !ew;

  if (hf) {
    hf.textContent = MS.floors.size
      ? `Выбрано ${MS.floors.size} · будет подтверждено: ${ef}`
      : 'Выберите этажи';
  }

  if (hw) {
    hw.textContent = MS.works.size
      ? `Выбрано ${MS.works.size} · будет подтверждено: ${ew}`
      : 'Выберите работы';
  }
}

/* ================= UNDO ================= */
function revertAction(action) {
  const list = action.type === 'single' ? [action] : action.changes;
  const touched = new Set();

  list.forEach(ch => {
    const sheet = ch.sheetName || S.currentSheet;
    const ws = S.wb.Sheets[sheet];
    const originals = getSheetOriginals(sheet);

    const snap = ch.originalSnapshot !== undefined
      ? ch.originalSnapshot
      : (originals.has(ch.cellRef) ? originals.get(ch.cellRef) : null);

    restoreCellSnapshot(ws, ch.cellRef, snap);

    getSheetPendingChanges(sheet).delete(ch.cellRef);

    if (sheet === S.currentSheet) {
      S.myCells[ch.workIdx + '_' + ch.colIndex] = ch.oldValue;
    }

    touched.add(sheet);
  });

  touched.forEach(n => {
    if (getSheetPendingChanges(n).size === 0) S.cachedSheets.delete(n);
  });

  refreshUI();
  scheduleAutoSave();
}

function undoLastChange() {
  const a = S.undoStack.pop();

  if (!a) {
    showToast('Нечего отменять', 'warning');
    return;
  }

  revertAction(a);
  showToast('Изменение отменено');
  vibrate([30, 50, 30]);
}

function undoFromToast(entry) {
  const i = S.undoStack.indexOf(entry);
  if (i < 0) return;

  S.undoStack.splice(i, 1);
  revertAction(entry);
  showToast('Изменение отменено');
}

function refreshUI() {
  updateSelectionPanel();
  renderMatrix();
  updateSummary();
  rebuildSheetSelect();
  updateUndoButton();
  renderDrawerMass();
}

/* ================= СОХРАНЕНИЕ / ЭКСПОРТ ================= */
async function buildPatchedBuffer() {
  let buffer = S.originalBuffer;
  const total = getTotalPendingChanges();

  if (total > 0 && buffer) {
    for (const n in S.pendingChanges) {
      const ch = S.pendingChanges[n];
      if (ch && ch.size) buffer = await patchExcelFile(buffer, ch, n);
    }
  } else if (!buffer && S.wb) {
    buffer = XLSX.write(S.wb, { bookType: 'xlsx', type: 'array' });
  }

  return buffer;
}

function clearAfterPersist(buffer) {
  S.originalBuffer = buffer;

  for (const k in S.pendingChanges) S.pendingChanges[k].clear();

  S.originalCells = {};
  S.undoStack = [];
  S.cachedSheets.clear();

  clearDraft();
  rebuildSheetSelect();
  updateUndoButton();
  updateSummary();
  renderDrawerMass();

  idbPut('lastFile', { name: S.fileName, buffer, timestamp: Date.now() }).catch(() => {});
}

async function saveToFile() {
  if (typeof XLSX === 'undefined') {
    showToast('Библиотека XLSX не загружена', 'error');
    return false;
  }

  if (!getTotalPendingChanges()) {
    showToast('Нет изменений', 'warning');
    return true;
  }

  try {
    const buffer = await buildPatchedBuffer();

    if (!buffer) {
      showToast('Нет данных', 'warning');
      return false;
    }

    if (S.saveHandle) {
      const w = await S.saveHandle.createWritable();
      await w.write(buffer);
      await w.close();

      clearAfterPersist(buffer);

      showToast('Файл сохранён ✓');
      vibrate([30, 50, 30]);

      return true;
    }

    const file = bufferToFile(buffer, S.fileName || 'audit.xlsx');

    if (isMobile() && navigator.canShare && navigator.canShare({ files: [file] })) {
      try {
        await navigator.share({
          files: [file],
          title: 'Аудит СМР',
          text: 'Сохраните файл через «Файлы» / Excel'
        });

        clearAfterPersist(buffer);
        showToast('Файл отправлен на сохранение ✓');

        return true;
      } catch (err) {
        if (err.name === 'AbortError') return false;
      }
    }

    downloadBuffer(buffer, S.fileName || 'audit.xlsx');
    clearAfterPersist(buffer);

    showToast('Файл скачан (браузер не умеет перезаписывать исходный)');
    vibrate([30, 50, 30]);

    return true;
  } catch (err) {
    console.error(err);
    showToast('Ошибка сохранения: ' + err.message, 'error');
    return false;
  }
}

function stamp() {
  const n = new Date();
  const p = x => String(x).padStart(2, '0');

  return `${p(n.getDate())}-${p(n.getMonth() + 1)}-${String(n.getFullYear()).slice(2)}_${p(n.getHours())}-${p(n.getMinutes())}-${p(n.getSeconds())}`;
}

function bufferToFile(buffer, name) {
  return new File([buffer], name, {
    type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
  });
}

async function exportToFile() {
  if (typeof XLSX === 'undefined') {
    showToast('Библиотека XLSX не загружена', 'error');
    return false;
  }

  if (!getTotalPendingChanges() && !S.originalBuffer && !S.wb) {
    showToast('Нет данных для экспорта', 'warning');
    return false;
  }

  try {
    const buffer = await buildPatchedBuffer();

    if (!buffer) {
      showToast('Нет данных', 'warning');
      return false;
    }

    downloadBuffer(buffer, `audit_${stamp()}.xlsx`);
    clearAfterPersist(buffer);

    showToast('Файл экспортирован ✓');
    vibrate([30, 50, 30]);

    return true;
  } catch (err) {
    console.error(err);
    showToast('Ошибка экспорта: ' + err.message, 'error');
    return false;
  }
}

function downloadBuffer(buffer, name) {
  const blob = new Blob([buffer], {
    type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
  });

  const url = URL.createObjectURL(blob);

  const a = document.createElement('a');
  a.href = url;
  a.download = name;

  document.body.appendChild(a);
  a.click();
  a.remove();

  setTimeout(() => URL.revokeObjectURL(url), 100);
}

async function shareFile() {
  if (!S.originalBuffer && !S.wb) {
    showToast('Нет файла', 'error');
    return;
  }

  try {
    const buffer = await buildPatchedBuffer();
    const name = `audit_${stamp()}.xlsx`;
    const file = bufferToFile(buffer, name);

    if (navigator.canShare && navigator.canShare({ files: [file] })) {
      try {
        await navigator.share({
          files: [file],
          title: 'Аудит СМР',
          text: 'Отчёт по факту выполнения работ'
        });

        showToast('Файл отправлен ✓');
      } catch (err) {
        if (err.name !== 'AbortError') showToast('Ошибка шеринга', 'error');
      }
    } else {
      downloadBuffer(buffer, name);
      showToast('Файл скачан ✓');
    }
  } catch (err) {
    showToast('Ошибка: ' + err.message, 'error');
  }
}

/* ================= ЧЕРНОВИК ================= */
function scheduleAutoSave() {
  S.lastChangeTime = Date.now();
  updateSaveIndicator();
}

function saveDraft() {
  if (!S.isLocalStorageAvailable || !S.fileName) return;

  try {
    const key = `draft_${S.fileName}`;
    const all = {};
    let total = 0;

    for (const n in S.pendingChanges) {
      if (S.pendingChanges[n].size) {
        all[n] = Array.from(S.pendingChanges[n].entries());
        total += S.pendingChanges[n].size;
      }
    }

    if (!total) {
      localStorage.removeItem(key);
      return;
    }

    localStorage.setItem(key, JSON.stringify({
      allChanges: all,
      total,
      timestamp: Date.now()
    }));
  } catch (e) {
    console.warn('draft', e);
  }
}

function checkDraft() {
  if (!S.isLocalStorageAvailable || !S.fileName) return;

  try {
    const key = `draft_${S.fileName}`;
    const raw = localStorage.getItem(key);

    if (!raw) return;

    const d = JSON.parse(raw);

    if (!d.allChanges || !d.total) {
      localStorage.removeItem(key);
      return;
    }

    const dt = new Date(d.timestamp);
    const p = x => String(x).padStart(2, '0');

    showAppModal(
      'Найден черновик',
      `Несохранённые правки (${d.total} ячеек) от ${p(dt.getDate())}.${p(dt.getMonth() + 1)} ${p(dt.getHours())}:${p(dt.getMinutes())}. Восстановить?`,
      [
        {
          label: 'Восстановить',
          cls: 'btn-gradient text-white',
          fn: () => {
            for (const n in d.allChanges) {
              S.pendingChanges[n] = new Map(d.allChanges[n]);
              S.cachedSheets.add(n);
            }

            applyPendingToCurrentSheet();
            refreshUI();
            scheduleAutoSave();

            showToast('Черновик восстановлен');
          }
        },
        {
          label: 'Удалить черновик',
          cls: 'bg-red-600 text-white',
          fn: () => localStorage.removeItem(key)
        },
        {
          label: 'Позже'
        }
      ]
    );
  } catch (e) {
    console.warn('draft', e);
  }
}

function clearDraft() {
  if (S.isLocalStorageAvailable && S.fileName) {
    try {
      localStorage.removeItem(`draft_${S.fileName}`);
    } catch (e) {}
  }

  updateSaveIndicator();
}

/* ================= СВОДКА (кнопка расхождений) ================= */
function updateSummary() {
  const cols = S.columns.filter(c => c.section === S.filterSection);

  let mismatch = 0;

  S.works.forEach((w, idx) => {
    const work = Object.assign({}, w, { idx });
    if (shouldHideRow(work, cols)) return;
    if (countMismatches(work, cols) > 0) mismatch++;
  });

  const cnt = document.getElementById('mismatchCount');
  const btn = document.getElementById('mismatchBtn');

  if (cnt) cnt.textContent = String(mismatch);

  if (btn) {
    const active = S.summaryFilter === 'mismatch';

    btn.classList.toggle('active', active);
    btn.setAttribute('aria-pressed', String(active));
    btn.disabled = mismatch === 0 && !active;

    btn.title = active
      ? 'Показать все работы'
      : `Показать только работы с расхождениями (${mismatch})`;
  }

  updateSaveIndicator();
  updateDrawerMassMeta();
}

function toggleMismatchFilter() {
  S.summaryFilter = S.summaryFilter === 'mismatch' ? null : 'mismatch';

  renderMatrix();
  rebuildWorksDropdown();
  renderDrawerMass();
  updateSummary();

  vibrate(20);
}

/* ================= ФАЙЛОВЫЕ СОБЫТИЯ ================= */
function handleDragOver(e) {
  e.preventDefault();
  e.stopPropagation();
  document.getElementById('dropZone').classList.add('drop-active');
}

function handleDragLeave(e) {
  e.preventDefault();
  e.stopPropagation();
  document.getElementById('dropZone').classList.remove('drop-active');
}

function handleDrop(e) {
  e.preventDefault();
  e.stopPropagation();

  document.getElementById('dropZone').classList.remove('drop-active');

  if (e.dataTransfer.files.length) {
    processFile(e.dataTransfer.files[0], null);
  }
}

function handleFileSelect(e) {
  if (e.target.files.length) processFile(e.target.files[0], null);
  e.target.value = '';
}

async function openFilePicker() {
  const proceed = async () => {
    if (window.showOpenFilePicker) {
      try {
        const [h] = await window.showOpenFilePicker({
          multiple: false,
          types: [{
            description: 'Excel',
            accept: {
              'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet': ['.xlsx']
            }
          }]
        });

        processFile(await h.getFile(), h);
      } catch (err) {
        if (err.name !== 'AbortError') {
          (document.getElementById('fileInputSecondary') || document.getElementById('fileInput')).click();
        }
      }
    } else {
      (document.getElementById('fileInputSecondary') || document.getElementById('fileInput')).click();
    }
  };

  if (getTotalPendingChanges() > 0) showWarningModal(proceed);
  else proceed();
}

function processFile(file, handle) {
  if (file.name.split('.').pop().toLowerCase() !== 'xlsx') {
    showError('Поддерживаются только .xlsx');
    return;
  }

  readExcelFile(file, handle);
}

function showFormStatus(msg, type) {
  const el = document.getElementById('formStatus');
  if (!el) return;

  el.textContent = msg;

  el.className = 'mt-3 p-3 rounded-xl text-sm font-medium ' + (
    type === 'success'
      ? 'bg-green-50 dark:bg-green-900/20 text-green-700 dark:text-green-400 border border-green-200 dark:border-green-800'
      : 'bg-red-50 dark:bg-red-900/20 text-red-700 dark:text-red-400 border border-red-200 dark:border-red-800'
  );

  el.classList.remove('hidden');
  setTimeout(() => el.classList.add('hidden'), 3000);
}

function showError(msg) {
  const el = document.getElementById('uploadError');

  el.textContent = msg;
  el.classList.remove('hidden');

  setTimeout(() => el.classList.add('hidden'), 6000);
}

function toggleTheme() {
  S.darkMode = !S.darkMode;

  document.documentElement.classList.toggle('dark', S.darkMode);

  if (S.isLocalStorageAvailable) {
    try {
      localStorage.setItem('ctrl-dark', S.darkMode ? '1' : '0');
    } catch (e) {}
  }
}

/* ================= ДЕМО ================= */
function buildDemo() {
  if (typeof XLSX === 'undefined') {
    showToast('Библиотека XLSX не загружена. Запустите START.bat', 'error');
    return;
  }

  const sections = ['Секция 1', 'Секция 2', 'Секция 3'];
  const floors = 5;

  const works = [
    'Монолитные работы',
    'Кладка перегородок',
    'Штукатурка стен',
    'Стяжка пола',
    'Электромонтаж',
    'Сантехника (ВК)',
    'Окна ПВХ',
    'Чистовая отделка'
  ];

  const aoa = [];

  const r0 = ['', '', ''];
  const r1 = ['', '', ''];
  const r2 = ['', '', 'Этаж'];

  let col = 3;

  sections.forEach(sec => {
    for (let f = 1; f <= floors; f++) {
      r0[col] = f === 1 ? sec : '';
      r1[col] = '';
      r2[col] = String(f);
      col++;
    }
  });

  aoa.push(r0, r1, r2);

  works.forEach((w, wi) => {
    const my = ['', '', 'w'.replace('w', w)];
    const fr = ['ПРОРАБ', '', ''];

    for (let c = 3; c < col; c++) {
      const fl = ((c - 3) % floors) + 1;

      const base = 100 - (fl - 1) * 18 - wi * 6;
      const v = Math.max(0, Math.min(100, base + ((c * 7 + wi * 13) % 15) - 7));

      fr[c] = v <= 5 ? null : v;
      my[c] = null;
    }

    aoa.push(my, fr);
  });

  const ws = XLSX.utils.aoa_to_sheet(aoa);
  const wb = XLSX.utils.book_new();

  XLSX.utils.book_append_sheet(wb, ws, 'Корпус 1');

  const buffer = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });

  S.originalBuffer = buffer;
  S.wb = XLSX.read(new Uint8Array(buffer), { type: 'array', cellDates: true });

  S.fileName = 'Демо — ЖК Северный парк.xlsx';
  S.saveHandle = null;
  S.demoMode = true;

  S.pendingChanges = {};
  S.originalCells = {};
  S.undoStack = [];
  S.summaryFilter = null;

  S.currentSheet = S.wb.SheetNames[0];
  S.sheetColors = { [S.currentSheet]: {} };
  S.cellColors = {};

  rebuildSheetSelect();
  parseSheet();

  showToast('Демо-проект загружен. Данные вымышленные.');
}

/* ================= INDEXEDDB ================= */
const IDB_NAME = 'audit-smr-store';
const IDB_STORE = 'kv';

function openIDB() {
  return new Promise((res, rej) => {
    if (!window.indexedDB) {
      rej(new Error('no idb'));
      return;
    }

    const r = indexedDB.open(IDB_NAME, 1);

    r.onupgradeneeded = () => {
      if (!r.result.objectStoreNames.contains(IDB_STORE)) {
        r.result.createObjectStore(IDB_STORE);
      }
    };

    r.onsuccess = () => res(r.result);
    r.onerror = () => rej(r.error);
  });
}

async function idbPut(k, v) {
  const db = await openIDB();

  return new Promise((res, rej) => {
    const tx = db.transaction(IDB_STORE, 'readwrite');

    tx.objectStore(IDB_STORE).put(v, k);

    tx.oncomplete = () => {
      db.close();
      res();
    };

    tx.onerror = () => {
      db.close();
      rej(tx.error);
    };
  });
}

async function idbGet(k) {
  const db = await openIDB();

  return new Promise((res, rej) => {
    const tx = db.transaction(IDB_STORE, 'readonly');
    const q = tx.objectStore(IDB_STORE).get(k);

    q.onsuccess = () => {
      db.close();
      res(q.result);
    };

    q.onerror = () => {
      db.close();
      rej(q.error);
    };
  });
}

async function idbDel(k) {
  const db = await openIDB();

  return new Promise((res, rej) => {
    const tx = db.transaction(IDB_STORE, 'readwrite');

    tx.objectStore(IDB_STORE).delete(k);

    tx.oncomplete = () => {
      db.close();
      res();
    };

    tx.onerror = () => {
      db.close();
      rej(tx.error);
    };
  });
}

async function tryRestoreLastFile() {
  let rec;

  try {
    rec = await idbGet('lastFile');
  } catch (e) {
    return;
  }

  if (!rec || !rec.buffer) return;

  const d = new Date(rec.timestamp);
  const p = x => String(x).padStart(2, '0');

  showAppModal(
    'Продолжить работу?',
    `Последний файл: <b>${esc(rec.name)}</b><br>Сохранён: ${p(d.getDate())}.${p(d.getMonth() + 1)} ${p(d.getHours())}:${p(d.getMinutes())}.<br>Открыть его заново без выбора файла?`,
    [
      {
        label: 'Открыть файл',
        cls: 'btn-gradient text-white',
        fn: () => readExcelFile(
          new File([rec.buffer], rec.name, {
            type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
          }),
          null
        )
      },
      {
        label: 'Не сейчас'
      },
      {
        label: 'Забыть файл',
        cls: 'bg-red-600 text-white',
        fn: () => idbDel('lastFile').catch(() => {})
      }
    ]
  );
}

/* ================= СТАРТ ================= */
(function init() {
  checkLocalStorage();

  try {
    if (S.isLocalStorageAvailable && localStorage.getItem('ctrl-dark') === '1') {
      S.darkMode = true;
      document.documentElement.classList.add('dark');
    }
  } catch (e) {}

  setTimeout(() => {
    const sp = document.getElementById('splashScreen');
    if (sp) sp.classList.add('hidden');
  }, 900);

  const mv = document.getElementById('manualValue');
  if (mv) {
    mv.addEventListener('keydown', e => {
      if (e.key === 'Enter') submitManualValue();
    });
  }

  initMatrixInteractions();

  document.addEventListener('keydown', e => {
    if (e.key === 'Escape') {
      closeBottomSheet();
      closeWorkSheet();
      closeDrawer();
      closeWarningModal();
      closeAppModal();
      return;
    }

    if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 's') {
      e.preventDefault();

      if (!document.getElementById('appScreen').classList.contains('hidden')) {
        saveToFile();
      }
    }

    if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 'z') {
      const t = (document.activeElement && document.activeElement.tagName) || '';

      if (t !== 'INPUT' && t !== 'TEXTAREA') {
        e.preventDefault();
        undoLastChange();
      }
    }
  });

  window.addEventListener('pagehide', () => {
    if (getTotalPendingChanges() > 0) saveDraft();
  });

  document.addEventListener('visibilitychange', () => {
    if (document.visibilityState === 'hidden' && getTotalPendingChanges() > 0) saveDraft();
  });

  window.addEventListener('resize', () => {
    S.isMobile = window.innerWidth < 1024;
  });

  setInterval(() => {
    if (getTotalPendingChanges() > 0) saveDraft();
  }, S.autoSaveInterval);

  if (typeof XLSX === 'undefined' || typeof JSZip === 'undefined') {
    showError('Библиотеки не загружены. Запустите START.bat в папке проекта и обновите страницу.');
  }

  const params = new URLSearchParams(location.search);

  if (params.get('demo') === '1') buildDemo();
  else tryRestoreLastFile();
})();