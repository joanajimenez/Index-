const CONFIG = {
  SHEET_NAME: 'Completo',
  HEADER_ROW: 1,

  HEADERS: {
    PAIS: 'PAIS',
    TIPO: 'Tipo',
    PVD: 'Nombre en Portal de Venta Directa (PVD)'
  },

  FALLBACK_INDEX: {
    PAIS: 0,
    TIPO: 1,
    PVD: 3
  },

  CACHE_DURATION: 300,
  MAX_ROWS_DISPLAY: 500
};

// ========== CACHE MANAGER ==========
const CACHE_KEYS = {
  SHEET_DATA: 'sheetData',
  OPTIONS: 'options'
};

function getCache(key) {
  const cache = CacheService.getScriptCache();
  const data = cache.get(key);
  return data ? JSON.parse(data) : null;
}

function setCache(key, data) {
  const cache = CacheService.getScriptCache();
  cache.put(key, JSON.stringify(data), CONFIG.CACHE_DURATION);
}

function clearCache() {
  const cache = CacheService.getScriptCache();
  cache.removeAll([CACHE_KEYS.SHEET_DATA, CACHE_KEYS.OPTIONS]);
}

// ========== INTERFAZ ==========
function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Buscador Productos')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ========== LECTURA DE DATOS ==========
function readSheet_() {
  const cached = getCache(CACHE_KEYS.SHEET_DATA);
  if (cached) return cached;

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(CONFIG.SHEET_NAME);
  if (!sh) throw new Error(`No existe la pestaña "${CONFIG.SHEET_NAME}". Revisá el nombre exacto.`);

  const lastRow = sh.getLastRow();
  const lastCol = sh.getLastColumn();
  if (lastRow <= CONFIG.HEADER_ROW || lastCol < 1) return { headers: [], data: [], colIndex: {} };

  const values = sh.getRange(CONFIG.HEADER_ROW, 1, lastRow - CONFIG.HEADER_ROW + 1, lastCol).getValues();
  const rawHeaders = values[0].map(h => String(h ?? '').trim());
  const data = values.slice(1);

  const colIndex = {};
  rawHeaders.forEach((h, i) => {
    const k = normalize_(h);
    if (!k) return;
    if (colIndex[k] === undefined) colIndex[k] = i;
  });

  const headers = rawHeaders.map((h, i) => h || `Col${i + 1}`);

  const result = { headers, data, colIndex };
  setCache(CACHE_KEYS.SHEET_DATA, result);
  return result;
}

function normalize_(s) {
  return String(s ?? '')
    .trim()
    .toLowerCase()
    .replace(/\s+/g, ' ')
    .normalize('NFD').replace(/[\u0300-\u036f]/g, '');
}

function getCol_(colIndex, headerText, fallbackIdx) {
  const key = normalize_(headerText);
  if (key && colIndex[key] !== undefined) return colIndex[key];
  return fallbackIdx;
}

function uniqSorted_(arr) {
  const s = new Set(
    arr.map(v => String(v ?? '').trim()).filter(v => v !== '')
  );
  return Array.from(s).sort((a, b) => a.localeCompare(b, 'es'));
}

// ========== OPCIONES - SOLO PAÍSES AL INICIO ==========
function getOptions() {
  const { data, colIndex } = readSheet_();

  const iPais = getCol_(colIndex, CONFIG.HEADERS.PAIS, CONFIG.FALLBACK_INDEX.PAIS);

  const paises = uniqSorted_(data.map(r => r[iPais]));

  return { paises };
}

// ========== OPCIONES FILTRADAS - TIPOS Y PVD EN CASCADA ==========
function getOptionsFiltered(filters) {
  const { data, colIndex } = readSheet_();

  const iPais = getCol_(colIndex, CONFIG.HEADERS.PAIS, CONFIG.FALLBACK_INDEX.PAIS);
  const iTipo = getCol_(colIndex, CONFIG.HEADERS.TIPO, CONFIG.FALLBACK_INDEX.TIPO);
  const iPvd = getCol_(colIndex, CONFIG.HEADERS.PVD, CONFIG.FALLBACK_INDEX.PVD);

  let filtered = data;

  // Filtrar por país
  if (filters?.pais) {
    filtered = filtered.filter(r => String(r[iPais]).trim() === String(filters.pais).trim());
  }

  const tipos = uniqSorted_(filtered.map(r => r[iTipo]));

  // Filtrar por tipo
  if (filters?.tipo) {
    filtered = filtered.filter(r => String(r[iTipo]).trim() === String(filters.tipo).trim());
  }

  const pvds = uniqSorted_(filtered.map(r => r[iPvd]));

  return { tipos, pvds };
}

// ========== BÚSQUEDA ==========
function search(filters) {
  const { headers, data, colIndex } = readSheet_();

  const iPais = getCol_(colIndex, CONFIG.HEADERS.PAIS, CONFIG.FALLBACK_INDEX.PAIS);
  const iTipo = getCol_(colIndex, CONFIG.HEADERS.TIPO, CONFIG.FALLBACK_INDEX.TIPO);
  const iPvd = getCol_(colIndex, CONFIG.HEADERS.PVD, CONFIG.FALLBACK_INDEX.PVD);

  let out = data;

  if (filters?.pais) out = out.filter(r => String(r[iPais]).trim() === String(filters.pais).trim());
  if (filters?.tipo) out = out.filter(r => String(r[iTipo]).trim() === String(filters.tipo).trim());
  if (filters?.pvd) out = out.filter(r => String(r[iPvd]).trim() === String(filters.pvd).trim());

  // Limitar resultados para mejor rendimiento
  if (out.length > CONFIG.MAX_ROWS_DISPLAY) {
    out = out.slice(0, CONFIG.MAX_ROWS_DISPLAY);
  }

  const rows = out.map(r => {
    const obj = {};
    headers.forEach((h, i) => obj[h] = r[i]);
    return obj;
  });

  return { headers, rows };
}

// ========== DIAGNÓSTICO ==========
function debugStatus() {
  const { headers, data, colIndex } = readSheet_();
  const iPais = getCol_(colIndex, CONFIG.HEADERS.PAIS, CONFIG.FALLBACK_INDEX.PAIS);
  const iTipo = getCol_(colIndex, CONFIG.HEADERS.TIPO, CONFIG.FALLBACK_INDEX.TIPO);
  const iPvd = getCol_(colIndex, CONFIG.HEADERS.PVD, CONFIG.FALLBACK_INDEX.PVD);

  return {
    sheet: CONFIG.SHEET_NAME,
    headers_count: headers.length,
    rows_count: data.length,
    detected_indexes: { pais: iPais, tipo: iTipo, pvd: iPvd },
    sample_paises: uniqSorted_(data.map(r => r[iPais])).slice(0, 10),
    sample_tipos: uniqSorted_(data.map(r => r[iTipo])).slice(0, 10),
    sample_pvds: uniqSorted_(data.map(r => r[iPvd])).slice(0, 10),
    cache_enabled: true,
    cache_duration_seconds: CONFIG.CACHE_DURATION
  };
}

// ========== UTILIDADES ==========
function invalidateCache() {
  clearCache();
  return { success: true, message: 'Cache limpiado' };
}
