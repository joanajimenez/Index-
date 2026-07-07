const CONFIG = {
  SHEET_NAME: 'Completo', // tu pestaña
  HEADER_ROW: 1,

  // Encabezados esperados (fila 1) para los motores de búsqueda:
  HEADERS: {
    PAIS: 'PAIS', // Col A
    TIPO: 'Tipo', // Col B
    PVD:  'Nombre en Portal de Venta Directa (PVD)' // Col D
  },

  // Fallback por posición (0-index) si no encuentra el encabezado:
  FALLBACK_INDEX: {
    PAIS: 0, // A
    TIPO: 1, // B
    PVD:  3  // D
  }
};

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Buscador Productos')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ---------- Lectura dinámica del Sheet ----------
function readSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(CONFIG.SHEET_NAME);
  if (!sh) throw new Error(`No existe la pestaña "${CONFIG.SHEET_NAME}". Revisá el nombre exacto.`);

  const lastRow = sh.getLastRow();
  const lastCol = sh.getLastColumn();
  if (lastRow <= CONFIG.HEADER_ROW || lastCol < 1) return { headers: [], data: [], colIndex: {} };

  const values = sh.getRange(CONFIG.HEADER_ROW, 1, lastRow - CONFIG.HEADER_ROW + 1, lastCol).getValues();
  const rawHeaders = values[0].map(h => String(h ?? '').trim());
  const data = values.slice(1);

  // Map normalizado header -> índice
  const colIndex = {};
  rawHeaders.forEach((h, i) => {
    const k = normalize_(h);
    if (!k) return;
    // Si hay headers duplicados, nos quedamos con el primero
    if (colIndex[k] === undefined) colIndex[k] = i;
  });

  // Headers seguros (si alguno viene vacío)
  const headers = rawHeaders.map((h, i) => h || `Col${i + 1}`);

  return { headers, data, colIndex };
}

function normalize_(s) {
  return String(s ?? '')
    .trim()
    .toLowerCase()
    .replace(/\s+/g, ' ')
    .normalize('NFD').replace(/[\u0300-\u036f]/g, ''); // quita tildes
}

function getCol_(colIndex, headerText, fallbackIdx) {
  const key = normalize_(headerText);
  if (key && colIndex[key] !== undefined) return colIndex[key];
  return fallbackIdx; // fallback por posición
}

function uniqSorted_(arr) {
  const s = new Set(
    arr.map(v => String(v ?? '').trim()).filter(v => v !== '')
  );
  return Array.from(s).sort((a, b) => a.localeCompare(b, 'es'));
}

// ---------- Opciones para desplegables ----------
function getOptions() {
  const { data, colIndex } = readSheet_();

  const iPais = getCol_(colIndex, CONFIG.HEADERS.PAIS, CONFIG.FALLBACK_INDEX.PAIS);
  const iTipo = getCol_(colIndex, CONFIG.HEADERS.TIPO, CONFIG.FALLBACK_INDEX.TIPO);
  const iPvd  = getCol_(colIndex, CONFIG.HEADERS.PVD,  CONFIG.FALLBACK_INDEX.PVD);

  const paises = uniqSorted_(data.map(r => r[iPais]));
  const tipos  = uniqSorted_(data.map(r => r[iTipo]));
  const pvds   = uniqSorted_(data.map(r => r[iPvd]));

  return { paises, tipos, pvds };
}

function getOptionsFiltered(filters) {
  const { data, colIndex } = readSheet_();

  const iPais = getCol_(colIndex, CONFIG.HEADERS.PAIS, CONFIG.FALLBACK_INDEX.PAIS);
  const iTipo = getCol_(colIndex, CONFIG.HEADERS.TIPO, CONFIG.FALLBACK_INDEX.TIPO);
  const iPvd  = getCol_(colIndex, CONFIG.HEADERS.PVD,  CONFIG.FALLBACK_INDEX.PVD);

  let filtered = data;

  if (filters?.pais) filtered = filtered.filter(r => String(r[iPais]).trim() === String(filters.pais).trim());
  const tipos = uniqSorted_(filtered.map(r => r[iTipo]));

  if (filters?.tipo) filtered = filtered.filter(r => String(r[iTipo]).trim() === String(filters.tipo).trim());
  const pvds = uniqSorted_(filtered.map(r => r[iPvd]));

  // Países completos
  const paises = uniqSorted_(data.map(r => r[iPais]));

  return { paises, tipos, pvds };
}

// ---------- Búsqueda ----------
function search(filters) {
  const { headers, data, colIndex } = readSheet_();

  const iPais = getCol_(colIndex, CONFIG.HEADERS.PAIS, CONFIG.FALLBACK_INDEX.PAIS);
  const iTipo = getCol_(colIndex, CONFIG.HEADERS.TIPO, CONFIG.FALLBACK_INDEX.TIPO);
  const iPvd  = getCol_(colIndex, CONFIG.HEADERS.PVD,  CONFIG.FALLBACK_INDEX.PVD);

  let out = data;

  if (filters?.pais) out = out.filter(r => String(r[iPais]).trim() === String(filters.pais).trim());
  if (filters?.tipo) out = out.filter(r => String(r[iTipo]).trim() === String(filters.tipo).trim());
  if (filters?.pvd)  out = out.filter(r => String(r[iPvd]).trim()  === String(filters.pvd).trim());

  // Convertimos filas a objetos usando headers actuales (dinámicos)
  const rows = out.map(r => {
    const obj = {};
    headers.forEach((h, i) => obj[h] = r[i]);
    return obj;
  });

  return { headers, rows };
}

// Diagnóstico rápido opcional
function debugStatus() {
  const { headers, data, colIndex } = readSheet_();
  const iPais = getCol_(colIndex, CONFIG.HEADERS.PAIS, CONFIG.FALLBACK_INDEX.PAIS);
  const iTipo = getCol_(colIndex, CONFIG.HEADERS.TIPO, CONFIG.FALLBACK_INDEX.TIPO);
  const iPvd  = getCol_(colIndex, CONFIG.HEADERS.PVD,  CONFIG.FALLBACK_INDEX.PVD);

  return {
    sheet: CONFIG.SHEET_NAME,
    headers_count: headers.length,
    rows_count: data.length,
    detected_indexes: { pais: iPais, tipo: iTipo, pvd: iPvd },
    sample_paises: uniqSorted_(data.map(r => r[iPais])).slice(0, 10),
    sample_tipos: uniqSorted_(data.map(r => r[iTipo])).slice(0, 10),
    sample_pvds: uniqSorted_(data.map(r => r[iPvd])).slice(0, 10),
  };
}

