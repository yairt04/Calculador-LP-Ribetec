function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('💼 Calculadora de Precios')
    .addItem('Abrir herramienta', 'showModal')
    .addItem('Reparar fórmulas NPI 2026', 'repararFormulasNpi2026')
    .addToUi();
}

function showModal() {
  const html = HtmlService.createTemplateFromFile('Sidebar').evaluate()
    .setWidth(900)
    .setHeight(780);
  SpreadsheetApp.getUi().showModalDialog(html, 'Calculadora de Precios');
}

// ===== Datos maestros
const TC_SOURCE_BANXICO = 'Banxico FIX';
const TC_SOURCE_GOOGLEFINANCE = 'GOOGLEFINANCE';
const TC_CACHE_KEY = 'TC_USD_MXN_DIA';
const TC_CACHE_SECONDS = 60 * 60 * 6;
const BANXICO_SERIE_FIX = 'SF43718';

function getTipoCambio() {
  const cache = CacheService.getScriptCache();
  const cached = cache.get(TC_CACHE_KEY);

  if (cached) {
    return JSON.parse(cached);
  }

  let result = null;

  try {
    result = getTipoCambioBanxico_();
  } catch (err) {
    console.warn('Banxico falló: ' + err.message);
  }

  if (!result || !result.tc) {
    result = getTipoCambioGoogleFinance_();
  }

  if (!result || !result.tc) {
    throw new Error('No se pudo obtener el tipo de cambio USD/MXN.');
  }

  cache.put(TC_CACHE_KEY, JSON.stringify(result), TC_CACHE_SECONDS);
  return result;
}

function getTipoCambioBanxico_() {
  const token = PropertiesService.getScriptProperties().getProperty('BANXICO_TOKEN');
  const url = 'https://www.banxico.org.mx/SieAPIRest/service/v1/series/'
    + BANXICO_SERIE_FIX
    + '/datos/oportuno';

  const options = {
    method: 'get',
    muteHttpExceptions: true,
    headers: {}
  };

  if (token) {
    options.headers['Bmx-Token'] = token;
  }

  const res = UrlFetchApp.fetch(url, options);
  const code = res.getResponseCode();

  if (code < 200 || code >= 300) {
    throw new Error('Banxico respondió HTTP ' + code);
  }

  const json = JSON.parse(res.getContentText());
  const serie = json.bmx && json.bmx.series && json.bmx.series[0];
  const dato = serie && serie.datos && serie.datos[0];

  if (!dato || !dato.dato) {
    throw new Error('Respuesta Banxico sin dato de TC.');
  }

  const tc = Number(String(dato.dato).replace(',', ''));

  if (!tc || isNaN(tc)) {
    throw new Error('TC Banxico inválido.');
  }

  return {
    tc,
    fecha: normalizarFechaTipoCambio_(dato.fecha),
    fuente: TC_SOURCE_BANXICO
  };
}


function normalizarFechaTipoCambio_(fecha) {
  const raw = String(fecha || '').trim();
  const match = raw.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
  if (match) {
    return match[3] + '-' + match[2].padStart(2, '0') + '-' + match[1].padStart(2, '0');
  }
  return raw || Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
}

function getTipoCambioGoogleFinance_() {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName('_TC_TEMP');

  if (!sh) sh = ss.insertSheet('_TC_TEMP');

  sh.hideSheet();
  sh.getRange('A1').setFormula('=GOOGLEFINANCE("CURRENCY:USDMXN")');

  SpreadsheetApp.flush();
  Utilities.sleep(1500);

  const tc = Number(sh.getRange('A1').getValue());

  if (!tc || isNaN(tc)) {
    throw new Error('GOOGLEFINANCE no pudo obtener TC.');
  }

  return {
    tc,
    fecha: Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd'),
    fuente: TC_SOURCE_GOOGLEFINANCE
  };
}

function getPlanDescuentos() {
  const rng = SpreadsheetApp.getActive().getRangeByName('PLAN_DESCUENTOS');
  if (!rng) return [];
  const values = rng.getValues().filter(r => String(r[0]).trim() !== '' && String(r[1]).trim() !== '');
  if (values.length <= 1) return [];
  const rows = values.slice(1);
  return rows.map(r => ({
    categoria: String(r[0]).trim(),
    descLabel: String(r[1]).trim(),
    descLista: parseCompoundDiscount(r[1])
  }));
}

function parseCompoundDiscount(input) {
  if (input === null || input === undefined) return 0;
  const parts = String(input).split("+").map(p => parseFloat(String(p).trim()));
  if (parts.some(isNaN)) {
    const n = parseFloat(String(input).trim());
    return isNaN(n) ? 0 : (n / 100);
  }
  const factor = parts.reduce((acc, p) => acc * (1 - (p / 100)), 1);
  return 1 - factor;
}

// ===== Utilidades
function toCol_(n) {
  let s = '';
  while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); }
  return s;
}
function categoryColor_(cat) {
  const map = { 'Elite':'#ffe5e5','Platino':'#fff2e5','Oro':'#fffbe5','Plata':'#e5f2ff','En desarrollo':'#e5ffe5' };
  return map[cat] || '#ffffff';
}

// ===== NPI 2026 en columnas
const NPI_2026_SHEET_NAME = 'NPI 2026';
const NPI_2026_RECORD_PREFIX = 'RIB-DES-';
const NPI_2026_IMAGE_ROW = 4;
const NPI_2026_IMAGE_ROW_HEIGHT = 130;
const NPI_2026_IMAGE_COL_WIDTH = 160;
const NPI_2026_IMAGE_WIDTH = 140;
const NPI_2026_IMAGE_HEIGHT = 110;

const NPI_2026_EDITABLE_LABELS = [
  'Producto',
  'Código',
  'MOQ',
  'Costo USD',
  'Tipo de cambio',
  'IVA (%)',
  'Arancel (%)',
  'Margen (%)',
  'Aplicar IVA en USD',
  'Usar certificación',
  'Costo certificación MXN',
  'ROI unidades',
  'ROI precio unitario USD',
  'ROI objetivo (%)',
  'Mercado bajo MXN',
  'Mercado promedio MXN',
  'Mercado alto MXN',
  'Estatus NPI',
  'Comentarios'
];

const NPI_2026_CALCULATED_LABELS = [
  'Costo puesto USD',
  'Lista de precios final (USD sin IVA)',
  'Lista USD (con IVA)',
  'Lista MXN (sin IVA)',
  'Lista MXN (con IVA)',
  'Costo certificación USD',
  'Ingreso total USD',
  'Inversión total USD',
  'Ganancia total USD',
  '% ROI',
  'Semáforo ROI',
  'Precio sugerido MXN',
  'Posición diamante',
  'Brecha vs mercado promedio (%)',
  'Recomendación precio'
];

function getNpi2026Labels_() {
  const baseRows = [
    'Registro','Fecha','Usuario','Imagen','Estatus NPI','Comentarios','Última actualización','Producto','Código','MOQ',
    'Costo USD','Tipo de cambio','Fecha TC','Fuente TC','IVA (%)','Arancel (%)','Margen (%)','Aplicar IVA en USD',
    'Costo puesto USD',
    'Lista de precios final (USD sin IVA)',
    'Lista USD (con IVA)',
    'Lista MXN (sin IVA)',
    'Lista MXN (con IVA)',
    // Diamante de precios
    'Mercado bajo MXN',
    'Mercado promedio MXN',
    'Mercado alto MXN',
    'Precio sugerido MXN',
    'Posición diamante',
    'Brecha vs mercado promedio (%)',
    'Recomendación precio',
    // Certificación
    'Usar certificación',
    'Costo certificación MXN',
    'Costo certificación USD',
    // ROI editable
    'ROI unidades',
    'ROI precio unitario USD',
    'ROI objetivo (%)',
    // Resultados ROI
    'Ingreso total USD',
    'Inversión total USD',
    'Ganancia total USD',
    '% ROI',
    'Semáforo ROI'
  ];

  const plan = getPlanDescuentos();
  const catRows = [];
  plan.forEach(p => {
    const c = p.categoria || '';
    catRows.push(`Desc. ${c} real (%)`);
    catRows.push(`Dist. ${c} USD (sin IVA)`);
    catRows.push(`Dist. ${c} USD (con IVA)`);
    catRows.push(`Dist. ${c} MXN (sin IVA)`);
    catRows.push(`Dist. ${c} MXN (con IVA)`);
  });

  return { labels: baseRows.concat(catRows), plan };
}

function syncNpi2026Labels_(sh, labels) {
  const aliases = {
    'Mercado bajo MXN': 'Mercado bajo USD',
    'Mercado promedio MXN': 'Mercado promedio USD',
    'Mercado alto MXN': 'Mercado alto USD',
    'Precio sugerido MXN': 'Precio sugerido USD'
  };

  labels.forEach((label, i) => {
    const targetRow = i + 1;
    const lastRow = Math.max(sh.getLastRow(), 1);
    const currentLabels = sh.getRange(1, 1, lastRow, 1).getValues().map(r => String(r[0] || ''));
    const alias = aliases[label];

    if (currentLabels[targetRow - 1] === label) return;
    if (alias && currentLabels[targetRow - 1] === alias) {
      sh.getRange(targetRow, 1).setValue(label);
      return;
    }

    let existingIndex = currentLabels.indexOf(label);
    if (existingIndex < 0 && alias) existingIndex = currentLabels.indexOf(alias);
    if (existingIndex >= 0) {
      const sourceRow = existingIndex + 1;
      sh.moveRows(sh.getRange(sourceRow, 1, 1, sh.getMaxColumns()), targetRow);
      sh.getRange(targetRow, 1).setValue(label);
      return;
    }

    if (targetRow <= sh.getLastRow()) {
      sh.insertRowsBefore(targetRow, 1);
    } else if (targetRow > sh.getMaxRows()) {
      sh.insertRowsAfter(sh.getMaxRows(), targetRow - sh.getMaxRows());
    }
    sh.getRange(targetRow, 1).setValue(label);
  });
}

function ensureNpi2026_() {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(NPI_2026_SHEET_NAME) || ss.insertSheet(NPI_2026_SHEET_NAME);
  const { labels, plan } = getNpi2026Labels_();

  syncNpi2026Labels_(sh, labels);

  sh.setFrozenColumns(1);
  sh.getRange(1,1,Math.max(sh.getLastRow(), labels.length),1).setFontWeight('bold').setBackground('#fafafa');

  const rowIndex = {};
  const finalRows = sh.getRange(1,1,Math.max(sh.getLastRow(), labels.length),1).getValues();
  for (let i=0;i<finalRows.length;i++) rowIndex[String(finalRows[i][0])] = i+1;

  return { sheet: sh, rowIndex, catOrder: plan.map(p=>p.categoria) };
}

/* ============ Pintado directo + semáforo sin fórmulas ============ */

// Pinta filas de categorías SIN fórmulas (background directo)
function paintCategoryRows_(sh, rowIndex, catOrder) {
  const lastCol = Math.max(sh.getLastColumn(), 2);
  catOrder.forEach(cat => {
    const rows = [
      rowIndex[`Desc. ${cat} real (%)`],
      rowIndex[`Dist. ${cat} USD (sin IVA)`],
      rowIndex[`Dist. ${cat} USD (con IVA)`],
      rowIndex[`Dist. ${cat} MXN (sin IVA)`],
      rowIndex[`Dist. ${cat} MXN (con IVA)`]
    ].filter(Boolean);

    const bg = categoryColor_(cat);
    rows.forEach(r => {
      sh.getRange(r, 2, 1, lastCol - 1).setBackground(bg);
    });
  });

  // Columna de etiquetas
  sh.getRange(1, 1, sh.getLastRow(), 1).setBackground('#fafafa').setFontWeight('bold');
}

// Reglas SOLO para el semáforo ROI (texto OK/BAJO)
function applySemaforoRules_(sh, rowIndex) {
  const lastCols = Math.max(sh.getMaxColumns() - 1, 1);
  const rules = [];
  const semRow = rowIndex['Semáforo ROI'];
  if (semRow) {
    const rng = sh.getRange(semRow, 2, 1, lastCols);
    rules.push(
      SpreadsheetApp.newConditionalFormatRule()
        .setRanges([rng]).whenTextEqualTo('OK')
        .setBackground('#e6f4ea').setFontColor('#137333').build(),
      SpreadsheetApp.newConditionalFormatRule()
        .setRanges([rng]).whenTextEqualTo('BAJO')
        .setBackground('#fce8e6').setFontColor('#a50e0e').build()
    );
  }

  const estatusRow = rowIndex['Estatus NPI'];
  if (estatusRow) {
    const rng = sh.getRange(estatusRow, 2, 1, lastCols);
    rules.push(
      SpreadsheetApp.newConditionalFormatRule().setRanges([rng]).whenTextEqualTo('Nuevo').setBackground('#f1f3f4').setFontColor('#3c4043').build(),
      SpreadsheetApp.newConditionalFormatRule().setRanges([rng]).whenTextEqualTo('En análisis').setBackground('#fef7e0').setFontColor('#b06000').build(),
      SpreadsheetApp.newConditionalFormatRule().setRanges([rng]).whenTextEqualTo('Aprobado').setBackground('#e6f4ea').setFontColor('#137333').build(),
      SpreadsheetApp.newConditionalFormatRule().setRanges([rng]).whenTextEqualTo('Rechazado').setBackground('#fce8e6').setFontColor('#a50e0e').build(),
      SpreadsheetApp.newConditionalFormatRule().setRanges([rng]).whenTextEqualTo('Lanzado').setBackground('#e8f0fe').setFontColor('#174ea6').build()
    );
  }

  const diamanteRow = rowIndex['Posición diamante'];
  if (diamanteRow) {
    const rng = sh.getRange(diamanteRow, 2, 1, lastCols);
    rules.push(
      SpreadsheetApp.newConditionalFormatRule().setRanges([rng]).whenTextEqualTo('Agresivo').setBackground('#e8f0fe').setFontColor('#174ea6').build(),
      SpreadsheetApp.newConditionalFormatRule().setRanges([rng]).whenTextEqualTo('Competitivo').setBackground('#e6f4ea').setFontColor('#137333').build(),
      SpreadsheetApp.newConditionalFormatRule().setRanges([rng]).whenTextEqualTo('Premium').setBackground('#f3e8fd').setFontColor('#6f2da8').build(),
      SpreadsheetApp.newConditionalFormatRule().setRanges([rng]).whenTextEqualTo('Fuera de rango').setBackground('#fce8e6').setFontColor('#a50e0e').build(),
      SpreadsheetApp.newConditionalFormatRule().setRanges([rng]).whenTextEqualTo('Sin datos').setBackground('#f1f3f4').setFontColor('#3c4043').build()
    );
  }

  sh.setConditionalFormatRules(rules);
}

function aplicarFormatoEstatusNpi_(sh, rowIndex, col, estatus) {
  const row = rowIndex['Estatus NPI'];
  if (!row) return;
  const colors = {
    'Nuevo': ['#f1f3f4', '#3c4043'],
    'En análisis': ['#fef7e0', '#b06000'],
    'Aprobado': ['#e6f4ea', '#137333'],
    'Rechazado': ['#fce8e6', '#a50e0e'],
    'Lanzado': ['#e8f0fe', '#174ea6']
  };
  const pair = colors[estatus] || colors['Nuevo'];
  sh.getRange(row, col).setBackground(pair[0]).setFontColor(pair[1]).setFontWeight('bold');
}

function aplicarFormatoDiamante_(sh, rowIndex, col, posicion) {
  const row = rowIndex['Posición diamante'];
  if (!row) return;

  const colors = {
    'Agresivo': ['#e8f0fe', '#174ea6'],
    'Competitivo': ['#e6f4ea', '#137333'],
    'Premium': ['#f3e8fd', '#6f2da8'],
    'Fuera de rango': ['#fce8e6', '#a50e0e'],
    'Sin datos': ['#f1f3f4', '#3c4043']
  };

  const pair = colors[posicion] || colors['Sin datos'];

  sh.getRange(row, col)
    .setBackground(pair[0])
    .setFontColor(pair[1])
    .setFontWeight('bold');
}




function eliminarImagenesEnCelda_(sh, row, col) {
  sh.getImages().forEach(image => {
    const anchor = image.getAnchorCell();
    if (anchor && anchor.getRow() === row && anchor.getColumn() === col) {
      image.remove();
    }
  });
}


function ajustarCeldaImagen_(sh, col) {
  sh.setRowHeight(NPI_2026_IMAGE_ROW, NPI_2026_IMAGE_ROW_HEIGHT);
  sh.setColumnWidth(col, Math.max(sh.getColumnWidth(col), NPI_2026_IMAGE_COL_WIDTH));
}

function guardarImagenRegistro_(sh, row, col, imagen) {
  if (!row) return;

  eliminarImagenesEnCelda_(sh, row, col);

  const cell = sh.getRange(row, col);
  cell.clearContent().clearNote();
  ajustarCeldaImagen_(sh, col);

  if (!imagen || !imagen.dataUrl) {
    cell.setValue('Sin imagen');
    return;
  }

  const contentType = String(imagen.type || '').toLowerCase();
  const name = String(imagen.name || 'imagen-producto');
  const source = String(imagen.source || 'file');
  const size = Number(imagen.size) || 0;
  const match = String(imagen.dataUrl).match(/^data:(image\/[a-z0-9.+-]+);base64,(.+)$/i);

  if (!contentType.startsWith('image/') || !match) {
    cell.setValue('Imagen inválida');
    cell.setNote('El archivo enviado no es una imagen válida.');
    return;
  }

  const maxBytes = 2 * 1024 * 1024;
  if (size > maxBytes) {
    cell.setValue('Imagen excede 2 MB');
    cell.setNote(`Archivo rechazado: ${name} (${Math.round(size / 1024)} KB).`);
    return;
  }

  const bytes = Utilities.base64Decode(match[2]);
  if (Math.max(size, bytes.length) > maxBytes) {
    cell.setValue('Imagen excede 2 MB');
    cell.setNote(`Archivo rechazado después de validar base64: ${name} (${Math.round(bytes.length / 1024)} KB).`);
    return;
  }

  const blob = Utilities.newBlob(bytes, match[1], name);
  cell.setValue(name);
  cell.setNote(`Imagen: ${name}\nTipo: ${match[1]}\nTamaño: ${Math.round(bytes.length / 1024)} KB\nOrigen: ${source}`);
  ajustarCeldaImagen_(sh, col);
  const image = sh.insertImage(blob, col, row);
  image.setWidth(NPI_2026_IMAGE_WIDTH).setHeight(NPI_2026_IMAGE_HEIGHT);
}

/* ===================== Guardado principal (columnas) ===================== */
function formatRegistroId_(n) {
  return NPI_2026_RECORD_PREFIX + String(n).padStart(3, '0');
}

function getUltimoNumeroRegistroNpi2026_(sh) {
  const lastCol = sh.getLastColumn();
  if (lastCol < 2) return 0;

  const ids = sh.getRange(1, 2, 1, lastCol - 1).getValues()[0];
  for (let i = ids.length - 1; i >= 0; i--) {
    const match = String(ids[i] || '').trim().match(/^RIB-DES-(\d+)$/i);
    if (match) return Number(match[1]) || 0;
  }
  return 0;
}

function getSiguienteRegistroNpi2026_(sh) {
  return formatRegistroId_(getUltimoNumeroRegistroNpi2026_(sh) + 1);
}

function normalizarRegistroNpi2026_(registroId) {
  return String(registroId || '').trim().toUpperCase();
}

function findRegistroNpi2026_(registroId) {
  const target = normalizarRegistroNpi2026_(registroId);
  if (!target) return null;

  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(NPI_2026_SHEET_NAME);
  if (!sh) return null;

  const display = sh.getDataRange().getDisplayValues();
  if (!display || !display.length) return null;

  const rowIndex = {};
  display.forEach((row, i) => {
    const label = String(row[0] || '').trim();
    if (label) rowIndex[label] = i + 1;
  });

  for (let r = 0; r < display.length; r++) {
    const label = String(display[r][0] || '').trim();
    if (label !== 'Registro') continue;

    for (let c = 1; c < display[r].length; c++) {
      const raw = String(display[r][c] || '').trim();
      const current = normalizarRegistroNpi2026_(raw);

      if (current === target) {
        return {
          sheet: sh,
          rowIndex,
          col: c + 1,
          registroId: raw
        };
      }
    }
  }

  return null;
}

function getRegistroNpi2026(registroId) {
  const found = findRegistroNpi2026_(registroId);

  if (!found) {
    return null;
  }

  const sh = found.sheet;
  const col = found.col;
  const display = sh.getDataRange().getDisplayValues();

  const record = {};

  for (let r = 0; r < display.length; r++) {
    const key = String(display[r][0] || '').trim();
    if (!key) continue;

    record[key] = display[r][col - 1];
  }

  return {
    registroId: found.registroId,
    col,
    record
  };
}

function listarRegistrosNpi2026() {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName('NPI 2026');

  if (!sh) return [];

  const display = sh.getDataRange().getDisplayValues();
  if (!display || !display.length) return [];

  let registroRow = -1;
  let productoRow = -1;
  let estatusRow = -1;
  let updateRow = -1;

  for (let r = 0; r < display.length; r++) {
    const label = String(display[r][0] || '').trim();

    if (label === 'Registro') registroRow = r;
    if (label === 'Producto') productoRow = r;
    if (label === 'Estatus NPI') estatusRow = r;
    if (label === 'Última actualización') updateRow = r;
  }

  if (registroRow === -1) return [];

  const data = [];

  for (let c = 1; c < display[registroRow].length; c++) {
    const registroId = String(display[registroRow][c] || '').trim();

    if (!/^RIB-DES-\d+/i.test(registroId)) continue;

    const producto = productoRow >= 0
      ? String(display[productoRow][c] || '').trim()
      : '';

    const estatus = estatusRow >= 0
      ? String(display[estatusRow][c] || '').trim()
      : '';

    const ultimaActualizacion = updateRow >= 0
      ? String(display[updateRow][c] || '').trim()
      : '';

    data.push({
      registroId,
      producto,
      estatus,
      ultimaActualizacion,
      col: c + 1,
      label: registroId + ' — ' + (producto || 'Sin producto')
    });
  }

  return data.reverse();
}

function getEditableRowsNpi2026_(rowIndex) {
  return NPI_2026_EDITABLE_LABELS.map(label => rowIndex[label]).filter(Boolean);
}

function getCalculatedRowsNpi2026_(rowIndex) {
  const rows = NPI_2026_CALCULATED_LABELS.map(label => rowIndex[label]).filter(Boolean);
  getPlanDescuentos().forEach(p => {
    [
      `Dist. ${p.categoria} USD (sin IVA)`,
      `Dist. ${p.categoria} USD (con IVA)`,
      `Dist. ${p.categoria} MXN (sin IVA)`,
      `Dist. ${p.categoria} MXN (con IVA)`
    ].forEach(label => {
      if (rowIndex[label]) rows.push(rowIndex[label]);
    });
  });
  return rows;
}

function aplicarFormatosRegistroNpi2026_(sh, rowIndex, col) {
  const plan = getPlanDescuentos();
  const money = [
    'Costo USD','Costo puesto USD',
    'Lista de precios final (USD sin IVA)','Lista USD (con IVA)',
    'Lista MXN (sin IVA)','Lista MXN (con IVA)',
    'Costo certificación MXN','Costo certificación USD',
    'Mercado bajo MXN','Mercado promedio MXN','Mercado alto MXN','Precio sugerido MXN',
    'ROI precio unitario USD','Ingreso total USD','Inversión total USD','Ganancia total USD'
  ];
  plan.forEach(p => {
    money.push(`Dist. ${p.categoria} USD (sin IVA)`);
    money.push(`Dist. ${p.categoria} USD (con IVA)`);
    money.push(`Dist. ${p.categoria} MXN (sin IVA)`);
    money.push(`Dist. ${p.categoria} MXN (con IVA)`);
  });
  money.forEach(label => { if (rowIndex[label]) sh.getRange(rowIndex[label], col).setNumberFormat('$#,##0.00'); });
  if (rowIndex['Tipo de cambio']) sh.getRange(rowIndex['Tipo de cambio'], col).setNumberFormat('0.0000');
  ['IVA (%)','Arancel (%)','Margen (%)','Brecha vs mercado promedio (%)','ROI objetivo (%)','% ROI'].forEach(label => {
    const r = rowIndex[label]; if (r) sh.getRange(r, col).setNumberFormat('0.00');
  });
}

function aplicarFormulasRegistroNpi2026_(sh, rowIndex, col) {
  if (!sh || !rowIndex || col < 2) return;

  const plan = getPlanDescuentos();
  const L = toCol_(col);
  const ref = label => `${L}${rowIndex[label]}`;
  const setFormula = (label, formula) => {
    const row = rowIndex[label];
    if (row) sh.getRange(row, col).setFormula(formula);
  };

  const requiredLabels = [
    'Costo USD','Arancel (%)','Margen (%)','Aplicar IVA en USD','IVA (%)','Tipo de cambio',
    'Lista de precios final (USD sin IVA)','Lista USD (con IVA)','Lista MXN (sin IVA)',
    'Lista MXN (con IVA)','Usar certificación','Costo certificación MXN','ROI precio unitario USD',
    'ROI unidades','ROI objetivo (%)','Mercado bajo MXN','Mercado promedio MXN','Mercado alto MXN'
  ];
  requiredLabels.forEach(label => {
    if (!rowIndex[label]) throw new Error('Falta la fila requerida en NPI 2026: ' + label);
  });

  setFormula('Costo puesto USD', `=IF(${ref('Costo USD')}<>"",${ref('Costo USD')}*1.30*(1+${ref('Arancel (%)')}/100),)`);

  const eliteDiscountRow = rowIndex['Desc. Elite real (%)'];
  const descFactor = eliteDiscountRow ? `*(1-${L}${eliteDiscountRow}/100)` : '';
  setFormula(
    'Lista de precios final (USD sin IVA)',
    `=IF((1-${ref('Margen (%)')}/100)${descFactor}>0,${ref('Costo puesto USD')}/(1-${ref('Margen (%)')}/100)${eliteDiscountRow ? `/(1-${L}${eliteDiscountRow}/100)` : ''},)`
  );
  setFormula('Lista USD (con IVA)', `=IF(${ref('Aplicar IVA en USD')}="Sí",${ref('Lista de precios final (USD sin IVA)')}*(1+${ref('IVA (%)')}/100),${ref('Lista de precios final (USD sin IVA)')})`);
  setFormula('Lista MXN (sin IVA)', `=${ref('Lista de precios final (USD sin IVA)')}*${ref('Tipo de cambio')}`);
  setFormula('Lista MXN (con IVA)', `=${ref('Lista MXN (sin IVA)')}*(1+${ref('IVA (%)')}/100)`);

  setFormula('Precio sugerido MXN', `=IF(${ref('Mercado promedio MXN')}>0,${ref('Mercado promedio MXN')},${ref('Lista MXN (con IVA)')})`);
  setFormula('Brecha vs mercado promedio (%)', `=IF(${ref('Mercado promedio MXN')}>0,(${ref('Lista MXN (con IVA)')}-${ref('Mercado promedio MXN')})/${ref('Mercado promedio MXN')}*100,)`);
  setFormula('Posición diamante', `=IF(OR(${ref('Lista MXN (con IVA)')}="",${ref('Mercado bajo MXN')}="",${ref('Mercado promedio MXN')}="",${ref('Mercado alto MXN')}=""),"Sin datos",IF(${ref('Lista MXN (con IVA)')}<${ref('Mercado bajo MXN')},"Agresivo",IF(${ref('Lista MXN (con IVA)')}<=${ref('Mercado promedio MXN')},"Competitivo",IF(${ref('Lista MXN (con IVA)')}<=${ref('Mercado alto MXN')},"Premium","Fuera de rango"))))`);
  setFormula('Recomendación precio', `=SWITCH(${ref('Posición diamante')},"Agresivo","Precio por debajo del mercado bajo","Competitivo","Precio competitivo vs ecommerce","Premium","Precio premium vs ecommerce","Fuera de rango","Fuera de rango vs mercado final","Captura precios de mercado final")`);

  plan.forEach(p => {
    const dRow = rowIndex[`Desc. ${p.categoria} real (%)`];
    const usdSin = rowIndex[`Dist. ${p.categoria} USD (sin IVA)`];
    const usdCon = rowIndex[`Dist. ${p.categoria} USD (con IVA)`];
    const mxnSin = rowIndex[`Dist. ${p.categoria} MXN (sin IVA)`];
    const mxnCon = rowIndex[`Dist. ${p.categoria} MXN (con IVA)`];
    if (!dRow || !usdSin || !usdCon || !mxnSin || !mxnCon) return;

    sh.getRange(usdSin, col).setFormula(`=${ref('Lista de precios final (USD sin IVA)')}*(1-${L}${dRow}/100)`);
    sh.getRange(usdCon, col).setFormula(`=IF(${ref('Aplicar IVA en USD')}="Sí",${L}${usdSin}*(1+${ref('IVA (%)')}/100),${L}${usdSin})`);
    sh.getRange(mxnSin, col).setFormula(`=${L}${usdSin}*${ref('Tipo de cambio')}`);
    sh.getRange(mxnCon, col).setFormula(`=${L}${mxnSin}*(1+${ref('IVA (%)')}/100)`);
  });

  setFormula('Costo certificación USD', `=IF(${ref('Usar certificación')}="Sí",${ref('Costo certificación MXN')}/${ref('Tipo de cambio')},0)`);
  setFormula('Ingreso total USD', `=${ref('ROI precio unitario USD')}*${ref('ROI unidades')}`);
  setFormula('Inversión total USD', `=${ref('Costo puesto USD')}*${ref('ROI unidades')}+${ref('Costo certificación USD')}`);
  setFormula('Ganancia total USD', `=${ref('Ingreso total USD')}-${ref('Inversión total USD')}`);
  setFormula('% ROI', `=IF(${ref('Inversión total USD')}>0,${ref('Ganancia total USD')}/${ref('Inversión total USD')}*100,)`);
  setFormula('Semáforo ROI', `=IF(${ref('Inversión total USD')}>0,IF(${ref('% ROI')}>=${ref('ROI objetivo (%)')},"OK","BAJO"),"")`);

  aplicarFormatosRegistroNpi2026_(sh, rowIndex, col);
}

function repararFormulasNpi2026() {
  const { sheet: sh, rowIndex, catOrder } = ensureNpi2026_();
  const lastCol = sh.getLastColumn();
  if (lastCol < 2) return { ok: true, repaired: 0, sheetName: NPI_2026_SHEET_NAME };

  for (let col = 2; col <= lastCol; col++) {
    aplicarFormulasRegistroNpi2026_(sh, rowIndex, col);
  }

  paintCategoryRows_(sh, rowIndex, catOrder);
  applySemaforoRules_(sh, rowIndex);
  sh.autoResizeColumn(1);

  return { ok: true, repaired: lastCol - 1, sheetName: NPI_2026_SHEET_NAME };
}

function onEdit(e) {
  if (!e || !e.range) return;

  const sh = e.range.getSheet();
  if (sh.getName() !== NPI_2026_SHEET_NAME) return;
  if (e.range.getColumn() < 2) return;
  if (e.range.getNumRows() !== 1 || e.range.getNumColumns() !== 1) return;

  const { rowIndex } = ensureNpi2026_();
  const editableRows = getEditableRowsNpi2026_(rowIndex);

  if (!editableRows.includes(e.range.getRow())) return;

  aplicarFormulasRegistroNpi2026_(sh, rowIndex, e.range.getColumn());

  if (rowIndex['Última actualización']) {
    sh.getRange(rowIndex['Última actualización'], e.range.getColumn()).setValue(new Date());
  }
}

function escribirRegistroNpi2026_(sh, rowIndex, col, payload, options) {
  options = options || {};
  payload = payload || {};
  const plan = getPlanDescuentos();
  const modo = options.modo || 'nuevo';
  const registroId = options.registroId;
  const L = toCol_(col);
  const ref = label => `${L}${rowIndex[label]}`;
  [
  'Registro',
  'Fecha',
  'Usuario',
  'Producto',
  'Código',
  'MOQ',
  'Costo USD',
  'Tipo de cambio',
  'IVA (%)',
  'Arancel (%)',
  'Margen (%)'
].forEach(label => {
  if (!rowIndex[label]) {
    throw new Error('Falta la fila requerida en NPI 2026: ' + label);
  }
});
  const now = new Date();
  const user = (Session.getActiveUser().getEmail && Session.getActiveUser().getEmail()) || '';
  const estatusNpi = payload.estatusNpi || 'Nuevo';

  if (modo === 'nuevo') {
    sh.getRange(rowIndex['Registro'], col).setValue(registroId);
    sh.getRange(rowIndex['Fecha'], col).setValue(now);
    sh.getRange(rowIndex['Usuario'], col).setValue(user);
  } else {
    sh.getRange(rowIndex['Registro'], col).setValue(registroId);
  }
  if (rowIndex['Última actualización']) sh.getRange(rowIndex['Última actualización'], col).setValue(now);

  if (options.eliminarImagen) {
    guardarImagenRegistro_(sh, NPI_2026_IMAGE_ROW, col, null);
  } else if (options.actualizarImagen) {
    guardarImagenRegistro_(sh, NPI_2026_IMAGE_ROW, col, payload.imagen);
  }

  sh.getRange(rowIndex['Estatus NPI'], col).setValue(estatusNpi);
  aplicarFormatoEstatusNpi_(sh, rowIndex, col, estatusNpi);
  sh.getRange(rowIndex['Comentarios'], col).setValue(payload.comentariosNpi || '').setWrap(true);
  sh.getRange(rowIndex['Producto'], col).setValue(payload.producto || '');
  sh.getRange(rowIndex['Código'], col).setValue(payload.codigo || '');
  sh.getRange(rowIndex['MOQ'], col).setValue(Number(payload.moq) || 0);
  sh.getRange(rowIndex['Costo USD'], col).setValue(Number(payload.costoUSD) || 0);
  sh.getRange(rowIndex['Tipo de cambio'], col).setValue(Number(payload.tc) || 0);
  sh.getRange(rowIndex['Fecha TC'], col).setValue(payload.tcFecha || '');
  sh.getRange(rowIndex['Fuente TC'], col).setValue(payload.tcFuente || '');
  sh.getRange(rowIndex['IVA (%)'], col).setValue((Number(payload.iva)*100) || 0);
  sh.getRange(rowIndex['Arancel (%)'], col).setValue((Number(payload.arancel)*100) || 0);
  sh.getRange(rowIndex['Margen (%)'], col).setValue((Number(payload.margen)*100) || 0);
  sh.getRange(rowIndex['Aplicar IVA en USD'], col).setValue(payload.aplicarIVAenUSD ? 'Sí' : 'No');

  // Certificación (valores de entrada)
  sh.getRange(rowIndex['Usar certificación'], col).setValue(payload.usarCert ? 'Sí' : 'No');
  sh.getRange(rowIndex['Costo certificación MXN'], col).setValue(Number(payload.certMXN) || 0);

  // Descuentos por categoría (valores)
  plan.forEach(p => {
    sh.getRange(rowIndex[`Desc. ${p.categoria} real (%)`], col).setValue((Number(p.descLista)*100) || 0);
  });

  sh.getRange(rowIndex['Mercado bajo MXN'], col).setValue(Number(payload.mercadoBajoUsd) || 0);
  sh.getRange(rowIndex['Mercado promedio MXN'], col).setValue(Number(payload.mercadoPromUsd) || 0);
  sh.getRange(rowIndex['Mercado alto MXN'], col).setValue(Number(payload.mercadoAltoUsd) || 0);

  // ROI editable
  sh.getRange(rowIndex['ROI unidades'], col).setValue(Number(payload.roiUnits) || 0);

  const elitePriceCell = rowIndex['Dist. Elite USD (sin IVA)'] ? `${L}${rowIndex['Dist. Elite USD (sin IVA)']}` : '';
  if (Number(payload.roiUnitPrice)) {
    sh.getRange(rowIndex['ROI precio unitario USD'], col).setValue(Number(payload.roiUnitPrice));
  } else if (elitePriceCell) {
    sh.getRange(rowIndex['ROI precio unitario USD'], col).setFormula(`=${elitePriceCell}`);
  } else {
    sh.getRange(rowIndex['ROI precio unitario USD'], col).setValue(0);
  }

  sh.getRange(rowIndex['ROI objetivo (%)'], col).setValue(Number(payload.roiTarget) || 0);

  aplicarFormulasRegistroNpi2026_(sh, rowIndex, col);

  // Formatos
  const money = [
    'Costo USD','Costo puesto USD',
    'Lista de precios final (USD sin IVA)','Lista USD (con IVA)',
    'Lista MXN (sin IVA)','Lista MXN (con IVA)',
    'Costo certificación MXN','Costo certificación USD',
    'Mercado bajo MXN','Mercado promedio MXN','Mercado alto MXN','Precio sugerido MXN',
    'ROI precio unitario USD','Ingreso total USD','Inversión total USD','Ganancia total USD'
  ];
  plan.forEach(p => {
    money.push(`Dist. ${p.categoria} USD (sin IVA)`);
    money.push(`Dist. ${p.categoria} USD (con IVA)`);
    money.push(`Dist. ${p.categoria} MXN (sin IVA)`);
    money.push(`Dist. ${p.categoria} MXN (con IVA)`);
  });
  money.forEach(label => { if (rowIndex[label]) sh.getRange(rowIndex[label], col).setNumberFormat('$#,##0.00'); });
  sh.getRange(rowIndex['Tipo de cambio'], col).setNumberFormat('0.0000');
  ['IVA (%)','Arancel (%)','Margen (%)','Brecha vs mercado promedio (%)','ROI objetivo (%)','% ROI'].forEach(label => {
    const r = rowIndex[label]; if (r) sh.getRange(r, col).setNumberFormat('0.00');
  });

  paintCategoryRows_(sh, rowIndex, plan.map(p => p.categoria));
  applySemaforoRules_(sh, rowIndex);

  sh.autoResizeColumn(1);
  sh.autoResizeColumn(col);
  ajustarCeldaImagen_(sh, col);
}

function saveRegistroColumnar(payload) {
  const { sheet: sh, rowIndex } = ensureNpi2026_();
  const col = Math.max(2, sh.getLastColumn() + 1);
  const registroId = getSiguienteRegistroNpi2026_(sh);
  escribirRegistroNpi2026_(sh, rowIndex, col, payload || {}, {
    modo: 'nuevo',
    registroId,
    actualizarImagen: true,
    eliminarImagen: false
  });
  return { ok: true, registroId, sheetName: NPI_2026_SHEET_NAME };
}

function updateRegistroNpi2026(payload) {
  payload = payload || {};
  const registroId = normalizarRegistroNpi2026_(payload.registroId);

  if (!registroId) {
    throw new Error('Falta registroId para actualizar el registro NPI.');
  }

  const ensured = ensureNpi2026_();
  const found = findRegistroNpi2026_(registroId);

  if (!found) {
    throw new Error(`Registro no encontrado: ${registroId}`);
  }

  escribirRegistroNpi2026_(ensured.sheet, ensured.rowIndex, found.col, payload, {
    modo: 'editar',
    registroId: found.registroId,
    actualizarImagen: !!(payload.imagen && payload.imagen.dataUrl),
    eliminarImagen: payload.imagenEliminar === true
  });

  return {
    ok: true,
    registroId: found.registroId,
    sheetName: NPI_2026_SHEET_NAME,
    updated: true
  };
}


// ===== Export escenario (sin cambios)
function exportEscenario(payload) {
  const { meta, rows } = payload;
  const ss = SpreadsheetApp.getActive();
  const pad = n => String(n).padStart(2, '0');
  const now = new Date();
  const name = `Escenario_${now.getFullYear()}${pad(now.getMonth()+1)}${pad(now.getDate())}_${pad(now.getHours())}${pad(now.getMinutes())}${pad(now.getSeconds())}`;
  const sh = ss.insertSheet(name);

  const metaLabels = [
    ['Producto', meta.producto || ''],
    ['Código', meta.codigo || ''],
    ['MOQ', meta.moq || ''],
    ['TC USD/MXN', meta.tc || ''],
    ['IVA %', (meta.iva*100) || 0],
    ['Margen %', (meta.margen*100) || 0],
    ['Arancel %', (meta.arancel*100) || 0],
    ['Lista USD (sin IVA)', meta.listaUsd || 0],
    ['Redondeo', meta.redondeo || ''],
    ['IVA en USD', meta.ivaEnUSD ? 'Sí' : 'No']
  ];
  sh.getRange(1,1,metaLabels.length,2).setValues(metaLabels);

  const headers = [['Categoria','Desc. leído','Desc. real %','USD (sin IVA)','USD (con IVA)','MXN (sin IVA)','MXN (con IVA)']];
  const startRow = metaLabels.length + 2;
  sh.getRange(startRow,1,1,headers[0].length).setValues(headers);
  if (rows && rows.length) sh.getRange(startRow+1,1,rows.length,rows[0].length).setValues(rows);

  sh.getRange(8,2,1,1).setNumberFormat('$#,##0.00');
  sh.getRange(5,2,3,1).setNumberFormat('0.00');
  sh.getRange(startRow,1,1,headers[0].length).setFontWeight('bold').setBackground('#fafafa');

  if (rows && rows.length) {
    sh.getRange(startRow+1,4,rows.length,4).setNumberFormat('$#,##0.00');
    sh.getRange(startRow+1,3,rows.length,1).setNumberFormat('0.00"%"');
  }

  sh.setFrozenRows(startRow);
  sh.autoResizeColumns(1,7);

  if (rows && rows.length) {
    const range = sh.getRange(startRow+1,1,rows.length,headers[0].length);
    const bgs = range.getBackgrounds();
    for (let i = 0; i < rows.length; i++) {
      const color = categoryColor_(String(rows[i][0]||'')); 
      for (let j = 0; j < headers[0].length; j++) bgs[i][j] = color;
    }
    range.setBackgrounds(bgs);
  }
  return { sheetName: name };
}

// Compatibilidad con llamadas antiguas
function ensureHistorialCols_() { return ensureNpi2026_(); }
function saveRegistroFilas(payload) { return saveRegistroColumnar(payload); }
