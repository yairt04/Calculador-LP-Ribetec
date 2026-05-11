function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('💼 Calculadora de Precios')
    .addItem('Abrir herramienta', 'showModal')
    .addToUi();
}

function showModal() {
  const html = HtmlService.createTemplateFromFile('Sidebar').evaluate()
    .setWidth(900)
    .setHeight(780);
  SpreadsheetApp.getUi().showModalDialog(html, 'Calculadora de Precios');
}

// ===== Datos maestros
function getTipoCambio() {
  const rng = SpreadsheetApp.getActive().getRangeByName('TC_USDMXN');
  const tc = rng ? rng.getValue() : 0;
  return Number(tc) || 0;
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

// ===== Historial en columnas
function ensureHistorialCols_() {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName('HistorialCols') || ss.insertSheet('HistorialCols');

  const baseRows = [
    'Fecha','Usuario','Producto','Código','MOQ',
    'Costo USD','Tipo de cambio','IVA (%)','Arancel (%)','Margen (%)','Aplicar IVA en USD',
    'Costo puesto USD',
    'Lista de precios final (USD sin IVA)',
    'Lista USD (con IVA)',
    'Lista MXN (sin IVA)',
    'Lista MXN (con IVA)',
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

  const labels = baseRows.concat(catRows);

  if (sh.getLastRow() < labels.length) {
    sh.getRange(1, 1, labels.length, 1).setValues(labels.map(x => [x]));
  } else {
    const current = sh.getRange(1,1,labels.length,1).getValues().map(r=>String(r[0]));
    let rewrite=false;
    for (let i=0;i<labels.length;i++){ if(current[i]!==labels[i]){ rewrite=true; break; } }
    if (rewrite) sh.getRange(1, 1, labels.length, 1).setValues(labels.map(x => [x]));
  }

  sh.setFrozenColumns(1);
  sh.getRange(1,1,sh.getLastRow(),1).setFontWeight('bold').setBackground('#fafafa');

  const rowIndex = {};
  const finalRows = sh.getRange(1,1,labels.length,1).getValues();
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
  const semRow = rowIndex['Semáforo ROI'];
  if (!semRow) return;

  const rng = sh.getRange(semRow, 2, 1, lastCols);
  const rules = [
    SpreadsheetApp.newConditionalFormatRule()
      .setRanges([rng]).whenTextEqualTo('OK')
      .setBackground('#e6f4ea').setFontColor('#137333').build(),
    SpreadsheetApp.newConditionalFormatRule()
      .setRanges([rng]).whenTextEqualTo('BAJO')
      .setBackground('#fce8e6').setFontColor('#a50e0e').build()
  ];

  sh.setConditionalFormatRules(rules);
}

/* ===================== Guardado principal (columnas) ===================== */
function saveRegistroColumnar(payload) {
  const {
    producto, codigo, moq,
    costoUSD, margen, arancel, iva, tc, aplicarIVAenUSD,
    listaUSD_sinIVA,
    certMXN, usarCert,
    roiUnits, roiUnitPrice, roiTarget
  } = payload;

  const { sheet: sh, rowIndex } = ensureHistorialCols_();
  const plan = getPlanDescuentos();

  // Columna nueva
  const col = Math.max(2, sh.getLastColumn() + 1);
  const L = toCol_(col);
  const ref = label => `${L}${rowIndex[label]}`;

  // 1) Bases (valores)
  const user = (Session.getActiveUser().getEmail && Session.getActiveUser().getEmail()) || '';
  sh.getRange(rowIndex['Fecha'], col).setValue(new Date());
  sh.getRange(rowIndex['Usuario'], col).setValue(user);
  sh.getRange(rowIndex['Producto'], col).setValue(producto || '');
  sh.getRange(rowIndex['Código'], col).setValue(codigo || '');
  sh.getRange(rowIndex['MOQ'], col).setValue(moq || 0);
  sh.getRange(rowIndex['Costo USD'], col).setValue(costoUSD || 0);
  sh.getRange(rowIndex['Tipo de cambio'], col).setValue(tc || 0);
  sh.getRange(rowIndex['IVA (%)'], col).setValue((Number(iva)*100) || 0);
  sh.getRange(rowIndex['Arancel (%)'], col).setValue((Number(arancel)*100) || 0);
  sh.getRange(rowIndex['Margen (%)'], col).setValue((Number(margen)*100) || 0);
  sh.getRange(rowIndex['Aplicar IVA en USD'], col).setValue(aplicarIVAenUSD ? 'Sí' : 'No');

  // Certificación (valores de entrada)
  sh.getRange(rowIndex['Usar certificación'], col).setValue(usarCert ? 'Sí' : 'No');
  sh.getRange(rowIndex['Costo certificación MXN'], col).setValue(Number(certMXN) || 0);

  // 2) Descuentos por categoría (valores)
  plan.forEach(p => {
    sh.getRange(rowIndex[`Desc. ${p.categoria} real (%)`], col).setValue((Number(p.descLista)*100) || 0);
  });

  // 3) Fórmulas de precios base
  sh.getRange(rowIndex['Costo puesto USD'], col)
    .setFormula(`=${ref('Costo USD')}*1.30*(1+${ref('Arancel (%)')}/100)`);

  const eliteRow = `Desc. Elite real (%)`;
  sh.getRange(rowIndex['Lista de precios final (USD sin IVA)'], col)
    .setFormula(`=IF((1-${ref('Margen (%)')}/100)*(1-${L}${rowIndex[eliteRow]}/100)>0, ${ref('Costo puesto USD')}/(1-${ref('Margen (%)')}/100)/(1-${L}${rowIndex[eliteRow]}/100), )`);

  sh.getRange(rowIndex['Lista USD (con IVA)'], col)
    .setFormula(`=IF(${ref('Aplicar IVA en USD')}="Sí", ${ref('Lista de precios final (USD sin IVA)')}*(1+${ref('IVA (%)')}/100), ${ref('Lista de precios final (USD sin IVA)')})`);

  sh.getRange(rowIndex['Lista MXN (sin IVA)'], col)
    .setFormula(`=${ref('Lista de precios final (USD sin IVA)')}*${ref('Tipo de cambio')}`);
  sh.getRange(rowIndex['Lista MXN (con IVA)'], col)
    .setFormula(`=${ref('Lista MXN (sin IVA)')}*(1+${ref('IVA (%)')}/100)`);

  plan.forEach(p => {
    const dRow = rowIndex[`Desc. ${p.categoria} real (%)`];
    const usdSin = rowIndex[`Dist. ${p.categoria} USD (sin IVA)`];
    const usdCon = rowIndex[`Dist. ${p.categoria} USD (con IVA)`];
    const mxnSin = rowIndex[`Dist. ${p.categoria} MXN (sin IVA)`];
    const mxnCon = rowIndex[`Dist. ${p.categoria} MXN (con IVA)`];

    sh.getRange(usdSin, col).setFormula(`=${ref('Lista de precios final (USD sin IVA)')}*(1-${L}${dRow}/100)`);
    sh.getRange(usdCon, col).setFormula(`=IF(${ref('Aplicar IVA en USD')}="Sí", ${L}${usdSin}*(1+${ref('IVA (%)')}/100), ${L}${usdSin})`);
    sh.getRange(mxnSin, col).setFormula(`=${L}${usdSin}*${ref('Tipo de cambio')}`);
    sh.getRange(mxnCon, col).setFormula(`=${L}${mxnSin}*(1+${ref('IVA (%)')}/100)`);
  });

  // 4) Certificación USD (SI/NO)
  sh.getRange(rowIndex['Costo certificación USD'], col)
    .setFormula(`=IF(${ref('Usar certificación')}="Sí", ${ref('Costo certificación MXN')}/${ref('Tipo de cambio')}, 0)`);

  // 5) ROI editable y fórmulas (INCLUYE CERTIFICACIÓN EN LA INVERSIÓN)
  sh.getRange(rowIndex['ROI unidades'], col).setValue(Number(roiUnits) || 0);

  const elitePriceCell = `${L}${rowIndex['Dist. Elite USD (sin IVA)']}`;
  if (Number(roiUnitPrice)) {
    sh.getRange(rowIndex['ROI precio unitario USD'], col).setValue(Number(roiUnitPrice));
  } else {
    sh.getRange(rowIndex['ROI precio unitario USD'], col).setFormula(`=${elitePriceCell}`);
  }

  sh.getRange(rowIndex['ROI objetivo (%)'], col).setValue(Number(roiTarget) || 0);

  sh.getRange(rowIndex['Ingreso total USD'], col)
    .setFormula(`=${ref('ROI precio unitario USD')}*${ref('ROI unidades')}`);

  // *** Aquí se incluye el costo de certificación ***
  sh.getRange(rowIndex['Inversión total USD'], col)
    .setFormula(`=${ref('Costo puesto USD')}*${ref('ROI unidades')} + ${ref('Costo certificación USD')}`);

  sh.getRange(rowIndex['Ganancia total USD'], col)
    .setFormula(`=${ref('Ingreso total USD')}-${ref('Inversión total USD')}`);

  sh.getRange(rowIndex['% ROI'], col)
    .setFormula(`=IF(${ref('Inversión total USD')}>0, ${ref('Ganancia total USD')}/${ref('Inversión total USD')}*100, )`);

  sh.getRange(rowIndex['Semáforo ROI'], col)
    .setFormula(`=IF(${ref('% ROI')}>=${ref('ROI objetivo (%)')},"OK","BAJO")`);

  // 6) Formatos
  const money = [
    'Costo USD','Costo puesto USD',
    'Lista de precios final (USD sin IVA)','Lista USD (con IVA)',
    'Lista MXN (sin IVA)','Lista MXN (con IVA)',
    'Costo certificación MXN','Costo certificación USD',
    'ROI precio unitario USD','Ingreso total USD','Inversión total USD','Ganancia total USD'
  ];
  plan.forEach(p => {
    money.push(`Dist. ${p.categoria} USD (sin IVA)`);
    money.push(`Dist. ${p.categoria} USD (con IVA)`);
    money.push(`Dist. ${p.categoria} MXN (sin IVA)`);
    money.push(`Dist. ${p.categoria} MXN (con IVA)`);
  });
  money.forEach(label => sh.getRange(rowIndex[label], col).setNumberFormat('$#,##0.00'));
  sh.getRange(rowIndex['Tipo de cambio'], col).setNumberFormat('0.0000');
  ['IVA (%)','Arancel (%)','Margen (%)','ROI objetivo (%)','% ROI'].forEach(label => {
    const r = rowIndex[label]; if (r) sh.getRange(r, col).setNumberFormat('0.00');
  });

  // Pintado directo + reglas del semáforo (sin fórmulas complejas)
  paintCategoryRows_(sh, rowIndex, plan.map(p => p.categoria));
  applySemaforoRules_(sh, rowIndex);

  sh.autoResizeColumn(1);
  sh.autoResizeColumn(col);
  return true;
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
function saveRegistroFilas(payload) { return saveRegistroColumnar(payload); }
