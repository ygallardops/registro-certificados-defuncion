/**
 * Registro de Certificados de Defunción – Área de Archivo, Registros Médicos.
 * Archivos del proyecto: Código.gs (este) y Correo.html (diálogo para preparar el correo).
 */

const H = { ins: 'Instrucciones', cert: 'Certificados', env: 'Envíos', sol: 'Solicitantes', res: 'Resumen', lis: 'Listas' };
const COL = { azul: '#1f4e78', azulSuave: '#dce6f1', banda: '#edf2f8', auto: '#eceff3', ejemplo: '#808080' };
const ZONA = 'America/Lima';
const DNI_CERT = 4, DNI_ENV = 8; // columna del DNI en cada hoja

function onOpen() {
  SpreadsheetApp.getUi().createMenu('Registro')
    .addItem('Nuevo certificado', 'irACertificados')
    .addItem('Preparar correo de envío', 'prepararCorreo')
    .addItem('Ver envíos', 'irAEnvios')
    .addToUi();
}

/**
 * Envíos y Solicitantes: avisan antes de cualquier edición manual (se llenan con el diálogo).
 * ponytail: aviso y no bloqueo, porque el diálogo escribe con la cuenta de quien envía;
 * un bloqueo real exige publicar el registro como aplicación web que se ejecute como el propietario.
 */
function protegerHojasDeRegistro() {
  const ss = SpreadsheetApp.getActive();
  [H.env, H.sol].forEach(n => {
    const h = ss.getSheetByName(n);
    h.getProtections(SpreadsheetApp.ProtectionType.SHEET).forEach(p => p.remove());
    h.protect().setDescription('Se llena con Registro > Preparar correo de envío. Editar a mano solo si es necesario.')
      .setWarningOnly(true);
  });
}

/** Al escribir un DNI en una fila nueva, completa la fecha de hoy. */
function onEdit(e) {
  const hoja = e.range.getSheet(), nombre = hoja.getName(), fila = e.range.getRow(), c = e.range.getColumn();
  if (fila < 2 || e.range.getNumRows() > 1 || !e.value) return;
  if (nombre === H.cert && c === DNI_CERT && !hoja.getRange(fila, 2).getValue()) hoja.getRange(fila, 2).setValue(hoy_());
  if (nombre === H.env && c === DNI_ENV && !hoja.getRange(fila, 3).getValue()) hoja.getRange(fila, 3).setValue(hoy_());
}

function irACertificados() { irAPrimeraFilaVacia_(H.cert, DNI_CERT); }
function irAEnvios() { irAPrimeraFilaVacia_(H.env, DNI_ENV); }

function irAPrimeraFilaVacia_(nombre, colDni) {
  const hoja = SpreadsheetApp.getActive().getSheetByName(nombre);
  hoja.activate();
  hoja.getRange(ultimaFila_(hoja, colDni) + 1, 2).activate();
}

function ultimaFila_(hoja, col) {
  const v = hoja.getRange(2, col, hoja.getMaxRows() - 1, 1).getValues();
  let ultima = 1;
  v.forEach((r, i) => { if (r[0] !== '') ultima = i + 2; });
  return ultima;
}

function hoy_() {
  const t = Utilities.formatDate(new Date(), ZONA, 'yyyy-MM-dd').split('-');
  return new Date(+t[0], +t[1] - 1, +t[2]);
}

// ---------------------------------------------------------------- Correo de envío

/** Código correlativo por año: ENV-2026-0001, ENV-2026-0002, ... */
function siguienteCodigo_() {
  const pref = 'ENV-' + Utilities.formatDate(new Date(), ZONA, 'yyyy') + '-';
  let max = 0;
  SpreadsheetApp.getActive().getSheetByName(H.env).getRange('B2:B').getValues().forEach(r => {
    const s = String(r[0]);
    if (s.indexOf(pref) === 0) max = Math.max(max, Number(s.slice(pref.length)) || 0);
  });
  return pref + String(max + 1).padStart(4, '0');
}

function prepararCorreo() {
  const ss = SpreadsheetApp.getActive(), ui = SpreadsheetApp.getUi(), hoja = ss.getActiveSheet();
  if (hoja.getName() !== H.cert) {
    ui.alert('En la hoja Certificados, seleccione las filas a enviar (Ctrl + clic para varias) y vuelva a elegir esta opción.');
    return;
  }
  const filas = [], vistos = {};
  hoja.getActiveRangeList().getRanges().forEach(r => {
    hoja.getRange(r.getRow(), 1, r.getNumRows(), 10).getDisplayValues().forEach((v, i) => {
      if (r.getRow() + i < 2 || !v[3] || vistos[v[3]]) return;
      vistos[v[3]] = true;
      filas.push({ dni: v[3], nombre: v[4], tipo: v[6], pdf: v[9] });
    });
  });
  if (!filas.length) {
    ui.alert('Las filas seleccionadas no tienen DNI. Seleccione los certificados a enviar.');
    return;
  }
  const lis = ss.getSheetByName(H.lis), sol = ss.getSheetByName(H.sol);
  const col = c => lis.getRange(c + '2:' + c + '30').getValues().map(r => r[0]).filter(String);
  const t = HtmlService.createTemplateFromFile('Correo');
  t.datos = JSON.stringify({
    codigo: siguienteCodigo_(),
    filas: filas,
    personal: col('E'),
    medios: col('D'),
    solicitantes: sol.getRange(2, 1, Math.max(sol.getLastRow() - 1, 1), 3).getDisplayValues()
      .filter(r => r[0]).map(r => ({ nombre: r[0], area: r[1], correo: r[2] })),
    cc: [lis.getRange('H2').getValue(), lis.getRange('H3').getValue()].filter(String).join('; '),
    hoy: Utilities.formatDate(new Date(), ZONA, 'yyyy-MM-dd'),
  }).replace(/</g, '\\x3c');
  ui.showModalDialog(t.evaluate().setWidth(760).setHeight(560), 'Preparar correo de envío');
}

/** Llamado desde el diálogo: anota en Envíos una fila por certificado y guarda al solicitante si es nuevo. */
function registrarEnvio(p) {
  const lock = LockService.getDocumentLock();
  lock.waitLock(15000);
  try {
    const ss = SpreadsheetApp.getActive(), hoja = ss.getSheetByName(H.env);
    // ponytail: el código se calcula al abrir el diálogo; si otro usuario registró antes con el mismo, se pide reabrir.
    if (hoja.getRange('B2:B').getValues().some(r => r[0] === p.codigo)) {
      throw new Error('El código ' + p.codigo + ' ya lo usó otro envío. Cierre y vuelva a preparar el correo.');
    }
    const inicio = ultimaFila_(hoja, DNI_ENV) + 1, n = p.dnis.length;
    if (inicio + n - 1 > hoja.getMaxRows()) hoja.insertRowsAfter(hoja.getMaxRows(), n + 100);
    const f = p.fecha ? p.fecha.split('-') : null;
    const fechaPedido = f ? new Date(+f[0], +f[1] - 1, +f[2]) : '';
    hoja.getRange(inicio, 2, n, 7).setValues(p.dnis.map(d => [p.codigo, hoy_(), p.sol, p.area, p.medio, fechaPedido, d]));
    hoja.getRange(inicio, 10, n, 2).setValues(p.dnis.map(() => [p.para, p.por]));

    const sol = ss.getSheetByName(H.sol);
    const existe = sol.getRange('A2:A').getValues().some(r => String(r[0]).trim().toLowerCase() === p.sol.trim().toLowerCase());
    if (!existe) sol.getRange(ultimaFila_(sol, 1) + 1, 1, 1, 3).setValues([[p.sol.trim(), p.area, p.para]]);
    return n;
  } finally {
    lock.releaseLock();
  }
}

/** Llamado desde el diálogo: deja el correo en Borradores de quien lo usa, con su firma de Gmail. */
function crearBorrador(p) {
  const esc = s => String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
  let firma = '';
  try {
    const cuenta = Gmail.Users.Settings.SendAs.list('me').sendAs.filter(s => s.isDefault)[0];
    firma = (cuenta && cuenta.signature) || '';
  } catch (e) { /* sin servicio Gmail o sin firma: el borrador sale sin firma */ }
  GmailApp.createDraft(p.para, p.asunto, p.cuerpo, {
    cc: p.cc.replace(/;/g, ','),
    htmlBody: '<div>' + esc(p.cuerpo).replace(/\n/g, '<br>') + '</div>' + (firma ? '<br>' + firma : ''),
  });
}

// ---------------------------------------------------------------- Configuración

function configurar() {
  const ss = SpreadsheetApp.getActive(), ui = SpreadsheetApp.getUi();
  const tieneDatos = (nombre, col) => {
    const h = ss.getSheetByName(nombre);
    return h && h.getMaxRows() > 2 && h.getRange(3, col, h.getMaxRows() - 2, 1).getValues().some(r => r[0] !== '');
  };
  if (tieneDatos(H.cert, DNI_CERT) || tieneDatos(H.env, 7) || tieneDatos(H.env, DNI_ENV)) {
    ui.alert('La plantilla ya tiene registros. No se hizo ningún cambio.');
    return;
  }
  ss.setSpreadsheetLocale('es_PE');
  ss.setSpreadsheetTimeZone(ZONA);
  const tmp = ss.insertSheet('tmp_' + Date.now());
  ss.getSheets().forEach(h => { if (h.getSheetId() !== tmp.getSheetId()) ss.deleteSheet(h); });
  SpreadsheetApp.flush();
  const nueva = nombre => { const h = ss.insertSheet(nombre); SpreadsheetApp.flush(); return h; };
  const ins = nueva(H.ins), cert = nueva(H.cert), env = nueva(H.env), sol = nueva(H.sol);
  const res = nueva(H.res), lis = nueva(H.lis);
  ss.deleteSheet(tmp);

  armarListas_(lis);
  armarSolicitantes_(sol);
  armarCertificados_(cert, lis);
  armarEnvios_(env, lis, sol);
  armarResumen_(res);
  armarInstrucciones_(ins);

  ins.setTabColor('#999999'); lis.setTabColor('#999999');
  cert.setTabColor(COL.azul); env.setTabColor('#38761d'); sol.setTabColor('#6aa84f'); res.setTabColor('#e69138');

  proteger_(ins);
  proteger_(lis);
  proteger_(res, [res.getRange('C4:C5')]);
  protegerHojasDeRegistro();
  ss.setActiveSheet(cert);
  ui.alert('Plantilla lista. Borre la fila de ejemplo (fila 2 de Certificados) al empezar a usarla.');
}

function armarListas_(lis) {
  const datos = [
    ['Tipo de documento', 'Origen', 'Sí / No', 'Medio del pedido', 'Personal de Archivo', '', 'Configuración', 'Valor'],
    ['Certificado de Defunción', 'Hospitalización', 'Sí', 'WhatsApp', 'Personal 1', '', 'Correo genérico de la Unidad (CC)', ''],
    ['Informe de Defunción', 'Emergencia', 'No', 'Correo institucional', 'Personal 2', '', 'Correo de la Jefatura (CC)', ''],
    ['Ambos', '', '', 'Presencial', 'Personal 3', '', '', ''],
    ['', '', '', 'Documento / Memorando', 'Personal 4', '', '', ''],
    ['', '', '', 'Teléfono', 'Personal 5', '', '', ''],
    ['', '', '', '', 'Personal 6', '', '', ''],
    ['', '', '', '', 'Practicante (P)', '', '', ''],
  ];
  lis.getRange(1, 1, datos.length, 8).setValues(datos);
  encabezado_(lis.getRange('A1:E1'));
  encabezado_(lis.getRange('G1:H1'));
  lis.getRange('H2:H3').setBackground('#fff2cc');
  lis.setColumnWidths(1, 5, 200).setColumnWidth(6, 30).setColumnWidth(7, 240).setColumnWidth(8, 260);
  lis.getRange('A10').setValue('Puede agregar opciones al final de cada columna. Los correos de CC se usan al preparar el correo de envío.')
    .setFontColor(COL.ejemplo).setFontStyle('italic');
  lis.getDataRange().setFontFamily('Arial').setVerticalAlignment('middle');
  lis.setRowHeights(1, 10, 24);
  lis.setFrozenRows(1);
}

function armarSolicitantes_(sol) {
  tabla_(sol, [['Solicitante (ej.: Dra. Ana López)', 260], ['Área / Servicio', 200], ['Correo institucional', 280]], 300);
  avisar_(sol.getRange('A1:C1'), 'Encabezados: no modificar');
  sol.getRange('A1').setNote('Se agregan solos al registrar un envío con un solicitante nuevo. También puede escribirlos aquí.');
}

function armarCertificados_(cert, lis) {
  const cols = [
    ['N.º', 45], ['Fecha de recepción', 105], ['N.º HC', 95], ['DNI', 90], ['Apellidos y nombres', 270],
    ['Fecha de fallecimiento', 115], ['Tipo de documento', 175], ['Origen', 120], ['¿Escaneado?', 90],
    ['Nombre del PDF (automático)', 360], ['Entregado a Epidemiología (fecha)', 135], ['Responsable', 160],
    ['Observaciones', 220],
  ];
  tabla_(cert, cols);
  cert.getRange('A1').setFormula('={"N.º";ARRAYFORMULA(IF(D2:D="","",ROW(D2:D)-1))}');
  // APELLIDOS-NOMBRES_HC_DNI_AAAA-MM-DD.pdf, sin tildes ni Ñ
  const sinTildes = 'SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(UPPER(TRIM(E2:E)),' +
    '"Á","A"),"É","E"),"Í","I"),"Ó","O"),"Ú","U"),"Ü","U"),"Ñ","N")';
  cert.getRange('J1').setFormula(
    '={"Nombre del PDF (automático)";ARRAYFORMULA(IF((D2:D="")+(E2:E="")+(F2:F=""),"",' +
    'REGEXREPLACE(REGEXREPLACE(' + sinTildes + ',"[^A-Z0-9]+","-"),"^-|-$","")' +
    '&IF(C2:C="","","_"&C2:C)&"_"&D2:D&"_"&YEAR(F2:F)&"-"&TEXT(MONTH(F2:F),"00")&"-"&TEXT(DAY(F2:F),"00")&".pdf"))}');
  cert.getRange('C2:D').setNumberFormat('@');
  ['B', 'F', 'K'].forEach(c => fecha_(cert, c));
  lista_(cert, 'G', lis.getRange('A2:A30'));
  lista_(cert, 'H', lis.getRange('B2:B30'));
  lista_(cert, 'I', lis.getRange('C2:C30'));
  lista_(cert, 'L', lis.getRange('E2:E30'));
  dni_(cert, 'D');
  ['A', 'J'].forEach(c => cert.getRange(c + '2:' + c).setBackground(COL.auto).setFontColor('#444444'));
  ['A', 'B', 'C', 'D', 'F', 'I', 'K'].forEach(c => cert.getRange(c + '2:' + c).setHorizontalAlignment('center'));
  cert.getRange('E1').setNote('Formato: APELLIDOS, Nombres (ej.: PÉREZ GÓMEZ, Juan)');
  cert.getRange('J1').setNote('Apellidos y nombres _ N.º HC _ DNI _ fecha de fallecimiento. Use este nombre al guardar el PDF.');
  cert.setConditionalFormatRules([
    SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('No')
      .setBackground('#f4cccc').setFontColor('#990000').setRanges([cert.getRange('I2:I')]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=AND($D2<>"",$K2="")')
      .setBackground('#fff2cc').setRanges([cert.getRange('K2:K')]).build(),
  ]);
  cert.getRange('B2:I2').setValues([[new Date(2026, 10, 20), '001234', '12345678', 'PÉREZ GÓMEZ, Juan',
    new Date(2026, 10, 18), 'Certificado de Defunción', 'Emergencia', 'Sí']]);
  cert.getRange('K2:M2').setValues([[new Date(2026, 10, 20), 'Personal 1', 'EJEMPLO – borrar esta fila']]);
  cert.getRange('B2:M2').setFontStyle('italic');
  cert.setFrozenColumns(5);
  avisar_(cert.getRange('A1:M1'), 'Encabezados y fórmulas: no modificar');
  avisar_(cert.getRange('J2:J'), 'Columna automática: no escribir');
}

function armarEnvios_(env, lis, sol) {
  const cols = [
    ['N.º', 45], ['Código de envío', 120], ['Fecha de envío', 100], ['Solicitante', 190], ['Área / Servicio', 150],
    ['Medio del pedido', 150], ['Fecha del pedido', 100], ['DNI', 90], ['Apellidos y nombres (automático)', 270],
    ['Correo de destino', 230], ['Enviado por', 160], ['Observaciones', 220],
  ];
  tabla_(env, cols);
  env.getRange('A1').setFormula('={"N.º";ARRAYFORMULA(IF(H2:H="","",ROW(H2:H)-1))}');
  env.getRange('I1').setFormula(
    '={"Apellidos y nombres (automático)";ARRAYFORMULA(IF(H2:H="","",IFERROR(VLOOKUP(H2:H,' +
    '{Certificados!D2:D,Certificados!E2:E},2,FALSE),"No está en Certificados")))}');
  env.getRange('H2:H').setNumberFormat('@');
  ['C', 'G'].forEach(c => fecha_(env, c));
  env.getRange('D2:D').setDataValidation(SpreadsheetApp.newDataValidation()
    .requireValueInRange(sol.getRange('A2:A300'), true).setAllowInvalid(true).build());
  lista_(env, 'F', lis.getRange('D2:D30'));
  lista_(env, 'K', lis.getRange('E2:E30'));
  dni_(env, 'H');
  ['A', 'I'].forEach(c => env.getRange(c + '2:' + c).setBackground(COL.auto).setFontColor('#444444'));
  ['A', 'B', 'C', 'G', 'H'].forEach(c => env.getRange(c + '2:' + c).setHorizontalAlignment('center'));
  env.getRange('B2:B').setFontWeight('bold').setFontColor(COL.azul);
  env.setConditionalFormatRules([
    SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('No está en Certificados')
      .setFontColor('#cc0000').setRanges([env.getRange('I2:I')]).build(),
  ]);
  env.setFrozenColumns(2);
  avisar_(env.getRange('A1:L1'), 'Encabezados y fórmulas: no modificar');
  avisar_(env.getRange('I2:I'), 'Columna automática: no escribir');
}

function armarResumen_(res) {
  res.setHiddenGridlines(true);
  res.setColumnWidth(1, 24).setColumnWidth(2, 400).setColumnWidth(3, 120);
  res.getRange('B2:C2').merge().setValue('Resumen mensual').setBackground(COL.azul).setFontColor('#ffffff')
    .setFontSize(14).setFontWeight('bold').setVerticalAlignment('middle');
  res.setRowHeight(2, 40);
  res.getRange('B4:C5').setValues([['Año', 2026], ['Mes (1 a 12)', 11]]);
  res.getRange('C4:C5').setBackground('#fff2cc').setFontWeight('bold').setHorizontalAlignment('center')
    .setBorder(true, true, true, true, true, true, '#bf9000', SpreadsheetApp.BorderStyle.SOLID);
  res.getRange('C5').setDataValidation(SpreadsheetApp.newDataValidation()
    .requireNumberBetween(1, 12).setAllowInvalid(false).build());
  res.getRange('B6').setValue('← Cambie el año y el mes en las celdas amarillas').setFontColor(COL.ejemplo).setFontStyle('italic');
  res.getRange('E1:F2').setValues([['Desde', ''], ['Hasta (antes de)', '']]);
  res.getRange('F1:F2').setFormulas([['=DATE(C4,C5,1)'], ['=DATE(C4,C5+1,1)']]);
  res.hideColumns(5, 2);
  const r = 'Certificados!B2:B,">="&$F$1,Certificados!B2:B,"<"&$F$2';
  const e = 'Envíos!C2:C,">="&$F$1,Envíos!C2:C,"<"&$F$2';
  const filas = [
    ['Certificados recibidos en Archivo', '=COUNTIFS(' + r + ')', true],
    ['Certificado de Defunción', '=COUNTIFS(' + r + ',Certificados!G2:G,"Certificado de Defunción")'],
    ['Solo Informe de Defunción', '=COUNTIFS(' + r + ',Certificados!G2:G,"Informe de Defunción")'],
    ['Ambos documentos', '=COUNTIFS(' + r + ',Certificados!G2:G,"Ambos")'],
    ['De Hospitalización', '=COUNTIFS(' + r + ',Certificados!H2:H,"Hospitalización")'],
    ['De Emergencia', '=COUNTIFS(' + r + ',Certificados!H2:H,"Emergencia")'],
    ['Escaneados', '=COUNTIFS(' + r + ',Certificados!I2:I,"Sí")'],
    ['Entregados a Epidemiología', '=COUNTIFS(' + r + ',Certificados!K2:K,"<>")'],
    ['Pendientes de entregar a Epidemiología', '=COUNTIFS(' + r + ',Certificados!K2:K,"")'],
    ['Certificados enviados por correo', '=COUNTIFS(' + e + ')', true],
    ['Correos enviados (códigos de envío)', '=COUNTUNIQUEIFS(Envíos!B2:B,' + e + ')'],
  ];
  filas.forEach((f, i) => {
    const fila = 8 + i;
    res.getRange(fila, 2).setValue(f[2] ? f[0] : '      ' + f[0]);
    res.getRange(fila, 3).setFormula(f[1]).setHorizontalAlignment('center');
    if (f[2]) res.getRange(fila, 2, 1, 2).setFontWeight('bold').setBackground(COL.azulSuave);
  });
  res.getRange(8, 2, filas.length, 2).setBorder(true, true, true, true, false, true, '#c9d3df', SpreadsheetApp.BorderStyle.SOLID);
  res.setRowHeights(8, filas.length, 26);
  res.getRange('B20').setValue('Mientras no se borre, la fila de ejemplo de Certificados también se cuenta.')
    .setFontColor(COL.ejemplo).setFontStyle('italic');
  res.getDataRange().setFontFamily('Arial').setVerticalAlignment('middle');
}

function armarInstrucciones_(ins) {
  ins.setHiddenGridlines(true);
  ins.setColumnWidth(1, 24).setColumnWidth(2, 880);
  const texto = [
    ['Registro de Certificados de Defunción – Área de Archivo, Registros Médicos', 't'],
    ['', ''],
    ['Registro diario', 's'],
    ['1. Menú Registro > Nuevo certificado: lleva a la primera fila libre de Certificados.', ''],
    ['2. Escriba el DNI: la fecha de recepción se completa sola con la fecha de hoy.', ''],
    ['3. Apellidos y nombres en una sola columna, con este formato: PÉREZ GÓMEZ, Juan.', ''],
    ['4. Guarde el PDF escaneado con el nombre que aparece en la columna "Nombre del PDF".', ''],
    ['', ''],
    ['Enviar certificados por correo', 's'],
    ['1. En Certificados, seleccione las filas a enviar (Ctrl + clic para elegir varias).', ''],
    ['2. Menú Registro > Preparar correo de envío. Se asigna un código único (ej.: ENV-2026-0001).', ''],
    ['3. Elija o escriba el solicitante: si ya está en la hoja Solicitantes, el área y el correo se completan solos.', ''],
    ['4. Pulse "1. Crear borrador en Gmail": el correo queda en Borradores con Para, CC, asunto y cuerpo.', ''],
    ['5. En Gmail, adjunte los PDF de la lista y envíelo. Luego pulse "2. Registrar envío": se anota solo, y si el solicitante es nuevo se guarda en Solicitantes.', ''],
    ['6. Para ubicar un envío, busque su código en la hoja Envíos o en el correo (Ctrl + F).', ''],
    ['', ''],
    ['Colores', 's'],
    ['Gris: se calcula solo, no escribir. Rojo: certificado no escaneado. Amarillo: falta entregar a Epidemiología.', ''],
    ['', ''],
    ['Reglas', 's'],
    ['• Envíos solo por correo institucional, con copia al correo genérico de la Unidad y a la Jefatura. Nunca por WhatsApp.', ''],
    ['• Contiene datos personales y de salud (Ley N.º 29733): compartir solo con el personal autorizado.', ''],
  ];
  ins.getRange(2, 2, texto.length, 1).setValues(texto.map(t => [t[0]]));
  ins.getRange(2, 2, texto.length, 1).setFontFamily('Arial').setFontSize(11).setVerticalAlignment('middle');
  ins.setRowHeights(2, texto.length, 24);
  texto.forEach((t, i) => {
    const c = ins.getRange(2 + i, 2);
    if (t[1] === 't') { c.setBackground(COL.azul).setFontColor('#ffffff').setFontSize(15).setFontWeight('bold'); ins.setRowHeight(2 + i, 44); }
    if (t[1] === 's') c.setBackground(COL.azulSuave).setFontColor(COL.azul).setFontWeight('bold');
  });
}

// ---------------------------------------------------------------- Utilidades

function encabezado_(rango) {
  rango.setBackground(COL.azul).setFontColor('#ffffff').setFontWeight('bold').setFontFamily('Arial')
    .setHorizontalAlignment('center').setVerticalAlignment('middle').setWrap(true);
}

function tabla_(hoja, cols, filas) {
  const n = cols.length, total = filas || 1000;
  if (hoja.getMaxColumns() > n) hoja.deleteColumns(n + 1, hoja.getMaxColumns() - n);
  if (hoja.getMaxRows() < total) hoja.insertRowsAfter(hoja.getMaxRows(), total - hoja.getMaxRows());
  if (hoja.getMaxRows() > total) hoja.deleteRows(total + 1, hoja.getMaxRows() - total);
  hoja.getRange(1, 1, 1, n).setValues([cols.map(c => c[0])]);
  cols.forEach((c, i) => hoja.setColumnWidth(i + 1, c[1]));
  const todo = hoja.getRange(1, 1, hoja.getMaxRows(), n);
  todo.applyRowBanding(SpreadsheetApp.BandingTheme.BLUE, true, false)
    .setHeaderRowColor(COL.azul).setFirstRowColor('#ffffff').setSecondRowColor(COL.banda);
  todo.setFontFamily('Arial').setFontSize(10).setVerticalAlignment('middle');
  encabezado_(hoja.getRange(1, 1, 1, n));
  hoja.setRowHeight(1, 44);
  hoja.setRowHeights(2, hoja.getMaxRows() - 1, 22);
  hoja.setFrozenRows(1);
  todo.createFilter();
}

function fecha_(hoja, col) {
  hoja.getRange(col + '2:' + col).setNumberFormat('dd/mm/yyyy').setDataValidation(
    SpreadsheetApp.newDataValidation().requireDate().setAllowInvalid(false)
      .setHelpText('Escriba una fecha dd/mm/aaaa.').build());
}

function lista_(hoja, col, rango) {
  hoja.getRange(col + '2:' + col).setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInRange(rango, true).setAllowInvalid(false).build());
}

function dni_(hoja, col) {
  hoja.getRange(col + '2:' + col).setDataValidation(
    SpreadsheetApp.newDataValidation().requireFormulaSatisfied('=OR(' + col + '2="",LEN(' + col + '2)=8)')
      .setAllowInvalid(true).setHelpText('El DNI tiene 8 dígitos.').build());
}

/** Bloquea la hoja: solo el propietario puede editarla. */
function proteger_(hoja, libres) {
  const p = hoja.protect().setDescription('Solo el propietario edita esta hoja');
  if (libres) p.setUnprotectedRanges(libres);
  p.removeEditors(p.getEditors());
  if (p.canDomainEdit()) p.setDomainEdit(false);
}

/** Muestra una advertencia antes de editar (no bloquea). */
function avisar_(rango, desc) {
  rango.protect().setDescription(desc).setWarningOnly(true);
}
