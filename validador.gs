/**
 * VALIDADOR BULL PUT · Hoja de cálculo
 * ------------------------------------
 * 1) Cambia CLAVE por una palabra tuya (la usará la app para leer los datos).
 * 2) Ejecuta "configurar" una vez: crea las pestañas H_xxx, Cálculos y Salida,
 *    y (solo si no existen) Earnings y FED, que rellenas tú a mano.
 * 3) Espera ~1 minuto a que carguen los datos y ejecuta "comprobar".
 */

const CLAVE = 'Opciones@2027!';

// Subyacente y tipo de reglas (AAPL usa reglas de índice)
const TICKERS = [
  ['IWM', 'Índice'],
  ['SPY', 'Índice'],
  ['QQQ', 'Índice'],
  ['AAPL', 'Índice'],
  ['NVDA', 'Growth'],
  ['INTC', 'Growth']
];

const FILAS_HIST = 170; // filas con fórmulas en cada pestaña H_xxx

// ===================== CONFIGURACIÓN =====================

function configurar() {
  const ss = SpreadsheetApp.getActive();
  // Las fórmulas están escritas en formato inglés (comas y punto decimal).
  // Cambiamos la hoja a inglés mientras se escriben y luego la devolvemos
  // a su idioma: Google traduce las fórmulas automáticamente.
  const idioma = ss.getSpreadsheetLocale();
  ss.setSpreadsheetLocale('en_US');
  SpreadsheetApp.flush();
  try {
    TICKERS.forEach(([t]) => crearHistorico_(ss, t));
    crearEarnings_(ss);
    crearFed_(ss);
    crearCalculos_(ss);
    crearSalida_(ss);
    SpreadsheetApp.flush();
  } finally {
    ss.setSpreadsheetLocale(idioma);
  }
  Logger.log('Hoja configurada. Espera un minuto y ejecuta "comprobar".');
}

// Histórico diario (~150 sesiones) + estocástico 10,3,3
function crearHistorico_(ss, t) {
  const nombre = 'H_' + t;
  let h = ss.getSheetByName(nombre);
  if (h) h.clear(); else h = ss.insertSheet(nombre);

  // A:F = Fecha, Apertura, Máximo, Mínimo, Cierre, Volumen
  h.getRange('A1').setFormula(
    '=GOOGLEFINANCE("' + t + '","all",WORKDAY(TODAY(),-150),TODAY())');

  h.getRange('H1:K1').setValues([['%K bruto (10)', '%K (3)', '%D (3)', 'Cruce al alza']]);

  // %K bruto 10 sesiones: (cierre - mín10) / (máx10 - mín10)
  h.getRange(11, 8, FILAS_HIST - 10, 1).setFormulaR1C1(
    '=IF(RC5="","",IFERROR((RC5-MIN(R[-9]C4:RC4))/(MAX(R[-9]C3:RC3)-MIN(R[-9]C4:RC4))*100,50))');
  // %K suavizado 3
  h.getRange(13, 9, FILAS_HIST - 12, 1).setFormulaR1C1(
    '=IF(RC8="","",AVERAGE(R[-2]C8:RC8))');
  // %D = media 3 del %K
  h.getRange(15, 10, FILAS_HIST - 14, 1).setFormulaR1C1(
    '=IF(RC9="","",AVERAGE(R[-2]C9:RC9))');
  // Cruce: hoy %K > %D y ayer %K <= %D
  h.getRange(16, 11, FILAS_HIST - 15, 1).setFormulaR1C1(
    '=IF(RC10="","",IF(AND(RC9>RC10,R[-1]C9<=R[-1]C10),1,0))');
}

// Fechas de earnings, escritas a mano. Esta pestaña NO se borra al volver a
// ejecutar "configurar": solo se crea si no existe.
function crearEarnings_(ss) {
  if (ss.getSheetByName('Earnings')) return;
  const e = ss.insertSheet('Earnings');
  e.getRange('A1:C1').setValues([['Ticker', 'Fecha earnings', 'Nota']]);
  TICKERS.forEach(([t], i) => {
    const etf = ['IWM', 'SPY', 'QQQ'].indexOf(t) >= 0;
    e.getRange(i + 2, 1, 1, 3).setValues([[t, '', etf ? 'ETF: sin earnings' : '']]);
  });
  e.getRange('B2:B' + (TICKERS.length + 1)).setNumberFormat('dd/mm/yyyy');
  e.setFrozenRows(1);
}

// Fechas de decisión de la FED (FOMC). La FED publica cada verano el
// calendario del año siguiente: federalreserve.gov/monetarypolicy/fomccalendars.htm
// Columna A = día del comunicado (2º día de la reunión, 14:00 h Nueva York).
// Igual que Earnings: solo se crea si no existe y luego la mantienes tú.
function crearFed_(ss) {
  if (ss.getSheetByName('FED')) return;
  const f = ss.insertSheet('FED');
  const filas = [
    ['Fecha decisión', 'Proyecciones (SEP)', 'Nota'],
    [new Date(2026, 9, 28), 'No', 'Reunión 27-28 oct'],
    [new Date(2026, 11, 9), 'Sí', 'Reunión 8-9 dic'],
    [new Date(2027, 0, 27), 'No', 'Reunión 26-27 ene · provisional'],
    [new Date(2027, 2, 17), 'Sí', 'Reunión 16-17 mar · provisional'],
    [new Date(2027, 3, 28), 'No', 'Reunión 27-28 abr · provisional'],
    [new Date(2027, 5, 9), 'Sí', 'Reunión 8-9 jun · provisional'],
    [new Date(2027, 6, 28), 'No', 'Reunión 27-28 jul · provisional'],
    [new Date(2027, 8, 15), 'Sí', 'Reunión 14-15 sep · provisional'],
    [new Date(2027, 9, 27), 'No', 'Reunión 26-27 oct · provisional'],
    [new Date(2027, 11, 8), 'Sí', 'Reunión 7-8 dic · provisional']
  ];
  f.getRange(1, 1, filas.length, 3).setValues(filas);
  f.getRange('A2:A60').setNumberFormat('dd/mm/yyyy');
  f.setFrozenRows(1);
}

// Lista de fechas FED para la app (se ignoran filas vacías o sin fecha)
function leerFed_() {
  const f = SpreadsheetApp.getActive().getSheetByName('FED');
  if (!f || f.getLastRow() < 2) return [];
  return f.getRange(2, 1, f.getLastRow() - 1, 2).getValues()
    .filter(r => r[0] instanceof Date)
    .map(r => ({ fecha: r[0], sep: String(r[1]).toLowerCase().startsWith('s') }));
}

// Una fila por subyacente con todos los indicadores
function crearCalculos_(ss) {
  let c = ss.getSheetByName('Cálculos');
  if (c) c.clear(); else c = ss.insertSheet('Cálculos');

  const cab = ['ticker', 'tipo', 'precio', 'fila', 'mm21', 'mm52',
    'sobreMM52', 'sobreMM21', 'tendencia', 'k', 'd', 'kAyer', 'dAyer',
    'estocasticoOK', 'max10', 'filaMax', 'minPosterior', 'retroceso',
    'pullbackOK', 'mm52semanas', 'ultimoCierre', 'earnings'];
  c.getRange(1, 1, 1, cab.length).setValues([cab]);

  TICKERS.forEach(([t, tipo], i) => {
    const r = i + 2;
    const h = "'H_" + t + "'!";
    const fila = [
      t,
      tipo,
      `=GOOGLEFINANCE(A${r},"price")`,
      `=COUNT(${h}A:A)+1`,
      `=AVERAGE(OFFSET(${h}E1,D${r}-21,0,21,1))`,
      `=AVERAGE(OFFSET(${h}E1,D${r}-52,0,52,1))`,
      `=C${r}>F${r}`,
      `=C${r}>E${r}`,
      // Tendencia: bajo MM52 = FALLO; sobre MM52 pero bajo MM21 = AVISO
      `=IF(NOT(G${r}),"FALLO",IF(H${r},"OK","AVISO"))`,
      `=INDEX(${h}I:I,D${r})`,
      `=INDEX(${h}J:J,D${r})`,
      `=INDEX(${h}I:I,D${r}-1)`,
      `=INDEX(${h}J:J,D${r}-1)`,
      // Cruce al alza en las 3 últimas sesiones, %K > %D y ambos subiendo
      `=AND(SUM(OFFSET(${h}K1,D${r}-3,0,3,1))>0,J${r}>K${r},J${r}>L${r},K${r}>M${r})`,
      // Pullback: máximo de 10 sesiones y mínimo posterior
      `=MAX(OFFSET(${h}C1,D${r}-10,0,10,1))`,
      `=D${r}-10+MATCH(O${r},OFFSET(${h}C1,D${r}-10,0,10,1),0)`,
      `=MIN(MIN(OFFSET(${h}D1,P${r}-1,0,D${r}-P${r}+1,1)),C${r})`,
      `=(O${r}-Q${r})/O${r}`,
      `=AND(R${r}>=0.05,C${r}<O${r},G${r})`,
      // Media de 52 semanas (regla INTC)
      `=AVERAGE(QUERY(GOOGLEFINANCE(A${r},"close",TODAY()-400,TODAY(),"WEEKLY"),"select Col2 order by Col1 desc limit 52",1))`,
      `=INDEX(${h}A:A,D${r})`,
      // Fecha de earnings desde la pestaña Earnings (vacío si no hay)
      `=IFERROR(1/(1/VLOOKUP(A${r},Earnings!A:B,2,FALSE)),"")`
    ];
    // Texto con setValues; fórmulas con setFormulas (sintaxis inglesa,
    // independiente de que la hoja esté configurada en español)
    c.getRange(r, 1, 1, 2).setValues([fila.slice(0, 2)]);
    c.getRange(r, 3, 1, fila.length - 2).setFormulas([fila.slice(2)]);
  });
  c.getRange(2, 22, TICKERS.length, 1).setNumberFormat('yyyy-mm-dd');
}

// Solo referencias a Cálculos: es lo que lee la app
function crearSalida_(ss) {
  let s = ss.getSheetByName('Salida');
  if (s) s.clear(); else s = ss.insertSheet('Salida');
  const filas = TICKERS.length + 1;
  const cols = 22;
  s.getRange(1, 1, filas, cols).setFormulaR1C1("='Cálculos'!RC");
  // La fecha de earnings debe salir como fecha para que la app la entienda
  s.getRange(2, 22, TICKERS.length, 1).setNumberFormat('yyyy-mm-dd');
}

// ===================== COMPROBACIÓN =====================

function comprobar() {
  const d = SpreadsheetApp.getActive().getSheetByName('Salida').getDataRange().getValues();
  const cab = d[0];
  d.slice(1).forEach(f => {
    Logger.log(
      f[0] + ' | precio ' + red_(f[2]) +
      ' | MM21 ' + red_(f[4]) + ' | MM52 ' + red_(f[5]) +
      ' | tendencia ' + f[8] +
      ' | %K ' + red_(f[9]) + ' %D ' + red_(f[10]) +
      ' | estocástico ' + (f[13] ? 'OK' : 'no') +
      ' | retroceso ' + red_(f[17] * 100) + '%' +
      ' | pullback ' + (f[18] ? 'sí' : 'no') +
      ' | MM52sem ' + red_(f[19]) +
      ' | earnings ' + (f[21] instanceof Date ? f[21].toLocaleDateString('es-ES') : '—'));
  });
}

function red_(x) {
  return (typeof x === 'number') ? x.toFixed(2) : String(x);
}

// ===================== DATOS PARA LA APP =====================
// Se activa al publicar el script como aplicación web (siguiente paso).

function doGet(e) {
  if (!e || !e.parameter || e.parameter.k !== CLAVE) {
    return json_({ error: 'clave incorrecta' });
  }
  const d = SpreadsheetApp.getActive().getSheetByName('Salida').getDataRange().getValues();
  const cab = d.shift();
  const subyacentes = d.map(f => Object.fromEntries(cab.map((k, i) =>
    // earnings como 'aaaa-mm-dd' para que la zona horaria no cambie el día
    [k, k === 'earnings' ? dia_(f[i]) : f[i]])));
  const fed = leerFed_().map(x => ({ fecha: dia_(x.fecha), sep: x.sep }));
  return json_({ actualizado: new Date(), subyacentes: subyacentes, fed: fed });
}

function dia_(v) {
  if (!(v instanceof Date)) return '';
  return Utilities.formatDate(v, SpreadsheetApp.getActive().getSpreadsheetTimeZone(), 'yyyy-MM-dd');
}

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o))
    .setMimeType(ContentService.MimeType.JSON);
}