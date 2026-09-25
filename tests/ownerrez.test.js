/**
 * Pruebas locales del script OwnerRez (node tests/ownerrez.test.js).
 * Carga src/OwnerRez.js en un contexto con stubs de los servicios de Apps Script.
 */
const assert = require('assert');
const fs = require('fs');
const path = require('path');
const vm = require('vm');

const TEMPLATE = fs.readFileSync(path.join(__dirname, 'fixtures/new-reservation.html'), 'utf8');
const SENDER = 'La Paz Bay Rentals <oru5b14b666b9x@inquiryspot.com>';

function fillTemplate(values) {
  return TEMPLATE.replace(/\{([A-Z]+)\}/g, (_, key) => values[key] !== undefined ? values[key] : '');
}

/* ---------- Stubs de Sheets ---------- */
function createFakeSheet(name) {
  const cells = [];   // cells[r][c], 0-indexed
  const colors = [];
  const sheet = {
    name,
    cells,
    colors,
    getLastRow: () => {
      for (let r = cells.length - 1; r >= 0; r--) {
        if ((cells[r] || []).some(v => v !== '' && v !== undefined)) return r + 1;
      }
      return 0;
    },
    getMaxRows: () => Math.max(cells.length, 1000),
    setFrozenRows: () => {},
    getRange: (row, col, numRows = 1, numCols = 1) => ({
      getValues: () => Array.from({ length: numRows }, (_, i) =>
        Array.from({ length: numCols }, (_, j) => {
          const v = (cells[row - 1 + i] || [])[col - 1 + j];
          return v === undefined ? '' : v;
        })),
      setValues: values => values.forEach((rowValues, i) => rowValues.forEach((v, j) => {
        cells[row - 1 + i] = cells[row - 1 + i] || [];
        cells[row - 1 + i][col - 1 + j] = v;
      })),
      setBackgrounds: values => values.forEach((rowValues, i) => { colors[row - 1 + i] = rowValues[0]; }),
      setFontColors: () => {},
      setFontWeight: () => {},
      setNumberFormat: () => {}
    })
  };
  return sheet;
}

function createContext({ threads = [], sheets = {} } = {}) {
  const props = {};
  const spreadsheet = {
    getSheetByName: name => sheets[name] || null,
    insertSheet: name => (sheets[name] = createFakeSheet(name)),
    toast: () => {}
  };
  const context = {
    console,
    Logger: { log: () => {} },
    LockService: { getScriptLock: () => ({ tryLock: () => true, releaseLock: () => {} }) },
    PropertiesService: {
      getScriptProperties: () => ({
        getProperty: k => (k in props ? props[k] : null),
        setProperty: (k, v) => { props[k] = v; },
        deleteProperty: k => { delete props[k]; }
      })
    },
    SpreadsheetApp: { getActiveSpreadsheet: () => spreadsheet },
    GmailApp: {
      queries: [],
      search(query, start, max) {
        this.queries.push(query);
        return threads.slice(start, start + max);
      },
      getMessagesForThreads: ts => ts.map(t => t.messages)
    }
  };
  vm.createContext(context);
  vm.runInContext(fs.readFileSync(path.join(__dirname, '../src/OwnerRez.js'), 'utf8'), context);
  context.__props = props;
  context.__sheets = sheets;
  return context;
}

function message({ id, subject, date, html, from = SENDER }) {
  return {
    getId: () => id,
    getSubject: () => subject,
    getDate: () => date,
    getBody: () => html,
    getPlainBody: () => '',
    getFrom: () => from
  };
}

const NEW_SUBJECT = 'Congrats! You have a new reservation by La Paz Bay';
const casaHispania = {
  PDISPNAME: 'Casa Hispania 334', BARR: 'Sep 23, 2026', BDEP: 'Sep 26, 2026',
  CFULL: 'Julio Cesar Nuño Becerra', BNGUEST: '4', BNNGHTS: '3',
  BLISTINGTOKEN: 'HMN8MSB98K', BSOURCE: 'Airbnb', MYCO: 'La Paz Bay Rentals'
};

let passed = 0;
function test(name, fn) {
  fn();
  passed++;
  console.log('✓ ' + name);
}

/* ---------- Parser ---------- */
test('extrae todos los campos del template de nueva reserva', () => {
  const ctx = createContext();
  const r = ctx.ownerRezParseEmail_(NEW_SUBJECT, fillTemplate(casaHispania), '');
  assert.strictEqual(r.unidad, 'Casa Hispania 334');
  assert.strictEqual(r.guest, 'Julio Cesar Nuño Becerra');
  assert.strictEqual(r.numberOfGuests, 4);
  assert.strictEqual(r.numberOfNights, 3);
  assert.strictEqual(r.confirmationCode, 'HMN8MSB98K');
  assert.strictEqual(r.source, 'Airbnb');
  assert.strictEqual(r.status, 'Confirmada');
  assert.strictEqual(r.checkIn.getFullYear(), 2026);
  assert.strictEqual(r.checkIn.getMonth(), 8);
  assert.strictEqual(r.checkIn.getDate(), 23);
  assert.strictEqual(r.checkOut.getDate(), 26);
});

test('decodifica entidades HTML y detecta status por asunto', () => {
  const ctx = createContext();
  const html = fillTemplate({ ...casaHispania, CFULL: 'Julio Nu&ntilde;o &amp; Co', PDISPNAME: 'Depto Legaspy 1' });
  assert.strictEqual(ctx.ownerRezParseEmail_(NEW_SUBJECT, html, '').guest, 'Julio Nuño & Co');
  assert.strictEqual(ctx.ownerRezParseEmail_(NEW_SUBJECT, html, '').unidad, 'Depto Legaspy 1');
  assert.strictEqual(ctx.ownerRezParseEmail_('Reservation update notification by La Paz Bay', html, '').status, 'Modificado');
  assert.strictEqual(ctx.ownerRezParseEmail_('Reservation Cancelled notification by La Paz Bay', html, '').status, 'Cancelada');
});

test('omite correos que no son reservas', () => {
  const ctx = createContext();
  assert.strictEqual(ctx.ownerRezParseEmail_('New inquiry', '<p>Hello, is it available?</p>', ''), null);
  const unknownSubject = fillTemplate(casaHispania)
    .replace('You have a new reservation', 'Monthly summary')
    .replace('New reservation at', 'Summary for');
  assert.strictEqual(ctx.ownerRezParseEmail_('Monthly summary', unknownSubject, ''), null);
});

test('usa el texto plano si el HTML no sirve', () => {
  const ctx = createContext();
  const plain = 'Good News!\nYou have a new reservation\nCasa X\nCheck-in\nOct 13, 2026\nCheck-out\nOct 22, 2026\n' +
    'Guest Name: Blanca Estela Vera\nNumber of Guests: 4\nNumber of Nights: 9\nConfirmation Code: HMPQMPD4JC\nSource: Airbnb';
  const r = ctx.ownerRezParseEmail_(NEW_SUBJECT, '', plain);
  assert.strictEqual(r.unidad, 'Casa X');
  assert.strictEqual(r.guest, 'Blanca Estela Vera');
  assert.strictEqual(r.confirmationCode, 'HMPQMPD4JC');
  assert.strictEqual(r.numberOfNights, 9);
});

test('formatos de fecha', () => {
  const ctx = createContext();
  const iso = d => d && typeof d.getFullYear === 'function' ? `${d.getFullYear()}-${d.getMonth() + 1}-${d.getDate()}` : d;
  assert.strictEqual(iso(ctx.ownerRezParseDate_('Sep 23, 2026')), '2026-9-23');
  assert.strictEqual(iso(ctx.ownerRezParseDate_('Sept 3rd, 2026')), '2026-9-3');
  assert.strictEqual(iso(ctx.ownerRezParseDate_('Wed, Jan 21st, 2026')), '2026-1-21');
  assert.strictEqual(iso(ctx.ownerRezParseDate_('23 Sep 2026')), '2026-9-23');
  assert.strictEqual(iso(ctx.ownerRezParseDate_('2026-10-13')), '2026-10-13');
  assert.strictEqual(ctx.ownerRezParseDate_('pronto'), 'pronto');
});

/* ---------- Sincronización completa ---------- */
function buildThreads() {
  const legaspy = { ...casaHispania, PDISPNAME: 'Depto Legaspy 1', BARR: 'Oct 13, 2026', BDEP: 'Oct 22, 2026',
    CFULL: 'Blanca Estela Vera', BNNGHTS: '9', BLISTINGTOKEN: 'HMPQMPD4JC' };
  // Gmail devuelve de más nuevo a más viejo
  return [
    { messages: [message({ id: 'm3', subject: 'Reservation Cancelled notification by La Paz Bay',
      date: new Date(2026, 8, 24, 10), html: fillTemplate({ ...casaHispania, BNNGHTS: '' }) })] },
    { messages: [message({ id: 'm2', subject: NEW_SUBJECT, date: new Date(2026, 8, 22, 18, 25), html: fillTemplate(legaspy) })] },
    { messages: [
      message({ id: 'm1', subject: NEW_SUBJECT, date: new Date(2026, 8, 20, 9), html: fillTemplate(casaHispania) }),
      message({ id: 'r1', subject: 'Re: ' + NEW_SUBJECT, date: new Date(2026, 8, 20, 10), html: '<p>gracias</p>', from: 'alex@lapazbay.com' })
    ] }
  ];
}

test('backfill: crea la hoja, inserta y aplica la cancelación aunque llegue primero', () => {
  const ctx = createContext({ threads: buildThreads() });
  ctx.ownerRezSync();

  const sheet = ctx.__sheets.OwnerRez;
  assert.ok(sheet, 'la hoja OwnerRez debe crearse');
  assert.deepStrictEqual(sheet.cells[0].slice(0, 3), ['Unidad', 'Guest', 'Check-in']);
  assert.strictEqual(sheet.getLastRow(), 3, 'encabezado + 2 reservas');

  const byCode = Object.fromEntries(sheet.cells.slice(1).map((row, i) => [row[6], { row, color: sheet.colors[i + 1] }]));
  assert.strictEqual(byCode.HMN8MSB98K.row[8], 'Cancelada');
  assert.strictEqual(byCode.HMN8MSB98K.row[4], '', 'la cancelación sin noches no borra dato (no había previo)');
  assert.strictEqual(byCode.HMN8MSB98K.row.length, 10, 'sin columna messageId');
  assert.strictEqual(byCode.HMN8MSB98K.color, 'red');
  assert.strictEqual(byCode.HMPQMPD4JC.row[0], 'Depto Legaspy 1');
  assert.strictEqual(byCode.HMPQMPD4JC.row[8], 'Confirmada');
  assert.strictEqual(byCode.HMPQMPD4JC.color, 'white');
  assert.strictEqual(ctx.__props.OWNERREZ_BACKFILL_DONE, 'true');
});

test('orden cronológico normal: la cancelación conserva campos que no trae', () => {
  const threads = buildThreads().reverse();
  const ctx = createContext({ threads });
  ctx.ownerRezSync();
  const row = ctx.__sheets.OwnerRez.cells.slice(1).find(r => r[6] === 'HMN8MSB98K');
  assert.strictEqual(row[8], 'Cancelada');
  assert.strictEqual(row[4], 3, 'conserva Number of nights del correo original');
});

test('segunda ejecución: incremental, sin duplicados', () => {
  const threads = buildThreads();
  const ctx = createContext({ threads });
  ctx.ownerRezSync();
  ctx.GmailApp.queries.length = 0;
  ctx.ownerRezSync();

  assert.ok(ctx.GmailApp.queries.every(q => /after:\d+/.test(q)), 'usa búsqueda incremental');
  assert.strictEqual(ctx.__sheets.OwnerRez.getLastRow(), 3, 'sin filas duplicadas');
});

test('si se vacía la hoja, vuelve a recorrer todo el histórico', () => {
  const ctx = createContext({ threads: buildThreads() });
  ctx.ownerRezSync();
  const sheet = ctx.__sheets.OwnerRez;
  sheet.cells.splice(1); // el usuario borra todas las reservas
  ctx.GmailApp.queries.length = 0;
  ctx.ownerRezSync();

  assert.ok(ctx.GmailApp.queries.some(q => !/after:/.test(q)), 'hace búsqueda completa');
  assert.strictEqual(sheet.getLastRow(), 3, 'recupera las 2 reservas');
});

test('respeta encabezados y filas existentes en la hoja', () => {
  const existing = createFakeSheet('OwnerRez');
  existing.cells[0] = ['Unidad', 'Guest', 'Check-in', 'Check-out', 'Number of nights', 'Number of guests',
    'Confirmation code', 'Source', 'status', 'fechaMail'];
  existing.cells[1] = ['Casa Hispania 334', 'Julio', '', '', 3, 4, 'HMN8MSB98K', 'Airbnb', 'Confirmada', '19/09/2026'];
  const ctx = createContext({ threads: buildThreads(), sheets: { OwnerRez: existing } });
  ctx.ownerRezSync();

  assert.strictEqual(existing.cells[0][10], undefined, 'no agrega columnas extra');
  assert.strictEqual(existing.getLastRow(), 3);
  assert.strictEqual(existing.cells[1][8], 'Cancelada', 'actualiza la fila existente en su lugar');
});

console.log(`\n${passed} pruebas OK`);
