/**
 * Prueba de equivalencia: src/Reservations.js (optimizado) debe dejar la hoja
 * exactamente igual que el script original (tests/fixtures/legacy-reservations.js),
 * pero leyendo la hoja completa una sola vez.
 *
 *   node tests/reservations.test.js
 */
const assert = require('assert');
const fs = require('fs');
const path = require('path');
const vm = require('vm');

const TZ = 'America/Mazatlan';

/* ---------- Stubs de Apps Script ---------- */
function pad(n) { return String(n).padStart(2, '0'); }
function formatDate(date, tz, fmt) {
  return fmt
    .replace('yyyy', date.getFullYear())
    .replace('MM', pad(date.getMonth() + 1))
    .replace('dd', pad(date.getDate()));
}

function createSheet(initialRows) {
  const cells = initialRows.map(r => r.slice());
  const format = {}; // fila -> {background, font, weight}
  const stats = { fullReads: 0, rowReads: 0 };
  const lastRow = () => {
    for (let r = cells.length - 1; r >= 0; r--) {
      if ((cells[r] || []).some(v => v !== '' && v !== undefined)) return r + 1;
    }
    return 0;
  };
  const lastCol = () => Math.max(0, ...cells.map(r => {
    for (let c = (r || []).length - 1; c >= 0; c--) if (r[c] !== '' && r[c] !== undefined) return c + 1;
    return 0;
  }));
  const read = (row, col, nr, nc) => Array.from({ length: nr }, (_, i) =>
    Array.from({ length: nc }, (_, j) => {
      const v = (cells[row - 1 + i] || [])[col - 1 + j];
      return v === undefined ? '' : v;
    }));
  const sheet = {
    cells, format, stats,
    getLastRow: lastRow,
    getDataRange: () => ({
      getValues: () => { stats.fullReads++; return read(1, 1, lastRow(), lastCol()); }
    }),
    appendRow: values => { cells[lastRow()] = Array.from(values); }, // copia en arrays de este contexto
    setFrozenRows: () => {},
    autoResizeColumn: () => {},
    getRange: (row, col, nr = 1, nc = 1) => ({
      getValues: () => { stats.rowReads++; return read(row, col, nr, nc); },
      setValues: values => values.forEach((rv, i) => rv.forEach((v, j) => {
        cells[row - 1 + i] = cells[row - 1 + i] || [];
        cells[row - 1 + i][col - 1 + j] = v;
      })),
      setBackground: v => { (format[row] = format[row] || {}).background = v; },
      setFontColor: v => { (format[row] = format[row] || {}).font = v; },
      setFontWeight: v => { (format[row] = format[row] || {}).weight = v; }
    })
  };
  return sheet;
}

function thread(messages) {
  return { getMessages: () => messages, messages };
}

function msg(id, date, html, plain = '') {
  return { getId: () => id, getDate: () => date, getBody: () => html, getPlainBody: () => plain };
}

function runScript(file, { sheetRows, threadsBySubject }) {
  const sheet = createSheet(sheetRows);
  const toasts = [];
  const context = {
    console,
    Logger: { log: () => {} },
    Session: { getScriptTimeZone: () => TZ },
    Utilities: { formatDate },
    SpreadsheetApp: {
      getActiveSpreadsheet: () => ({
        getSheetByName: name => (name === 'RESERVAS' ? sheet : null),
        insertSheet: () => sheet,
        toast: m => toasts.push(m)
      })
    },
    GmailApp: {
      search: query => {
        const key = Object.keys(threadsBySubject).find(subject => query.includes(subject));
        return key ? threadsBySubject[key] : [];
      },
      getMessagesForThreads: threads => threads.map(t => t.messages)
    }
  };
  vm.createContext(context);
  vm.runInContext(fs.readFileSync(file, 'utf8'), context);
  context.extractReservations();
  return { sheet, toasts };
}

/* ---------- Datos de prueba ---------- */
// Template anterior (el que entiende extractData)
function legacyEmail({ header = 'You have a new reservation', unit, checkIn, checkOut, guest, guests, nights, code, source }) {
  return `<div class="header"><p>${header}</p></div><div class="content">
    <div class="property-name">${unit}</div>
    <div class="dates"><div>CHECK-IN</div><div>${checkIn}</div><div>CHECK-OUT</div><div>${checkOut}</div></div>
    <div><span>Guest Name:</span><span>${guest}</span></div>
    <div><span>Number of Guests:</span><span>${guests}</span></div>
    <div><span>Number of Nights:</span><span>${nights}</span></div>
    <div><span>Confirmation Code:</span><span>${code}</span></div>
    <div><span>Source:</span><span>${source}</span></div></div>`;
}

function guestyReport(rows) {
  const header = '<tr><td>CONFIRMATION DATE</td>' + '<td>x</td>'.repeat(11) + '</tr>';
  return '<table>' + header + rows.map(r => '<tr>' + r.map(c => `<td><span>${c}</span></td>`).join('') + '</tr>').join('') + '</table>';
}

const HEADERS = ['Unidad', 'Guest', 'Check-in', 'Check-out', 'Number of nights', 'Number of guests',
  'Confirmation code', 'Source', 'status', 'fechaMail'];

function buildSheetRows(n) {
  const rows = [HEADERS];
  for (let i = 0; i < n; i++) {
    rows.push([`A0${i}-LAP-UNIT`, `Guest ${i}`, '01/01/2026', '03/01/2026', 2, 2, `HMCODE${i}`, 'Airbnb', 'Confirmada', '01/12/2025']);
  }
  rows[50][6] = 'HMCODE10';   // código duplicado: debe usarse la PRIMERA fila (11)
  rows[70][6] = '';           // fila sin código
  rows[71][6] = '';
  return rows;
}

const ownerRezTemplate = fs.readFileSync(path.join(__dirname, 'fixtures/new-reservation.html'), 'utf8')
  .replace(/\{BLISTINGTOKEN\}/g, 'HMNEWOWNERREZ').replace(/\{[A-Z]+\}/g, 'x');

const threadsBySubject = {
  'Congrats! You have a new reservation by La Paz Bay': [
    thread([
      msg('m1', new Date(2026, 8, 20), legacyEmail({ unit: 'Casa Hispania 334', checkIn: 'Sep 23rd, 2026', checkOut: 'Sep 26, 2026',
        guest: 'Julio Cesar', guests: '4', nights: '3', code: 'HMN8MSB98K', source: 'Airbnb' })),
      msg('m2', new Date(2026, 8, 21), legacyEmail({ unit: 'Depto Legaspy 1', checkIn: 'Oct 13, 2026', checkOut: 'Oct 22, 2026',
        guest: 'Blanca Estela Vera', guests: '4', nights: '9', code: 'HMCODE10', source: 'Airbnb' }))
    ]),
    // Mismo código dos veces en la misma ejecución: la segunda debe actualizar la fila recién agregada
    thread([msg('m3', new Date(2026, 8, 22), legacyEmail({ unit: 'Casa Hispania 334', checkIn: 'Sep 23, 2026', checkOut: 'Sep 27, 2026',
      guest: 'Julio Cesar', guests: '5', nights: '4', code: 'HMN8MSB98K', source: 'Airbnb' }))]),
    // Template nuevo de OwnerRez: el extractor anterior no lo entiende (código vacío)
    thread([msg('m4', new Date(2026, 8, 23), ownerRezTemplate)])
  ],
  'Reservation update notification by La Paz Bay': [
    thread([msg('m5', new Date(2026, 8, 23), legacyEmail({ header: 'Your reservation was updated', unit: 'Unit 5', checkIn: 'Jan 5, 2026',
      checkOut: 'Jan 9, 2026', guest: 'Guest 5', guests: '3', nights: '4', code: 'HMCODE5', source: 'VRBO' }))])
  ],
  'Reservation Cancelled notification by La Paz Bay': [
    thread([msg('m6', new Date(2026, 8, 24), legacyEmail({ header: 'Your reservation was cancelled', unit: 'Unit 7', checkIn: 'Jan 1, 2026',
      checkOut: 'Jan 3, 2026', guest: 'Guest 7', guests: '2', nights: '2', code: 'HMCODE7', source: 'Airbnb' }))])
  ],
  'Guesty shared A096 - Check In Next 7 Days (Tiffany)': [
    thread([msg('g1', new Date(2026, 8, 25), guestyReport([
      ['2026-09-01 10:00 AM', 'A096-GUESTY-1', '2026-09-28 03:00 PM', '2026-09-30 11:00 AM', '2', 'Ana', '2', '555', 'GY-NEW-1', '', 'Booking.com', 'Ana'],
      // Reserva que ya existe en la hoja
      ['2026-08-01 10:00 AM', 'A020-LAP-UNIT', '2026-01-01 03:00 PM', '2026-01-03 11:00 AM', '2', 'Guest 20', '2', '555', 'HMCODE20', '', 'Airbnb', ''],
      // Reserva que modificó un correo en esta misma ejecución (debe leer el valor actualizado)
      ['2026-08-02 10:00 AM', 'Unit 5', '2026-01-05 03:00 PM', '2026-01-09 11:00 AM', '4', 'Guest 5', '3', '555', 'HMCODE5', '', 'VRBO', ''],
      // Nueva reserva agregada por un correo en esta misma ejecución
      ['2026-09-20 10:00 AM', 'Casa Hispania 334', '2026-09-23 03:00 PM', '2026-09-27 11:00 AM', '4', 'Julio Cesar', '5', '555', 'HMN8MSB98K', '', 'Airbnb', ''],
      // Dos veces el mismo código nuevo en el reporte
      ['2026-09-02 10:00 AM', 'A096-GUESTY-2', '2026-09-29 03:00 PM', '2026-10-01 11:00 AM', '2', 'Luis', '2', '555', 'GY-NEW-2', '', 'Airbnb', ''],
      ['2026-09-02 10:00 AM', 'A096-GUESTY-2', '2026-09-29 03:00 PM', '2026-10-02 11:00 AM', '3', 'Luis', '2', '555', 'GY-NEW-2', '', 'Airbnb', '']
    ]))])
  ]
};

/* ---------- Ejecución ---------- */
const ROWS = 10000;
const legacyFile = path.join(__dirname, 'fixtures/legacy-reservations.js');
const optimizedFile = path.join(__dirname, '../src/Reservations.js');

let t = Date.now();
const legacy = runScript(legacyFile, { sheetRows: buildSheetRows(ROWS), threadsBySubject });
const legacyMs = Date.now() - t;
t = Date.now();
const optimized = runScript(optimizedFile, { sheetRows: buildSheetRows(ROWS), threadsBySubject });
const optimizedMs = Date.now() - t;

assert.deepStrictEqual(optimized.sheet.cells, legacy.sheet.cells, 'la hoja debe quedar idéntica');
assert.deepStrictEqual(optimized.sheet.format, legacy.sheet.format, 'los colores deben quedar idénticos');
assert.deepStrictEqual(optimized.toasts, legacy.toasts, 'el resumen debe ser idéntico');
assert.strictEqual(optimized.sheet.stats.fullReads, 1, 'la versión optimizada lee la hoja completa una sola vez');

// Sanidad: los casos borde realmente se ejercitaron
const byCode = code => legacy.sheet.cells.findIndex(r => r[6] === code) + 1;
assert.strictEqual(legacy.sheet.cells[11][1], 'Blanca Estela Vera', 'duplicado: se actualizó la primera fila (HMCODE10)');
assert.strictEqual(legacy.sheet.cells[50][1], 'Guest 49', 'la fila duplicada posterior no se toca');
assert.notStrictEqual(legacy.sheet.cells[70][1], 'Guest 69', 'el correo sin código sobrescribió la primera fila sin código (comportamiento original)');
assert.ok(byCode('GY-NEW-1') > ROWS, 'Guesty agregó al final');
assert.strictEqual(legacy.sheet.format[byCode('HMCODE7')].background, 'red');

console.log('✓ hoja, colores y resumen idénticos al script original');
console.log(`  Lecturas completas de la hoja: original ${legacy.sheet.stats.fullReads}, optimizado ${optimized.sheet.stats.fullReads}`);
console.log(`  Tiempo local con ${ROWS} filas: original ${legacyMs} ms, optimizado ${optimizedMs} ms`);
console.log(`  Resumen: ${optimized.toasts[0].replace('\n', ' | ')}`);
