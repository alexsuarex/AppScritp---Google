/**
 * OwnerRez → Google Sheets
 *
 * Rastrea TODOS los correos enviados por el remitente de notificaciones de
 * OwnerRez (La Paz Bay Rentals), extrae los datos de la reserva y los guarda
 * en la hoja "OwnerRez".
 *
 * - Primera ejecución: recorre todo el histórico del remitente (backfill) en
 *   bloques; si no termina dentro del límite de tiempo, continúa en la
 *   siguiente ejecución donde se quedó.
 * - Ejecuciones siguientes: solo revisa los correos recientes (incremental).
 * - Una reserva se identifica por su Confirmation Code: si llega un correo
 *   más nuevo (modificación / cancelación) se actualiza la misma fila.
 *
 * Todas las funciones y constantes llevan el prefijo "ownerRez" para que este
 * archivo pueda convivir en el mismo proyecto de Apps Script con otros
 * scripts sin choques de nombres.
 */

const OWNERREZ_CONFIG = {
  SENDER: 'oru5b14b666b9x@inquiryspot.com',
  SHEET_NAME: 'OwnerRez',
  HEADERS: [
    'Unidad',
    'Guest',
    'Check-in',
    'Check-out',
    'Number of nights',
    'Number of guests',
    'Confirmation code',
    'Source',
    'status',
    'fechaMail'
  ],
  STATUS: {
    NEW: 'Confirmada',
    MODIFIED: 'Modificado',
    CANCELLED: 'Cancelada'
  },
  // Colores por status (los que no aparecen aquí quedan blanco / texto negro)
  COLORS: {
    Modificado: { background: 'orange', font: 'white' },
    Cancelada: { background: 'red', font: 'white' }
  },
  DATE_FORMAT: 'dd/mm/yyyy',
  PAGE_SIZE: 100,                   // hilos por página de búsqueda en Gmail
  TIME_BUDGET_MS: 4.5 * 60 * 1000,  // Apps Script corta a los 6 min
  INCREMENTAL_BUFFER_DAYS: 2,       // margen al buscar correos recientes
  TRIGGER_EVERY_HOURS: 1
};

const OWNERREZ_COL = {
  UNIDAD: 0,
  GUEST: 1,
  CHECK_IN: 2,
  CHECK_OUT: 3,
  NIGHTS: 4,
  GUESTS: 5,
  CODE: 6,
  SOURCE: 7,
  STATUS: 8,
  FECHA_MAIL: 9
};

const OWNERREZ_PROPS = {
  BACKFILL_OFFSET: 'OWNERREZ_BACKFILL_OFFSET',
  BACKFILL_DONE: 'OWNERREZ_BACKFILL_DONE',
  LAST_SYNC: 'OWNERREZ_LAST_SYNC'
};

const OWNERREZ_MONTHS = {
  jan: 0, feb: 1, mar: 2, apr: 3, may: 4, jun: 5,
  jul: 6, aug: 7, sep: 8, oct: 9, nov: 10, dec: 11,
  ene: 0, abr: 3, ago: 7, dic: 11
};

/* ======================================================================
 * Punto de entrada
 * ==================================================================== */

/**
 * Sincroniza los correos de OwnerRez con la hoja. Es la función que ejecuta
 * el trigger y el menú.
 */
function ownerRezSync() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30 * 1000)) {
    Logger.log('OwnerRez: ya hay una sincronización en curso, se omite esta ejecución.');
    return;
  }

  try {
    const startedAt = Date.now();
    const props = PropertiesService.getScriptProperties();
    const store = ownerRezLoadStore_(ownerRezGetSheet_());
    const stats = { read: 0, added: 0, updated: 0, unchanged: 0, skipped: 0, errors: 0 };

    let backfillDone = props.getProperty(OWNERREZ_PROPS.BACKFILL_DONE) === 'true';
    if (!backfillDone) {
      backfillDone = ownerRezBackfill_(store, props, stats, startedAt);
    }
    if (backfillDone) {
      ownerRezIncremental_(store, props, stats, startedAt);
    }

    const summary =
      `${stats.added} nuevas, ${stats.updated} actualizadas, ${stats.unchanged} sin cambios, ` +
      `${stats.skipped} omitidas, ${stats.errors} errores (${stats.read} correos leídos)` +
      (backfillDone ? '' : '. Histórico incompleto: se continuará en la próxima ejecución.');
    Logger.log('OwnerRez: ' + summary);
    ownerRezToast_(summary);
  } finally {
    lock.releaseLock();
  }
}

/**
 * Recorre el histórico completo del remitente, de más nuevo a más viejo.
 * Devuelve true cuando terminó todo el histórico.
 */
function ownerRezBackfill_(store, props, stats, startedAt) {
  // Lo que llegue mientras dura el backfill lo recoge después el incremental.
  if (!props.getProperty(OWNERREZ_PROPS.LAST_SYNC)) {
    props.setProperty(OWNERREZ_PROPS.LAST_SYNC, String(startedAt));
  }

  const query = `from:${OWNERREZ_CONFIG.SENDER}`;
  let offset = Number(props.getProperty(OWNERREZ_PROPS.BACKFILL_OFFSET) || 0);

  while (Date.now() - startedAt < OWNERREZ_CONFIG.TIME_BUDGET_MS) {
    const threads = GmailApp.search(query, offset, OWNERREZ_CONFIG.PAGE_SIZE);
    if (threads.length === 0) {
      props.setProperty(OWNERREZ_PROPS.BACKFILL_DONE, 'true');
      props.deleteProperty(OWNERREZ_PROPS.BACKFILL_OFFSET);
      Logger.log('OwnerRez: histórico completo.');
      return true;
    }

    ownerRezProcessThreads_(threads, store, stats);
    ownerRezFlush_(store);

    offset += threads.length;
    props.setProperty(OWNERREZ_PROPS.BACKFILL_OFFSET, String(offset));
    Logger.log(`OwnerRez: histórico, ${offset} hilos revisados.`);
  }

  return false;
}

/**
 * Revisa solo los correos desde la última sincronización (con un margen).
 */
function ownerRezIncremental_(store, props, stats, startedAt) {
  const lastSync = Number(props.getProperty(OWNERREZ_PROPS.LAST_SYNC) || startedAt);
  const bufferMs = OWNERREZ_CONFIG.INCREMENTAL_BUFFER_DAYS * 24 * 60 * 60 * 1000;
  const afterSeconds = Math.floor((lastSync - bufferMs) / 1000);
  const query = `from:${OWNERREZ_CONFIG.SENDER} after:${afterSeconds}`;

  let start = 0;
  while (true) {
    const threads = GmailApp.search(query, start, OWNERREZ_CONFIG.PAGE_SIZE);
    if (threads.length === 0) break;
    ownerRezProcessThreads_(threads, store, stats);
    start += threads.length;
  }

  ownerRezFlush_(store);
  props.setProperty(OWNERREZ_PROPS.LAST_SYNC, String(startedAt));
}

function ownerRezProcessThreads_(threads, store, stats) {
  const sender = OWNERREZ_CONFIG.SENDER.toLowerCase();

  GmailApp.getMessagesForThreads(threads).forEach(messages => {
    messages.forEach(message => {
      // Un hilo puede contener respuestas propias u otros remitentes
      if (message.getFrom().toLowerCase().indexOf(sender) === -1) return;
      stats.read++;

      try {
        const reservation = ownerRezParseEmail_(
          message.getSubject(),
          message.getBody(),
          message.getPlainBody()
        );
        if (!reservation) {
          Logger.log(`OwnerRez: omitido (no es una reserva reconocible): "${message.getSubject()}" ${message.getId()}`);
          stats.skipped++;
          return;
        }

        reservation.fechaMail = message.getDate();
        stats[ownerRezUpsert_(store, reservation)]++;
      } catch (e) {
        stats.errors++;
        Logger.log(`OwnerRez: error en el correo ${message.getId()}: ${e && e.stack || e}`);
      }
    });
  });
}

/* ======================================================================
 * Extracción de datos del correo
 * ==================================================================== */

/**
 * Convierte un correo en un objeto de reserva, o null si no se reconoce.
 * Busca por etiquetas ("Guest Name:", "Check-in", ...) en lugar de depender
 * de la estructura HTML exacta, así que tolera cambios de diseño del template.
 */
function ownerRezParseEmail_(subject, htmlBody, plainBody) {
  let reservation = ownerRezParseLines_(subject, ownerRezHtmlToLines_(htmlBody || ''));
  if (!reservation && plainBody) {
    reservation = ownerRezParseLines_(subject, ownerRezTextToLines_(plainBody));
  }
  return reservation;
}

function ownerRezParseLines_(subject, lines) {
  const confirmationCode = ownerRezValueAfter_(lines, /^confirmation code\s*:?\s*(.*)$/i);
  if (!confirmationCode) return null;

  const status = ownerRezDetectStatus_(subject, lines);
  if (!status) return null;

  return {
    unidad: ownerRezFindUnit_(lines),
    guest: ownerRezValueAfter_(lines, /^guest name\s*:?\s*(.*)$/i),
    checkIn: ownerRezParseDate_(ownerRezValueAfter_(lines, /^check[\s-]?in\s*:?\s*(.*)$/i)),
    checkOut: ownerRezParseDate_(ownerRezValueAfter_(lines, /^check[\s-]?out\s*:?\s*(.*)$/i)),
    numberOfNights: ownerRezParseNumber_(ownerRezValueAfter_(lines, /^number of nights\s*:?\s*(.*)$/i)),
    numberOfGuests: ownerRezParseNumber_(ownerRezValueAfter_(lines, /^number of guests\s*:?\s*(.*)$/i)),
    confirmationCode: confirmationCode,
    source: ownerRezValueAfter_(lines, /^source\s*:?\s*(.*)$/i),
    status: status
  };
}

/**
 * El status sale del asunto; si el asunto no lo dice, del encabezado del correo.
 */
function ownerRezDetectStatus_(subject, lines) {
  const S = OWNERREZ_CONFIG.STATUS;
  const classify = text => {
    if (/cancel/i.test(text)) return S.CANCELLED;
    if (/updat|modif|chang|alter/i.test(text)) return S.MODIFIED;
    if (/new reservation|new booking|congrats/i.test(text)) return S.NEW;
    return null;
  };

  const fromSubject = classify(subject || '');
  if (fromSubject) return fromSubject;

  // Encabezado: todo lo que aparece antes de la tabla de fechas
  const checkInIndex = lines.findIndex(line => /^check[\s-]?in\b/i.test(line));
  const header = lines.slice(0, checkInIndex > 0 ? checkInIndex : 6).join('\n');
  return classify(header);
}

/**
 * El nombre de la propiedad es la línea justo antes de "Check-in".
 * Si no se encuentra, se usa el preheader: "New reservation at X: ... to ...".
 */
function ownerRezFindUnit_(lines) {
  const checkInIndex = lines.findIndex(line => /^check[\s-]?in\s*:?$/i.test(line));
  if (checkInIndex > 0) {
    const candidate = lines[checkInIndex - 1];
    if (!/reservation|good news|cancel|updat/i.test(candidate)) return candidate;
  }

  for (const line of lines) {
    const match = line.match(/\bat (.+?):\s.+\bto\b/i);
    if (match) return match[1].trim();
  }
  return '';
}

/**
 * Devuelve el valor de una etiqueta: el texto en la misma línea ("Source: Airbnb")
 * o, si la etiqueta está sola, la línea siguiente (caso de las tablas HTML).
 */
function ownerRezValueAfter_(lines, labelRegex) {
  for (let i = 0; i < lines.length; i++) {
    const match = lines[i].match(labelRegex);
    if (!match) continue;
    if (match[1] && match[1].trim()) return match[1].trim();
    const next = (lines[i + 1] || '').trim();
    // Campo vacío: la línea siguiente ya es otra etiqueta, no un valor
    return OWNERREZ_LABEL_REGEX.test(next) ? '' : next;
  }
  return '';
}

const OWNERREZ_LABEL_REGEX =
  /^(check[\s-]?in|check[\s-]?out|guest name|number of guests|number of nights|confirmation code|source|guest details|reservation details)\b/i;

function ownerRezHtmlToLines_(html) {
  const text = html
    .replace(/<(script|style|head)[\s\S]*?<\/\1>/gi, ' ')
    .replace(/<br\s*\/?>/gi, '\n')
    .replace(/<\/(p|div|td|th|tr|li|table|h[1-6])\s*>/gi, '\n')
    .replace(/<[^>]+>/g, ' ');
  return ownerRezTextToLines_(ownerRezDecodeEntities_(text));
}

function ownerRezTextToLines_(text) {
  return text
    .split(/\r?\n/)
    .map(line => line.replace(/\s+/g, ' ').trim())
    .filter(line => line.length > 0);
}

function ownerRezDecodeEntities_(text) {
  const named = {
    amp: '&', lt: '<', gt: '>', quot: '"', apos: "'", nbsp: ' ',
    aacute: 'á', eacute: 'é', iacute: 'í', oacute: 'ó', uacute: 'ú', ntilde: 'ñ', uuml: 'ü',
    Aacute: 'Á', Eacute: 'É', Iacute: 'Í', Oacute: 'Ó', Uacute: 'Ú', Ntilde: 'Ñ', Uuml: 'Ü'
  };
  return text.replace(/&(#x[0-9a-f]+|#\d+|[a-z]+);/gi, (entity, name) => {
    if (name[0] === '#') {
      const isHex = name[1] === 'x' || name[1] === 'X';
      const code = isHex ? parseInt(name.slice(2), 16) : parseInt(name.slice(1), 10);
      return isNaN(code) ? entity : String.fromCodePoint(code);
    }
    return Object.prototype.hasOwnProperty.call(named, name) ? named[name] : entity;
  });
}

/**
 * "Sep 23, 2026" / "Sep 23rd, 2026" / "Wed, Sep 23, 2026" / "23 Sep 2026" /
 * "2026-09-23" → Date. Si no se reconoce, devuelve el texto tal cual para no
 * perder el dato.
 */
function ownerRezParseDate_(value) {
  if (!value) return '';
  const clean = value
    .replace(/(\d+)(st|nd|rd|th)\b/i, '$1')
    .replace(/^[A-Za-z]+,\s*/, '')
    .trim();

  let match = clean.match(/^([A-Za-z]{3,})\.?\s+(\d{1,2}),?\s+(\d{4})/);
  if (match && OWNERREZ_MONTHS[match[1].slice(0, 3).toLowerCase()] !== undefined) {
    return new Date(Number(match[3]), OWNERREZ_MONTHS[match[1].slice(0, 3).toLowerCase()], Number(match[2]));
  }

  match = clean.match(/^(\d{1,2})\s+(?:de\s+)?([A-Za-z]{3,})\.?,?\s+(?:de\s+)?(\d{4})/i);
  if (match && OWNERREZ_MONTHS[match[2].slice(0, 3).toLowerCase()] !== undefined) {
    return new Date(Number(match[3]), OWNERREZ_MONTHS[match[2].slice(0, 3).toLowerCase()], Number(match[1]));
  }

  match = clean.match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (match) return new Date(Number(match[1]), Number(match[2]) - 1, Number(match[3]));

  return value;
}

function ownerRezParseNumber_(value) {
  const number = parseInt(value, 10);
  return isNaN(number) ? value : number;
}

/* ======================================================================
 * Hoja de cálculo
 * ==================================================================== */

function ownerRezGetSheet_() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const headers = OWNERREZ_CONFIG.HEADERS;
  let sheet = spreadsheet.getSheetByName(OWNERREZ_CONFIG.SHEET_NAME);

  if (!sheet) {
    sheet = spreadsheet.insertSheet(OWNERREZ_CONFIG.SHEET_NAME);
  }

  // Escribe los encabezados que falten (sin tocar los que ya existen)
  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  const current = headerRange.getValues()[0];
  if (current.some((value, i) => value === '' && headers[i])) {
    headerRange.setValues([current.map((value, i) => value === '' ? headers[i] : value)]);
    headerRange.setFontWeight('bold');
    sheet.setFrozenRows(1);
  }

  const maxRows = sheet.getMaxRows();
  if (maxRows > 1) {
    const format = OWNERREZ_CONFIG.DATE_FORMAT;
    sheet.getRange(2, OWNERREZ_COL.CHECK_IN + 1, maxRows - 1, 2).setNumberFormat(format);
    sheet.getRange(2, OWNERREZ_COL.FECHA_MAIL + 1, maxRows - 1, 1).setNumberFormat(format);
  }

  return sheet;
}

/**
 * Carga la hoja una sola vez en memoria, indexada por Confirmation Code.
 */
function ownerRezLoadStore_(sheet) {
  const width = OWNERREZ_CONFIG.HEADERS.length;
  const lastRow = sheet.getLastRow();
  const rows = lastRow > 1 ? sheet.getRange(2, 1, lastRow - 1, width).getValues() : [];

  const index = new Map();
  rows.forEach((values, i) => {
    const code = String(values[OWNERREZ_COL.CODE]).trim();
    if (code) index.set(code, { row: i + 2, values: values });
  });

  return { sheet: sheet, index: index, pending: [], dirty: new Set() };
}

/**
 * Inserta o actualiza una reserva en memoria. Solo un correo MÁS NUEVO que el
 * último aplicado puede modificar la fila, así el orden en que se leen los
 * correos no importa (el backfill va de más nuevo a más viejo).
 * Devuelve 'added' | 'updated' | 'unchanged'.
 */
function ownerRezUpsert_(store, reservation) {
  const values = ownerRezToRow_(reservation);
  const entry = store.index.get(reservation.confirmationCode);

  if (!entry) {
    const created = { row: null, values: values };
    store.index.set(reservation.confirmationCode, created);
    store.pending.push(created);
    return 'added';
  }

  const previousMail = ownerRezToTime_(entry.values[OWNERREZ_COL.FECHA_MAIL]);
  if (reservation.fechaMail.getTime() <= previousMail) return 'unchanged';

  // Si el correo nuevo trae un campo vacío (p. ej. una cancelación con menos
  // datos) se conserva el valor anterior.
  entry.values = values.map((value, i) => value === '' || value === null ? entry.values[i] : value);
  if (entry.row) store.dirty.add(entry);
  return 'updated';
}

function ownerRezToRow_(r) {
  const row = [];
  row[OWNERREZ_COL.UNIDAD] = r.unidad;
  row[OWNERREZ_COL.GUEST] = r.guest;
  row[OWNERREZ_COL.CHECK_IN] = r.checkIn;
  row[OWNERREZ_COL.CHECK_OUT] = r.checkOut;
  row[OWNERREZ_COL.NIGHTS] = r.numberOfNights;
  row[OWNERREZ_COL.GUESTS] = r.numberOfGuests;
  row[OWNERREZ_COL.CODE] = r.confirmationCode;
  row[OWNERREZ_COL.SOURCE] = r.source;
  row[OWNERREZ_COL.STATUS] = r.status;
  row[OWNERREZ_COL.FECHA_MAIL] = r.fechaMail;
  return row;
}

function ownerRezToTime_(value) {
  if (value && typeof value.getTime === 'function') return value.getTime();
  if (!value) return 0;
  const match = String(value).match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})/); // dd/MM/yyyy
  return match ? new Date(Number(match[3]), Number(match[2]) - 1, Number(match[1])).getTime() : 0;
}

/**
 * Escribe en la hoja los cambios acumulados en memoria.
 */
function ownerRezFlush_(store) {
  const sheet = store.sheet;
  const width = OWNERREZ_CONFIG.HEADERS.length;

  store.dirty.forEach(entry => {
    const range = sheet.getRange(entry.row, 1, 1, width);
    range.setValues([entry.values]);
    ownerRezApplyColors_(range, [entry.values]);
  });
  store.dirty.clear();

  if (store.pending.length > 0) {
    const firstRow = sheet.getLastRow() + 1;
    const rows = store.pending.map(entry => entry.values);
    const range = sheet.getRange(firstRow, 1, rows.length, width);
    range.setValues(rows);
    ownerRezApplyColors_(range, rows);
    store.pending.forEach((entry, i) => { entry.row = firstRow + i; });
    store.pending = [];
  }
}

function ownerRezApplyColors_(range, rows) {
  const width = OWNERREZ_CONFIG.HEADERS.length;
  const backgrounds = [];
  const fonts = [];

  rows.forEach(values => {
    const color = OWNERREZ_CONFIG.COLORS[values[OWNERREZ_COL.STATUS]] || { background: 'white', font: '#000000' };
    backgrounds.push(new Array(width).fill(color.background));
    fonts.push(new Array(width).fill(color.font));
  });

  range.setBackgrounds(backgrounds);
  range.setFontColors(fonts);
}

function ownerRezToast_(message) {
  try {
    SpreadsheetApp.getActiveSpreadsheet().toast(message, 'OwnerRez', 8);
  } catch (e) {
    // Sin interfaz (ejecución por trigger): basta con el log
  }
}

/* ======================================================================
 * Utilidades: trigger, menú, depuración
 * ==================================================================== */

/**
 * Ejecutar UNA vez: programa ownerRezSync cada N horas.
 */
function ownerRezInstallTrigger() {
  ScriptApp.getProjectTriggers()
    .filter(trigger => trigger.getHandlerFunction() === 'ownerRezSync')
    .forEach(trigger => ScriptApp.deleteTrigger(trigger));

  ScriptApp.newTrigger('ownerRezSync')
    .timeBased()
    .everyHours(OWNERREZ_CONFIG.TRIGGER_EVERY_HOURS)
    .create();

  Logger.log(`OwnerRez: trigger creado, se ejecutará cada ${OWNERREZ_CONFIG.TRIGGER_EVERY_HOURS} hora(s).`);
}

/**
 * Vuelve a recorrer todo el histórico en la próxima ejecución. No borra filas:
 * las reservas existentes solo se actualizan si hay un correo más nuevo.
 */
function ownerRezResetBackfill() {
  const props = PropertiesService.getScriptProperties();
  Object.keys(OWNERREZ_PROPS).forEach(key => props.deleteProperty(OWNERREZ_PROPS[key]));
  Logger.log('OwnerRez: estado reiniciado, la próxima ejecución recorrerá todo el histórico.');
}

/**
 * Muestra en el log lo que se extraería de los últimos correos, sin escribir
 * nada en la hoja. Útil para validar el parser.
 */
function ownerRezDebugLatest() {
  const threads = GmailApp.search(`from:${OWNERREZ_CONFIG.SENDER}`, 0, 5);
  GmailApp.getMessagesForThreads(threads).forEach(messages => {
    messages.forEach(message => {
      const reservation = ownerRezParseEmail_(message.getSubject(), message.getBody(), message.getPlainBody());
      Logger.log(`${message.getDate()} | ${message.getSubject()}\n${JSON.stringify(reservation, null, 2)}`);
    });
  });
}

/**
 * Agrega el menú "OwnerRez". Llamar desde tu onOpen existente:
 *   function onOpen() { ...; ownerRezAddMenu(); }
 */
function ownerRezAddMenu() {
  SpreadsheetApp.getUi()
    .createMenu('OwnerRez')
    .addItem('Sincronizar reservas', 'ownerRezSync')
    .addItem('Probar extracción (solo log)', 'ownerRezDebugLatest')
    .addSeparator()
    .addItem('Instalar ejecución automática', 'ownerRezInstallTrigger')
    .addItem('Reprocesar todo el histórico', 'ownerRezResetBackfill')
    .addToUi();
}
