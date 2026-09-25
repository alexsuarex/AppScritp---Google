/**
 * Script para extraer información de reservaciones desde Gmail
 * Busca correos con asunto: "Congrats! You have a new reservation by La Paz Bay"
 *
 * Optimización: la hoja RESERVAS se lee UNA sola vez por ejecución y se
 * mantiene un índice en memoria (Confirmation Code -> fila). Antes se leía la
 * hoja completa varias veces por cada correo. La extracción de datos de los
 * correos (extractData, convertDateFormat, extractGuestyTableData,
 * convertGuestyDate) no cambió.
 *
 * Guesty: una reserva existente solo se reescribe si algún dato cambió de
 * verdad (se comparan valores normalizados, ver normalizeForCompare).
 */

function extractReservations() {
  // Calcular la fecha de hace 5 días
  const fiveDaysAgo = new Date();
  fiveDaysAgo.setDate(fiveDaysAgo.getDate() - 5);
  const dateString = Utilities.formatDate(fiveDaysAgo, Session.getScriptTimeZone(), 'yyyy/MM/dd');

  // Buscar correos de nuevas reservaciones
  const newReservationsQuery = `subject:"Congrats! You have a new reservation by La Paz Bay" after:${dateString}`;
  const newThreads = GmailApp.search(newReservationsQuery);

  // Buscar correos de modificaciones
  const modificationsQuery = `subject:"Reservation update notification by La Paz Bay" after:${dateString}`;
  const modThreads = GmailApp.search(modificationsQuery);

  // Buscar correos de cancelaciones
  const cancellationsQuery = `subject:"Reservation Cancelled notification by La Paz Bay" after:${dateString}`;
  const cancelThreads = GmailApp.search(cancellationsQuery);

  // Obtener o crear la hoja de cálculo y leerla UNA sola vez
  const sheet = getOrCreateSheet();
  const index = loadReservationIndex(sheet);

  // Procesar cada tipo de correo
  const newReservations = processThreads(index, newThreads, 'Confirmada');
  const modifications = processThreads(index, modThreads, 'Modificada');
  const cancellations = processThreads(index, cancelThreads, 'Cancelada');

  const totalProcessed = newReservations + modifications + cancellations;
  Logger.log(`Proceso LPB completado. ${newReservations} confirmadas, ${modifications} modificadas, ${cancellations} canceladas.`);

  // Después de procesar los correos de La Paz Bay, procesar Guesty
  Logger.log('Iniciando extracción de reservas de Guesty...');
  const guestyResult = extractGuestyReservations(index);

  // Mostrar notificación de resultado combinado
  SpreadsheetApp.getActiveSpreadsheet().toast(
    `La Paz Bay: ${totalProcessed} procesadas (${newReservations} confirmadas, ${modifications} modificadas, ${cancellations} canceladas)\n` +
    `Guesty: ${guestyResult.processed} procesadas (${guestyResult.added} nuevas, ${guestyResult.updated} actualizadas)`,
    'Proceso Completado',
    8
  );
}

/**
 * Procesar todos los mensajes de una lista de hilos.
 * Los mensajes se obtienen en una sola llamada para todos los hilos.
 */
function processThreads(index, threads, status) {
  if (threads.length === 0) return 0;

  let count = 0;
  GmailApp.getMessagesForThreads(threads).forEach(messages => {
    messages.forEach(message => {
      if (processMessage(index, message, status)) {
        count++;
      }
    });
  });
  return count;
}

/**
 * Procesar un mensaje individual y extraer los datos
 * (acepta el índice o, por compatibilidad, la hoja)
 */
function processMessage(sheetOrIndex, message, status) {
  const index = asReservationIndex(sheetOrIndex);
  const messageId = message.getId();

  // Verificar si este correo ya fue procesado
  if (isAlreadyProcessed(index, messageId)) {
    Logger.log('Omitiendo correo ya procesado: ' + messageId);
    return false;
  }

  // Extraer datos del correo
  const body = message.getPlainBody();
  const htmlBody = message.getBody();
  const date = message.getDate();

  const reservationData = extractData(body, htmlBody, date, messageId, status);

  if (reservationData) {
    addToSheet(index, reservationData);
    Logger.log(`Reservación ${status}: ` + reservationData.confirmationCode);
    return true;
  }

  return false;
}

function extractData(plainBody, htmlBody, emailDate, messageId, status) {
  const data = {
    unidad: '',
    guest: '',
    checkIn: '',
    checkOut: '',
    numberOfNights: '',
    numberOfGuests: '',
    confirmationCode: '',
    source: '',
    status: status, // Usar el status pasado como parámetro
    fechaMail: Utilities.formatDate(emailDate, Session.getScriptTimeZone(), 'dd/MM/yyyy'),
    messageId: messageId
  };

  // Usar el cuerpo HTML para extraer datos
  const body = htmlBody || plainBody;

  // Log para debugging - mostrar parte del HTML
  Logger.log('=== Extrayendo datos del correo ===');
  Logger.log('Status: ' + status);
  Logger.log('HTML Body (primeros 500 caracteres): ' + body.substring(0, 500));

  // Extraer Propiedad/Unidad - buscar solo el nombre entre el header y CHECK-IN
  // Método 1: Buscar por clase property-name
  let unidadMatch = body.match(/<div[^>]*class="property-name"[^>]*>([^<]+)<\/div>/i);
  if (unidadMatch) {
    data.unidad = unidadMatch[1].trim();
    Logger.log('Unidad encontrada (método 1 - property-name): ' + data.unidad);
  }

  // Método 2: Si no encontró, buscar entre el header y CHECK-IN
  if (!data.unidad) {
    unidadMatch = body.match(/new reservation<\/p>[\s\S]*?<div[^>]*>\s*<div[^>]*>\s*([A-Za-z0-9\s()]+?)\s*<\/div>\s*<div[^>]*>\s*CHECK-IN/i);
    if (unidadMatch) {
      data.unidad = unidadMatch[1].trim();
      Logger.log('Unidad encontrada (método 2 - entre header y CHECK-IN): ' + data.unidad);
    }
  }

  // Método 3: Si no encontró o tiene texto extra, intentar con texto plano
  if (!data.unidad || data.unidad.includes('You have') || data.unidad.includes('reservation') || data.unidad.length > 100) {
    const plainMatch = plainBody.match(/(?:reservation|updated|cancelled)\s*\n\s*([A-Za-z0-9\s()]+?)\s*\n/i);
    if (plainMatch) {
      data.unidad = plainMatch[1].trim();
      Logger.log('Unidad encontrada (método 3 - texto plano): ' + data.unidad);
    }
  }

  // Método 4: Buscar después del contenido y antes de dates-container
  if (!data.unidad) {
    const altMatch = body.match(/<div[^>]*class="content"[^>]*>[\s\S]*?<div[^>]*class="property-name"[^>]*>\s*([^<]+?)\s*<\/div>/i);
    if (altMatch) {
      data.unidad = altMatch[1].trim();
      Logger.log('Unidad encontrada (método 4): ' + data.unidad);
    }
  }

  if (!data.unidad) {
    Logger.log('Unidad NO encontrada - Buscando en los primeros 2000 caracteres del HTML');
    // Mostrar más HTML para debugging
    const htmlSnippet = body.substring(0, 2000);
    Logger.log('HTML extendido: ' + htmlSnippet);
  }

  // Extraer Check-in (ahora con el nuevo formato que tiene ":")
  const checkinMatch = body.match(/CHECK-IN<\/div>\s*<div[^>]*>\s*([^<]+?)\s*<\/div>/i);
  if (checkinMatch) {
    data.checkIn = convertDateFormat(checkinMatch[1].trim());
    Logger.log('Check-in encontrado: ' + data.checkIn);
  }

  // Extraer Check-out
  const checkoutMatch = body.match(/CHECK-OUT<\/div>\s*<div[^>]*>\s*([^<]+?)\s*<\/div>/i);
  if (checkoutMatch) {
    data.checkOut = convertDateFormat(checkoutMatch[1].trim());
    Logger.log('Check-out encontrado: ' + data.checkOut);
  }

  // Extraer Nombre del Huésped (ahora busca "Guest Name:" con dos puntos)
  let guestMatch = body.match(/Guest Name:\s*<\/span>\s*<span[^>]*>\s*([^<]+?)\s*<\/span>/i);
  if (!guestMatch) {
    // Intento alternativo sin los dos puntos
    guestMatch = body.match(/Guest Name<\/span>\s*<span[^>]*>\s*([^<]+?)\s*<\/span>/i);
  }
  if (!guestMatch) {
    // Tercer intento: buscar en detail-value después de Guest Name
    guestMatch = body.match(/Guest Name[:\s]*<\/[^>]+>[\s\S]{0,100}?<[^>]*class="detail-value"[^>]*>([^<]+)<\//i);
  }
  if (guestMatch) {
    data.guest = guestMatch[1].trim();
    Logger.log('Guest encontrado: ' + data.guest);
  } else {
    Logger.log('Guest NO encontrado');
  }

  // Extraer Número de Huéspedes
  let guestsMatch = body.match(/Number of Guests:\s*<\/span>\s*<span[^>]*>\s*([^<]+?)\s*<\/span>/i);
  if (!guestsMatch) {
    guestsMatch = body.match(/Number of Guests<\/span>\s*<span[^>]*>\s*([^<]+?)\s*<\/span>/i);
  }
  if (!guestsMatch) {
    guestsMatch = body.match(/Number of Guests[:\s]*<\/[^>]+>[\s\S]{0,100}?<[^>]*class="detail-value"[^>]*>([^<]+)<\//i);
  }
  if (guestsMatch) {
    data.numberOfGuests = guestsMatch[1].trim();
    Logger.log('Number of Guests encontrado: ' + data.numberOfGuests);
  } else {
    Logger.log('Number of Guests NO encontrado');
  }

  // Extraer Número de Noches
  let nightsMatch = body.match(/Number of Nights:\s*<\/span>\s*<span[^>]*>\s*([^<]+?)\s*<\/span>/i);
  if (!nightsMatch) {
    nightsMatch = body.match(/Number of Nights<\/span>\s*<span[^>]*>\s*([^<]+?)\s*<\/span>/i);
  }
  if (!nightsMatch) {
    nightsMatch = body.match(/Number of Nights[:\s]*<\/[^>]+>[\s\S]{0,100}?<[^>]*class="detail-value"[^>]*>([^<]+)<\//i);
  }
  if (nightsMatch) {
    data.numberOfNights = nightsMatch[1].trim();
    Logger.log('Number of Nights encontrado: ' + data.numberOfNights);
  } else {
    Logger.log('Number of Nights NO encontrado');
  }

  // Extraer Código de Confirmación
  let confirmationMatch = body.match(/Confirmation Code:\s*<\/span>\s*<span[^>]*>\s*([^<]+?)\s*<\/span>/i);
  if (!confirmationMatch) {
    confirmationMatch = body.match(/Confirmation Code<\/span>\s*<span[^>]*>\s*([^<]+?)\s*<\/span>/i);
  }
  if (!confirmationMatch) {
    confirmationMatch = body.match(/Confirmation Code[:\s]*<\/[^>]+>[\s\S]{0,100}?<[^>]*class="detail-value"[^>]*>([^<]+)<\//i);
  }
  if (confirmationMatch) {
    data.confirmationCode = confirmationMatch[1].trim();
    Logger.log('Confirmation Code encontrado: ' + data.confirmationCode);
  } else {
    Logger.log('Confirmation Code NO encontrado');
  }

  // Extraer Fuente
  let sourceMatch = body.match(/Source:\s*<\/span>\s*<span[^>]*>\s*([^<]+?)\s*<\/span>/i);
  if (!sourceMatch) {
    sourceMatch = body.match(/Source<\/span>\s*<span[^>]*>\s*([^<]+?)\s*<\/span>/i);
  }
  if (!sourceMatch) {
    sourceMatch = body.match(/Source[:\s]*<\/[^>]+>[\s\S]{0,100}?<[^>]*class="detail-value"[^>]*>([^<]+)<\//i);
  }
  if (sourceMatch) {
    data.source = sourceMatch[1].trim();
    Logger.log('Source encontrado: ' + data.source);
  } else {
    Logger.log('Source NO encontrado');
  }

  // Log para debugging
  Logger.log('Datos extraídos - Unidad: "' + data.unidad + '", Confirmación: ' + data.confirmationCode);
  Logger.log('=================================');

  return data;
}

/**
 * Convierte fechas del formato "Jan 21st, 2026" a "21/01/2026"
 */
function convertDateFormat(dateString) {
  // Remover sufijos como st, nd, rd, th
  dateString = dateString.replace(/(\d+)(st|nd|rd|th)/, '$1');

  try {
    // Parsear la fecha
    const date = new Date(dateString);

    // Verificar si la fecha es válida
    if (isNaN(date.getTime())) {
      Logger.log('Fecha inválida: ' + dateString);
      return dateString;
    }

    // Formatear a DD/MM/YYYY
    const formatted = Utilities.formatDate(date, Session.getScriptTimeZone(), 'dd/MM/yyyy');
    Logger.log('Fecha convertida: ' + dateString + ' -> ' + formatted);
    return formatted;
  } catch (e) {
    Logger.log('Error al convertir fecha: ' + dateString + ' - ' + e);
    return dateString;
  }
}

function getOrCreateSheet() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = spreadsheet.getSheetByName('RESERVAS');

  if (!sheet) {
    sheet = spreadsheet.insertSheet('RESERVAS');

    // Agregar encabezados
    const headers = [
      'Unidad',
      'Guest',
      'Check-in',
      'Check-out',
      'Number of Nights',
      'Number of Guests',
      'Confirmation Code',
      'Source',
      'Status',
      'Fecha Mail'
    ];

    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
    sheet.setFrozenRows(1);

    // Auto-redimensionar columnas
    for (let i = 1; i <= headers.length; i++) {
      sheet.autoResizeColumn(i);
    }
  }

  return sheet;
}

/**
 * Índice en memoria de la hoja RESERVAS (una sola lectura):
 *  - rowByCode: Confirmation Code (columna G) -> número de fila. Guarda la
 *    PRIMERA aparición, igual que el recorrido original de arriba hacia abajo.
 *  - rowValues: columnas A–J de cada fila. Cuando el script escribe una fila se
 *    descarta su copia y, si se vuelve a necesitar, se relee de la hoja.
 */
function loadReservationIndex(sheet) {
  const data = sheet.getDataRange().getValues();
  const rowByCode = new Map();
  const rowValues = new Map();

  for (let i = 1; i < data.length; i++) {
    const rowNumber = i + 1;
    if (!rowByCode.has(data[i][6])) {
      rowByCode.set(data[i][6], rowNumber);
    }
    const values = data[i].slice(0, 10);
    while (values.length < 10) values.push('');
    rowValues.set(rowNumber, values);
  }

  return { sheet: sheet, rowByCode: rowByCode, rowValues: rowValues };
}

/**
 * Permite seguir llamando a estas funciones con la hoja (como antes) desde
 * otros archivos del proyecto; en ese caso se construye el índice al vuelo.
 */
function asReservationIndex(sheetOrIndex) {
  return sheetOrIndex && sheetOrIndex.rowByCode ? sheetOrIndex : loadReservationIndex(sheetOrIndex);
}

/**
 * Columnas A–J de una fila (desde memoria; se relee si el script la modificó).
 */
function getIndexedRowValues(index, rowNumber) {
  if (!index.rowValues.has(rowNumber)) {
    index.rowValues.set(rowNumber, index.sheet.getRange(rowNumber, 1, 1, 10).getValues()[0]);
  }
  return index.rowValues.get(rowNumber);
}

/**
 * Registrar en el índice una fila recién agregada al final de la hoja.
 */
function registerAppendedRow(index, rowNumber, confirmationCode) {
  if (!index.rowByCode.has(confirmationCode)) {
    index.rowByCode.set(confirmationCode, rowNumber);
  }
  index.rowValues.delete(rowNumber);
}

function isAlreadyProcessed(sheetOrIndex, messageId) {
  // Verificar si el messageId existe en la columna Confirmation Code
  return asReservationIndex(sheetOrIndex).rowByCode.has(messageId);
}

/**
 * Buscar si una reserva ya existe por Confirmation Code
 * Retorna el número de fila si existe, o -1 si no existe
 */
function findReservationRow(sheetOrIndex, confirmationCode) {
  // Buscar por Confirmation Code (columna 7, índice 6) en el índice en memoria
  const rowNumber = asReservationIndex(sheetOrIndex).rowByCode.get(confirmationCode);
  return rowNumber === undefined ? -1 : rowNumber; // -1 = No encontrada
}

function addToSheet(sheetOrIndex, data) {
  const index = asReservationIndex(sheetOrIndex);
  const sheet = index.sheet;

  // Buscar si la reserva ya existe
  const existingRow = findReservationRow(index, data.confirmationCode);

  const rowData = [
    data.unidad,
    data.guest,
    data.checkIn,
    data.checkOut,
    data.numberOfNights,
    data.numberOfGuests,
    data.confirmationCode,
    data.source,
    data.status,
    data.fechaMail
  ];

  if (existingRow > 0) {
    // Actualizar fila existente
    Logger.log('Actualizando reserva existente en fila ' + existingRow + ' con status: ' + data.status);
    sheet.getRange(existingRow, 1, 1, rowData.length).setValues([rowData]);

    // Aplicar formato según el status
    applyRowFormatting(sheet, existingRow, data.status);
    index.rowValues.delete(existingRow);
  } else {
    // Agregar nueva fila
    Logger.log('Agregando nueva reserva con status: ' + data.status);
    sheet.appendRow(rowData);

    // Aplicar formato a la última fila agregada
    const lastRow = sheet.getLastRow();
    applyRowFormatting(sheet, lastRow, data.status);
    registerAppendedRow(index, lastRow, data.confirmationCode);
  }
}

/**
 * Aplicar formato de color a la fila según el status
 */
function applyRowFormatting(sheet, rowNumber, status) {
  const numColumns = 12; // Total de columnas
  const range = sheet.getRange(rowNumber, 1, 1, numColumns);

  if (status === 'Modificada') {
    // Color naranja con texto blanco
    range.setBackground('orange'); // Naranja
    range.setFontColor('white'); // Texto blanco
    range.setFontWeight('normal');
  } else if (status === 'Cancelada') {
    // Color rojo con texto blanco
    range.setBackground('red'); // Rojo
    range.setFontColor('white'); // Texto blanco
    range.setFontWeight('normal');
  } else {
    // Status Confirmada - formato normal (sin color especial)
    range.setBackground('white'); // Blanco
    range.setFontColor('#000000'); // Texto negro
    range.setFontWeight('normal');
  }
}

/**
 * Crear un trigger programado para ejecutarse automáticamente
 * Ejecutar esta función una vez para configurar la ejecución automática
 */
function createTimeDrivenTrigger() {
  // Eliminar triggers existentes primero
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => {
    if (trigger.getHandlerFunction() === 'extractReservations') {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  // Crear nuevo trigger para ejecutarse diariamente a las 4 AM hora del Pacífico (México)
  ScriptApp.newTrigger('extractReservations')
    .timeBased()
    .atHour(4)
    .everyDays(1)
    .inTimezone('America/Mazatlan') // Hora del Pacífico México
    .create();

  Logger.log('Trigger creado exitosamente. El script se ejecutará diariamente a las 4:00 AM hora del Pacífico (México).');
}

/**
 * Crear un menú en Google Sheets
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Reservations')
    .addItem('Extract Reservations by LPB & Guesty', 'extractReservations')
    .addItem('Extract Reservations by AH', 'reservation_Modified')
    .addToUi();
}

/**
 * Extraer reservaciones del reporte diario de Guesty
 * Asunto: "Guesty shared A096 - Check In Next 7 Days (Tiffany) YYYY-MM-DD report with you"
 * Solo busca el correo con la fecha de HOY
 * (recibe el índice de extractReservations para no volver a leer la hoja)
 */
function extractGuestyReservations(index) {
  // Obtener la fecha de hoy en formato YYYY-MM-DD
  const today = new Date();
  const todayString = Utilities.formatDate(today, Session.getScriptTimeZone(), 'yyyy-MM-dd');

  // Buscar solo el correo de hoy
  const searchQuery = `subject:"Guesty shared A096 - Check In Next 7 Days (Tiffany) ${todayString} report with you"`;
  const threads = GmailApp.search(searchQuery);

  if (threads.length === 0) {
    Logger.log('No se encontró el reporte de Guesty de hoy: ' + todayString);
    return { processed: 0, added: 0, updated: 0 };
  }

  if (!index) {
    index = loadReservationIndex(getOrCreateSheet());
  }
  let processed = 0;
  let updated = 0;
  let added = 0;

  GmailApp.getMessagesForThreads(threads).forEach(messages => {
    messages.forEach(message => {
      const htmlBody = message.getBody();

      // Extraer las filas de la tabla HTML
      const reservations = extractGuestyTableData(htmlBody);

      reservations.forEach(reservation => {
        if (reservation.confirmationCode) {
          const result = processGuestyReservation(index, reservation);
          if (result === 'added') added++;
          if (result === 'updated') updated++;
          processed++;
        }
      });
    });
  });

  Logger.log(`Proceso Guesty completado. ${processed} reservaciones procesadas: ${added} nuevas, ${updated} actualizadas.`);

  return { processed: processed, added: added, updated: updated };
}

/**
 * Extraer datos de la tabla HTML del reporte de Guesty
 */
function extractGuestyTableData(htmlBody) {
  const reservations = [];

  // Buscar todas las filas de la tabla (tr) que contienen datos
  const rowRegex = /<tr[^>]*>([\s\S]*?)<\/tr>/gi;
  const rows = htmlBody.match(rowRegex);

  if (!rows) return reservations;

  rows.forEach(row => {
    // Extraer todas las celdas (td) de la fila
    const cellRegex = /<td[^>]*>([\s\S]*?)<\/td>/gi;
    const cells = [];
    let match;

    while ((match = cellRegex.exec(row)) !== null) {
      // Limpiar el contenido HTML y obtener solo el texto
      let cellText = match[1]
        .replace(/<[^>]+>/g, '') // Remover tags HTML
        .replace(/&nbsp;/g, ' ') // Reemplazar &nbsp;
        .trim();
      cells.push(cellText);
    }

    // Si la fila tiene al menos 11 columnas, procesarla
    if (cells.length >= 11) {
      // Verificar que no sea la fila de encabezados
      if (!cells[0].includes('CONFIRMATION DATE') && cells[8]) {
        const reservation = {
          confirmationDate: cells[0] || '',
          listingNickname: cells[1] || '',
          checkIn: cells[2] || '',
          checkOut: cells[3] || '',
          numberOfNights: cells[4] || '',
          guestName: cells[5] || '',
          numberOfGuests: cells[6] || '',
          guestPhone: cells[7] || '',
          confirmationCode: cells[8] || '',
          otherNotes: cells[9] || '',
          source: cells[10] || '',
          allGuests: cells[11] || ''
        };

        reservations.push(reservation);
      }
    }
  });

  return reservations;
}

/**
 * Procesar una reservación de Guesty
 * (acepta el índice o, por compatibilidad, la hoja)
 */
function processGuestyReservation(sheetOrIndex, guestyData) {
  const index = asReservationIndex(sheetOrIndex);
  const sheet = index.sheet;
  
  // Convertir fechas al formato DD/MM/YYYY
  const checkIn = convertGuestyDate(guestyData.checkIn);
  const checkOut = convertGuestyDate(guestyData.checkOut);
  const fechaMail = convertGuestyDate(guestyData.confirmationDate);

  const data = {
    unidad: guestyData.listingNickname,
    guest: guestyData.guestName,
    checkIn: checkIn,
    checkOut: checkOut,
    numberOfNights: guestyData.numberOfNights,
    numberOfGuests: guestyData.numberOfGuests,
    confirmationCode: guestyData.confirmationCode,
    source: guestyData.source,
    status: 'Confirmada',
    fechaMail: fechaMail
  };

  // Buscar si la reserva ya existe
  const existingRow = findReservationRow(index, data.confirmationCode);

  if (existingRow > 0) {
    // Verificar si hay cambios en los datos
    const existingData = getIndexedRowValues(index, existingRow);

    // Se comparan valores normalizados: la hoja devuelve números y fechas,
    // mientras que Guesty entrega todo como texto ('3' vs 3, '23/09/2026' vs Date)
    const differs = (sheetValue, guestyValue) =>
      normalizeForCompare(index, sheetValue) !== normalizeForCompare(index, guestyValue);

    const hasChanges = (
      differs(existingData[0], data.unidad) ||
      differs(existingData[1], data.guest) ||
      differs(existingData[2], data.checkIn) ||
      differs(existingData[3], data.checkOut) ||
      differs(existingData[4], data.numberOfNights) ||
      differs(existingData[5], data.numberOfGuests) ||
      differs(existingData[7], data.source)
    );

    if (hasChanges) {
      // Actualizar la fila manteniendo el status actual (no lo cambiamos a menos que venga un correo específico)
      data.status = existingData[8]; // Mantener el status actual

      const rowData = [
        data.unidad,
        data.guest,
        data.checkIn,
        data.checkOut,
        data.numberOfNights,
        data.numberOfGuests,
        data.confirmationCode,
        data.source,
        data.status,
        data.fechaMail
      ];

      sheet.getRange(existingRow, 1, 1, rowData.length).setValues([rowData]);
      index.rowValues.delete(existingRow);
      Logger.log('Actualizada reserva de Guesty: ' + data.confirmationCode);
      return 'updated';
    }

    return 'unchanged';
  } else {
    // Agregar nueva fila
    const rowData = [
      data.unidad,
      data.guest,
      data.checkIn,
      data.checkOut,
      data.numberOfNights,
      data.numberOfGuests,
      data.confirmationCode,
      data.source,
      data.status,
      data.fechaMail
    ];

    sheet.appendRow(rowData);
    const lastRow = sheet.getLastRow();
    applyRowFormatting(sheet, lastRow, data.status);
    registerAppendedRow(index, lastRow, data.confirmationCode);

    Logger.log('Nueva reserva de Guesty agregada: ' + data.confirmationCode);
    return 'added';
  }
}

/**
 * Normaliza un valor para compararlo sin importar su tipo:
 *  - Fecha (como la devuelve la hoja)  -> 'dd/MM/yyyy'
 *  - Texto con fecha '2/3/2025'        -> '02/03/2025'
 *  - Número 3 o texto '3'              -> '3'
 *  - Vacío / null                      -> ''
 */
function normalizeForCompare(index, value) {
  if (value === null || value === undefined) return '';
  
  if (typeof value.getTime === 'function') {
    // Las fechas de la hoja se interpretan en la zona horaria de la hoja
    if (!index.timeZone) {
      index.timeZone = index.sheet.getParent().getSpreadsheetTimeZone();
    }
    return Utilities.formatDate(value, index.timeZone, 'dd/MM/yyyy');
  }
  
  const text = String(value).trim();
  
  const dateMatch = text.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
  if (dateMatch) {
    return dateMatch[1].padStart(2, '0') + '/' + dateMatch[2].padStart(2, '0') + '/' + dateMatch[3];
  }
  
  if (text !== '' && !isNaN(Number(text))) {
    return String(Number(text)); // '3', '3.0' y 3 -> '3'
  }
  
  return text;
}

/**
 * Convertir fechas de Guesty "2026-01-02 03:00 PM" a "02/01/2026"
 */
function convertGuestyDate(dateString) {
  if (!dateString) return '';

  try {
    // Extraer solo la parte de la fecha (antes del espacio)
    const datePart = dateString.split(' ')[0];

    // Parsear la fecha en formato YYYY-MM-DD
    const parts = datePart.split('-');
    if (parts.length === 3) {
      const year = parts[0];
      const month = parts[1];
      const day = parts[2];

      return `${day}/${month}/${year}`;
    }

    return dateString;
  } catch (e) {
    Logger.log('Error al convertir fecha de Guesty: ' + dateString + ' - ' + e);
    return dateString;
  }
}
