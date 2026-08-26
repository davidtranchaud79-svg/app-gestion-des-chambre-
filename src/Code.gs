const CONFIG = {
  spreadsheetId: '1mB_4WNuqw1sysQIa4LUPqktFOngqOTS61pj6v3PPtnE',
  sheets: {
    params: 'Parametres',
    rooms: 'Chambres',
    registry: 'Registre',
    requests: 'Reservations_Public',
  },
  roomStartRow: 6,
  registryStartRow: 6,
  requestStartRow: 2,
  requestCols: 20,
};

function doGet() {
  const template = HtmlService.createTemplateFromFile('Index');
  template.bootstrap = JSON.stringify({
    title: 'Gestion chambre',
    today: toIso_(new Date()),
  });

  return template.evaluate()
    .setTitle('Gestion chambre')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function checkAvailability(payload) {
  try {
    const stay = normalizeStay_(payload || {});
    const ss = SpreadsheetApp.openById(CONFIG.spreadsheetId);
    const rooms = getRooms_(ss);
    const registry = getRegistry_(ss);

    const availability = rooms.map((room) => {
      const closed = norm_(room.status) !== norm_('Ouverte');
      const conflicts = registry.filter((entry) => (
        entry.room === room.room &&
        entry.arrival < stay.departure &&
        entry.departure > stay.arrival
      ));

      return {
        ...room,
        available: !closed && conflicts.length === 0,
        reason: closed ? 'Chambre fermée' : conflicts.length ? 'Déjà occupée' : 'Libre',
      };
    });

    const publicRooms = availability
      .filter((room) => room.available)
      .map((room) => ({
        room: room.room,
        type: room.type,
        capacity: room.capacity,
        nightRate: room.nightRate,
        reason: room.reason,
      }));

    return {
      ok: true,
      arrival: formatFr_(stay.arrival),
      departure: formatFr_(stay.departure),
      nights: stay.nights,
      rooms: publicRooms,
      availableCount: publicRooms.length,
      totalCount: rooms.length,
    };
  } catch (err) {
    return { ok: false, error: err.message || String(err) };
  }
}

function createReservation(payload) {
  const lock = LockService.getScriptLock();

  if (!lock.tryLock(10000)) {
    return {
      ok: false,
      error: 'Une autre demande est en cours. Réessaie dans quelques secondes.',
    };
  }

  try {
    payload = payload || {};

    const stay = normalizeStay_(payload);
    const name = clean_(payload.occupant);
    const phone = clean_(payload.phone);
    const email = clean_(payload.email);
    const comment = clean_(payload.comment);
    const requestedRoom = clean_(payload.requestedRoom);

    if (!name) throw new Error('Le nom est obligatoire.');
    if (!phone && !email) throw new Error('Ajoute au moins un téléphone ou un email.');

    const ss = SpreadsheetApp.openById(CONFIG.spreadsheetId);
    const availability = checkAvailability(payload);
    if (!availability.ok) return availability;

    let selected = null;

    if (requestedRoom) {
      selected = availability.rooms.find((room) => room.room === requestedRoom);
      if (!selected) {
        return { ok: false, error: 'Cette chambre vient de devenir indisponible.' };
      }
    } else {
      selected = availability.rooms[0];
    }

    if (!selected) {
      return { ok: false, error: 'Aucune chambre disponible sur cette période.' };
    }

    const id = createReservationId_();
    const amount = stay.nights * Number(selected.nightRate || 0);

    const row = [
      id,
      new Date(),
      'En attente',
      requestedRoom || '',
      selected.room,
      name,
      phone,
      email,
      stay.arrival,
      stay.departure,
      stay.nights,
      selected.nightRate || '',
      amount,
      '',
      'Micro-app',
      comment,
      'À valider réception',
      'Non',
      '',
      selected.reason || '',
    ];

    const sheet = ss.getSheetByName(CONFIG.sheets.requests);
    if (!sheet) throw new Error('Onglet Reservations_Public introuvable.');

    const targetRow = firstEmptyRow_(sheet, CONFIG.requestStartRow, 1);
    sheet.getRange(targetRow, 1, 1, CONFIG.requestCols).setValues([row]);

    return {
      ok: true,
      id,
      room: selected.room,
      nights: stay.nights,
      amount,
      message: 'Demande envoyée. La réception doit valider.',
    };
  } catch (err) {
    return { ok: false, error: err.message || String(err) };
  } finally {
    lock.releaseLock();
  }
}

function getRooms_(ss) {
  const sheet = ss.getSheetByName(CONFIG.sheets.rooms);
  if (!sheet) throw new Error('Onglet Chambres introuvable.');

  const last = sheet.getLastRow();
  if (last < CONFIG.roomStartRow) return [];

  return sheet.getRange(CONFIG.roomStartRow, 1, last - CONFIG.roomStartRow + 1, 8)
    .getValues()
    .map((row) => ({
      room: clean_(row[0]),
      floor: clean_(row[1]),
      type: clean_(row[2]),
      capacity: Number(row[3] || 1),
      mode: clean_(row[4]),
      nightRate: Number(row[5] || 0),
      monthRate: Number(row[6] || 0),
      status: clean_(row[7]) || 'Ouverte',
    }))
    .filter((room) => room.room);
}

function getRegistry_(ss) {
  const sheet = ss.getSheetByName(CONFIG.sheets.registry);
  if (!sheet) throw new Error('Onglet Registre introuvable.');

  const last = sheet.getLastRow();
  if (last < CONFIG.registryStartRow) return [];

  return sheet.getRange(CONFIG.registryStartRow, 1, last - CONFIG.registryStartRow + 1, 15)
    .getValues()
    .map((row) => ({
      room: clean_(row[2]),
      arrival: toDate_(row[6]),
      departure: toDate_(row[7]),
      status: clean_(row[14]),
    }))
    .filter((entry) => entry.room && entry.arrival && entry.departure && !isCancelled_(entry.status));
}

function normalizeStay_(payload) {
  const arrival = parseDate_(payload.arrival, 'arrivée');
  const departure = parseDate_(payload.departure, 'départ');

  if (departure <= arrival) {
    throw new Error('La date de départ doit être après la date d’arrivée.');
  }

  const nights = Math.round((departure - arrival) / 86400000);
  if (nights > 90) {
    throw new Error('Séjour trop long. Passe par la réception.');
  }

  return { arrival, departure, nights };
}

function parseDate_(value, label) {
  if (!value) throw new Error('Date ' + label + ' obligatoire.');

  if (value instanceof Date && !isNaN(value.getTime())) {
    return new Date(value.getFullYear(), value.getMonth(), value.getDate());
  }

  const text = clean_(value);
  const iso = text.match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (iso) return new Date(Number(iso[1]), Number(iso[2]) - 1, Number(iso[3]));

  const fr = text.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
  if (fr) return new Date(Number(fr[3]), Number(fr[2]) - 1, Number(fr[1]));

  throw new Error('Date ' + label + ' invalide.');
}

function toDate_(value) {
  if (!value) return null;

  if (value instanceof Date && !isNaN(value.getTime())) {
    return new Date(value.getFullYear(), value.getMonth(), value.getDate());
  }

  if (typeof value === 'number') {
    const base = new Date(Date.UTC(1899, 11, 30));
    const date = new Date(base.getTime() + value * 86400000);
    return new Date(date.getUTCFullYear(), date.getUTCMonth(), date.getUTCDate());
  }

  try {
    return parseDate_(value, 'registre');
  } catch (e) {
    return null;
  }
}

function firstEmptyRow_(sheet, startRow, col) {
  const values = sheet.getRange(startRow, col, sheet.getMaxRows() - startRow + 1, 1).getValues();

  for (let i = 0; i < values.length; i++) {
    if (!clean_(values[i][0])) return startRow + i;
  }

  const maxRows = sheet.getMaxRows();
  sheet.insertRowsAfter(maxRows, 50);
  return maxRows + 1;
}

function createReservationId_() {
  const stamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd-HHmmss');
  const rand = Math.floor(Math.random() * 9000 + 1000);
  return 'RES-' + stamp + '-' + rand;
}

function isCancelled_(status) {
  return ['annule', 'annulee', 'no-show', 'refuse', 'refusee'].includes(norm_(status));
}

function clean_(value) {
  return value === null || value === undefined ? '' : String(value).trim();
}

function norm_(value) {
  return clean_(value)
    .toLowerCase()
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '');
}

function formatFr_(date) {
  return Utilities.formatDate(date, Session.getScriptTimeZone(), 'dd/MM/yyyy');
}

function toIso_(date) {
  return Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy-MM-dd');
}

function healthCheck_() {
  const ss = SpreadsheetApp.openById(CONFIG.spreadsheetId);
  return {
    spreadsheet: ss.getName(),
    rooms: getRooms_(ss).length,
    registryEntries: getRegistry_(ss).length,
    requestsSheet: Boolean(ss.getSheetByName(CONFIG.sheets.requests)),
  };
}
