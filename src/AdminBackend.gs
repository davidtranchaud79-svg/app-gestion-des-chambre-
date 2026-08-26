// Module administrateur : validation, dashboard et notifications.
function setupAdminModule() {
  const ss = SpreadsheetApp.openById(CONFIG.spreadsheetId);
  ensureAdminSheets_(ss);
  refreshRegistryOperationalFlags_(ss);
  repairRoomStatusFormulas_(ss);

  const props = PropertiesService.getScriptProperties();
  const effectiveEmail = clean_(Session.getEffectiveUser().getEmail());
  if (!clean_(props.getProperty('ADMIN_EMAIL')) && effectiveEmail) {
    props.setProperty('ADMIN_EMAIL', effectiveEmail);
  }

  return {
    ok: true,
    message: 'Module administrateur initialisé.',
    adminEmail: getAdminEmail_() || 'ADMIN_EMAIL à renseigner',
    pinConfigured: Boolean(clean_(props.getProperty('ADMIN_PIN'))),
    adminUrl: getAdminUrl_(),
  };
}

function testAdminNotification() {
  const email = getAdminEmail_();
  if (!email) {
    throw new Error('ADMIN_EMAIL absent. Ajoute ADMIN_EMAIL dans Propriétés du script.');
  }

  MailApp.sendEmail({
    to: email,
    subject: 'Test notification admin - Gestion chambre',
    body: 'Notification admin opérationnelle.',
    htmlBody: '<p>Notification admin opérationnelle.</p>',
    name: 'Gestion chambre',
  });

  logAdminNotification_('TEST_ADMIN', 'info', '', 'Email de test envoyé à ' + email);
  return 'Email de test envoyé à ' + email;
}

function installAdminTriggers() {
  ScriptApp.getProjectTriggers().forEach((trigger) => {
    if (trigger.getHandlerFunction() === 'adminScanAndNotifyAnomalies') {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  ScriptApp.newTrigger('adminScanAndNotifyAnomalies')
    .timeBased()
    .everyHours(1)
    .create();

  return 'Déclencheur installé : scan anomalies toutes les heures.';
}

function adminScanAndNotifyAnomalies() {
  const ss = SpreadsheetApp.openById(CONFIG.spreadsheetId);
  const anomalies = scanReservationAnomalies_(ss);
  writeAdminAnomalies_(ss, anomalies);

  if (anomalies.length) {
    notifyAdminAnomalies_(anomalies, true);
  }

  return anomalies.length + ' anomalie(s) détectée(s).';
}

function adminLogin(pin) {
  assertAdminPin_(pin);
  const token = Utilities.getUuid();
  CacheService.getScriptCache().put(adminTokenKey_(token), 'ok', 21600);

  return {
    ok: true,
    token,
    expiresInSeconds: 21600,
  };
}

function getAdminDashboard(token) {
  assertAdminToken_(token);

  const ss = SpreadsheetApp.openById(CONFIG.spreadsheetId);
  ensureAdminSheets_(ss);
  refreshRegistryOperationalFlags_(ss);

  const requests = getAdminRequests_(ss);
  const rooms = getAdminRoomDashboard_(ss);
  const anomalies = scanReservationAnomalies_(ss, requests, rooms);
  writeAdminAnomalies_(ss, anomalies);

  const pending = requests.filter(adminIsPendingRequest_);
  const occupied = rooms.filter((room) => room.status === 'Occupée' || room.status === 'Rotation du jour');
  const free = rooms.filter((room) => room.status === 'Libre');
  const arrivalsToday = rooms.filter((room) => room.movement === 'Arrivée du jour' || room.movement === 'Rotation du jour');
  const departuresToday = rooms.filter((room) => room.movement === 'Départ du jour' || room.movement === 'Rotation du jour');

  return {
    ok: true,
    generatedAt: formatDateTimeFr_(new Date()),
    adminUrl: getAdminUrl_(),
    adminEmailConfigured: Boolean(getAdminEmail_()),
    stats: {
      pending: pending.length,
      anomalies: anomalies.length,
      roomsTotal: rooms.length,
      roomsFree: free.length,
      roomsOccupied: occupied.length,
      arrivalsToday: arrivalsToday.length,
      departuresToday: departuresToday.length,
    },
    pending: pending.slice(0, 60),
    requests: requests.slice(0, 100),
    rooms,
    anomalies: anomalies.slice(0, 100),
  };
}

function adminApproveReservation(token, reference, note) {
  assertAdminToken_(token);

  const lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) {
    return { ok: false, error: 'Une autre action admin est en cours. Réessaie dans quelques secondes.' };
  }

  try {
    const ss = SpreadsheetApp.openById(CONFIG.spreadsheetId);
    ensureAdminSheets_(ss);

    const found = findAdminRequestById_(ss, reference);
    if (!found) throw new Error('Demande introuvable : ' + reference);

    const request = found.record;
    if (isCancelled_(request.status)) {
      throw new Error('Cette demande est déjà refusée ou annulée.');
    }

    if (!request.room) throw new Error('Aucune chambre attribuée sur cette demande.');
    if (!request.arrivalDate || !request.departureDate || request.departureDate <= request.arrivalDate) {
      throw new Error('Dates invalides sur la demande.');
    }

    const conflict = adminFindRoomConflict_(ss, request.room, request.arrivalDate, request.departureDate, request.registryId || request.id);
    if (conflict) {
      const message = 'Conflit avec ' + conflict.id + ' / ' + conflict.occupant + ' du ' + formatFr_(conflict.arrival) + ' au ' + formatFr_(conflict.departure);
      logAdminNotification_('ANOMALIE_VALIDATION', 'critical', request.id, message);
      throw new Error(message);
    }

    const registry = ensureRegistryForRequest_(ss, request, 'Prévu', note);
    const decision = appendDecisionText_('Validée par admin', note);

    found.sheet.getRange(found.rowNumber, 3).setValue('Validée');
    found.sheet.getRange(found.rowNumber, 17).setValue(decision);
    found.sheet.getRange(found.rowNumber, 18).setValue('Oui');
    found.sheet.getRange(found.rowNumber, 19).setValue(registry.id);

    registry.sheet.getRange(registry.rowNumber, 15).setValue('Prévu');
    appendAdminNote_(registry.sheet, registry.rowNumber, 19, decision);

    refreshRegistryOperationalFlags_(ss);
    SpreadsheetApp.flush();

    const mailResult = sendAdminDecisionEmail_(request, 'approved', note);
    logAdminNotification_('VALIDATION', 'info', request.id, 'Réservation validée' + (mailResult.ok ? ' et email envoyé.' : '. Email non envoyé : ' + (mailResult.error || 'aucun email')));

    return {
      ok: true,
      message: 'Réservation validée.',
      emailSent: Boolean(mailResult.ok),
      emailWarning: mailResult.ok ? '' : (mailResult.error || 'Email non envoyé.'),
    };
  } catch (err) {
    return { ok: false, error: err.message || String(err) };
  } finally {
    lock.releaseLock();
  }
}

function adminRejectReservation(token, reference, note) {
  assertAdminToken_(token);

  const lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) {
    return { ok: false, error: 'Une autre action admin est en cours. Réessaie dans quelques secondes.' };
  }

  try {
    const ss = SpreadsheetApp.openById(CONFIG.spreadsheetId);
    ensureAdminSheets_(ss);

    const found = findAdminRequestById_(ss, reference);
    if (!found) throw new Error('Demande introuvable : ' + reference);

    const request = found.record;
    const decision = appendDecisionText_('Refusée par admin', note);

    found.sheet.getRange(found.rowNumber, 3).setValue('Refusée');
    found.sheet.getRange(found.rowNumber, 17).setValue(decision);

    const registry = findAdminRegistryById_(ss, request.registryId || request.id) || findAdminRegistryByRequest_(ss, request);
    if (registry) {
      registry.sheet.getRange(registry.rowNumber, 15).setValue('Refusé');
      appendAdminNote_(registry.sheet, registry.rowNumber, 19, decision);
    }

    refreshRegistryOperationalFlags_(ss);
    SpreadsheetApp.flush();

    const mailResult = sendAdminDecisionEmail_(request, 'rejected', note);
    logAdminNotification_('REFUS', 'info', request.id, 'Réservation refusée' + (mailResult.ok ? ' et email envoyé.' : '. Email non envoyé : ' + (mailResult.error || 'aucun email')));

    return {
      ok: true,
      message: 'Demande refusée.',
      emailSent: Boolean(mailResult.ok),
      emailWarning: mailResult.ok ? '' : (mailResult.error || 'Email non envoyé.'),
    };
  } catch (err) {
    return { ok: false, error: err.message || String(err) };
  } finally {
    lock.releaseLock();
  }
}

function assertAdminPin_(pin) {
  const expected = clean_(PropertiesService.getScriptProperties().getProperty('ADMIN_PIN'));
  if (!expected) {
    throw new Error('ADMIN_PIN non configuré. Va dans Propriétés du script et ajoute ADMIN_PIN.');
  }

  if (clean_(pin) !== expected) {
    throw new Error('PIN administrateur incorrect.');
  }
}

function assertAdminToken_(token) {
  const cleanToken = clean_(token);
  if (!cleanToken || CacheService.getScriptCache().get(adminTokenKey_(cleanToken)) !== 'ok') {
    throw new Error('Session admin expirée. Reconnecte-toi avec le PIN.');
  }

  CacheService.getScriptCache().put(adminTokenKey_(cleanToken), 'ok', 21600);
}

function adminTokenKey_(token) {
  return 'admin:' + token;
}

function ensureAdminSheets_(ss) {
  ensureSheetWithHeaders_(ss, CONFIG.sheets.requests, [
    'Référence',
    'Date demande',
    'Statut',
    'Chambre demandée',
    'Chambre attribuée',
    'Demandeur',
    'Téléphone',
    'Email',
    'Arrivée',
    'Départ',
    'Nuits',
    'Tarif nuit',
    'Montant',
    'Paiement',
    'Source',
    'Commentaire',
    'Décision admin',
    'Intégrée Registre',
    'ID Registre',
    'Disponibilité',
  ]);

  ensureSheetWithHeaders_(ss, CONFIG.sheets.adminNotifications, [
    'Date',
    'Type',
    'Niveau',
    'Référence',
    'Message',
    'Statut',
  ]);

  ensureSheetWithHeaders_(ss, CONFIG.sheets.adminAnomalies, [
    'Date scan',
    'Niveau',
    'Référence',
    'Chambre',
    'Type',
    'Message',
  ]);
}

function ensureSheetWithHeaders_(ss, name, headers) {
  let sheet = ss.getSheetByName(name);
  if (!sheet) sheet = ss.insertSheet(name);

  const current = sheet.getRange(1, 1, 1, headers.length).getValues()[0];
  const empty = current.every((value) => !clean_(value));
  if (empty) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.setFrozenRows(1);
  }

  return sheet;
}

function getAdminRequests_(ss) {
  const sheet = ss.getSheetByName(CONFIG.sheets.requests);
  if (!sheet) return [];

  const last = sheet.getLastRow();
  if (last < CONFIG.requestStartRow) return [];

  return sheet.getRange(CONFIG.requestStartRow, 1, last - CONFIG.requestStartRow + 1, CONFIG.requestCols)
    .getValues()
    .map((row, index) => adminRequestFromRow_(row, CONFIG.requestStartRow + index))
    .filter((request) => request.id)
    .sort((a, b) => (b.createdAtTime || 0) - (a.createdAtTime || 0));
}

function adminRequestFromRow_(row, rowNumber) {
  const arrivalDate = toDate_(row[8]);
  const departureDate = toDate_(row[9]);
  const createdAtDate = row[1] instanceof Date && !isNaN(row[1].getTime()) ? row[1] : null;

  return {
    rowNumber,
    id: clean_(row[0]),
    createdAt: formatMaybeDateTime_(row[1]),
    createdAtTime: createdAtDate ? createdAtDate.getTime() : 0,
    status: clean_(row[2]) || 'À valider',
    requestedRoom: clean_(row[3]),
    room: clean_(row[4]) || clean_(row[3]),
    occupant: clean_(row[5]),
    phone: clean_(row[6]),
    email: clean_(row[7]).toLowerCase(),
    arrival: arrivalDate ? formatFr_(arrivalDate) : clean_(row[8]),
    departure: departureDate ? formatFr_(departureDate) : clean_(row[9]),
    arrivalDate,
    departureDate,
    nights: Number(row[10] || 0),
    nightRate: Number(row[11] || 0),
    amount: Number(row[12] || 0),
    payment: clean_(row[13]),
    source: clean_(row[14]),
    comment: clean_(row[15]),
    adminDecision: clean_(row[16]),
    integratedRegistry: clean_(row[17]),
    registryId: clean_(row[18]) || clean_(row[0]),
    availabilityNote: clean_(row[19]),
  };
}

function getAdminRegistryRows_(ss) {
  const sheet = ss.getSheetByName(CONFIG.sheets.registry);
  if (!sheet) return [];

  const last = sheet.getLastRow();
  if (last < CONFIG.registryStartRow) return [];

  return sheet.getRange(CONFIG.registryStartRow, 1, last - CONFIG.registryStartRow + 1, CONFIG.registryCols)
    .getValues()
    .map((row, index) => {
      const arrival = toDate_(row[6]);
      const departure = toDate_(row[7]);
      return {
        rowNumber: CONFIG.registryStartRow + index,
        id: clean_(row[0]),
        segment: clean_(row[1]),
        room: clean_(row[2]),
        occupant: clean_(row[3]),
        organisme: clean_(row[4]),
        phone: clean_(row[5]),
        arrival,
        departure,
        unit: clean_(row[8]),
        quantity: Number(row[9] || 0),
        unitRate: Number(row[10] || 0),
        amount: Number(row[11] || 0),
        paid: Number(row[12] || 0),
        status: clean_(row[14]),
        checkIn: clean_(row[15]),
        checkOut: clean_(row[16]),
        channel: clean_(row[17]),
        notes: clean_(row[18]),
        anomaly: clean_(row[21]),
        movement: clean_(row[22]),
        financial: clean_(row[23]),
      };
    })
    .filter((entry) => entry.id || entry.room || entry.occupant);
}

function getAdminRoomDashboard_(ss) {
  const rooms = getRooms_(ss);
  const registry = getAdminRegistryRows_(ss);
  const today = today_();

  return rooms.map((room) => {
    const closed = norm_(room.status) !== norm_('Ouverte');
    const active = registry.find((entry) => (
      norm_(entry.room) === norm_(room.room) &&
      entry.arrival &&
      entry.departure &&
      !isCancelled_(entry.status) &&
      today >= entry.arrival &&
      today < entry.departure
    ));
    const arrivalToday = registry.find((entry) => (
      norm_(entry.room) === norm_(room.room) &&
      entry.arrival &&
      sameDate_(entry.arrival, today) &&
      !isCancelled_(entry.status)
    ));
    const departureToday = registry.find((entry) => (
      norm_(entry.room) === norm_(room.room) &&
      entry.departure &&
      sameDate_(entry.departure, today) &&
      !isCancelled_(entry.status)
    ));

    let status = 'Libre';
    if (closed) status = 'Fermée';
    else if (arrivalToday && departureToday) status = 'Rotation du jour';
    else if (active) status = 'Occupée';

    let movement = '';
    if (arrivalToday && departureToday) movement = 'Rotation du jour';
    else if (arrivalToday) movement = 'Arrivée du jour';
    else if (departureToday) movement = 'Départ du jour';

    const occupant = active ? active.occupant : arrivalToday ? arrivalToday.occupant : '';

    return {
      room: room.room,
      type: room.type,
      capacity: room.capacity,
      nightRate: room.nightRate,
      status,
      movement,
      occupant,
      manualStatus: room.status,
    };
  });
}

function scanReservationAnomalies_(ss, requests, rooms) {
  ss = ss || SpreadsheetApp.openById(CONFIG.spreadsheetId);
  requests = requests || getAdminRequests_(ss);
  rooms = rooms || getAdminRoomDashboard_(ss);

  const registry = getAdminRegistryRows_(ss);
  const roomByKey = {};
  rooms.forEach((room) => {
    roomByKey[norm_(room.room)] = room;
  });

  const anomalies = [];
  const requestIds = {};

  requests.forEach((request) => {
    if (requestIds[request.id]) {
      addAdminAnomaly_(anomalies, 'critical', request.id, request.room, 'Doublon demande', 'Référence présente plusieurs fois dans Reservations_Public.');
    }
    requestIds[request.id] = true;

    if (!request.occupant) {
      addAdminAnomaly_(anomalies, 'warning', request.id, request.room, 'Demandeur manquant', 'La demande n’a pas de nom.');
    }
    if (!request.phone && !request.email) {
      addAdminAnomaly_(anomalies, 'warning', request.id, request.room, 'Contact manquant', 'Aucun téléphone ni email.');
    }
    if (!request.arrivalDate || !request.departureDate || request.departureDate <= request.arrivalDate) {
      addAdminAnomaly_(anomalies, 'critical', request.id, request.room, 'Dates invalides', 'Arrivée ou départ invalide.');
    }
    if (request.room && !roomByKey[norm_(request.room)]) {
      addAdminAnomaly_(anomalies, 'critical', request.id, request.room, 'Chambre inconnue', 'La chambre n’existe pas dans l’onglet Chambres.');
    }

    if (!isCancelled_(request.status)) {
      const linked = registry.some((entry) => norm_(entry.id) === norm_(request.registryId || request.id));
      if (!linked) {
        addAdminAnomaly_(anomalies, 'critical', request.id, request.room, 'Non reliée au Registre', 'La demande existe dans Reservations_Public mais pas dans Registre.');
      }
    }

    if (adminIsPendingRequest_(request) && request.createdAtTime) {
      const ageHours = (new Date().getTime() - request.createdAtTime) / 3600000;
      if (ageHours >= 24) {
        addAdminAnomaly_(anomalies, 'warning', request.id, request.room, 'Validation en attente', 'Demande en attente depuis plus de 24 heures.');
      }
    }
  });

  registry.forEach((entry) => {
    if (!entry.id) {
      addAdminAnomaly_(anomalies, 'warning', '', entry.room, 'ID manquant', 'Une ligne Registre n’a pas de référence.');
    }
    if (entry.room && !roomByKey[norm_(entry.room)]) {
      addAdminAnomaly_(anomalies, 'critical', entry.id, entry.room, 'Chambre inconnue', 'Ligne Registre reliée à une chambre absente de Chambres.');
    }
    if (!entry.arrival || !entry.departure || entry.departure <= entry.arrival) {
      addAdminAnomaly_(anomalies, 'critical', entry.id, entry.room, 'Dates Registre invalides', 'Arrivée ou départ invalide dans Registre.');
    }
    if (!entry.occupant) {
      addAdminAnomaly_(anomalies, 'warning', entry.id, entry.room, 'Occupant manquant', 'Ligne Registre sans occupant.');
    }
  });

  const activeRegistry = registry
    .filter((entry) => entry.room && entry.arrival && entry.departure && !isCancelled_(entry.status))
    .sort((a, b) => norm_(a.room).localeCompare(norm_(b.room)) || a.arrival - b.arrival);

  for (let i = 0; i < activeRegistry.length; i++) {
    for (let j = i + 1; j < activeRegistry.length; j++) {
      const a = activeRegistry[i];
      const b = activeRegistry[j];
      if (norm_(a.room) !== norm_(b.room)) break;
      if (b.arrival >= a.departure) break;
      if (b.arrival < a.departure && b.departure > a.arrival) {
        addAdminAnomaly_(anomalies, 'critical', a.id, a.room, 'Chevauchement', 'Conflit avec ' + b.id + ' du ' + formatFr_(b.arrival) + ' au ' + formatFr_(b.departure) + '.');
      }
    }
  }

  requests.forEach((request) => {
    const registryEntry = registry.find((entry) => norm_(entry.id) === norm_(request.registryId || request.id));
    if (!registryEntry) return;

    if (norm_(request.status) === 'validee' && isCancelled_(registryEntry.status)) {
      addAdminAnomaly_(anomalies, 'critical', request.id, request.room, 'Statuts incohérents', 'Demande validée mais ligne Registre annulée/refusée.');
    }
    if (isCancelled_(request.status) && !isCancelled_(registryEntry.status)) {
      addAdminAnomaly_(anomalies, 'critical', request.id, request.room, 'Statuts incohérents', 'Demande refusée mais ligne Registre encore active.');
    }
  });

  return anomalies;
}

function addAdminAnomaly_(list, level, reference, room, type, message) {
  list.push({
    level,
    reference: clean_(reference),
    room: clean_(room),
    type,
    message,
  });
}

function writeAdminAnomalies_(ss, anomalies) {
  ensureAdminSheets_(ss);
  const sheet = ss.getSheetByName(CONFIG.sheets.adminAnomalies);
  const last = sheet.getLastRow();
  if (last > 1) sheet.getRange(2, 1, last - 1, 6).clearContent();
  if (!anomalies.length) return;

  const now = new Date();
  const values = anomalies.map((item) => [
    now,
    item.level,
    item.reference,
    item.room,
    item.type,
    item.message,
  ]);
  sheet.getRange(2, 1, values.length, 6).setValues(values);
}

function findAdminRequestById_(ss, reference) {
  const target = norm_(reference);
  const sheet = ss.getSheetByName(CONFIG.sheets.requests);
  if (!sheet || !target) return null;

  const last = sheet.getLastRow();
  if (last < CONFIG.requestStartRow) return null;

  const values = sheet.getRange(CONFIG.requestStartRow, 1, last - CONFIG.requestStartRow + 1, CONFIG.requestCols).getValues();
  for (let i = 0; i < values.length; i++) {
    const row = values[i];
    const request = adminRequestFromRow_(row, CONFIG.requestStartRow + i);
    if (norm_(request.id) === target || norm_(request.registryId) === target) {
      return {
        sheet,
        rowNumber: CONFIG.requestStartRow + i,
        values: row,
        record: request,
      };
    }
  }

  return null;
}

function findAdminRegistryById_(ss, reference) {
  const target = norm_(reference);
  if (!target) return null;

  const sheet = ss.getSheetByName(CONFIG.sheets.registry);
  if (!sheet) return null;

  const rows = getAdminRegistryRows_(ss);
  const found = rows.find((entry) => norm_(entry.id) === target);
  return found ? { sheet, rowNumber: found.rowNumber, id: found.id, record: found } : null;
}

function findAdminRegistryByRequest_(ss, request) {
  const sheet = ss.getSheetByName(CONFIG.sheets.registry);
  if (!sheet) return null;

  const rows = getAdminRegistryRows_(ss);
  const found = rows.find((entry) => (
    norm_(entry.room) === norm_(request.room) &&
    norm_(entry.occupant) === norm_(request.occupant) &&
    entry.arrival &&
    entry.departure &&
    request.arrivalDate &&
    request.departureDate &&
    entry.arrival.getTime() === request.arrivalDate.getTime() &&
    entry.departure.getTime() === request.departureDate.getTime()
  ));

  return found ? { sheet, rowNumber: found.rowNumber, id: found.id, record: found } : null;
}

function ensureRegistryForRequest_(ss, request, status, note) {
  const existing = findAdminRegistryById_(ss, request.registryId || request.id) || findAdminRegistryByRequest_(ss, request);
  if (existing) return existing;

  const sheet = ss.getSheetByName(CONFIG.sheets.registry);
  if (!sheet) throw new Error('Onglet Registre introuvable.');

  const id = request.registryId || request.id || createRegistryId_();
  const notes = [
    request.comment,
    request.email ? 'Email : ' + request.email : '',
    'Créé par validation admin',
    clean_(note) ? 'Note admin : ' + clean_(note) : '',
  ].filter(Boolean).join(' | ');

  const row = buildRegistryRow_({
    id,
    segment: 'Court séjour',
    room: request.room,
    occupant: request.occupant,
    organisme: '',
    phone: request.phone,
    arrival: request.arrivalDate,
    departure: request.departureDate,
    unit: 'Nuit',
    quantity: request.nights || Math.round((request.departureDate - request.arrivalDate) / 86400000),
    unitRate: request.nightRate,
    amount: request.amount,
    paid: 0,
    status,
    channel: request.email ? 'Mail' : 'Téléphone',
    notes,
    createdAt: new Date(),
  });

  const rowNumber = firstEmptyRow_(sheet, CONFIG.registryStartRow, 1);
  sheet.getRange(rowNumber, 1, 1, CONFIG.registryCols).setValues([row]);
  return { sheet, rowNumber, id, record: adminRequestFromRegistryRow_(row, rowNumber) };
}

function adminRequestFromRegistryRow_(row, rowNumber) {
  return {
    rowNumber,
    id: clean_(row[0]),
    room: clean_(row[2]),
    occupant: clean_(row[3]),
    arrival: toDate_(row[6]),
    departure: toDate_(row[7]),
    status: clean_(row[14]),
  };
}

function adminFindRoomConflict_(ss, room, arrival, departure, exceptId) {
  const targetRoom = norm_(room);
  const targetId = norm_(exceptId);

  return getAdminRegistryRows_(ss).find((entry) => (
    norm_(entry.room) === targetRoom &&
    norm_(entry.id) !== targetId &&
    entry.arrival &&
    entry.departure &&
    !isCancelled_(entry.status) &&
    entry.arrival < departure &&
    entry.departure > arrival
  )) || null;
}

function adminIsPendingRequest_(request) {
  const status = norm_(request.status);
  const decision = norm_(request.adminDecision);
  return ['a valider', 'en attente', 'nouvelle', 'demande'].includes(status) ||
    (status === 'integree' && !decision.includes('validee'));
}

function appendDecisionText_(decision, note) {
  const parts = [decision + ' le ' + formatDateTimeFr_(new Date())];
  if (clean_(note)) parts.push('Note : ' + clean_(note));
  return parts.join(' - ');
}

function appendAdminNote_(sheet, rowNumber, colNumber, note) {
  const previous = clean_(sheet.getRange(rowNumber, colNumber).getValue());
  const next = [previous, note].filter(Boolean).join(' | ');
  sheet.getRange(rowNumber, colNumber).setValue(next);
}

function sendAdminDecisionEmail_(request, decision, note) {
  try {
    if (!request.email) return { ok: false, skipped: true, error: 'Aucun email sur la demande.' };
    if (!isValidEmail_(request.email)) return { ok: false, error: 'Email demandeur invalide.' };

    const approved = decision === 'approved';
    const subject = (approved ? 'Réservation validée - ' : 'Demande de réservation refusée - ') + request.id;
    const title = approved ? 'Votre réservation est validée' : 'Votre demande de réservation est refusée';
    const intro = approved
      ? 'Votre réservation a été validée par la réception.'
      : 'Votre demande ne peut pas être validée par la réception.';

    const rows = [
      receiptEmailRow_('Référence', request.id),
      receiptEmailRow_('Statut', approved ? 'Validée' : 'Refusée'),
      receiptEmailRow_('Demandeur', request.occupant),
      receiptEmailRow_('Téléphone', request.phone || '-'),
      receiptEmailRow_('Email', request.email || '-'),
      receiptEmailRow_('Chambre', request.room),
      receiptEmailRow_('Arrivée', request.arrival),
      receiptEmailRow_('Départ', request.departure),
      receiptEmailRow_('Nombre de nuits', request.nights),
      receiptEmailRow_('Montant prévisionnel', moneyText_(request.amount)),
    ].join('');

    const noteHtml = clean_(note)
      ? '<p style="margin-top:14px"><strong>Message réception :</strong> ' + escapeHtml_(note) + '</p>'
      : '';

    const htmlBody = [
      '<div style="font-family:Arial,sans-serif;color:#222;max-width:680px;margin:auto;border:1px solid #ddd;border-radius:8px;padding:18px">',
      '<h2 style="margin:0 0 8px;color:#5a0d0d">' + escapeHtml_(title) + '</h2>',
      '<p>' + escapeHtml_(intro) + '</p>',
      '<table style="width:100%;border-collapse:collapse;margin-top:14px">',
      rows,
      '</table>',
      approved ? accessInstructionsHtml_() : '',
      noteHtml,
      '</div>',
    ].join('');

    const body = [
      title,
      '',
      intro,
      '',
      'Référence : ' + request.id,
      'Statut : ' + (approved ? 'Validée' : 'Refusée'),
      'Demandeur : ' + request.occupant,
      'Téléphone : ' + (request.phone || '-'),
      'Email : ' + (request.email || '-'),
      'Chambre : ' + request.room,
      'Arrivée : ' + request.arrival,
      'Départ : ' + request.departure,
      'Nombre de nuits : ' + request.nights,
      'Montant prévisionnel : ' + moneyText_(request.amount),
      approved ? '\n' + accessInstructionsText_() : '',
      clean_(note) ? '\nMessage réception : ' + clean_(note) : '',
    ].join('\n');

    MailApp.sendEmail({
      to: request.email,
      subject,
      body,
      htmlBody,
      name: 'Gestion chambre',
    });

    return { ok: true };
  } catch (err) {
    return { ok: false, error: err.message || String(err) };
  }
}

function notifyAdminNewReservation_(receipt, comment, emailWarning) {
  const email = getAdminEmail_();
  if (!email) {
    logAdminNotification_('NOUVELLE_RESERVATION', 'warning', receipt.id, 'ADMIN_EMAIL absent : notification admin non envoyée.');
    return { ok: false, error: 'ADMIN_EMAIL absent.' };
  }

  const adminUrl = getAdminUrl_();
  const subject = 'Nouvelle réservation à valider - ' + receipt.id;
  const htmlBody = [
    '<div style="font-family:Arial,sans-serif;color:#222;max-width:680px;margin:auto;border:1px solid #ddd;border-radius:8px;padding:18px">',
    '<h2 style="margin:0 0 8px;color:#5a0d0d">Nouvelle réservation à valider</h2>',
    '<table style="width:100%;border-collapse:collapse;margin-top:14px">',
    receiptEmailRow_('Référence', receipt.id),
    receiptEmailRow_('Demandeur', receipt.occupant),
    receiptEmailRow_('Téléphone', receipt.phone || '-'),
    receiptEmailRow_('Email', receipt.email || '-'),
    receiptEmailRow_('Chambre', receipt.room),
    receiptEmailRow_('Arrivée', receipt.arrival),
    receiptEmailRow_('Départ', receipt.departure),
    receiptEmailRow_('Montant', moneyText_(receipt.amount)),
    '</table>',
    clean_(comment) ? '<p><strong>Commentaire :</strong> ' + escapeHtml_(comment) + '</p>' : '',
    clean_(emailWarning) ? '<p style="color:#991b1b"><strong>Alerte email client :</strong> ' + escapeHtml_(emailWarning) + '</p>' : '',
    adminUrl ? '<p><a href="' + escapeHtml_(adminUrl) + '" style="display:inline-block;background:#5a0d0d;color:#fff;padding:10px 14px;border-radius:8px;text-decoration:none">Ouvrir le module admin</a></p>' : '',
    '</div>',
  ].join('');

  MailApp.sendEmail({
    to: email,
    subject,
    body: 'Nouvelle réservation à valider : ' + receipt.id + '\nAdmin : ' + adminUrl,
    htmlBody,
    name: 'Gestion chambre',
  });

  logAdminNotification_('NOUVELLE_RESERVATION', 'info', receipt.id, 'Notification admin envoyée à ' + email);
  return { ok: true };
}

function notifyAdminAnomalies_(anomalies, avoidDuplicate) {
  const email = getAdminEmail_();
  if (!email) {
    logAdminNotification_('ANOMALIES', 'warning', '', 'ADMIN_EMAIL absent : anomalies non envoyées.');
    return { ok: false, error: 'ADMIN_EMAIL absent.' };
  }

  const payload = anomalies.map((item) => [item.level, item.reference, item.room, item.type, item.message].join('|')).join('\n');
  const hash = Utilities.base64EncodeWebSafe(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, payload));
  const props = PropertiesService.getScriptProperties();
  if (avoidDuplicate && clean_(props.getProperty('LAST_ADMIN_ANOMALY_HASH')) === hash) {
    return { ok: true, skipped: true };
  }
  props.setProperty('LAST_ADMIN_ANOMALY_HASH', hash);

  const adminUrl = getAdminUrl_();
  const rows = anomalies.slice(0, 20).map((item) => (
    '<tr>' +
    '<td style="padding:8px;border-bottom:1px solid #eee">' + escapeHtml_(item.level) + '</td>' +
    '<td style="padding:8px;border-bottom:1px solid #eee">' + escapeHtml_(item.reference || '-') + '</td>' +
    '<td style="padding:8px;border-bottom:1px solid #eee">' + escapeHtml_(item.room || '-') + '</td>' +
    '<td style="padding:8px;border-bottom:1px solid #eee"><strong>' + escapeHtml_(item.type) + '</strong><br>' + escapeHtml_(item.message) + '</td>' +
    '</tr>'
  )).join('');

  MailApp.sendEmail({
    to: email,
    subject: 'Anomalies réservation détectées - ' + anomalies.length,
    body: 'Anomalies détectées : ' + anomalies.length + '\nAdmin : ' + adminUrl,
    htmlBody: [
      '<div style="font-family:Arial,sans-serif;color:#222;max-width:760px;margin:auto;border:1px solid #ddd;border-radius:8px;padding:18px">',
      '<h2 style="margin:0 0 8px;color:#991b1b">Anomalies réservation détectées</h2>',
      '<p>' + anomalies.length + ' anomalie(s) détectée(s). Les 20 premières sont listées ci-dessous.</p>',
      '<table style="width:100%;border-collapse:collapse;margin-top:14px">',
      '<tr><th align="left">Niveau</th><th align="left">Référence</th><th align="left">Chambre</th><th align="left">Anomalie</th></tr>',
      rows,
      '</table>',
      adminUrl ? '<p><a href="' + escapeHtml_(adminUrl) + '" style="display:inline-block;background:#991b1b;color:#fff;padding:10px 14px;border-radius:8px;text-decoration:none">Ouvrir le module admin</a></p>' : '',
      '</div>',
    ].join(''),
    name: 'Gestion chambre',
  });

  logAdminNotification_('ANOMALIES', 'critical', '', anomalies.length + ' anomalie(s) envoyée(s) à ' + email);
  return { ok: true };
}

function logAdminNotification_(type, level, reference, message) {
  try {
    const ss = SpreadsheetApp.openById(CONFIG.spreadsheetId);
    ensureAdminSheets_(ss);
    ss.getSheetByName(CONFIG.sheets.adminNotifications).appendRow([
      new Date(),
      type,
      level,
      reference || '',
      message || '',
      '',
    ]);
  } catch (e) {
    // Journalisation secondaire volontairement silencieuse.
  }
}

function getAdminEmail_() {
  return clean_(PropertiesService.getScriptProperties().getProperty('ADMIN_EMAIL')) ||
    clean_(Session.getEffectiveUser().getEmail());
}

function getAdminUrl_() {
  try {
    const url = ScriptApp.getService().getUrl();
    if (!url) return '';
    return url + (url.indexOf('?') === -1 ? '?view=admin' : '&view=admin');
  } catch (e) {
    return '';
  }
}

function formatMaybeDateTime_(value) {
  if (value instanceof Date && !isNaN(value.getTime())) return formatDateTimeFr_(value);
  return clean_(value);
}
