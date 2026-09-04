// Sidebar Google Sheets pour ajout manuel dans le Registre.
// Menu : Gestion chambres > Ajout manuel résident

const RESIDENT_INPUT_CONFIG = {
  sheetName: 'Saisie résidents',
  startRow: 8,
  cols: 11,
};

function openResidentSidebar() {
  const html = HtmlService.createTemplateFromFile('ResidentSidebar')
    .evaluate()
    .setTitle('Ajout manuel résident');
  SpreadsheetApp.getUi().showSidebar(html);
}

function getResidentSidebarBootstrap() {
  try {
    const ss = SpreadsheetApp.openById(CONFIG.spreadsheetId);
    const rooms = getRooms_(ss)
      .filter((room) => norm_(room.status) === norm_('Ouverte'))
      .map((room) => ({
        room: room.room,
        floor: room.floor,
        type: room.type,
        capacity: room.capacity,
        nightRate: room.nightRate,
        monthRate: room.monthRate,
        status: room.status,
      }));

    return {
      ok: true,
      today: toIso_(new Date()),
      rooms,
      statuses: ['Prévu', 'En cours', 'Parti', 'Annulé', 'No-show'],
      channels: ['Direct', 'Téléphone', 'Mail', 'Association', 'Organisme', 'Partenaire'],
      units: ['Mois', 'Nuit', 'Forfait'],
    };
  } catch (err) {
    return { ok: false, error: err && err.message ? err.message : String(err) };
  }
}

function createManualResidentEntry(payload) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) {
    return { ok: false, error: 'Une autre écriture est en cours. Réessaie dans quelques secondes.' };
  }

  try {
    payload = payload || {};

    const ss = SpreadsheetApp.openById(CONFIG.spreadsheetId);
    const sheet = ss.getSheetByName(CONFIG.sheets.registry);
    if (!sheet) throw new Error('Onglet Registre introuvable.');

    const room = clean_(payload.room);
    const occupant = clean_(payload.occupant);
    const organisme = clean_(payload.organisme);
    const phone = clean_(payload.phone);
    const arrival = parseDate_(payload.arrival, 'arrivée');
    const departure = parseDate_(payload.departure, 'départ');
    const unit = clean_(payload.unit) || 'Mois';
    const quantity = Number(payload.quantity || 0);
    const unitRate = Number(payload.unitRate || 0);
    const paid = Number(payload.paid || 0);
    const channel = clean_(payload.channel) || 'Direct';
    const status = clean_(payload.status) || manualDefaultStatus_(arrival, departure);
    const notes = [
      clean_(payload.notes),
      'Ajout manuel sidebar',
    ].filter(Boolean).join(' | ');

    if (!room) throw new Error('Choisis une chambre.');
    if (!occupant) throw new Error('Le nom du résident est obligatoire.');
    if (departure <= arrival) throw new Error('La date de départ doit être après la date d’arrivée.');
    if (!['prévu', 'prevu', 'en cours', 'parti', 'annulé', 'annule', 'no-show'].includes(norm_(status))) {
      throw new Error('Statut invalide. Utilise Prévu, En cours, Parti, Annulé ou No-show.');
    }

    const rooms = getRooms_(ss);
    if (!rooms.some((item) => norm_(item.room) === norm_(room))) {
      throw new Error('Chambre inconnue dans l’onglet Chambres : ' + room);
    }

    const conflict = manualFindRoomConflict_(ss, room, arrival, departure);
    if (conflict) {
      throw new Error('Chambre déjà occupée par ' + conflict.occupant + ' du ' + formatFr_(conflict.arrival) + ' au ' + formatFr_(conflict.departure) + ' (' + conflict.id + ').');
    }

    const amount = payload.amount === '' || payload.amount == null
      ? quantity * unitRate
      : Number(payload.amount || 0);

    const row = buildRegistryRow_({
      id: manualCreateRegistryId_(),
      segment: 'Résident',
      room,
      occupant,
      organisme,
      phone,
      arrival,
      departure,
      unit,
      quantity,
      unitRate,
      amount,
      paid,
      status,
      channel,
      notes,
      createdAt: new Date(),
    });

    const targetRow = firstEmptyRow_(sheet, CONFIG.registryStartRow, 1);
    sheet.getRange(targetRow, 1, 1, CONFIG.registryCols).setValues([row]);

    refreshRegistryOperationalFlags_(ss);
    repairRoomStatusFormulas_(ss);
    if (typeof refreshVisualPlanning_ === 'function') refreshVisualPlanning_(ss);
    SpreadsheetApp.flush();

    return {
      ok: true,
      id: row[0],
      rowNumber: targetRow,
      message: 'Résident ajouté dans le Registre : ' + row[0],
    };
  } catch (err) {
    return { ok: false, error: err && err.message ? err.message : String(err) };
  } finally {
    lock.releaseLock();
  }
}

function integrateResidentInputRows() {
  const result = syncResidentInputRows_();
  const ss = SpreadsheetApp.openById(CONFIG.spreadsheetId);
  ss.toast(result.message, 'Gestion chambres', 6);
  return result.message;
}

function setupSheetSyncTrigger() {
  ScriptApp.getProjectTriggers().forEach((trigger) => {
    if (trigger.getHandlerFunction() === 'handleSheetEditSync') {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  ScriptApp.newTrigger('handleSheetEditSync')
    .forSpreadsheet(CONFIG.spreadsheetId)
    .onEdit()
    .create();

  return 'Synchro automatique installée : Saisie résidents, Registre et Chambres seront recalculés après modification.';
}

function handleSheetEditSync(e) {
  try {
    const range = e && e.range;
    if (!range) return;

    const sheet = range.getSheet();
    const sheetName = sheet.getName();
    const row = range.getRow();
    const col = range.getColumn();
    const ss = sheet.getParent();

    if (
      sheetName === RESIDENT_INPUT_CONFIG.sheetName &&
      row >= RESIDENT_INPUT_CONFIG.startRow &&
      col <= RESIDENT_INPUT_CONFIG.cols
    ) {
      syncResidentInputRows_(ss);
      return;
    }

    if (
      sheetName === CONFIG.sheets.registry ||
      (sheetName === CONFIG.sheets.rooms && row >= CONFIG.roomStartRow && col <= 8)
    ) {
      refreshRegistryOperationalFlags_(ss);
      repairRoomStatusFormulas_(ss);
      if (typeof refreshVisualPlanning_ === 'function') refreshVisualPlanning_(ss);
      SpreadsheetApp.flush();
    }
  } catch (err) {
    console.error(err && err.stack ? err.stack : err);
  }
}

function syncResidentInputRows_(ss) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) {
    return { ok: false, added: 0, updated: 0, skipped: 0, errors: [], message: 'Synchronisation déjà en cours.' };
  }

  try {
    ss = ss || SpreadsheetApp.openById(CONFIG.spreadsheetId);
    const inputSheet = ss.getSheetByName(RESIDENT_INPUT_CONFIG.sheetName);
    const registrySheet = ss.getSheetByName(CONFIG.sheets.registry);
    if (!inputSheet) throw new Error('Onglet Saisie résidents introuvable.');
    if (!registrySheet) throw new Error('Onglet Registre introuvable.');

    const inputLast = inputSheet.getLastRow();
    if (inputLast < RESIDENT_INPUT_CONFIG.startRow) {
      return { ok: true, added: 0, updated: 0, skipped: 0, errors: [], message: 'Aucune ligne à intégrer.' };
    }

    const rooms = getRooms_(ss);
    const roomsByName = rooms.reduce((index, room) => {
      index[norm_(room.room)] = room;
      return index;
    }, {});

    const registryLast = registrySheet.getLastRow();
    const registryValues = registryLast >= CONFIG.registryStartRow
      ? registrySheet.getRange(CONFIG.registryStartRow, 1, registryLast - CONFIG.registryStartRow + 1, CONFIG.registryCols).getValues()
      : [];

    const inputRows = inputSheet
      .getRange(
        RESIDENT_INPUT_CONFIG.startRow,
        1,
        inputLast - RESIDENT_INPUT_CONFIG.startRow + 1,
        RESIDENT_INPUT_CONFIG.cols
      )
      .getValues();

    let added = 0;
    let updated = 0;
    let skipped = 0;
    let nextEmptyRow = firstEmptyRow_(registrySheet, CONFIG.registryStartRow, 1);
    const errors = [];

    inputRows.forEach((row, index) => {
      const sourceRow = RESIDENT_INPUT_CONFIG.startRow + index;
      if (residentInputRowIsEmpty_(row)) return;

      try {
        const data = residentDataFromInputRow_(row, roomsByName);
        const existingIndex = findResidentRegistryMatch_(registryValues, data);
        const existing = existingIndex >= 0 ? registryValues[existingIndex] : null;
        const registryRow = buildRegistryRow_({
          id: clean_(existing && existing[0]) || manualCreateRegistryId_(),
          segment: 'Résident',
          room: data.room,
          occupant: data.occupant,
          organisme: data.organisme || clean_(existing && existing[4]),
          phone: data.phone || clean_(existing && existing[5]),
          arrival: data.arrival,
          departure: data.departure,
          unit: 'Mois',
          quantity: 1,
          unitRate: data.monthRate,
          amount: data.monthRate,
          paid: data.paid,
          status: data.status,
          channel: data.channel,
          notes: mergeResidentInputNotes_(data.notes, clean_(existing && existing[18])),
          createdAt: existing && existing[24] ? existing[24] : new Date(),
        });

        if (existingIndex >= 0) {
          registrySheet.getRange(CONFIG.registryStartRow + existingIndex, 1, 1, CONFIG.registryCols).setValues([registryRow]);
          registryValues[existingIndex] = registryRow;
          updated++;
        } else {
          registrySheet.getRange(nextEmptyRow, 1, 1, CONFIG.registryCols).setValues([registryRow]);
          registryValues.push(registryRow);
          nextEmptyRow++;
          added++;
        }
      } catch (err) {
        skipped++;
        errors.push('Ligne ' + sourceRow + ' : ' + (err && err.message ? err.message : String(err)));
      }
    });

    if (added || updated) {
      refreshRegistryOperationalFlags_(ss);
      repairRoomStatusFormulas_(ss);
      if (typeof refreshVisualPlanning_ === 'function') refreshVisualPlanning_(ss);
      SpreadsheetApp.flush();
    }

    return {
      ok: errors.length === 0,
      added,
      updated,
      skipped,
      errors,
      message: residentSyncMessage_(added, updated, skipped, errors),
    };
  } finally {
    lock.releaseLock();
  }
}

function residentDataFromInputRow_(row, roomsByName) {
  const room = clean_(row[0]);
  const occupant = clean_(row[1]);
  const arrival = toDate_(row[4]);
  const departure = toDate_(row[5]);
  const roomInfo = roomsByName[norm_(room)];
  const status = clean_(row[9]) || manualDefaultStatus_(arrival, departure);

  if (!room) throw new Error('chambre obligatoire.');
  if (!occupant) throw new Error('nom du résident obligatoire.');
  if (!arrival) throw new Error('date d’entrée invalide.');
  if (!departure) throw new Error('date de sortie invalide.');
  if (departure <= arrival) throw new Error('la sortie doit être après l’entrée.');
  if (!roomInfo) throw new Error('chambre inconnue dans l’onglet Chambres : ' + room);
  if (!['prévu', 'prevu', 'en cours', 'parti', 'annulé', 'annule', 'no-show'].includes(norm_(status))) {
    throw new Error('statut invalide. Utilise Prévu, En cours, Parti, Annulé ou No-show.');
  }

  return {
    room,
    occupant,
    organisme: clean_(row[2]),
    phone: clean_(row[3]),
    arrival,
    departure,
    monthRate: Number(row[6] || roomInfo.monthRate || 0),
    paid: Number(row[7] || 0),
    channel: clean_(row[8]) || 'Direct',
    status,
    notes: clean_(row[10]),
  };
}

function findResidentRegistryMatch_(registryValues, data) {
  const targetOccupant = norm_(data.occupant);
  const targetPhone = residentDigits_(data.phone);
  let sameOccupant = -1;

  for (let i = 0; i < registryValues.length; i++) {
    const row = registryValues[i];
    if (norm_(row[1]) !== norm_('Résident')) continue;
    if (isCancelled_(row[14])) continue;
    if (norm_(row[3]) !== targetOccupant) continue;

    const rowRoom = clean_(row[2]);
    const rowArrival = toDate_(row[6]);
    const rowDeparture = toDate_(row[7]);
    const rowPhone = residentDigits_(row[5]);

    if (
      norm_(rowRoom) === norm_(data.room) &&
      sameDate_(rowArrival, data.arrival) &&
      sameDate_(rowDeparture, data.departure)
    ) {
      return i;
    }

    if (targetPhone && rowPhone === targetPhone) return i;
    if (rowArrival && rowDeparture && rowArrival < data.departure && rowDeparture > data.arrival) return i;
    if (sameOccupant === -1) sameOccupant = i;
  }

  return sameOccupant;
}

function residentInputRowIsEmpty_(row) {
  return row.every((cell) => !clean_(cell));
}

function mergeResidentInputNotes_(inputNotes, existingNotes) {
  const marker = 'Synchronisé depuis Saisie résidents';
  return [inputNotes, existingNotes, marker]
    .filter(Boolean)
    .filter((value, index, values) => values.indexOf(value) === index)
    .join(' | ');
}

function residentDigits_(value) {
  return clean_(value).replace(/\D/g, '');
}

function residentSyncMessage_(added, updated, skipped, errors) {
  const parts = [
    added + ' ajout(s)',
    updated + ' mise(s) à jour',
  ];

  if (skipped) parts.push(skipped + ' ligne(s) ignorée(s)');
  if (errors.length) parts.push(errors.slice(0, 3).join(' / '));
  return 'Saisie résidents synchronisée : ' + parts.join(', ') + '.';
}

function manualDefaultStatus_(arrival, departure) {
  const today = today_();
  return today >= arrival && today < departure ? 'En cours' : 'Prévu';
}

function manualCreateRegistryId_() {
  return 'MAN-' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd-HHmmss');
}

function manualFindRoomConflict_(ss, room, arrival, departure) {
  const sheet = ss.getSheetByName(CONFIG.sheets.registry);
  const last = sheet.getLastRow();
  if (last < CONFIG.registryStartRow) return null;

  const rows = sheet.getRange(CONFIG.registryStartRow, 1, last - CONFIG.registryStartRow + 1, CONFIG.registryCols).getValues();
  const targetRoom = norm_(room);

  for (let i = 0; i < rows.length; i++) {
    const row = rows[i];
    const rowRoom = norm_(row[2]);
    const rowArrival = toDate_(row[6]);
    const rowDeparture = toDate_(row[7]);
    const rowStatus = clean_(row[14]);

    if (rowRoom !== targetRoom) continue;
    if (!rowArrival || !rowDeparture) continue;
    if (isCancelled_(rowStatus)) continue;
    if (rowArrival < departure && rowDeparture > arrival) {
      return {
        id: clean_(row[0]) || 'sans référence',
        occupant: clean_(row[3]) || 'occupant non renseigné',
        arrival: rowArrival,
        departure: rowDeparture,
      };
    }
  }

  return null;
}
