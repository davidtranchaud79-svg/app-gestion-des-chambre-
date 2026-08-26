const CONFIG = {
  spreadsheetId: '1mB_4WNuqw1sysQIa4LUPqktFOngqOTS61pj6v3PPtnE',
  sheets: {
    params: 'Parametres',
    rooms: 'Chambres',
    registry: 'Registre',
    requests: 'Reservations_Public',
    adminNotifications: 'Admin_Notifications',
    adminAnomalies: 'Anomalies_Reservation',
  },
  roomStartRow: 6,
  registryStartRow: 6,
  requestStartRow: 2,
  requestCols: 20,
  registryCols: 25,
};

function doGet(e) {
  const params = (e && e.parameter) || {};
  const view = norm_(params.view || params.page || '');
  const isAdmin = view === 'admin';
  const template = HtmlService.createTemplateFromFile(isAdmin ? 'Admin' : 'Index');
  template.bootstrap = JSON.stringify({
    title: isAdmin ? 'Administration réservations' : 'Gestion chambre',
    today: toIso_(new Date()),
    view: isAdmin ? 'admin' : 'public',
  });

  return template.evaluate()
    .setTitle(isAdmin ? 'Administration réservations' : 'Gestion chambre')
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
    const email = clean_(payload.email).toLowerCase();
    const comment = clean_(payload.comment);
    const requestedRoom = clean_(payload.requestedRoom);

    if (!name) throw new Error('Le nom est obligatoire.');
    if (!phone && !email) throw new Error('Ajoute au moins un téléphone ou un email.');
    if (email && !isValidEmail_(email)) throw new Error('Adresse email invalide.');

    const ss = SpreadsheetApp.openById(CONFIG.spreadsheetId);
    const availability = checkAvailability(payload);
    if (!availability.ok) return availability;

    const selected = selectAvailableRoom_(availability.rooms, requestedRoom);
    if (!selected) {
      return { ok: false, error: 'Aucune chambre disponible sur cette période.' };
    }

    const registrySheet = ss.getSheetByName(CONFIG.sheets.registry);
    const requestSheet = ss.getSheetByName(CONFIG.sheets.requests);
    if (!registrySheet) throw new Error('Onglet Registre introuvable.');
    if (!requestSheet) throw new Error('Onglet Reservations_Public introuvable.');

    if (hasRegistryDuplicate_(registrySheet, selected.room, name, stay.arrival, stay.departure)) {
      return { ok: false, error: 'Cette demande existe déjà dans le Registre.' };
    }

    const registryId = createRegistryId_();
    const createdAt = new Date();
    const channel = email ? 'Mail' : 'Téléphone';
    const nightRate = Number(selected.nightRate || 0);
    const amount = stay.nights * nightRate;
    const publicStatus = 'À valider';
    const registryStatus = 'À valider';
    const notes = [
      comment,
      'Demande micro-app',
      email ? 'Email : ' + email : '',
      'En attente validation réception',
    ].filter(Boolean).join(' | ');

    const requestRow = [
      registryId,
      createdAt,
      publicStatus,
      requestedRoom || '',
      selected.room,
      name,
      phone,
      email,
      stay.arrival,
      stay.departure,
      stay.nights,
      nightRate,
      amount,
      '',
      'Micro-app',
      comment,
      'En attente réception',
      'Non',
      registryId,
      selected.reason || '',
    ];

    const registryRow = buildRegistryRow_({
      id: registryId,
      segment: 'Court séjour',
      room: selected.room,
      occupant: name,
      organisme: '',
      phone,
      arrival: stay.arrival,
      departure: stay.departure,
      unit: 'Nuit',
      quantity: stay.nights,
      unitRate: nightRate,
      amount,
      paid: 0,
      status: registryStatus,
      channel,
      notes,
      createdAt,
    });

    const requestTargetRow = firstEmptyRow_(requestSheet, CONFIG.requestStartRow, 1);
    requestSheet.getRange(requestTargetRow, 1, 1, CONFIG.requestCols).setValues([requestRow]);

    const registryTargetRow = firstEmptyRow_(registrySheet, CONFIG.registryStartRow, 1);
    registrySheet.getRange(registryTargetRow, 1, 1, CONFIG.registryCols).setValues([registryRow]);

    refreshRegistryOperationalFlags_(ss);
    repairRoomStatusFormulas_(ss);
    SpreadsheetApp.flush();

    const receipt = {
      id: registryId,
      status: registryStatus,
      occupant: name,
      phone,
      email,
      room: selected.room,
      arrival: formatFr_(stay.arrival),
      departure: formatFr_(stay.departure),
      nights: stay.nights,
      nightRate,
      amount,
      createdAt: formatDateTimeFr_(createdAt),
      paymentNotice: 'Montant prévisionnel, validation par la réception.',
    };

    const mailResult = email ? sendReservationReceiptEmail_(email, receipt) : { ok: false, skipped: true };
    const emailWarning = email && !mailResult.ok ? mailResult.error : '';

    try {
      notifyAdminNewReservation_(receipt, comment, emailWarning);
      const anomalies = scanReservationAnomalies_(ss);
      if (anomalies.length) notifyAdminAnomalies_(anomalies, true);
    } catch (notifyErr) {
      logAdminNotification_('ERREUR_NOTIFICATION', 'warning', registryId, notifyErr.message || String(notifyErr));
    }

    return {
      ok: true,
      id: registryId,
      room: selected.room,
      nights: stay.nights,
      amount,
      status: registryStatus,
      emailSent: Boolean(email && mailResult.ok),
      emailWarning,
      message: emailWarning
        ? 'Demande enregistrée et intégrée au Registre. Reçu affiché, mais email non envoyé : ' + emailWarning
        : email
          ? 'Demande enregistrée et intégrée au Registre. Reçu envoyé par email.'
          : 'Demande enregistrée et intégrée au Registre. Reçu disponible ci-dessous.',
      receipt,
    };
  } catch (err) {
    return { ok: false, error: err.message || String(err) };
  } finally {
    lock.releaseLock();
  }
}

function testMailAuthorization() {
  const quota = MailApp.getRemainingDailyQuota();
  return 'MailApp autorisé. Quota restant aujourd’hui : ' + quota;
}

function repairReservationSystem() {
  const ss = SpreadsheetApp.openById(CONFIG.spreadsheetId);
  refreshRegistryOperationalFlags_(ss);
  repairRoomStatusFormulas_(ss);
  return 'Système réservation recâblé : Registre recalculé et Chambres reliées au Registre complet.';
}

function sendReservationReceiptEmail_(recipientEmail, receipt) {
  try {
    if (!recipientEmail) return { ok: false, skipped: true, error: 'Aucun email renseigné.' };
    if (!isValidEmail_(recipientEmail)) return { ok: false, error: 'Adresse email invalide.' };

    const subject = 'Reçu de demande de réservation - ' + receipt.id;
    const htmlBody = receiptEmailHtml_(receipt);
    const body = plainReceiptText_(receipt);

    MailApp.sendEmail({
      to: recipientEmail,
      subject,
      body,
      htmlBody,
      name: 'Gestion chambre',
    });

    return { ok: true };
  } catch (err) {
    return { ok: false, error: err && err.message ? err.message : String(err) };
  }
}

function receiptEmailHtml_(receipt) {
  return [
    '<div style="font-family:Arial,sans-serif;color:#222;max-width:680px;margin:auto;border:1px solid #ddd;border-radius:8px;padding:18px">',
    '<h2 style="margin:0 0 8px;color:#5a0d0d">Reçu de demande de réservation</h2>',
    '<p>Votre demande a bien été enregistrée. La réservation reste soumise à validation par la réception.</p>',
    '<table style="width:100%;border-collapse:collapse;margin-top:14px">',
    receiptEmailRow_('Référence', receipt.id),
    receiptEmailRow_('Statut', receipt.status),
    receiptEmailRow_('Date de demande', receipt.createdAt),
    receiptEmailRow_('Demandeur', receipt.occupant),
    receiptEmailRow_('Téléphone', receipt.phone || '-'),
    receiptEmailRow_('Email', receipt.email || '-'),
    receiptEmailRow_('Chambre', receipt.room),
    receiptEmailRow_('Arrivée', receipt.arrival),
    receiptEmailRow_('Départ', receipt.departure),
    receiptEmailRow_('Nombre de nuits', receipt.nights),
    receiptEmailRow_('Tarif nuit', moneyText_(receipt.nightRate)),
    receiptEmailRow_('Montant prévisionnel', moneyText_(receipt.amount)),
    '</table>',
    shouldShowAccessInstructions_(receipt.status) ? accessInstructionsHtml_() : '',
    '<p style="margin-top:16px;color:#667085">' + escapeHtml_(receipt.paymentNotice || 'Validation par la réception.') + '</p>',
    '</div>',
  ].join('');
}

function receiptEmailRow_(label, value) {
  return '<tr>' +
    '<td style="padding:8px;border-bottom:1px solid #eee;color:#667085">' + escapeHtml_(label) + '</td>' +
    '<td style="padding:8px;border-bottom:1px solid #eee;font-weight:bold;text-align:right">' + escapeHtml_(value) + '</td>' +
    '</tr>';
}

function plainReceiptText_(receipt) {
  return [
    'Reçu de demande de réservation',
    '',
    'Référence : ' + receipt.id,
    'Statut : ' + receipt.status,
    'Date de demande : ' + receipt.createdAt,
    'Demandeur : ' + receipt.occupant,
    'Téléphone : ' + (receipt.phone || '-'),
    'Email : ' + (receipt.email || '-'),
    'Chambre : ' + receipt.room,
    'Arrivée : ' + receipt.arrival,
    'Départ : ' + receipt.departure,
    'Nombre de nuits : ' + receipt.nights,
    'Tarif nuit : ' + moneyText_(receipt.nightRate),
    'Montant prévisionnel : ' + moneyText_(receipt.amount),
    '',
    shouldShowAccessInstructions_(receipt.status) ? accessInstructionsText_() : '',
    '',
    receipt.paymentNotice || 'Validation par la réception.',
  ].filter(function(line) { return line !== null && line !== undefined; }).join('\n');
}

function accessInstructionsHtml_() {
  const info = getAccessInfo_();
  const parts = [];

  if (info.mailboxCode) {
    parts.push('<p style="margin:6px 0"><strong>Code boîte aux lettres :</strong> ' + escapeHtml_(info.mailboxCode) + '</p>');
  }
  if (info.keyBoxCode) {
    parts.push('<p style="margin:6px 0"><strong>Code boîte à clé :</strong> ' + escapeHtml_(info.keyBoxCode) + '</p>');
  }
  if (info.badgeNotice) {
    parts.push('<p style="margin:8px 0 0">' + escapeHtml_(info.badgeNotice) + '</p>');
  }

  if (!parts.length) return '';

  return [
    '<div style="margin-top:16px;padding:12px;border-radius:8px;background:#fff7ed;border:1px solid #fed7aa;color:#7c2d12">',
    '<h3 style="margin:0 0 8px;color:#7c2d12">Informations d’arrivée</h3>',
    parts.join(''),
    '</div>',
  ].join('');
}

function accessInstructionsText_() {
  const info = getAccessInfo_();
  const lines = [];

  if (info.mailboxCode || info.keyBoxCode || info.badgeNotice) {
    lines.push('Informations d’arrivée');
  }
  if (info.mailboxCode) lines.push('Code boîte aux lettres : ' + info.mailboxCode);
  if (info.keyBoxCode) lines.push('Code boîte à clé : ' + info.keyBoxCode);
  if (info.badgeNotice) lines.push(info.badgeNotice);

  return lines.join('\n');
}

function shouldShowAccessInstructions_(status) {
  const normalized = norm_(status);
  return ['prevu', 'validee', 'en cours'].includes(normalized);
}

function getAccessInfo_() {
  const props = PropertiesService.getScriptProperties();
  return {
    mailboxCode: clean_(props.getProperty('MAILBOX_CODE')),
    keyBoxCode: clean_(props.getProperty('KEYBOX_CODE')),
    badgeNotice: clean_(props.getProperty('BADGE_NOTICE')) || 'Le badge permet d’ouvrir la porte. Merci de le remettre aussitôt dans la boîte à clé pour les prochains.',
  };
}

function selectAvailableRoom_(rooms, requestedRoom) {
  if (!rooms || !rooms.length) return null;
  if (!requestedRoom) return rooms[0];
  return rooms.find((room) => room.room === requestedRoom) || null;
}

function buildRegistryRow_(data) {
  const today = today_();
  const activeToday = !isCancelled_(data.status) && today >= data.arrival && today < data.departure;
  const balance = Number(data.amount || 0) - Number(data.paid || 0);

  return [
    data.id,
    data.segment,
    data.room,
    data.occupant,
    data.organisme || '',
    data.phone || '',
    data.arrival,
    data.departure,
    data.unit,
    data.quantity,
    data.unitRate,
    data.amount,
    data.paid,
    balance,
    data.status,
    '',
    '',
    data.channel,
    data.notes || '',
    activeToday ? 'Oui' : 'Non',
    exploitationStatus_(today, data.arrival, data.departure, data.status, '', ''),
    '',
    movementOfDay_(today, data.arrival, data.departure),
    financialStatus_(data.paid, data.amount),
    data.createdAt || new Date(),
  ];
}

function refreshRegistryOperationalFlags_(ss) {
  const sheet = ss.getSheetByName(CONFIG.sheets.registry);
  if (!sheet) return;

  const last = sheet.getLastRow();
  if (last < CONFIG.registryStartRow) return;

  const count = last - CONFIG.registryStartRow + 1;
  const data = sheet.getRange(CONFIG.registryStartRow, 1, count, CONFIG.registryCols).getValues();
  const today = today_();
  const roomItems = {};

  const output = data.map((row, index) => {
    const id = clean_(row[0]);
    const room = clean_(row[2]);
    const arrival = toDate_(row[6]);
    const departure = toDate_(row[7]);
    const status = clean_(row[14]);
    const checkIn = row[15];
    const checkOut = row[16];

    if (!id || !room || !arrival || !departure) return ['', '', '', '', ''];

    const active = !isCancelled_(status) && today >= arrival && today < departure;
    if (!isCancelled_(status)) {
      if (!roomItems[room]) roomItems[room] = [];
      roomItems[room].push({ index, arrival, departure });
    }

    return [
      active ? 'Oui' : 'Non',
      exploitationStatus_(today, arrival, departure, status, checkIn, checkOut),
      '',
      movementOfDay_(today, arrival, departure),
      financialStatus_(row[12], row[11]),
    ];
  });

  Object.keys(roomItems).forEach((room) => {
    const items = roomItems[room].sort((a, b) => a.arrival - b.arrival);
    for (let i = 0; i < items.length; i++) {
      for (let j = i + 1; j < items.length; j++) {
        if (items[j].arrival >= items[i].departure) break;
        if (items[j].arrival < items[i].departure && items[j].departure > items[i].arrival) {
          output[items[i].index][2] = 'Chevauchement';
          output[items[j].index][2] = 'Chevauchement';
        }
      }
    }
  });

  sheet.getRange(CONFIG.registryStartRow, 20, count, 5).setValues(output);
}

function repairRoomStatusFormulas_(ss) {
  const sheet = ss.getSheetByName(CONFIG.sheets.rooms);
  if (!sheet) return;

  const last = sheet.getLastRow();
  if (last < CONFIG.roomStartRow) return;

  for (let row = CONFIG.roomStartRow; row <= last; row++) {
    const roomCell = '$A' + row;
    const manualCell = '$H' + row;
    sheet.getRange(row, 9).setFormula(
      '=IF(' + roomCell + '="","",IF(COUNTIFS(Registre!$C:$C,' + roomCell + ',Registre!$B:$B,"Résident",Registre!$G:$G,"<="&TODAY(),Registre!$H:$H,">"&TODAY(),Registre!$O:$O,"<>Annulé",Registre!$O:$O,"<>No-show")>0,"Résident",IF(COUNTIFS(Registre!$C:$C,' + roomCell + ',Registre!$B:$B,"Court séjour",Registre!$G:$G,"<="&TODAY(),Registre!$H:$H,">"&TODAY(),Registre!$O:$O,"<>Annulé",Registre!$O:$O,"<>No-show")>0,"Court séjour","")))'
    );
    sheet.getRange(row, 10).setFormula(
      '=IF(' + roomCell + '="","",IF(' + manualCell + '="Hors service","Hors service",IF(AND(COUNTIFS(Registre!$C:$C,' + roomCell + ',Registre!$G:$G,TODAY(),Registre!$O:$O,"<>Annulé",Registre!$O:$O,"<>No-show")>0,COUNTIFS(Registre!$C:$C,' + roomCell + ',Registre!$H:$H,TODAY(),Registre!$O:$O,"<>Annulé",Registre!$O:$O,"<>No-show")>0),"Rotation du jour",IF(COUNTIFS(Registre!$C:$C,' + roomCell + ',Registre!$G:$G,TODAY(),Registre!$O:$O,"<>Annulé",Registre!$O:$O,"<>No-show")>0,"Arrivée du jour",IF(COUNTIFS(Registre!$C:$C,' + roomCell + ',Registre!$H:$H,TODAY(),Registre!$O:$O,"<>Annulé",Registre!$O:$O,"<>No-show")>0,"Départ du jour",IF(COUNTIFS(Registre!$C:$C,' + roomCell + ',Registre!$G:$G,"<="&TODAY(),Registre!$H:$H,">"&TODAY(),Registre!$O:$O,"<>Annulé",Registre!$O:$O,"<>No-show")>0,"Occupée","Libre"))))))'
    );
    sheet.getRange(row, 11).setFormula(
      '=IF(' + roomCell + '="","",IF(J' + row + '="Libre","Attribuer",IF(J' + row + '="Arrivée du jour","Préparer accueil",IF(J' + row + '="Départ du jour","Contrôler départ",IF(J' + row + '="Rotation du jour","Nettoyage + accueil",IF(J' + row + '="Hors service","Maintenance","Suivi"))))))'
    );
  }
}

function hasRegistryDuplicate_(sheet, room, occupant, arrival, departure) {
  const last = sheet.getLastRow();
  if (last < CONFIG.registryStartRow) return false;

  const data = sheet.getRange(CONFIG.registryStartRow, 1, last - CONFIG.registryStartRow + 1, CONFIG.registryCols).getValues();
  const key = [norm_(room), norm_(occupant), arrival.getTime(), departure.getTime()].join('|');

  return data.some((row) => {
    const rowArrival = toDate_(row[6]);
    const rowDeparture = toDate_(row[7]);
    if (!rowArrival || !rowDeparture || isCancelled_(row[14])) return false;
    const rowKey = [norm_(row[2]), norm_(row[3]), rowArrival.getTime(), rowDeparture.getTime()].join('|');
    return rowKey === key;
  });
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

  return sheet.getRange(CONFIG.registryStartRow, 1, last - CONFIG.registryStartRow + 1, CONFIG.registryCols)
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

function createRegistryId_() {
  const stamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd-HHmmss');
  const rand = Math.floor(Math.random() * 900 + 100);
  return 'FJT-' + stamp + '-' + rand;
}

function exploitationStatus_(today, arrival, departure, status, checkIn, checkOut) {
  const normalized = norm_(status);
  if (normalized === 'prevu' && arrival < today && !checkIn) return 'Arrivée dépassée';
  if (normalized === 'en cours' && departure < today && !checkOut) return 'Départ dépassé';
  if (normalized === 'parti' && !checkOut) return 'Sortie à régulariser';
  if (normalized === 'en cours') return 'Occupé';
  if (normalized === 'prevu') return 'À venir';
  return status || '';
}

function movementOfDay_(today, arrival, departure) {
  const isArrival = sameDate_(today, arrival);
  const isDeparture = sameDate_(today, departure);
  if (isArrival && isDeparture) return 'Rotation du jour';
  if (isArrival) return 'Arrivée du jour';
  if (isDeparture) return 'Départ du jour';
  return '';
}

function financialStatus_(paid, amount) {
  const paidNumber = Number(paid || 0);
  const amountNumber = Number(amount || 0);
  if (amountNumber <= 0) return '';
  if (paidNumber >= amountNumber) return 'Soldé';
  if (paidNumber > 0) return 'Partiel';
  return 'Impayé';
}

function isValidEmail_(email) {
  return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(clean_(email));
}

function moneyText_(value) {
  return Utilities.formatString('%.2f €', Number(value || 0)).replace('.', ',');
}

function escapeHtml_(value) {
  return clean_(value)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#039;');
}

function isCancelled_(status) {
  return ['annule', 'annulee', 'no-show', 'refuse', 'refusee', 'annulé', 'annulée'].includes(norm_(status));
}

function sameDate_(a, b) {
  return a && b && a.getFullYear() === b.getFullYear() && a.getMonth() === b.getMonth() && a.getDate() === b.getDate();
}

function today_() {
  const now = new Date();
  return new Date(now.getFullYear(), now.getMonth(), now.getDate());
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

function formatDateTimeFr_(date) {
  return Utilities.formatDate(date, Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm');
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
