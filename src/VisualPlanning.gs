// Couleurs et vues exploitation pour Chambres, Registre et Calendrier.
// Ce module est compatible avec le pack réservation + admin.

const PLANNING_THEME = {
  navy: '#17324D',
  blue: '#2F75B5',
  paleBlue: '#D9EAF7',
  green: '#70AD47',
  paleGreen: '#E2F0D9',
  orange: '#ED7D31',
  paleOrange: '#FCE4D6',
  red: '#C00000',
  paleRed: '#F4CCCC',
  purple: '#7030A0',
  palePurple: '#EAD1DC',
  yellow: '#FFD966',
  paleYellow: '#FFF2CC',
  grey: '#7F8C8D',
  paleGrey: '#E7E6E6',
  white: '#FFFFFF',
  text: '#1F2937',
};

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Gestion chambres')
    .addItem('Ajout manuel résident', 'openResidentSidebar')
    .addItem('Intégrer Saisie résidents', 'integrateResidentInputRows')
    .addItem('Installer synchro automatique', 'setupSheetSyncTrigger')
    .addSeparator()
    .addItem('Actualiser couleurs + planning', 'setupVisualPlanning')
    .addItem('Reconstruire le plan chambres', 'buildRoomPlanVisual')
    .addItem('Reconstruire le calendrier', 'buildCalendarVisual')
    .addSeparator()
    .addItem('Calendrier : mois précédent', 'planningCalendarPreviousMonth')
    .addItem('Calendrier : mois actuel', 'planningCalendarCurrentMonth')
    .addItem('Calendrier : mois suivant', 'planningCalendarNextMonth')
    .addToUi();
}

function setupVisualPlanning() {
  const ss = planningOpenSpreadsheet_();
  if (typeof refreshRegistryOperationalFlags_ === 'function') refreshRegistryOperationalFlags_(ss);
  if (typeof repairRoomStatusFormulas_ === 'function') repairRoomStatusFormulas_(ss);
  refreshVisualPlanning_(ss);
  ss.toast('Couleurs, plan chambres et calendrier actualisés.', 'Gestion chambres', 5);
  return 'Couleurs, plan chambres et calendrier actualisés.';
}

function refreshVisualPlanning_(ss) {
  ss = ss || planningOpenSpreadsheet_();
  planningApplyRegistryColors_(ss);
  planningApplyRoomsColors_(ss);
  buildRoomPlanVisual_(ss);
  buildCalendarVisual_(ss);
}

function buildRoomPlanVisual() {
  const ss = planningOpenSpreadsheet_();
  if (typeof refreshRegistryOperationalFlags_ === 'function') refreshRegistryOperationalFlags_(ss);
  buildRoomPlanVisual_(ss);
  ss.setActiveSheet(planningEnsureSheet_(ss, 'Plan chambres'));
  return 'Plan chambres reconstruit.';
}

function buildCalendarVisual(requestedDate) {
  const ss = planningOpenSpreadsheet_();
  if (typeof refreshRegistryOperationalFlags_ === 'function') refreshRegistryOperationalFlags_(ss);
  buildCalendarVisual_(ss, requestedDate);
  ss.setActiveSheet(planningEnsureSheet_(ss, 'Calendrier'));
  return 'Calendrier reconstruit.';
}

function planningCalendarPreviousMonth() {
  planningShiftCalendarMonth_(-1);
}

function planningCalendarCurrentMonth() {
  buildCalendarVisual(new Date());
}

function planningCalendarNextMonth() {
  planningShiftCalendarMonth_(1);
}

function planningShiftCalendarMonth_(offset) {
  const ss = planningOpenSpreadsheet_();
  const sheet = planningEnsureSheet_(ss, 'Calendrier');
  const current = toDate_(sheet.getRange('B1').getValue()) || today_();
  const target = new Date(current.getFullYear(), current.getMonth() + Number(offset || 0), 1);
  buildCalendarVisual_(ss, target);
  ss.setActiveSheet(sheet);
}

function buildRoomPlanVisual_(ss) {
  const sheet = planningEnsureSheet_(ss, 'Plan chambres');
  const rooms = planningGetRooms_(ss);
  const records = planningGetRegistryRecords_(ss);
  const today = today_();

  sheet.clear();
  sheet.setHiddenGridlines(true);
  sheet.setFrozenRows(4);
  sheet.getRange('A1:L1').merge().setValue('PLAN VISUEL DES CHAMBRES');
  sheet.getRange('A2:L2').merge().setValue('Situation au ' + formatFr_(today));
  sheet.getRange('A3:B3').merge().setValue('Libre');
  sheet.getRange('C3:D3').merge().setValue('Occupée');
  sheet.getRange('E3:F3').merge().setValue('Arrivée');
  sheet.getRange('G3:H3').merge().setValue('Départ');
  sheet.getRange('I3:J3').merge().setValue('Rotation');
  sheet.getRange('K3:L3').merge().setValue('Hors service');

  rooms.forEach(function(room, index) {
    const roomRecords = records.filter(function(record) { return record.room === room.room; });
    const active = roomRecords.find(function(record) {
      return today >= record.arrival && today < record.departure && !planningIsInactiveStatus_(record.status);
    });
    const arrival = roomRecords.find(function(record) {
      return sameDate_(record.arrival, today) && !planningIsInactiveStatus_(record.status);
    });
    const departure = roomRecords.find(function(record) {
      return sameDate_(record.departure, today) && !planningIsInactiveStatus_(record.status);
    });
    const pending = roomRecords.find(function(record) {
      return planningIsPendingStatus_(record.status) && record.departure >= today;
    });

    let state = 'LIBRE';
    let color = PLANNING_THEME.green;
    let occupant = 'Disponible';
    let note = 'Chambre disponible';

    if (norm_(room.manualStatus) === 'hors service') {
      state = 'HORS SERVICE';
      color = PLANNING_THEME.grey;
      occupant = 'Maintenance';
      note = 'Chambre fermée manuellement';
    } else if (arrival && departure) {
      state = 'ROTATION';
      color = PLANNING_THEME.purple;
      occupant = arrival.occupant || departure.occupant || '';
      note = planningRoomNote_(arrival || departure);
    } else if (arrival) {
      state = 'ARRIVÉE';
      color = PLANNING_THEME.orange;
      occupant = arrival.occupant;
      note = planningRoomNote_(arrival);
    } else if (departure) {
      state = 'DÉPART';
      color = PLANNING_THEME.blue;
      occupant = departure.occupant;
      note = planningRoomNote_(departure);
    } else if (active) {
      state = planningIsPendingStatus_(active.status) ? 'À VALIDER' : 'OCCUPÉE';
      color = planningIsPendingStatus_(active.status) ? PLANNING_THEME.orange : PLANNING_THEME.red;
      occupant = active.occupant;
      note = planningRoomNote_(active);
    } else if (pending) {
      state = 'À VALIDER';
      color = PLANNING_THEME.orange;
      occupant = pending.occupant;
      note = planningRoomNote_(pending);
    }

    const row = 5 + Math.floor(index / 4) * 4;
    const col = 1 + (index % 4) * 3;
    const range = sheet.getRange(row, col, 3, 3);
    range.merge();
    range.setValue('CHAMBRE ' + room.room + '\n' + state + '\n' + occupant)
      .setBackground(color)
      .setFontColor(PLANNING_THEME.white)
      .setFontWeight('bold')
      .setFontSize(12)
      .setWrap(true)
      .setHorizontalAlignment('center')
      .setVerticalAlignment('middle')
      .setBorder(true, true, true, true, false, false, PLANNING_THEME.white, SpreadsheetApp.BorderStyle.SOLID_THICK)
      .setNote(note);
  });

  sheet.setColumnWidths(1, 12, 95);
  for (let row = 5; row <= Math.max(sheet.getLastRow(), 5); row += 4) {
    sheet.setRowHeights(row, 3, 30);
  }

  sheet.getRange('A1:L1')
    .setBackground(PLANNING_THEME.navy)
    .setFontColor(PLANNING_THEME.white)
    .setFontWeight('bold')
    .setFontSize(20)
    .setHorizontalAlignment('center');
  sheet.getRange('A2:L2')
    .setBackground(PLANNING_THEME.paleBlue)
    .setFontColor(PLANNING_THEME.navy)
    .setFontWeight('bold')
    .setHorizontalAlignment('center');
  sheet.getRange('A3:B3').setBackground(PLANNING_THEME.paleGreen);
  sheet.getRange('C3:D3').setBackground(PLANNING_THEME.paleRed);
  sheet.getRange('E3:F3').setBackground(PLANNING_THEME.paleOrange);
  sheet.getRange('G3:H3').setBackground(PLANNING_THEME.paleBlue);
  sheet.getRange('I3:J3').setBackground(PLANNING_THEME.palePurple);
  sheet.getRange('K3:L3').setBackground(PLANNING_THEME.paleGrey);
  sheet.getRange('A3:L3').setFontWeight('bold').setHorizontalAlignment('center');
}

function buildCalendarVisual_(ss, requestedDate) {
  const sheet = planningEnsureSheet_(ss, 'Calendrier');
  const rooms = planningGetRooms_(ss);
  const records = planningGetRegistryRecords_(ss);
  const savedDate = toDate_(sheet.getRange('B1').getValue());
  const baseDate = toDate_(requestedDate) || savedDate || today_();
  const year = baseDate.getFullYear();
  const month = baseDate.getMonth();
  const monthStart = new Date(year, month, 1);
  const monthEnd = new Date(year, month + 1, 0);
  const daysInMonth = monthEnd.getDate();

  sheet.clear();
  sheet.setHiddenGridlines(true);
  sheet.setFrozenRows(4);
  sheet.setFrozenColumns(2);

  sheet.getRange('A1').setValue('Calendrier d’occupation');
  sheet.getRange('B1')
    .setValue(monthStart)
    .setNumberFormat('mmmm yyyy')
    .setNote('Mets ici une date du mois voulu, puis lance buildCalendarVisual.');
  sheet.getRange('C1').setValue('Mois précédent');
  sheet.getRange('D1').setValue('Aujourd’hui');
  sheet.getRange('E1').setValue('Mois suivant');
  sheet.getRange('A2').setValue('Légende');
  sheet.getRange('B2:H2').setValues([[
    'A arrivée',
    'D départ',
    'A/D rotation',
    'C court séjour',
    'R résident',
    'V à valider',
    '! chevauchement',
  ]]);
  sheet.getRange(3, 1, 1, 2).setValues([['Chambre', 'Type']]);

  const headers = Array.from({ length: daysInMonth }, function(_, index) {
    return new Date(year, month, index + 1);
  });
  sheet.getRange(3, 3, 1, daysInMonth).setValues([headers]).setNumberFormat('dd');

  if (rooms.length) {
    sheet.getRange(4, 1, rooms.length, 2).setValues(rooms.map(function(room) {
      return [room.room, room.type];
    }));
  }

  const roomIndex = {};
  rooms.forEach(function(room, index) {
    roomIndex[room.room] = index;
  });

  const values = Array.from({ length: rooms.length }, function() { return Array(daysInMonth).fill(''); });
  const backgrounds = Array.from({ length: rooms.length }, function() { return Array(daysInMonth).fill(PLANNING_THEME.white); });
  const notes = Array.from({ length: rooms.length }, function() { return Array(daysInMonth).fill(''); });

  records.forEach(function(record) {
    if (!record.room || !record.arrival || !record.departure) return;
    if (planningIsInactiveStatus_(record.status)) return;

    const resolvedRoom = planningResolveRoom_(record.room, roomIndex);
    if (!resolvedRoom) return;

    const rowIndex = roomIndex[resolvedRoom];
    const from = record.arrival < monthStart ? monthStart : record.arrival;
    const to = record.departure > monthEnd ? monthEnd : record.departure;
    if (to < from) return;

    for (let day = from.getDate(); day <= to.getDate(); day++) {
      const date = new Date(year, month, day);
      const colIndex = day - 1;
      let label = planningIsPendingStatus_(record.status) ? 'V' : (record.segment === 'Résident' ? 'R' : 'C');

      if (sameDate_(record.arrival, date) && sameDate_(record.departure, date)) label = 'A/D';
      else if (sameDate_(record.arrival, date)) label = 'A';
      else if (sameDate_(record.departure, date)) label = 'D';

      if (values[rowIndex][colIndex] && values[rowIndex][colIndex] !== label) {
        values[rowIndex][colIndex] = '!';
        backgrounds[rowIndex][colIndex] = PLANNING_THEME.paleRed;
      } else {
        values[rowIndex][colIndex] = label;
        backgrounds[rowIndex][colIndex] = planningCalendarColor_(label, record);
      }

      notes[rowIndex][colIndex] = [
        record.occupant,
        record.segment,
        record.status,
        record.id,
        formatFr_(record.arrival) + ' -> ' + formatFr_(record.departure),
      ].filter(Boolean).join(' - ');
    }
  });

  if (rooms.length) {
    const grid = sheet.getRange(4, 3, rooms.length, daysInMonth);
    grid.setValues(values);
    grid.setBackgrounds(backgrounds);
    grid.setNotes(notes);
    grid.setHorizontalAlignment('center')
      .setVerticalAlignment('middle')
      .setFontWeight('bold');
  }

  const lastCol = daysInMonth + 2;
  const lastRow = Math.max(4, rooms.length + 3);
  sheet.getRange(1, 1, 1, lastCol)
    .setBackground(PLANNING_THEME.navy)
    .setFontColor(PLANNING_THEME.white)
    .setFontWeight('bold');
  sheet.getRange(2, 1, 1, lastCol)
    .setBackground(PLANNING_THEME.paleBlue)
    .setFontColor(PLANNING_THEME.navy);
  sheet.getRange(3, 1, 1, lastCol)
    .setBackground(PLANNING_THEME.paleBlue)
    .setFontWeight('bold')
    .setHorizontalAlignment('center');
  sheet.getRange(3, 1, lastRow - 2, lastCol)
    .setBorder(true, true, true, true, true, true, '#D1D5DB', SpreadsheetApp.BorderStyle.SOLID);
  sheet.autoResizeColumns(1, 2);
  for (let col = 3; col <= lastCol; col++) sheet.setColumnWidth(col, 36);
}

function planningCalendarColor_(label, record) {
  if (label === 'A') return PLANNING_THEME.paleGreen;
  if (label === 'D') return PLANNING_THEME.paleYellow;
  if (label === 'A/D') return PLANNING_THEME.palePurple;
  if (label === 'V') return PLANNING_THEME.paleOrange;
  if (record.segment === 'Résident') return PLANNING_THEME.paleBlue;
  return '#D9D2E9';
}

function planningApplyRegistryColors_(ss) {
  const sheet = ss.getSheetByName(CONFIG.sheets.registry);
  if (!sheet) return;

  const headerRow = CONFIG.registryStartRow - 1;
  const startRow = CONFIG.registryStartRow;
  const maxRows = Math.max(sheet.getMaxRows() - startRow + 1, 1);
  const full = sheet.getRange(startRow, 1, maxRows, CONFIG.registryCols);

  sheet.getRange(headerRow, 1, 1, Math.max(CONFIG.registryCols, sheet.getLastColumn()))
    .setBackground(PLANNING_THEME.navy)
    .setFontColor(PLANNING_THEME.white)
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setWrap(true);
  sheet.setFrozenRows(headerRow);

  const rules = [
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=OR($O6="Annulé",$O6="Annulée",$O6="Refusé",$O6="Refusée",$O6="No-show")')
      .setBackground(PLANNING_THEME.paleGrey)
      .setFontColor(PLANNING_THEME.grey)
      .setRanges([full])
      .build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=$V6="Chevauchement"')
      .setBackground(PLANNING_THEME.paleRed)
      .setRanges([full])
      .build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=$X6="Impayé"')
      .setBackground(PLANNING_THEME.paleYellow)
      .setRanges([full])
      .build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=$O6="À valider"')
      .setBackground(PLANNING_THEME.paleOrange)
      .setRanges([full])
      .build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=OR($W6="Arrivée du jour",$W6="Rotation du jour")')
      .setBackground(PLANNING_THEME.paleOrange)
      .setRanges([full])
      .build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=$W6="Départ du jour"')
      .setBackground(PLANNING_THEME.paleBlue)
      .setRanges([full])
      .build(),
  ];

  sheet.setConditionalFormatRules(rules);
}

function planningApplyRoomsColors_(ss) {
  const sheet = ss.getSheetByName(CONFIG.sheets.rooms);
  if (!sheet) return;

  const headerRow = CONFIG.roomStartRow - 1;
  const startRow = CONFIG.roomStartRow;
  const lastCol = Math.max(sheet.getLastColumn(), 11);
  const maxRows = Math.max(sheet.getMaxRows() - startRow + 1, 1);
  const full = sheet.getRange(startRow, 1, maxRows, lastCol);

  sheet.getRange(headerRow, 1, 1, lastCol)
    .setBackground(PLANNING_THEME.navy)
    .setFontColor(PLANNING_THEME.white)
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setWrap(true);
  sheet.setFrozenRows(headerRow);

  const rules = [
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=$H6="Hors service"')
      .setBackground(PLANNING_THEME.paleGrey)
      .setRanges([full])
      .build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=$J6="Libre"')
      .setBackground(PLANNING_THEME.paleGreen)
      .setRanges([full])
      .build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=$J6="Occupée"')
      .setBackground(PLANNING_THEME.paleRed)
      .setRanges([full])
      .build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=$J6="Arrivée du jour"')
      .setBackground(PLANNING_THEME.paleOrange)
      .setRanges([full])
      .build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=$J6="Départ du jour"')
      .setBackground(PLANNING_THEME.paleBlue)
      .setRanges([full])
      .build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=$J6="Rotation du jour"')
      .setBackground(PLANNING_THEME.palePurple)
      .setRanges([full])
      .build(),
  ];

  sheet.setConditionalFormatRules(rules);
  sheet.autoResizeColumns(1, Math.min(lastCol, 11));
}

function planningGetRooms_(ss) {
  const sheet = ss.getSheetByName(CONFIG.sheets.rooms);
  if (!sheet) return [];

  const lastRow = sheet.getLastRow();
  if (lastRow < CONFIG.roomStartRow) return [];

  return sheet.getRange(CONFIG.roomStartRow, 1, lastRow - CONFIG.roomStartRow + 1, Math.max(11, sheet.getLastColumn()))
    .getValues()
    .map(function(row) {
      return {
        room: clean_(row[0]),
        floor: clean_(row[1]),
        type: clean_(row[2]),
        capacity: Number(row[3] || 1),
        manualStatus: clean_(row[7]) || 'Ouverte',
        segment: clean_(row[8]),
        status: clean_(row[9]),
        action: clean_(row[10]),
      };
    })
    .filter(function(room) { return room.room; });
}

function planningGetRegistryRecords_(ss) {
  const sheet = ss.getSheetByName(CONFIG.sheets.registry);
  if (!sheet) return [];

  const lastRow = sheet.getLastRow();
  if (lastRow < CONFIG.registryStartRow) return [];

  return sheet.getRange(CONFIG.registryStartRow, 1, lastRow - CONFIG.registryStartRow + 1, CONFIG.registryCols)
    .getValues()
    .map(function(row) {
      const arrival = toDate_(row[6]);
      const departure = toDate_(row[7]);
      return {
        id: clean_(row[0]),
        segment: planningNormalizeSegment_(row[1]),
        room: clean_(row[2]),
        occupant: clean_(row[3]),
        phone: clean_(row[5]),
        arrival: arrival,
        departure: departure,
        amount: Number(row[11] || 0),
        paid: Number(row[12] || 0),
        balance: Number(row[13] || 0),
        status: clean_(row[14]),
        overlap: clean_(row[21]),
        movement: clean_(row[22]),
        financial: clean_(row[23]),
      };
    })
    .filter(function(record) {
      return record.id && record.room && record.arrival && record.departure;
    });
}

function planningNormalizeSegment_(value) {
  const normalized = norm_(value);
  if (normalized.indexOf('resident') !== -1) return 'Résident';
  return 'Court séjour';
}

function planningResolveRoom_(room, roomIndex) {
  if (roomIndex[room] !== undefined) return room;

  const simplified = norm_(room).replace(/^chambre\s*/, '').replace(/\s+/g, '');
  const keys = Object.keys(roomIndex);
  for (let i = 0; i < keys.length; i++) {
    const key = keys[i];
    const keySimplified = norm_(key).replace(/^chambre\s*/, '').replace(/\s+/g, '');
    if (keySimplified === simplified) return key;
  }

  return '';
}

function planningIsInactiveStatus_(status) {
  const normalized = norm_(status);
  return [
    'annule',
    'annulee',
    'refuse',
    'refusee',
    'no-show',
    'noshow',
  ].indexOf(normalized) !== -1;
}

function planningIsPendingStatus_(status) {
  const normalized = norm_(status);
  return normalized === 'a valider' || normalized === 'en attente' || normalized === 'attente validation';
}

function planningRoomNote_(record) {
  if (!record) return '';
  return [
    record.occupant,
    record.segment,
    record.status,
    record.id,
    formatFr_(record.arrival) + ' -> ' + formatFr_(record.departure),
    record.balance ? 'Solde : ' + moneyText_(record.balance) : '',
  ].filter(Boolean).join('\n');
}

function planningEnsureSheet_(ss, name) {
  return ss.getSheetByName(name) || ss.insertSheet(name);
}

function planningOpenSpreadsheet_() {
  try {
    return SpreadsheetApp.openById(CONFIG.spreadsheetId);
  } catch (e) {
    return SpreadsheetApp.getActiveSpreadsheet();
  }
}
