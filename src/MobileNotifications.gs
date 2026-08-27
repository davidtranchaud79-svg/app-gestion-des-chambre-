// Notifications téléphone via OneSignal.
//
// Propriétés du script à configurer :
// ONESIGNAL_APP_ID = identifiant de l'application OneSignal
// ONESIGNAL_REST_API_KEY = clé REST API OneSignal
//
// Optionnel :
// ONESIGNAL_PLAYER_IDS = ids d'appareils OneSignal séparés par des virgules
// ONESIGNAL_EXTERNAL_USER_IDS = external user ids séparés par des virgules
// ONESIGNAL_INCLUDED_SEGMENTS = segment OneSignal, par défaut "Subscribed Users"

function setupMobileNotifications() {
  const settings = mobileGetOneSignalSettings_();
  if (!settings.ok) {
    throw new Error(settings.error);
  }

  const ss = SpreadsheetApp.openById(CONFIG.spreadsheetId);
  if (typeof ensureAdminSheets_ === 'function') {
    ensureAdminSheets_(ss);
  }

  const sheet = ss.getSheetByName(CONFIG.sheets.adminNotifications);
  const lastRow = sheet ? Math.max(1, sheet.getLastRow()) : 1;
  PropertiesService.getScriptProperties().setProperty('MOBILE_LAST_NOTIFICATION_ROW', String(lastRow));

  installMobileNotificationTrigger();

  const test = mobileSendOneSignal_('Test notification portable', [
    'Notification OneSignal opérationnelle.',
    mobileGetAdminUrl_() ? 'Admin : ' + mobileGetAdminUrl_() : '',
  ]);

  if (!test.ok) {
    throw new Error(test.error || 'Notification portable non envoyée.');
  }

  return 'Notifications téléphone OneSignal activées. Scan automatique toutes les minutes.';
}

function installMobileNotificationTrigger() {
  ScriptApp.getProjectTriggers().forEach((trigger) => {
    if (trigger.getHandlerFunction() === 'mobileScanAndNotify') {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  ScriptApp.newTrigger('mobileScanAndNotify')
    .timeBased()
    .everyMinutes(1)
    .create();

  return 'Déclencheur installé : notifications téléphone toutes les minutes.';
}

function testMobileNotification() {
  const result = mobileSendOneSignal_('Test notification portable', [
    'Notification OneSignal opérationnelle.',
    mobileGetAdminUrl_() ? 'Admin : ' + mobileGetAdminUrl_() : '',
  ]);

  if (!result.ok) {
    throw new Error(result.error || 'Notification portable non envoyée.');
  }

  return 'Notification portable OneSignal envoyée.';
}

function mobileScanAndNotify() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) {
    return 'Scan mobile déjà en cours.';
  }

  try {
    const ss = SpreadsheetApp.openById(CONFIG.spreadsheetId);
    if (typeof ensureAdminSheets_ === 'function') {
      ensureAdminSheets_(ss);
    }

    const sheet = ss.getSheetByName(CONFIG.sheets.adminNotifications);
    if (!sheet) return 'Onglet Admin_Notifications introuvable.';

    const props = PropertiesService.getScriptProperties();
    const lastRow = sheet.getLastRow();
    let lastDone = Number(props.getProperty('MOBILE_LAST_NOTIFICATION_ROW') || 1);

    if (lastRow < 2) {
      props.setProperty('MOBILE_LAST_NOTIFICATION_ROW', '1');
      return 'Aucune notification à traiter.';
    }

    if (lastDone >= lastRow) {
      return 'Aucune nouvelle notification.';
    }

    if (lastDone < 1 || lastDone > lastRow) {
      lastDone = 1;
    }

    const startRow = Math.max(2, lastDone + 1);
    const values = sheet.getRange(startRow, 1, lastRow - startRow + 1, 6).getValues();
    let sent = 0;

    values.forEach((row, index) => {
      const rowNumber = startRow + index;
      const item = {
        date: row[0],
        type: mobileClean_(row[1]),
        level: mobileClean_(row[2]),
        reference: mobileClean_(row[3]),
        message: mobileClean_(row[4]),
        status: mobileClean_(row[5]),
      };

      const notification = mobileBuildNotification_(item);
      if (notification) {
        const result = mobileSendOneSignal_(notification.title, notification.lines);
        if (!result.ok) {
          throw new Error(result.error || 'Notification OneSignal non envoyée.');
        }

        sheet.getRange(rowNumber, 6).setValue('Mobile envoyée ' + mobileFormatDateTime_(new Date()));
        sent++;
      }

      props.setProperty('MOBILE_LAST_NOTIFICATION_ROW', String(rowNumber));
    });

    return sent + ' notification(s) téléphone envoyée(s).';
  } finally {
    lock.releaseLock();
  }
}

function mobileBuildNotification_(item) {
  const type = mobileClean_(item.type).toUpperCase();

  if (type === 'NOUVELLE_RESERVATION') {
    return {
      title: 'Nouvelle réservation à valider',
      lines: mobileNotificationLines_(item),
    };
  }

  return null;
}

function mobileNotificationLines_(item) {
  const lines = [
    item.date ? 'Date : ' + mobileFormatDateTime_(item.date) : '',
    item.level ? 'Niveau : ' + item.level : '',
    item.reference ? 'Référence : ' + item.reference : '',
    item.message ? 'Message : ' + mobileLimitText_(item.message, 900) : '',
  ];

  const adminUrl = mobileGetAdminUrl_();
  if (adminUrl) lines.push('Admin : ' + adminUrl);

  return lines.filter(Boolean);
}

function mobileSendOneSignal_(title, lines) {
  const settings = mobileGetOneSignalSettings_();
  if (!settings.ok) {
    mobileLog_('warning', settings.error);
    return { ok: false, skipped: true, error: settings.error };
  }

  const text = mobileOneSignalText_(title, lines);
  const payload = {
    app_id: settings.appId,
    headings: {
      en: title,
      fr: title,
    },
    contents: {
      en: text,
      fr: text,
    },
    url: mobileGetAdminUrl_() || undefined,
  };

  Object.assign(payload, mobileOneSignalTarget_(settings));

  const response = UrlFetchApp.fetch('https://onesignal.com/api/v1/notifications', {
    method: 'post',
    contentType: 'application/json',
    muteHttpExceptions: true,
    headers: {
      Authorization: 'Basic ' + settings.restApiKey,
    },
    payload: JSON.stringify(payload),
  });

  const code = response.getResponseCode();
  const body = response.getContentText();
  if (code < 200 || code >= 300) {
    const error = 'OneSignal erreur HTTP ' + code + ' : ' + body;
    mobileLog_('warning', error);
    return { ok: false, error };
  }

  mobileLog_('info', 'Notification portable OneSignal envoyée : ' + body);
  return { ok: true, response: body };
}

function mobileOneSignalTarget_(settings) {
  if (settings.playerIds.length) {
    return { include_player_ids: settings.playerIds };
  }

  if (settings.externalUserIds.length) {
    return { include_external_user_ids: settings.externalUserIds };
  }

  return { included_segments: settings.includedSegments };
}

function mobileGetOneSignalSettings_() {
  const props = PropertiesService.getScriptProperties();
  const appId = mobileClean_(props.getProperty('ONESIGNAL_APP_ID'));
  const restApiKey = mobileClean_(props.getProperty('ONESIGNAL_REST_API_KEY'));
  const playerIds = mobileSplitList_(props.getProperty('ONESIGNAL_PLAYER_IDS'));
  const externalUserIds = mobileSplitList_(props.getProperty('ONESIGNAL_EXTERNAL_USER_IDS'));
  const includedSegments = mobileSplitList_(props.getProperty('ONESIGNAL_INCLUDED_SEGMENTS'));

  if (!appId || !restApiKey) {
    return {
      ok: false,
      error: 'ONESIGNAL_APP_ID ou ONESIGNAL_REST_API_KEY absent dans Propriétés du script.',
    };
  }

  return {
    ok: true,
    appId,
    restApiKey,
    playerIds,
    externalUserIds,
    includedSegments: includedSegments.length ? includedSegments : ['Subscribed Users'],
  };
}

function mobileOneSignalText_(title, lines) {
  return ['Gestion chambre - ' + title, '']
    .concat((lines || []).filter(Boolean))
    .join('\n');
}

function mobileGetAdminUrl_() {
  if (typeof getAdminUrl_ === 'function') {
    return getAdminUrl_();
  }

  try {
    const url = ScriptApp.getService().getUrl();
    if (!url) return '';
    return url + (url.indexOf('?') === -1 ? '?view=admin' : '&view=admin');
  } catch (e) {
    return '';
  }
}

function mobileLog_(level, message) {
  try {
    if (typeof logAdminNotification_ === 'function') {
      logAdminNotification_('MOBILE_NOTIFICATION', level, '', message);
    }
  } catch (e) {
    // Évite qu'une erreur de journalisation bloque la réservation.
  }
}

function mobileSplitList_(value) {
  return mobileClean_(value)
    .split(',')
    .map(function(item) { return mobileClean_(item); })
    .filter(Boolean);
}

function mobileLimitText_(value, maxLength) {
  const text = mobileClean_(value);
  if (text.length <= maxLength) return text;
  return text.slice(0, maxLength - 1) + '…';
}

function mobileClean_(value) {
  return String(value == null ? '' : value).trim();
}

function mobileFormatDateTime_(value) {
  if (value instanceof Date && !isNaN(value.getTime())) {
    return Utilities.formatDate(value, Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm');
  }
  return mobileClean_(value);
}
