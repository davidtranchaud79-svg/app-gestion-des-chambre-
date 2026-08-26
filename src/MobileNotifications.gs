// Notifications téléphone via OneSignal Web Push.
// À configurer dans Propriétés du script :
// ONESIGNAL_REST_API_KEY = clé REST API OneSignal (NE PAS la mettre dans le code)

const ONESIGNAL_APP_ID = 'c1c76342-2445-4ad8-9335-9563c23e9854';
const ONESIGNAL_ADMIN_URL = 'https://davidtranchaud79-svg.github.io/app-gestion-des-chambre-/admin.html';

function setupMobileNotifications() {
  const settings = mobileGetOneSignalSettings_();
  if (!settings.ok) throw new Error(settings.error);

  const ss = SpreadsheetApp.openById(CONFIG.spreadsheetId);
  if (typeof ensureAdminSheets_ === 'function') ensureAdminSheets_(ss);

  const sheet = ss.getSheetByName(CONFIG.sheets.adminNotifications);
  const lastRow = sheet ? Math.max(1, sheet.getLastRow()) : 1;
  PropertiesService.getScriptProperties().setProperty('MOBILE_LAST_NOTIFICATION_ROW', String(lastRow));

  installMobileNotificationTrigger();

  const test = mobileSendOneSignal_('Test notification portable', 'Notification portable opérationnelle.', ONESIGNAL_ADMIN_URL);
  if (!test.ok) throw new Error(test.error || 'Notification portable non envoyée.');

  return 'Notifications téléphone OneSignal activées. Scan automatique toutes les minutes.';
}

function installMobileNotificationTrigger() {
  ScriptApp.getProjectTriggers().forEach((trigger) => {
    if (trigger.getHandlerFunction() === 'mobileScanAndNotify') ScriptApp.deleteTrigger(trigger);
  });

  ScriptApp.newTrigger('mobileScanAndNotify')
    .timeBased()
    .everyMinutes(1)
    .create();

  return 'Déclencheur installé : notifications téléphone toutes les minutes.';
}

function testMobileNotification() {
  const result = mobileSendOneSignal_('Test notification portable', 'Notification portable opérationnelle.', ONESIGNAL_ADMIN_URL);
  if (!result.ok) throw new Error(result.error || 'Notification portable non envoyée.');
  return 'Notification portable OneSignal envoyée.';
}

function mobileScanAndNotify() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) return 'Scan mobile déjà en cours.';

  try {
    const ss = SpreadsheetApp.openById(CONFIG.spreadsheetId);
    if (typeof ensureAdminSheets_ === 'function') ensureAdminSheets_(ss);

    const sheet = ss.getSheetByName(CONFIG.sheets.adminNotifications);
    if (!sheet) return 'Onglet Admin_Notifications introuvable.';

    const props = PropertiesService.getScriptProperties();
    const lastRow = sheet.getLastRow();
    let lastDone = Number(props.getProperty('MOBILE_LAST_NOTIFICATION_ROW') || 1);

    if (lastRow < 2) {
      props.setProperty('MOBILE_LAST_NOTIFICATION_ROW', '1');
      return 'Aucune notification à traiter.';
    }

    if (lastDone >= lastRow) return 'Aucune nouvelle notification.';
    if (lastDone < 1 || lastDone > lastRow) lastDone = 1;

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
        const result = mobileSendOneSignal_(notification.title, notification.body, ONESIGNAL_ADMIN_URL);
        if (!result.ok) throw new Error(result.error || 'Notification OneSignal non envoyée.');
        sheet.getRange(rowNumber, 6).setValue('Push envoyé ' + mobileFormatDateTime_(new Date()));
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
      body: mobileNotificationBody_(item),
    };
  }

  if (type === 'ANOMALIES') {
    return {
      title: 'Anomalie réservation détectée',
      body: mobileNotificationBody_(item),
    };
  }

  if (type === 'ANOMALIE_VALIDATION') {
    return {
      title: 'Blocage validation réservation',
      body: mobileNotificationBody_(item),
    };
  }

  return null;
}

function mobileNotificationBody_(item) {
  return [
    item.reference ? 'Réf. ' + item.reference : '',
    item.message ? mobileLimitText_(item.message, 180) : '',
  ].filter(Boolean).join(' — ');
}

function mobileSendOneSignal_(title, body, url) {
  const settings = mobileGetOneSignalSettings_();
  if (!settings.ok) {
    mobileLog_('warning', settings.error);
    return { ok: false, skipped: true, error: settings.error };
  }

  const payload = {
    app_id: ONESIGNAL_APP_ID,
    included_segments: ['Subscribed Users'],
    headings: { en: title, fr: title },
    contents: { en: body || title, fr: body || title },
    url: url || ONESIGNAL_ADMIN_URL,
  };

  const response = UrlFetchApp.fetch('https://api.onesignal.com/notifications?c=push', {
    method: 'post',
    contentType: 'application/json',
    headers: {
      Authorization: 'Key ' + settings.apiKey,
    },
    muteHttpExceptions: true,
    payload: JSON.stringify(payload),
  });

  const code = response.getResponseCode();
  const text = response.getContentText();
  if (code < 200 || code >= 300) {
    const error = 'OneSignal erreur HTTP ' + code + ' : ' + text;
    mobileLog_('warning', error);
    return { ok: false, error };
  }

  mobileLog_('info', 'Notification portable OneSignal envoyée.');
  return { ok: true, response: text };
}

function mobileGetOneSignalSettings_() {
  const apiKey = mobileClean_(PropertiesService.getScriptProperties().getProperty('ONESIGNAL_REST_API_KEY'));
  if (!apiKey) {
    return { ok: false, error: 'ONESIGNAL_REST_API_KEY absent dans Propriétés du script.' };
  }
  return { ok: true, apiKey };
}

function mobileLog_(level, message) {
  try {
    if (typeof logAdminNotification_ === 'function') {
      logAdminNotification_('MOBILE_NOTIFICATION', level, '', message);
    }
  } catch (e) {}
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
