// Notifications téléphone via Telegram.
// À configurer dans Propriétés du script :
// TELEGRAM_BOT_TOKEN = token donné par BotFather
// TELEGRAM_CHAT_ID = identifiant du chat Telegram à notifier

function setupMobileNotifications() {
  const settings = mobileGetTelegramSettings_();
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

  const test = mobileSendTelegram_('Test notification portable', [
    'Notification portable opérationnelle.',
    mobileGetAdminUrl_() ? 'Admin : ' + mobileGetAdminUrl_() : '',
  ]);

  if (!test.ok) {
    throw new Error(test.error || 'Notification portable non envoyée.');
  }

  return 'Notifications téléphone activées. Scan automatique toutes les minutes.';
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
  const result = mobileSendTelegram_('Test notification portable', [
    'Notification portable opérationnelle.',
    mobileGetAdminUrl_() ? 'Admin : ' + mobileGetAdminUrl_() : '',
  ]);

  if (!result.ok) {
    throw new Error(result.error || 'Notification portable non envoyée.');
  }

  return 'Notification portable envoyée.';
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
        const result = mobileSendTelegram_(notification.title, notification.lines);
        if (!result.ok) {
          throw new Error(result.error || 'Notification Telegram non envoyée.');
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

  if (type === 'ANOMALIES') {
    return {
      title: 'Anomalie réservation détectée',
      lines: mobileNotificationLines_(item),
    };
  }

  if (type === 'ANOMALIE_VALIDATION') {
    return {
      title: 'Blocage validation réservation',
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

function mobileSendTelegram_(title, lines) {
  const settings = mobileGetTelegramSettings_();
  if (!settings.ok) {
    mobileLog_('warning', settings.error);
    return { ok: false, skipped: true, error: settings.error };
  }

  const response = UrlFetchApp.fetch('https://api.telegram.org/bot' + settings.token + '/sendMessage', {
    method: 'post',
    contentType: 'application/json',
    muteHttpExceptions: true,
    payload: JSON.stringify({
      chat_id: settings.chatId,
      text: mobileTelegramText_(title, lines),
      disable_web_page_preview: true,
    }),
  });

  const code = response.getResponseCode();
  if (code < 200 || code >= 300) {
    const error = 'Telegram erreur HTTP ' + code + ' : ' + response.getContentText();
    mobileLog_('warning', error);
    return { ok: false, error };
  }

  mobileLog_('info', 'Notification portable envoyée.');
  return { ok: true };
}

function mobileGetTelegramSettings_() {
  const props = PropertiesService.getScriptProperties();
  const token = mobileClean_(props.getProperty('TELEGRAM_BOT_TOKEN'));
  const chatId = mobileClean_(props.getProperty('TELEGRAM_CHAT_ID'));

  if (!token || !chatId) {
    return { ok: false, error: 'TELEGRAM_BOT_TOKEN ou TELEGRAM_CHAT_ID absent dans Propriétés du script.' };
  }

  return { ok: true, token, chatId };
}

function mobileTelegramText_(title, lines) {
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
