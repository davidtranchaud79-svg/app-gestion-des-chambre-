// Actualisation automatique du Sheet.
// À lancer une fois : setupAutoRefresh

function setupAutoRefresh() {
  installAutoRefreshTrigger();
  autoRefreshSheet();
  return 'Actualisation automatique installée : toutes les 5 minutes.';
}

function installAutoRefreshTrigger() {
  ScriptApp.getProjectTriggers().forEach((trigger) => {
    if (trigger.getHandlerFunction() === 'autoRefreshSheet') {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  ScriptApp.newTrigger('autoRefreshSheet')
    .timeBased()
    .everyMinutes(5)
    .create();

  return 'Déclencheur installé : autoRefreshSheet toutes les 5 minutes.';
}

function removeAutoRefreshTrigger() {
  let removed = 0;
  ScriptApp.getProjectTriggers().forEach((trigger) => {
    if (trigger.getHandlerFunction() === 'autoRefreshSheet') {
      ScriptApp.deleteTrigger(trigger);
      removed++;
    }
  });
  return removed + ' déclencheur(s) autoRefreshSheet supprimé(s).';
}

function autoRefreshSheet() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) {
    return 'Actualisation déjà en cours.';
  }

  try {
    const ss = SpreadsheetApp.openById(CONFIG.spreadsheetId);

    if (typeof planningSetTodayReference_ === 'function') {
      planningSetTodayReference_(ss);
    }

    if (typeof refreshRegistryOperationalFlags_ === 'function') {
      refreshRegistryOperationalFlags_(ss);
    }

    if (typeof repairRoomStatusFormulas_ === 'function') {
      repairRoomStatusFormulas_(ss);
    }

    if (typeof refreshVisualPlanning_ === 'function') {
      refreshVisualPlanning_(ss);
    }

    SpreadsheetApp.flush();
    return 'Sheet actualisé : ' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm:ss');
  } catch (err) {
    if (typeof logAdminNotification_ === 'function') {
      logAdminNotification_('AUTO_REFRESH', 'warning', '', err && err.message ? err.message : String(err));
    }
    throw err;
  } finally {
    lock.releaseLock();
  }
}
