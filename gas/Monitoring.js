var ADMIN_ALERT_THROTTLE_SECONDS_ = 60 * 60;
var DEFAULT_ADMIN_ALERT_EMAIL_ = 'tryharddancers@gmail.com';
var RETRY_FALLBACK_INTERVAL_MS_ = 10 * 60 * 1000;
var HEALTH_FALLBACK_INTERVAL_MS_ = 24 * 60 * 60 * 1000;

function getAdminAlertEmail_() {
  const configured = String(
    PropertiesService.getScriptProperties().getProperty('ADMIN_ALERT_EMAIL') || ''
  ).trim();
  if (configured) return configured;
  try {
    const effectiveUser = String(Session.getEffectiveUser().getEmail() || '').trim();
    return effectiveUser || DEFAULT_ADMIN_ALERT_EMAIL_;
  } catch (err) {
    return DEFAULT_ADMIN_ALERT_EMAIL_;
  }
}

function ensureAdminAlertEmail_() {
  const props = PropertiesService.getScriptProperties();
  const configured = String(props.getProperty('ADMIN_ALERT_EMAIL') || '').trim();
  if (configured) return configured;

  const email = getAdminAlertEmail_();
  if (email) props.setProperty('ADMIN_ALERT_EMAIL', email);
  return email;
}

/** Send a throttled operational alert without masking the original failure. */
function reportSystemError_(context, err) {
  try {
    const email = getAdminAlertEmail_();
    if (!email) {
      Logger.log('管理者通知先が未設定: ' + context + ' / ' + err.message);
      return;
    }

    const signature = String(context) + '|' + String(err && err.message || err);
    const digest = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, signature)
      .map(function(byte) { return ('0' + ((byte + 256) % 256).toString(16)).slice(-2); })
      .join('');
    const cache = CacheService.getScriptCache();
    const cacheKey = 'admin-alert-' + digest.slice(0, 32);
    if (cache.get(cacheKey)) return;
    cache.put(cacheKey, '1', ADMIN_ALERT_THROTTLE_SECONDS_);

    const stack = err && err.stack ? '\n\n' + err.stack : '';
    GmailApp.sendEmail(
      email,
      '【要確認】請求書自動作成システムでエラーが発生しました',
      '発生箇所: ' + context + '\n内容: ' + String(err && err.message || err) + stack,
      { name: '請求書自動作成システム' }
    );
  } catch (alertErr) {
    Logger.log('管理者エラー通知に失敗: ' + alertErr.message);
  }
}

function dailySystemHealthCheck() {
  const result = runSystemHealthCheck_();
  Logger.log('システム点検: ' + JSON.stringify(result));
  if (!result.ok) {
    reportSystemError_('日次システム点検', new Error(result.issues.join('\n')));
  }
  cleanupDurableSubmissionHistory_();
  return result;
}

function setupReliabilityTriggers() {
  ensureAdminAlertEmail_();
  const status = reconcileReliabilityTriggers_();
  Logger.log('信頼性向上トリガーの設定が完了しました: ' + JSON.stringify(status));
  return status;
}

function reconcileReliabilityTriggers_() {
  const required = {
    retryPendingNotifications: function() {
      return ScriptApp.newTrigger('retryPendingNotifications').timeBased().everyMinutes(10).create();
    },
    dailySystemHealthCheck: function() {
      return ScriptApp.newTrigger('dailySystemHealthCheck').timeBased().everyDays(1).atHour(6).create();
    }
  };
  const found = {};

  ScriptApp.getProjectTriggers().forEach(function(trigger) {
    const handler = trigger.getHandlerFunction();
    if (!required[handler]) return;
    if (!found[handler]) {
      found[handler] = trigger;
      return;
    }
    ScriptApp.deleteTrigger(trigger);
  });

  if (!found.retryPendingNotifications) {
    ScriptApp.newTrigger('retryPendingNotifications').timeBased().everyMinutes(10).create();
  }
  if (!found.dailySystemHealthCheck) {
    ScriptApp.newTrigger('dailySystemHealthCheck').timeBased().everyDays(1).atHour(6).create();
  }

  const activeHandlers = ScriptApp.getProjectTriggers().map(function(trigger) {
    return trigger.getHandlerFunction();
  });
  if (activeHandlers.indexOf('retryPendingNotifications') === -1
      || activeHandlers.indexOf('dailySystemHealthCheck') === -1) {
    throw new Error('必要な自動トリガーを作成できませんでした');
  }

  return {
    retryPendingNotifications: true,
    dailySystemHealthCheck: true,
    adminAlertEmail: ensureAdminAlertEmail_()
  };
}

function claimMaintenanceWindow_(propertyKey, intervalMs) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(1000)) return false;
  try {
    const props = PropertiesService.getScriptProperties();
    const now = Date.now();
    const lastRun = Number(props.getProperty(propertyKey)) || 0;
    if (now - lastRun < intervalMs) return false;
    props.setProperty(propertyKey, String(now));
    return true;
  } finally {
    lock.releaseLock();
  }
}

/**
 * Trigger-free fallback. Normal form traffic keeps retries and health checks
 * alive even when no installable trigger is configured. Trigger management is
 * intentionally excluded because web-app requests cannot authorize ScriptApp.
 */
function runOpportunisticMaintenance_(includeHealthCheck) {
  try {
    ensureAdminAlertEmail_();

    if (claimMaintenanceWindow_('OPS_LAST_RETRY_FALLBACK_AT', RETRY_FALLBACK_INTERVAL_MS_)) {
      try {
        const retryResult = retryPendingNotifications_(3);
        Logger.log('メール再送の予備処理: ' + JSON.stringify(retryResult));
      } catch (err) {
        reportSystemError_('メール再送の予備処理', err);
      }
    }

    if (includeHealthCheck === true
        && claimMaintenanceWindow_('OPS_LAST_HEALTH_FALLBACK_AT', HEALTH_FALLBACK_INTERVAL_MS_)) {
      try {
        dailySystemHealthCheck();
      } catch (err) {
        reportSystemError_('日次点検の予備処理', err);
      }
    }

  } catch (err) {
    Logger.log('予備メンテナンス処理に失敗: ' + err.message);
  }
}
