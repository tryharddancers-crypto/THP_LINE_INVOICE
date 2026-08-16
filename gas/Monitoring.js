var ADMIN_ALERT_THROTTLE_SECONDS_ = 60 * 60;

function getAdminAlertEmail_() {
  const configured = String(
    PropertiesService.getScriptProperties().getProperty('ADMIN_ALERT_EMAIL') || ''
  ).trim();
  if (configured) return configured;
  try {
    return String(Session.getEffectiveUser().getEmail() || '').trim();
  } catch (err) {
    return '';
  }
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
  const handlers = ScriptApp.getProjectTriggers().map(function(trigger) {
    return trigger.getHandlerFunction();
  });
  if (handlers.indexOf('retryPendingNotifications') === -1) {
    ScriptApp.newTrigger('retryPendingNotifications').timeBased().everyMinutes(10).create();
  }
  if (handlers.indexOf('dailySystemHealthCheck') === -1) {
    ScriptApp.newTrigger('dailySystemHealthCheck').timeBased().everyDays(1).atHour(6).create();
  }
  Logger.log('信頼性向上トリガーの設定が完了しました');
}
