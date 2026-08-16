var SUBMISSION_GUARD_PREFIX_ = 'submission.guard.';
var SUBMISSION_GUARD_RETENTION_MS_ = 7 * 24 * 60 * 60 * 1000;

/**
 * Script-wide lock used while choosing a monthly file and appending rows.
 * @param {Function} callback
 * @returns {*}
 */
function withSubmissionLock_(callback) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) {
    throw new Error('現在ほかの送信を処理しています。少し待ってから再度お試しください');
  }

  try {
    return callback();
  } finally {
    lock.releaseLock();
  }
}

/**
 * @param {*} value
 * @returns {string}
 */
function normalizeSubmissionId_(value) {
  const id = String(value || '').trim();
  if (!id) return '';
  if (id.length > 128 || !/^[A-Za-z0-9._:-]+$/.test(id)) {
    throw new Error('submissionId is invalid');
  }
  return id;
}

/**
 * @param {string} submissionId
 * @returns {Object|null}
 */
function getSavedSubmissionResult_(submissionId) {
  const raw = PropertiesService.getScriptProperties()
    .getProperty(SUBMISSION_GUARD_PREFIX_ + submissionId);
  if (!raw) return null;

  try {
    const record = JSON.parse(raw);
    return record && record.result ? record.result : null;
  } catch (err) {
    Logger.log('送信ID記録の読み取り失敗: ' + submissionId + ' / ' + err.message);
    return null;
  }
}

/**
 * @param {string} submissionId
 * @param {Object} result
 */
function saveSubmissionResult_(submissionId, result) {
  cleanupSavedSubmissionResults_();
  PropertiesService.getScriptProperties().setProperty(
    SUBMISSION_GUARD_PREFIX_ + submissionId,
    JSON.stringify({ savedAt: Date.now(), result: result })
  );
}

/** Remove old idempotency records so script properties do not grow forever. */
function cleanupSavedSubmissionResults_() {
  const props = PropertiesService.getScriptProperties();
  const all = props.getProperties();
  const cutoff = Date.now() - SUBMISSION_GUARD_RETENTION_MS_;

  Object.keys(all).forEach(function(key) {
    if (key.indexOf(SUBMISSION_GUARD_PREFIX_) !== 0) return;
    try {
      const record = JSON.parse(all[key]);
      if (!record.savedAt || Number(record.savedAt) < cutoff) {
        props.deleteProperty(key);
      }
    } catch (err) {
      props.deleteProperty(key);
    }
  });
}
