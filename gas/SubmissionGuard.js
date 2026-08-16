var SUBMISSION_GUARD_PREFIX_ = 'submission.guard.';
var SUBMISSION_GUARD_RETENTION_MS_ = 7 * 24 * 60 * 60 * 1000;
var SUBMISSION_HISTORY_SHEET_ = '_送信履歴';
var SUBMISSION_HISTORY_RETENTION_MS_ = 180 * 24 * 60 * 60 * 1000;

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
  const props = PropertiesService.getScriptProperties();
  const key = SUBMISSION_GUARD_PREFIX_ + submissionId;
  const raw = props.getProperty(key);
  if (!raw) return getDurableSubmissionResult_(submissionId);

  try {
    const record = JSON.parse(raw);
    const cutoff = Date.now() - SUBMISSION_GUARD_RETENTION_MS_;
    if (!record.savedAt || Number(record.savedAt) < cutoff) {
      props.deleteProperty(key);
      return getDurableSubmissionResult_(submissionId);
    }
    return record && record.result ? record.result : null;
  } catch (err) {
    props.deleteProperty(key);
    Logger.log('送信ID記録の読み取り失敗: ' + submissionId + ' / ' + err.message);
    return getDurableSubmissionResult_(submissionId);
  }
}

/**
 * @param {string} submissionId
 * @param {Object} result
 */
function saveSubmissionResult_(submissionId, result) {
  let propertySaved = false;
  let historySaved = false;
  try {
    cleanupSavedSubmissionResults_();
    PropertiesService.getScriptProperties().setProperty(
      SUBMISSION_GUARD_PREFIX_ + submissionId,
      JSON.stringify({ savedAt: Date.now(), result: result })
    );
    propertySaved = true;
  } catch (err) {
    Logger.log('送信IDの一時記録に失敗: ' + err.message);
  }

  try {
    saveDurableSubmissionResult_(submissionId, result);
    historySaved = true;
  } catch (err) {
    Logger.log('送信IDの永続記録に失敗: ' + err.message);
  }

  if (!propertySaved && !historySaved) {
    reportSystemError_('送信重複防止記録', new Error('送信IDを記録できませんでした: ' + submissionId));
  }
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

function getSubmissionHistorySheet_() {
  const ss = getMasterSpreadsheet();
  let sheet = ss.getSheetByName(SUBMISSION_HISTORY_SHEET_);
  if (!sheet) {
    sheet = ss.insertSheet(SUBMISSION_HISTORY_SHEET_);
    sheet.getRange(1, 1, 1, 3).setValues([['submissionId', 'savedAt', 'resultJson']]);
    sheet.setFrozenRows(1);
    sheet.hideSheet();
  }
  return sheet;
}

function getDurableSubmissionResult_(submissionId) {
  try {
    const sheet = getSubmissionHistorySheet_();
    const match = sheet.createTextFinder(submissionId).matchEntireCell(true).findNext();
    if (!match || match.getRow() < 2) return null;
    const values = sheet.getRange(match.getRow(), 1, 1, 3).getValues()[0];
    const savedAt = values[1] instanceof Date ? values[1].getTime() : new Date(values[1]).getTime();
    if (!savedAt || savedAt < Date.now() - SUBMISSION_HISTORY_RETENTION_MS_) return null;
    return JSON.parse(String(values[2] || 'null'));
  } catch (err) {
    Logger.log('永続送信履歴の読み取り失敗: ' + err.message);
    return null;
  }
}

function saveDurableSubmissionResult_(submissionId, result) {
  const sheet = getSubmissionHistorySheet_();
  const match = sheet.createTextFinder(submissionId).matchEntireCell(true).findNext();
  const values = [[submissionId, new Date(), JSON.stringify(result)]];
  if (match && match.getRow() >= 2) {
    sheet.getRange(match.getRow(), 1, 1, 3).setValues(values);
  } else {
    sheet.appendRow(values[0]);
  }
}

function cleanupDurableSubmissionHistory_() {
  try {
    const sheet = getSubmissionHistorySheet_();
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return;
    const dates = sheet.getRange(2, 2, lastRow - 1, 1).getValues();
    const cutoff = Date.now() - SUBMISSION_HISTORY_RETENTION_MS_;
    for (let i = dates.length - 1; i >= 0; i--) {
      const value = dates[i][0];
      const time = value instanceof Date ? value.getTime() : new Date(value).getTime();
      if (!time || time < cutoff) sheet.deleteRow(i + 2);
    }
  } catch (err) {
    Logger.log('永続送信履歴の整理に失敗: ' + err.message);
  }
}
