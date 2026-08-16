var NOTIFICATION_QUEUE_SHEET_ = '_通知送信キュー';
var NOTIFICATION_MAX_ATTEMPTS_ = 5;
var NOTIFICATION_QUEUE_HEADERS_ = [
  'queueKey', 'status', 'attempts', 'nextRetryAt', 'createdAt', 'updatedAt',
  'spreadsheetId', 'personName', 'rowNumbers', 'lastError', 'sentAt'
];

function getNotificationQueueSheet_() {
  const ss = getMasterSpreadsheet();
  let sheet = ss.getSheetByName(NOTIFICATION_QUEUE_SHEET_);
  if (!sheet) {
    sheet = ss.insertSheet(NOTIFICATION_QUEUE_SHEET_);
    sheet.getRange(1, 1, 1, NOTIFICATION_QUEUE_HEADERS_.length)
      .setValues([NOTIFICATION_QUEUE_HEADERS_]);
    sheet.setFrozenRows(1);
    sheet.hideSheet();
  }
  return sheet;
}

function withNotificationQueueLock_(callback) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) {
    throw new Error('通知再送キューを処理中です。しばらくしてから再実行してください。');
  }
  try {
    return callback();
  } finally {
    lock.releaseLock();
  }
}

function enqueueNotificationRetry_(submissionId, spreadsheetId, personName, rowNumbers, error) {
  return withNotificationQueueLock_(function() {
    const sheet = getNotificationQueueSheet_();
    const key = String(submissionId || Utilities.getUuid()) + '|' + String(personName);
    const match = sheet.createTextFinder(key).matchEntireCell(true).findNext();
    if (match) {
      const status = String(sheet.getRange(match.getRow(), 2).getValue() || '');
      if (status !== 'SENT') {
        sheet.getRange(match.getRow(), 2, 1, 9).setValues([[
          'PENDING',
          Number(sheet.getRange(match.getRow(), 3).getValue()) || 0,
          new Date(Date.now() + 10 * 60 * 1000),
          sheet.getRange(match.getRow(), 5).getValue() || new Date(),
          new Date(),
          spreadsheetId,
          personName,
          JSON.stringify(rowNumbers || []),
          String(error && error.message || error || '')
        ]]);
      }
      return key;
    }

    const now = new Date();
    sheet.appendRow([
      key,
      'PENDING',
      0,
      new Date(now.getTime() + 10 * 60 * 1000),
      now,
      now,
      spreadsheetId,
      personName,
      JSON.stringify(rowNumbers || []),
      String(error && error.message || error || ''),
      ''
    ]);
    return key;
  });
}

function notificationRetryDelayMs_(attempts) {
  return Math.min(24 * 60 * 60 * 1000, Math.pow(2, Math.max(0, attempts - 1)) * 10 * 60 * 1000);
}

/** Time-trigger entry point. */
function retryPendingNotifications() {
  return retryPendingNotifications_(10);
}

function retryPendingNotifications_(limit) {
  return withNotificationQueueLock_(function() {
    const sheet = getNotificationQueueSheet_();
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return { processed: 0, sent: 0, failed: 0 };

    const values = sheet.getRange(2, 1, lastRow - 1, NOTIFICATION_QUEUE_HEADERS_.length).getValues();
    const now = new Date();
    let processed = 0;
    let sent = 0;
    let failed = 0;

    for (let i = 0; i < values.length && processed < Number(limit || 10); i++) {
      const row = values[i];
      const status = String(row[1] || '');
      const nextRetryAt = row[3] instanceof Date ? row[3] : new Date(row[3] || 0);
      if (status !== 'PENDING' || nextRetryAt.getTime() > now.getTime()) continue;

      processed += 1;
      const sheetRow = i + 2;
      const attempts = (Number(row[2]) || 0) + 1;
      try {
        const monthlySs = SpreadsheetApp.openById(String(row[6]));
        const rowNumbers = JSON.parse(String(row[8] || '[]'));
        const allRows = readInsertedInputRows_(monthlySs, rowNumbers);
        const personRows = allRows.filter(function(inputRow) {
          return inputRow.name === String(row[7]);
        });
        if (personRows.length === 0) {
          throw new Error('再送対象の入力行が見つかりません');
        }

        sendNotificationForPerson_(monthlySs, String(row[7]), personRows);
        sheet.getRange(sheetRow, 2, 1, 10).setValues([[
          'SENT', attempts, '', row[4], new Date(), row[6], row[7], row[8], '', new Date()
        ]]);
        sent += 1;
      } catch (err) {
        const exhausted = attempts >= NOTIFICATION_MAX_ATTEMPTS_;
        sheet.getRange(sheetRow, 2, 1, 9).setValues([[
          exhausted ? 'FAILED' : 'PENDING',
          attempts,
          exhausted ? '' : new Date(now.getTime() + notificationRetryDelayMs_(attempts)),
          row[4],
          new Date(),
          row[6],
          row[7],
          row[8],
          String(err.message || err)
        ]]);
        failed += 1;
        if (exhausted) {
          reportSystemError_('メール通知の再送上限超過: ' + row[7], err);
        }
      }
    }

    return { processed: processed, sent: sent, failed: failed };
  });
}
