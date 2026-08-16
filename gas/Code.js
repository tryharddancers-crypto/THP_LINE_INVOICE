/**
 * LIFFフォームからのPOSTリクエストを受け取るエントリーポイント
 */
function doPost(e) {
  try {
    const result = handleSubmission(e);
    runOpportunisticMaintenance_(true);
    return ContentService
      .createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    Logger.log('doPost error: ' + err.message);
    if (!err.isClientError) {
      reportSystemError_('フォーム送信処理', err);
    }
    return ContentService
      .createTextOutput(JSON.stringify({
        ok: false,
        error: err.isClientError
          ? err.message
          : 'システムでエラーが発生しました。時間を置いて再度お試しください。'
      }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * LIFFからのGETリクエスト（マスタデータ取得・フォームHTML配信）
 */
function doGet(e) {
  const action = e && e.parameter && e.parameter.action;

  if (action === 'getMaster') {
    const data = getMasterData();
    runOpportunisticMaintenance_(false);
    return ContentService
      .createTextOutput(JSON.stringify({ ok: true, data }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  // LIFFフォームHTMLを配信
  const props = PropertiesService.getScriptProperties();
  const template = HtmlService.createTemplateFromFile('liff/index');
  template.liffId = props.getProperty('LIFF_ID') || '';
  template.gasUrl = ScriptApp.getService().getUrl();
  return template.evaluate().setTitle('案件入力フォーム');
}

/**
 * フォーム送信の実処理
 * @param {object} e GASイベントオブジェクト
 * @returns {{ ok: boolean, count: number, date: string }}
 */
function handleSubmission(e) {
  let body;
  try {
    body = JSON.parse(e && e.postData && e.postData.contents || '');
  } catch (err) {
    throw createClientError_('送信データを読み取れませんでした。フォームを開き直してください。');
  }

  const userId = verifyLineIdToken_(body.idToken);
  const enrichedRows = validateSubmissionRows_(body.rows);

  // The current frontend supplies this ID so uncertain retries cannot duplicate rows.
  const submissionId = normalizeSubmissionId_(body.submissionId) || Utilities.getUuid();

  // Monthly-file selection, row selection and write must be one locked operation.
  const writeResult = withSubmissionLock_(function() {
    const savedResult = getSavedSubmissionResult_(submissionId);
    if (savedResult) {
      return {
        duplicate: true,
        response: Object.assign({}, savedResult, { duplicate: true })
      };
    }

    const date = parseDate(enrichedRows[0].date);
    const ss = getOrCreateMonthlySpreadsheet(date);
    const appendResult = appendRowsToInputSheet(ss, enrichedRows);
    SpreadsheetApp.flush();

    const response = {
      ok: true,
      count: enrichedRows.length,
      date: enrichedRows[0].date,
      submissionId: submissionId,
      duplicate: false
    };
    saveSubmissionResult_(submissionId, response);

    return {
      duplicate: false,
      response: response,
      spreadsheetId: ss.getId(),
      rowNumbers: appendResult.rowNumbers
    };
  });

  if (writeResult.duplicate) {
    Logger.log('重複送信をスキップ: ' + submissionId);
    return writeResult.response;
  }

  const ss = SpreadsheetApp.openById(writeResult.spreadsheetId);
  Utilities.sleep(1000);

  // 入力されたC〜N列の内容を、該当担当者のメールアドレスへ通知する
  try {
    sendInputRowsNotification_(ss, writeResult.rowNumbers, submissionId);
  } catch (err) {
    Logger.log('入力内容メール通知中にエラー: ' + err.message);
  }

  // LINE通知（送信内容の詳細を含む）
  const liffId = PropertiesService.getScriptProperties().getProperty('LIFF_ID') || '2009725727-3jvV9g52';
  const liffUrl = 'https://liff.line.me/' + liffId;
  const message = buildSubmissionMessage(enrichedRows, liffUrl);
  if (userId) {
    try {
      sendLineMessage(userId, message);
    } catch (err) {
      Logger.log('LINE通知中にエラー: ' + err.message);
    }
  }

  // PDF送信は手動メニューからのみ行う。フォーム送信では添付しない。
  return writeResult.response;
}

/**
 * 送信内容を人物・案件ごとに整形したLINEメッセージを生成する
 * @param {object[]} rows
 * @param {string} liffUrl
 * @returns {string}
 */
function buildSubmissionMessage(rows, liffUrl) {
  // 人物・日付ごとにグループ化
  const groups = {};
  rows.forEach(function(row) {
    const key = row.date + '__' + row.name;
    if (!groups[key]) groups[key] = { date: row.date, name: row.name, jobs: [] };
    groups[key].jobs.push(row);
  });

  const lines = [];
  lines.push('✅ ' + rows.length + '件を追加しました\n');

  Object.values(groups).forEach(function(g) {
    lines.push('【' + g.date + '】' + g.name);
    g.jobs.forEach(function(j) {
      lines.push('・' + j.jobName + ' ×' + j.qty);
    });
    lines.push('');
  });

  lines.push('次回の入力もよろしくお願いします。');
  lines.push('▼ 入力フォーム(LIFF)のURL');
  lines.push(liffUrl);

  return lines.join('\n').trim();
}
