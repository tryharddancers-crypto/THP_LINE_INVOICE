/**
 * LIFFフォームからのPOSTリクエストを受け取るエントリーポイント
 */
function doPost(e) {
  try {
    const result = handleSubmission(e);
    return ContentService
      .createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    Logger.log('doPost error: ' + err.message);
    return ContentService
      .createTextOutput(JSON.stringify({ ok: false, error: err.message }))
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
  const body = JSON.parse(e.postData.contents);
  const { userId, rows } = body;

  if (!rows || rows.length === 0) {
    throw new Error('rows is empty');
  }

  // 単価はフロントエンドから送られてきたものをそのまま使用する
  const enrichedRows = rows.map(row => ({
    ...row,
    unitPrice: row.unitPrice || 0
  }));

  // 当月スプレッドシートに追記
  const date = parseDate(rows[0].date);
  const ss = getOrCreateMonthlySpreadsheet(date);
  appendRowsToInputSheet(ss, enrichedRows);

  // LINE通知（送信内容の詳細を含む）
  const message = buildSubmissionMessage(enrichedRows, ss.getUrl());
  if (userId) {
    sendLineMessage(userId, message);
  }

  return { ok: true, count: rows.length, date: enrichedRows[0].date };
}

/**
 * 送信内容を人物・案件ごとに整形したLINEメッセージを生成する
 * @param {object[]} rows
 * @param {string} sheetUrl
 * @returns {string}
 */
function buildSubmissionMessage(rows, sheetUrl) {
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

  return lines.join('\n').trim();
}
