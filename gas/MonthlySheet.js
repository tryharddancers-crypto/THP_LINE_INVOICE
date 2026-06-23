/**
 * 当月の月次スプレッドシートを取得する（なければ作成する）
 * @param {Date} date
 * @returns {GoogleAppsScript.Spreadsheet.Spreadsheet}
 */
function getOrCreateMonthlySpreadsheet(date) {
  const props = PropertiesService.getScriptProperties();
  const rootFolderId = props.getProperty('MONTHLY_FOLDER_ID');
  const templateId = props.getProperty('TEMPLATE_SPREADSHEET_ID');
  if (!rootFolderId) throw new Error('MONTHLY_FOLDER_ID is not set');
  if (!templateId) throw new Error('TEMPLATE_SPREADSHEET_ID is not set');

  const rootFolder = DriveApp.getFolderById(rootFolderId);
  const yearName = date.getFullYear() + '年';

  // 年フォルダを取得、なければ作成
  let yearFolder;
  const yearFolders = rootFolder.getFoldersByName(yearName);
  if (yearFolders.hasNext()) {
    yearFolder = yearFolders.next();
  } else {
    yearFolder = rootFolder.createFolder(yearName);
  }

  const fileName = getMonthlyFileName(date);

  // 年フォルダ内で既存ファイルを検索
  const files = yearFolder.getFilesByName(fileName);
  if (files.hasNext()) {
    const existingSs = SpreadsheetApp.open(files.next());
    syncMonthlyJobMaster_(existingSs);
    return existingSs;
  }

  // テンプレートをコピーして新規作成（保存先は年フォルダ）
  const templateFile = DriveApp.getFileById(templateId);
  const newFile = templateFile.makeCopy(fileName, yearFolder);
  const newSs = SpreadsheetApp.open(newFile);

  // 外注連絡票シートのタイトルと日付を設定
  const invoiceSheet = newSs.getSheetByName('外注連絡票');
  if (invoiceSheet) {
    invoiceSheet.getRange('A1').setValue(`≪${toWareki(date)}分 外注費支払表≫`);
    fillDancerMasterToSheet(invoiceSheet);
  }

  syncMonthlyJobMaster_(newSs);

  return newSs;
}

/**
 * 入力表シートに複数行を追記する
 * B/E/F/M/N 列は月次シート内の数式で自動計算する
 * 列: B=team, C=日程, D=曜日, E=現場, F=項目, G=案件名, J=詳細, K=名前, L=数量, M=単価, N=合計
 * @returns {{startRow: number, endRow: number, rowNumbers: number[]}} 追記した行情報
 */
function appendRowsToInputSheet(ss, rows) {
  const sheet = ss.getSheetByName('2.入力表');
  if (!sheet) {
    throw new Error('「2.入力表」シートが見つかりません');
  }

  var DAY = ['日','月','火','水','木','金','土'];
  
  // 本当の最終行（C列が空欄の最初の行）を探す
  var cValues = sheet.getRange('C1:C1000').getValues();
  var startRow = 8; // 8行目からデータ入力開始
  for (var j = 7; j < cValues.length; j++) { // j=7 は8行目(C8)を指す
    if (cValues[j][0] === '' || cValues[j][0] === null) {
      startRow = j + 1;
      break;
    }
  }

  rows.forEach(function(row, i) {
    var r       = startRow + i;
    var d       = new Date(row.date.replace(/\//g, '-'));
    var weekday = DAY[d.getDay()];
    var jobName = String(row.jobName || '').trim();
    var name    = String(row.name || '').trim();

    sheet.getRange(r,  3).clearDataValidations();
    sheet.getRange(r,  3).setValue(row.date);            // C: 日程

    sheet.getRange(r,  4).clearDataValidations();
    sheet.getRange(r,  4).setValue(weekday);             // D: 曜日

    sheet.getRange(r,  7).clearDataValidations();
    sheet.getRange(r,  7).setValue(jobName);             // G: 案件名

    sheet.getRange(r, 10).clearDataValidations();
    sheet.getRange(r, 10).setValue(row.detail   || '');  // J: 詳細

    sheet.getRange(r, 11).clearDataValidations();
    sheet.getRange(r, 11).setValue(name);                // K: 名前

    sheet.getRange(r, 12).clearDataValidations();
    sheet.getRange(r, 12).setValue(row.qty);             // L: 数量
  });

  // 書き込みを確実にシートに反映させる
  SpreadsheetApp.flush();

  return {
    startRow: startRow,
    endRow: startRow + rows.length - 1,
    rowNumbers: rows.map(function(_, i) { return startRow + i; })
  };
}

/**
 * 月次シート内の案件マスタを、最新の本体マスタに同期する
 * 2.入力表の数式は月次内「案件マスタ」A3:D を参照しているため、データ開始行を3行目に揃える
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} monthlySs
 */
function syncMonthlyJobMaster_(monthlySs) {
  const masterSs = getMasterSpreadsheet();
  const srcSheet = masterSs.getSheetByName('案件マスタ');
  if (!srcSheet) {
    throw new Error('マスタスプレッドシートに「案件マスタ」シートがありません');
  }

  let dstSheet = monthlySs.getSheetByName('案件マスタ');
  if (!dstSheet) {
    dstSheet = monthlySs.insertSheet('案件マスタ');
  }

  const srcLastRow = srcSheet.getLastRow();
  const jobs = srcLastRow >= 2
    ? srcSheet.getRange(2, 1, srcLastRow - 1, 4).getValues()
      .filter(function(row) { return String(row[0] || '').trim() !== ''; })
    : [];

  dstSheet.getRange(1, 1, 1, 4).setValues([['案件名', '現場コード', '単価', '項目区分']]);

  const clearRows = Math.max(dstSheet.getLastRow() - 1, jobs.length + 1, 1);
  dstSheet.getRange(2, 1, clearRows, 4).clearContent();

  if (jobs.length > 0) {
    dstSheet.getRange(3, 1, jobs.length, 4).setValues(jobs);
  }
}

/**
 * マスタの外注連絡票から、芸名ごとのメールアドレスを取得する
 * @returns {Object<string, string>} {芸名: メールアドレス}
 */
function getDancerEmailMap_() {
  return getOutsourceContactEmailMap_();
}

/**
 * 追記した行のC〜N列を読み取る
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss 対象の月次スプレッドシート
 * @param {number[]} rowNumbers 追記した行番号
 * @returns {Object[]} C〜N列の値を持つ行データ
 */
function readInsertedInputRows_(ss, rowNumbers) {
  const sheet = ss.getSheetByName('2.入力表');
  if (!sheet) {
    throw new Error('「2.入力表」シートが見つかりません');
  }
  if (!rowNumbers || rowNumbers.length === 0) {
    return [];
  }

  return rowNumbers.map(function(rowNumber) {
    const values = sheet.getRange(rowNumber, 3, 1, 12).getValues()[0]; // C:N
    return {
      rowNumber: rowNumber,
      date: values[0],
      weekday: values[1],
      venue: values[2],
      category: values[3],
      jobName: values[4],
      h: values[5],
      i: values[6],
      detail: values[7],
      name: String(values[8] || '').trim(),
      qty: values[9],
      unitPrice: values[10],
      total: values[11]
    };
  }).filter(function(row) {
    return row.name !== '';
  });
}

/**
 * 追記した入力行の内容を、担当者ごとにメール通知する
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss 対象の月次スプレッドシート
 * @param {number[]} rowNumbers 追記した行番号
 */
function sendInputRowsNotification_(ss, rowNumbers) {
  const insertedRows = readInsertedInputRows_(ss, rowNumbers);
  if (insertedRows.length === 0) {
    Logger.log('入力内容メール通知: 通知対象行がありません');
    return;
  }

  const emailMap = getDancerEmailMap_();
  const grouped = {};
  insertedRows.forEach(function(row) {
    if (!grouped[row.name]) grouped[row.name] = [];
    grouped[row.name].push(row);
  });

  Object.keys(grouped).forEach(function(name) {
    const email = emailMap[name];
    if (!email) {
      Logger.log('入力内容メール通知スキップ: ' + name + ' のメールアドレスが外注連絡票U列にありません');
      return;
    }

    const rows = grouped[name];
    const subject = '【確認】案件入力内容を受け付けました';
    const body = buildInputRowsNotificationBody_(ss, name, rows);
    try {
      GmailApp.sendEmail(email, subject, body, { name: '管理事務局' });
      Logger.log('入力内容メール通知送信: ' + name + ' <' + email + '> ' + rows.length + '件');
    } catch (e) {
      Logger.log('入力内容メール通知失敗: ' + name + ' / ' + e.message);
    }
  });
}

/**
 * 入力内容通知メールの本文を作成する
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss 対象の月次スプレッドシート
 * @param {string} name 担当者名
 * @param {Object[]} rows 担当者ごとの入力行
 * @returns {string} メール本文
 */
function buildInputRowsNotificationBody_(ss, name, rows) {
  const lines = [];
  lines.push(name + ' 様');
  lines.push('');
  lines.push('お疲れ様です。');
  lines.push('フォームより送信された案件入力内容を受け付けました。');
  lines.push('');

  rows.forEach(function(row, index) {
    lines.push('【入力内容 ' + (index + 1) + '】');
    lines.push('日程: ' + formatMailValue_(row.date));
    lines.push('曜日: ' + formatMailValue_(row.weekday));
    lines.push('現場: ' + formatMailValue_(row.venue));
    lines.push('項目: ' + formatMailValue_(row.category));
    lines.push('案件名: ' + formatMailValue_(row.jobName));
    lines.push('詳細: ' + formatMailValue_(row.detail));
    lines.push('名前: ' + formatMailValue_(row.name));
    lines.push('数量: ' + formatMailValue_(row.qty));
    lines.push('単価: ' + formatMailValue_(row.unitPrice));
    lines.push('合計: ' + formatMailValue_(row.total));
    lines.push('');
  });

  lines.push('内容をご確認ください。');
  return lines.join('\n');
}

/**
 * メール本文用にセル値を表示文字列へ変換する
 * @param {*} value セル値
 * @returns {string}
 */
function formatMailValue_(value) {
  if (value === '' || value === null || typeof value === 'undefined') return '-';
  if (value instanceof Date) {
    return Utilities.formatDate(value, 'Asia/Tokyo', 'yyyy/MM/dd');
  }
  return String(value);
}

/**
 * 毎月1日0時に実行されるトリガー関数
 * 当月の月次スプレッドシートを作成する
 */
function createMonthlySheetTrigger() {
  const now = new Date();
  Logger.log(`月次シート作成開始: ${getMonthlyFileName(now)}`);
  const ss = getOrCreateMonthlySpreadsheet(now);
  Logger.log(`月次シート作成完了: ${ss.getUrl()}`);
}
