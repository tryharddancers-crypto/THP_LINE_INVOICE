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
    return SpreadsheetApp.open(files.next());
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

  return newSs;
}

/**
 * 入力表シートに複数行を追記する
 * B/D/E/F/M/N 列をすべてGASで計算して書き込む（シート数式に依存しない）
 * 列: B=team, C=日程, D=曜日, E=現場, F=項目, G=案件名, J=詳細, K=名前, L=数量, M=単価, N=合計
 */
function appendRowsToInputSheet(ss, rows) {
  const sheet = ss.getSheetByName('2.入力表');

  // ── 金額シート: N=案件名, O=現場, P=単価, Q=項目 ──
  const kinSheet = ss.getSheetByName('金額');
  const kinData  = kinSheet.getRange('N3:Q62').getValues();
  const kinMap   = {};
  kinData.forEach(function(r) {
    if (r[0]) kinMap[String(r[0])] = { venue: String(r[1]||''), unitPrice: Number(r[2])||0, category: String(r[3]||'') };
  });

  // ── 判定表シート: B=名前, G=team(6列目) ──
  const teamSheet = ss.getSheetByName('判定表');
  const teamData  = teamSheet.getRange('B3:G108').getValues();
  const teamMap   = {};
  teamData.forEach(function(r) {
    if (r[0]) teamMap[String(r[0])] = String(r[5]||'');
  });

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
    var kin     = kinMap[row.jobName] || {};
    var unit    = kin.unitPrice || row.unitPrice || 0;
    var total   = row.qty * unit;
    var team    = teamMap[row.name] || '';
    var d       = new Date(row.date.replace(/\//g, '-'));
    var weekday = DAY[d.getDay()];

    // シートの数式（ARRAYFORMULA等）を上書きしないよう、黄色い範囲（入力専用列）にのみ書き込む
    // sheet.getRange(r,  2).setValue(team);               // B: team(判定) -> 自動計算

    sheet.getRange(r,  3).clearDataValidations();
    sheet.getRange(r,  3).setValue(row.date);            // C: 日程

    sheet.getRange(r,  4).clearDataValidations();
    sheet.getRange(r,  4).setValue(weekday);             // D: 曜日

    // sheet.getRange(r,  5).setValue(kin.venue    || ''); // E: 現場 -> 自動計算
    // sheet.getRange(r,  6).setValue(kin.category || ''); // F: 項目 -> 自動計算

    sheet.getRange(r,  7).clearDataValidations();
    sheet.getRange(r,  7).setValue(row.jobName);         // G: 案件名

    sheet.getRange(r, 10).clearDataValidations();
    sheet.getRange(r, 10).setValue(row.detail   || '');  // J: 詳細

    sheet.getRange(r, 11).clearDataValidations();
    sheet.getRange(r, 11).setValue(row.name);            // K: 名前

    sheet.getRange(r, 12).clearDataValidations();
    sheet.getRange(r, 12).setValue(row.qty);             // L: 数量

    // sheet.getRange(r, 13).setValue(unit);                // M: 単価 -> 自動計算
    // sheet.getRange(r, 14).setValue(total);               // N: 合計金額 -> 自動計算
  });

  // 書き込みを確実にシートに反映させる
  SpreadsheetApp.flush();
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
