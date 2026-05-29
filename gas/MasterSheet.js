/**
 * スクリプトプロパティからマスタスプレッドシートを取得する
 * @returns {GoogleAppsScript.Spreadsheet.Spreadsheet}
 */
function getMasterSpreadsheet() {
  const id = PropertiesService.getScriptProperties().getProperty('MASTER_SPREADSHEET_ID');
  if (!id) throw new Error('MASTER_SPREADSHEET_ID is not set in script properties');
  return SpreadsheetApp.openById(id);
}

/**
 * 案件名から単価を検索する
 * @param {string} jobName
 * @returns {number} 単価（見つからない場合は0）
 */
function lookupUnitPrice(jobName) {
  const ss = getMasterSpreadsheet();
  const sheet = ss.getSheetByName('案件マスタ');
  const data = sheet.getDataRange().getValues();
  // 1行目はヘッダー、A列=案件名、C列=単価
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === jobName) {
      return Number(data[i][2]) || 0;
    }
  }
  return 0;
}

/**
 * LIFFフォームのドロップダウン用にマスタデータを返す
 * jobList: [{ name, billing, unitPrice }]
 * @returns {{ jobList: object[], dancerNames: string[] }}
 */
function getMasterData() {
  const ss = getMasterSpreadsheet();

  // 案件マスタ: A=案件名, B=請求元, C=単価
  const jobSheet = ss.getSheetByName('案件マスタ');
  const jobData = jobSheet.getDataRange().getValues();
  const jobList = jobData.slice(1)
    .filter(row => row[0] !== '')
    .map(row => ({
      name:      String(row[0]),
      billing:   String(row[1] || 'その他'),
      unitPrice: Number(row[2]) || 0
    }));

  // ダンサーマスタ: A=芸名
  const dancerSheet = ss.getSheetByName('ダンサーマスタ');
  const dancerData = dancerSheet.getDataRange().getValues();
  let dancerNames = dancerData.slice(1).map(row => String(row[0])).filter(v => v !== '');

  // 判定表シートが存在する場合は、両方に合致する人物のみを抽出する
  const judgeSheet = ss.getSheetByName('判定表');
  if (judgeSheet) {
    const judgeData = judgeSheet.getDataRange().getValues();
    const judgeNames = new Set(judgeData.map(row => String(row[1] || '')).filter(v => v !== '' && v !== '芸名'));
    dancerNames = dancerNames.filter(name => judgeNames.has(name));
  }

  return { jobList, dancerNames };
}

/**
 * 外注連絡票シートにダンサーマスタの情報を書き込む
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet 外注連絡票シート
 */
function fillDancerMasterToSheet(sheet) {
  const ss = getMasterSpreadsheet();
  const dancerSheet = ss.getSheetByName('ダンサーマスタ');
  const dancers = dancerSheet.getDataRange().getValues().slice(1).filter(row => row[0] !== '');

  // 4行目から書き込む（1-3行目はヘッダー）
  dancers.forEach((d, i) => {
    const row = 4 + i;
    // A:在籍, B:コード, C:芸名, D:氏名, M:振込先口座, N:口座番号, O:口座名義, P:住所
    sheet.getRange(row, 1).setValue(d[5]);  // 在籍
    sheet.getRange(row, 2).setValue(d[2]);  // コード
    sheet.getRange(row, 3).setValue(d[0]);  // 芸名
    sheet.getRange(row, 4).setValue(d[1]);  // 氏名
    sheet.getRange(row, 13).setValue(d[6]); // 金融機関
    sheet.getRange(row, 14).setValue(d[7]); // 口座番号
    sheet.getRange(row, 15).setValue(d[8]); // 口座名義
    sheet.getRange(row, 16).setValue(d[9]); // 住所
  });
}


