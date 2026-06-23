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
 * 外注連絡票を管理しているスプレッドシートを取得する
 * @returns {GoogleAppsScript.Spreadsheet.Spreadsheet}
 */
function getOutsourceMasterSpreadsheet_() {
  const props = PropertiesService.getScriptProperties();
  const id = props.getProperty('OUTSOURCE_MASTER_SPREADSHEET_ID')
    || '1zGicgFAXB-P_xEfU3u54Uk6yOpCf937lyIlm0F5rQAo'
    || props.getProperty('MASTER_SPREADSHEET_ID');
  if (!id) throw new Error('OUTSOURCE_MASTER_SPREADSHEET_ID is not set in script properties');
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
 * マスタスプレッドシートの外注連絡票から人物情報を取得する
 * 列: B=在籍, C=コード, D=芸名, E=氏名, N=金融機関, O=口座番号, P=口座名義, R=郵便番号, S=住所, U=メールアドレス
 * @returns {Object[]} 外注連絡票の人物情報
 */
function getOutsourceContacts_() {
  const ss = getOutsourceMasterSpreadsheet_();
  const sheet = ss.getSheetByName('外注連絡票');
  if (!sheet) {
    throw new Error('マスタスプレッドシートに「外注連絡票」シートがありません');
  }

  const values = sheet.getDataRange().getValues();
  let startIndex = 3; // 通常は4行目からデータ
  for (let i = 0; i < values.length; i++) {
    if (String(values[i][1] || '').trim() === '在籍' && String(values[i][3] || '').trim() === '芸名') {
      startIndex = i + 1;
      break;
    }
  }

  return values.slice(startIndex).map(function(row) {
    const zip = row[17] ? String(row[17]).replace(/\.0$/, '').trim() : '';
    const address = String(row[18] || '').trim();
    return {
      membership: String(row[1] || '').trim(),
      code: String(row[2] || '').replace(/\.0$/, '').trim(),
      stageName: String(row[3] || '').trim(),
      realName: String(row[4] || '').trim(),
      bank: String(row[13] || '').trim(),
      accountNumber: String(row[14] || '').replace(/\.0$/, '').trim(),
      accountName: String(row[15] || '').trim(),
      address: zip ? '〒' + zip + ' ' + address : address,
      email: String(row[20] || '').trim()
    };
  }).filter(function(person) {
    return person.stageName !== '';
  });
}

/**
 * 外注連絡票の芸名からメールアドレスを取得する
 * @returns {Object<string, string>} {芸名: メールアドレス}
 */
function getOutsourceContactEmailMap_() {
  const emailMap = {};
  getOutsourceContacts_().forEach(function(person) {
    if (person.stageName && person.email) {
      emailMap[person.stageName] = person.email;
    }
  });
  return emailMap;
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

  // 外注連絡票: D=芸名
  let dancerNames = getOutsourceContacts_().map(person => person.stageName).filter(v => v !== '');

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
 * 外注連絡票シートにマスタ外注連絡票の人物情報を書き込む
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet 外注連絡票シート
 */
function fillDancerMasterToSheet(sheet) {
  const dancers = getOutsourceContacts_();

  if (sheet.getLastRow() >= 4) {
    const clearRows = Math.max(sheet.getLastRow() - 3, dancers.length, 1);
    sheet.getRange(4, 1, clearRows, 4).clearContent();   // A:D
    sheet.getRange(4, 13, clearRows, 9).clearContent();  // M:U
  }

  // 4行目から書き込む（1-3行目はヘッダー）
  dancers.forEach((d, i) => {
    const row = 4 + i;
    // A:在籍, B:コード, C:芸名, D:氏名, M:金融機関, N:口座番号, O:口座名義, P:住所, U:メール
    sheet.getRange(row, 1).setValue(d.membership);
    sheet.getRange(row, 2).setValue(d.code);
    sheet.getRange(row, 3).setValue(d.stageName);
    sheet.getRange(row, 4).setValue(d.realName);
    sheet.getRange(row, 13).setValue(d.bank);
    sheet.getRange(row, 14).setValue(d.accountNumber);
    sheet.getRange(row, 15).setValue(d.accountName);
    sheet.getRange(row, 16).setValue(d.address);
    sheet.getRange(row, 21).setValue(d.email);
  });
}
