/**
 * スプレッドシートを開いたときにメニューを追加する
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('PDF送信')
    .addItem('明細PDFを一括送信', 'sendPdfsToAll')
    .addToUi();
}

/**
 * 2.入力表のデータを元にPDFを生成し、メールで送信するメイン処理
 */
function sendPdfsToAll() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  
  // 1. ダンサーマスタからメールアドレスの取得 (L列: 12列目を想定)
  let masterSs;
  try {
    masterSs = getMasterSpreadsheet();
  } catch (e) {
    ui.alert('マスタスプレッドシートの取得に失敗しました: ' + e.message);
    return;
  }
  
  const dancerSheet = masterSs.getSheetByName('ダンサーマスタ');
  if (!dancerSheet) {
    ui.alert('エラー: ダンサーマスタシートが見つかりません。');
    return;
  }
  
  const dancerData = dancerSheet.getDataRange().getValues();
  const mailMap = {}; // {名前: メールアドレス}
  // 1行目はヘッダー
  for (let i = 1; i < dancerData.length; i++) {
    const name = String(dancerData[i][0]).trim(); // A列: 芸名
    const email = String(dancerData[i][11] || '').trim(); // L列: 12番目 (index 11)
    if (name && email) {
      mailMap[name] = email;
    }
  }

  // 2. 「2.入力表」からデータ取得
  const inputSheet = ss.getSheetByName('2.入力表');
  if (!inputSheet) {
    ui.alert('エラー: 2.入力表シートが見つかりません。月次スプレッドシート上で実行してください。');
    return;
  }
  
  const lastRow = inputSheet.getLastRow();
  if (lastRow < 8) {
    ui.alert('データがありません。');
    return;
  }
  
  const inputData = inputSheet.getRange(8, 3, lastRow - 7, 12).getValues();
  // 取得範囲 C~N列
  // 0:日程, 1:曜日, 2:現場, 3:項目, 4:案件名, 5:H, 6:I, 7:詳細, 8:名前, 9:数量, 10:単価, 11:合計
  
  const targetData = {}; // {名前: [行データの配列]}
  
  for (let i = 0; i < inputData.length; i++) {
    const row = inputData[i];
    const date = row[0];
    const name = String(row[8] || '').trim();
    if (!date || !name) continue;
    
    if (!targetData[name]) {
      targetData[name] = [];
    }
    targetData[name].push({
      date: date instanceof Date ? Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy/MM/dd') : date,
      venue: row[2],
      category: row[3],
      jobName: row[4],
      qty: row[9],
      unitPrice: row[10],
      total: row[11]
    });
  }
  
  // 3. テンプレートシートへの書き込みとPDF化、メール送信
  const templateSheet = ss.getSheetByName('明細書テンプレート');
  if (!templateSheet) {
    ui.alert('エラー: 「明細書テンプレート」シートが見つかりません。作成してください。');
    return;
  }
  
  let sentCount = 0;
  let errorMessages = [];

  Object.keys(targetData).forEach(name => {
    const email = mailMap[name];
    if (!email) {
      errorMessages.push(`・${name} 様のメールアドレスがマスタに登録されていません。`);
      return; // continue
    }
    
    const records = targetData[name];
    
    // テンプレートシートの初期化（例: B4に名前、A10以降に明細データとする）
    templateSheet.getRange('B4').setValue(name + ' 様');
    
    // 対象月（C4など、日付から推測）
    if (records.length > 0 && records[0].date) {
        const firstDateStr = String(records[0].date);
        const match = firstDateStr.match(/^(\d{4})\/(\d{1,2})/);
        if (match) {
            templateSheet.getRange('C4').setValue(match[1] + '年' + match[2] + '月分');
        }
    }

    // 既存の明細行をクリア (10行目から50行目までを想定)
    const startRow = 10;
    const maxRows = 50;
    templateSheet.getRange(startRow, 1, maxRows, 6).clearContent();
    
    // 明細書き込み
    const outData = [];
    let grandTotal = 0;
    records.forEach(rec => {
      outData.push([
        rec.date,        // A: 日程
        rec.venue,       // B: 現場
        rec.jobName,     // C: 案件名
        rec.qty,         // D: 数量
        rec.unitPrice,   // E: 単価
        rec.total        // F: 合計
      ]);
      grandTotal += Number(rec.total) || 0;
    });
    
    if (outData.length > 0) {
      templateSheet.getRange(startRow, 1, outData.length, 6).setValues(outData);
    }
    // 総合計
    templateSheet.getRange('B6').setValue(grandTotal);
    
    SpreadsheetApp.flush(); // 書き込みをシートに反映
    
    // 短時間のスリープを入れ、Google側の反映を待つ
    Utilities.sleep(1500);

    // PDF化
    try {
      const pdfFileName = `${name}_明細書.pdf`;
      const pdfBlob = createPdfFromSheet(ss, templateSheet, pdfFileName);
      
      // メール送信
      const subject = `【ご案内】明細書のご送付（${name}様）`;
      const body = `${name} 様\n\nお疲れ様です。\n今月分の明細書を添付ファイルにてお送りいたします。\n\nご確認のほどよろしくお願いいたします。`;
      
      GmailApp.sendEmail(email, subject, body, {
        attachments: [pdfBlob],
        name: '管理事務局'
      });
      sentCount++;
    } catch (e) {
      errorMessages.push(`・${name} 様の処理中にエラーが発生しました: ${e.message}`);
    }
  });
  
  let resultMsg = `${sentCount}件のメールを送信しました。`;
  if (errorMessages.length > 0) {
    resultMsg += `\n\n【エラー/スキップ】\n` + errorMessages.join('\n');
  }
  ui.alert(resultMsg);
}

/**
 * 指定したシートをPDFとしてエクスポートし、Blobを返す
 */
function createPdfFromSheet(ss, sheet, pdfFileName) {
  const ssId = ss.getId();
  const sheetId = sheet.getSheetId();
  
  const url = "https://docs.google.com/spreadsheets/d/" + ssId + "/export"
            + "?format=pdf"
            + "&size=A4"
            + "&portrait=true"
            + "&fitw=true"
            + "&sheetnames=false"
            + "&printtitle=false"
            + "&pagenumbers=false"
            + "&gridlines=false"
            + "&fzr=false"
            + "&gid=" + sheetId;
            
  const token = ScriptApp.getOAuthToken();
  const response = UrlFetchApp.fetch(url, {
    headers: {
      'Authorization': 'Bearer ' + token
    },
    muteHttpExceptions: true
  });
  
  if (response.getResponseCode() !== 200) {
    throw new Error('PDFのエクスポートに失敗しました。Status: ' + response.getResponseCode());
  }
  
  return response.getBlob().setName(pdfFileName);
}

/**
 * 特定の人物の明細PDFを生成し、メールで送信する（フォーム送信時などのバックグラウンド用）
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss 対象の月次スプレッドシート
 * @param {string} targetName 対象の人物名
 * @param {string} targetDate 抽出対象の日付（YYYY/MM/DD形式などを想定）
 */
function sendPdfForPerson(ss, targetName, targetDate) {
  if (!targetName) return;

  // 1. ダンサーマスタからメールアドレスの取得 (L列: 12列目を想定)
  let masterSs;
  try {
    masterSs = getMasterSpreadsheet();
  } catch (e) {
    console.error('マスタスプレッドシートの取得に失敗しました: ' + e.message);
    return;
  }
  
  const dancerSheet = masterSs.getSheetByName('ダンサーマスタ');
  if (!dancerSheet) {
    console.error('エラー: ダンサーマスタシートが見つかりません。');
    return;
  }
  
  const dancerData = dancerSheet.getDataRange().getValues();
  let email = '';
  for (let i = 1; i < dancerData.length; i++) {
    const name = String(dancerData[i][0]).trim();
    if (name === targetName) {
      email = String(dancerData[i][11] || '').trim(); // L列: 12番目 (index 11)
      break;
    }
  }

  if (!email) {
    console.warn(`${targetName} 様のメールアドレスがマスタに登録されていません。送信をスキップします。`);
    return;
  }

  // 2. 「2.入力表」からデータ取得
  const inputSheet = ss.getSheetByName('2.入力表');
  if (!inputSheet) {
    console.error('エラー: 2.入力表シートが見つかりません。');
    return;
  }
  
  const lastRow = inputSheet.getLastRow();
  if (lastRow < 8) return;
  
  const inputData = inputSheet.getRange(8, 3, lastRow - 7, 12).getValues();
  const records = [];
  
  for (let i = 0; i < inputData.length; i++) {
    const row = inputData[i];
    let dateObj = row[0];
    let dateStr = dateObj instanceof Date ? Utilities.formatDate(dateObj, Session.getScriptTimeZone(), 'yyyy/MM/dd') : String(dateObj);
    const name = String(row[8] || '').trim();
    
    // targetDateが指定されている場合はその日付のみを抽出
    if (name === targetName && dateStr && (!targetDate || dateStr === targetDate)) {
      records.push({
        date: dateStr,
        venue: row[2],
        category: row[3],
        jobName: row[4],
        qty: row[9],
        unitPrice: row[10],
        total: row[11]
      });
    }
  }

  if (records.length === 0) {
    console.warn(`${targetName} 様の対象データが見つかりませんでした。`);
    return;
  }
  
  // 3. テンプレートシートへの書き込みとPDF化、メール送信
  const templateSheet = ss.getSheetByName('明細書テンプレート');
  if (!templateSheet) {
    console.error('エラー: 「明細書テンプレート」シートが見つかりません。作成してください。');
    return;
  }
  
  // テンプレートシートの初期化
  templateSheet.getRange('B4').setValue(targetName + ' 様');
  
  // 対象月
  if (records[0].date) {
      const firstDateStr = String(records[0].date);
      const match = firstDateStr.match(/^(\d{4})\/(\d{1,2})/);
      if (match) {
          templateSheet.getRange('C4').setValue(match[1] + '年' + match[2] + '月分');
      }
  }

  const startRow = 10;
  const maxRows = 50;
  templateSheet.getRange(startRow, 1, maxRows, 6).clearContent();
  
  const outData = [];
  let grandTotal = 0;
  records.forEach(rec => {
    outData.push([
      rec.date,
      rec.venue,
      rec.jobName,
      rec.qty,
      rec.unitPrice,
      rec.total
    ]);
    grandTotal += Number(rec.total) || 0;
  });
  
  if (outData.length > 0) {
    templateSheet.getRange(startRow, 1, outData.length, 6).setValues(outData);
  }
  templateSheet.getRange('B6').setValue(grandTotal);
  
  SpreadsheetApp.flush();
  Utilities.sleep(1500);

  // PDF化
  try {
    const pdfFileName = `${targetName}_明細書.pdf`;
    const pdfBlob = createPdfFromSheet(ss, templateSheet, pdfFileName);
    
    const subject = `【ご案内】受領内容および明細書のご送付（${targetName}様）`;
    const body = `${targetName} 様\n\nお疲れ様です。\nフォームより送信された内容を受け付けました。\n最新の明細書を添付ファイルにてお送りいたします。\n\nご確認のほどよろしくお願いいたします。`;
    
    GmailApp.sendEmail(email, subject, body, {
      attachments: [pdfBlob],
      name: '管理事務局'
    });
    console.log(`${targetName} 様へPDFを送信しました。`);
  } catch (e) {
    console.error(`${targetName} 様の処理中にエラーが発生しました: ${e.message}`);
  }
}
