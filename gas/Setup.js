/**
 * ================================================================
 * 【初回セットアップ専用】マスタデータ自動転記スクリプト
 * ================================================================
 *
 * 使い方:
 *   1. 【ダンサー】編集用.xlsx を Google ドライブにアップロード
 *   2. Googleスプレッドシートとして開き、URLからIDをコピー
 *      例) https://docs.google.com/spreadsheets/d/<ここがID>/edit
 *   3. GASエディタ → プロジェクトの設定 → スクリプトプロパティ に追加:
 *      キー: DANCER_EXCEL_SPREADSHEET_ID  値: 上記のID
 *   4. この画面で「setupMasterFromExcel」を選択して「実行」ボタンを押す
 *   5. 実行ログに「✅ マスタセットアップ完了」と表示されれば成功
 * ================================================================
 */
function setupMasterFromExcel() {
  const props = PropertiesService.getScriptProperties();
  const excelId = props.getProperty('DANCER_EXCEL_SPREADSHEET_ID');
  if (!excelId) {
    throw new Error(
      'スクリプトプロパティ「DANCER_EXCEL_SPREADSHEET_ID」が設定されていません。\n' +
      'GASエディタ → プロジェクトの設定 → スクリプトプロパティ で設定してください。'
    );
  }

  const excelSs = SpreadsheetApp.openById(excelId);
  const masterSs = getMasterSpreadsheet();

  // ダンサーマスタへの転記
  const dancerCount = _setupDancerMaster(excelSs, masterSs);

  // 案件マスタへの転記
  const jobCount = _setupJobMaster(excelSs, masterSs);

  const msg = `✅ マスタセットアップ完了\n　・ダンサーマスタ: ${dancerCount}件\n　・案件マスタ: ${jobCount}件`;
  Logger.log(msg);
  SpreadsheetApp.getUi && SpreadsheetApp.getUi().alert(msg);
}

// ----------------------------------------------------------------
// ダンサーマスタへの転記
// Excelシート: 外注連絡票 + 判定表
// 転記先: ダンサーマスタ
//   A:芸名 B:本名 C:コード D:時給 E:交通費 F:在籍
//   G:金融機関 H:口座番号 I:口座名義(カナ) J:住所 K:連絡先
// ----------------------------------------------------------------
function _setupDancerMaster(excelSs, masterSs) {
  const srcSheet   = excelSs.getSheetByName('外注連絡票');
  const judgeSheet = excelSs.getSheetByName('判定表');

  if (!srcSheet)   throw new Error('Excelに「外注連絡票」シートが見つかりません');
  if (!judgeSheet) throw new Error('Excelに「判定表」シートが見つかりません');

  const srcData   = srcSheet.getDataRange().getValues();
  const judgeData = judgeSheet.getDataRange().getValues();

  // ── 判定表から 芸名 → {時給, 交通費} マップを作成 ──────────────
  // ヘッダー行: col2=時給, col8=キツネ交通費
  // データ行 : col1=芸名
  const judgeMap = {};
  for (let i = 0; i < judgeData.length; i++) {
    const name = judgeData[i][1];
    if (!name || name === '芸名' || name === '') continue;
    judgeMap[name] = {
      hourlyRate: Number(judgeData[i][2]) || 0,
      transport:  Number(judgeData[i][8]) || 0
    };
  }

  // ── 外注連絡票からデータ行の開始位置を探す ─────────────────────
  // ヘッダー行は「在籍」「コード」「芸名」が含まれる行
  let headerRow = -1;
  for (let i = 0; i < srcData.length; i++) {
    if (srcData[i][1] === '在籍' && srcData[i][3] === '芸名') {
      headerRow = i;
      break;
    }
  }
  if (headerRow === -1) throw new Error('外注連絡票のヘッダー行が見つかりません（「在籍」「芸名」の行を確認してください）');

  // ── ダンサーデータを抽出 ─────────────────────────────────────
  const dancers = [];
  for (let i = headerRow + 1; i < srcData.length; i++) {
    const row = srcData[i];
    const membership = row[1];  // A列: 在籍
    const code       = row[2];  // B列: コード
    const stageName  = row[3];  // C列: 芸名
    const realName   = row[4];  // D列: 氏名(本名)

    if (!stageName || stageName === '') continue; // 空行はスキップ

    // 住所: col16='〒', col17=郵便番号, col18=番地
    const zip     = row[17] ? String(row[17]).replace(/\.0$/, '') : '';
    const addr    = row[18] ? String(row[18]) : '';
    const fullAddr = zip ? `〒${zip} ${addr}` : addr;

    // 金融機関
    const bank    = row[13] || '';
    const bankNo  = row[14] ? String(row[14]).replace(/\.0$/, '') : '';
    const bankName = row[15] || '';
    const contact = row[19] || '';

    const judgeInfo = judgeMap[stageName] || { hourlyRate: 0, transport: 0 };

    dancers.push([
      stageName,              // A: 芸名
      realName,               // B: 本名
      code !== '-' ? String(code).replace(/\.0$/, '') : '-', // C: コード
      judgeInfo.hourlyRate,   // D: 時給
      judgeInfo.transport,    // E: 交通費
      membership,             // F: 在籍
      bank,                   // G: 金融機関
      bankNo,                 // H: 口座番号
      bankName,               // I: 口座名義(カナ)
      fullAddr,               // J: 住所
      contact                 // K: 連絡先
    ]);
  }

  // ── マスタシートに書き込み ─────────────────────────────────────
  const dstSheet = masterSs.getSheetByName('ダンサーマスタ');
  if (!dstSheet) throw new Error('マスタスプレッドシートに「ダンサーマスタ」シートがありません');

  // ヘッダー(1行目)は残してデータ行のみクリア
  if (dstSheet.getLastRow() > 1) {
    dstSheet.getRange(2, 1, dstSheet.getLastRow() - 1, 11).clearContent();
  }
  if (dancers.length > 0) {
    dstSheet.getRange(2, 1, dancers.length, 11).setValues(dancers);
  }

  Logger.log(`ダンサーマスタ: ${dancers.length}件 転記完了`);
  return dancers.length;
}

// ----------------------------------------------------------------
// 案件マスタへの転記
// Excelシート: 金額
//   列13=案件名, 列14=現場名, 列15=単価, 列16=項目
//   列18=案件名, 列19=現場名, 列20=単価, 列21=項目  (右側にも同形式で続く)
// 転記先: 案件マスタ
//   A:案件名 B:現場コード C:単価 D:項目区分
// ----------------------------------------------------------------
function _setupJobMaster(excelSs, masterSs) {
  const srcSheet = excelSs.getSheetByName('金額');
  if (!srcSheet) throw new Error('Excelに「金額」シートが見つかりません');

  const srcData = srcSheet.getDataRange().getValues();

  // ヘッダー行を探す（'案件名' が含まれる行）
  let headerRow = -1;
  for (let i = 0; i < srcData.length; i++) {
    if (srcData[i][13] === '案件名') {
      headerRow = i;
      break;
    }
  }
  if (headerRow === -1) throw new Error('金額シートのヘッダー行が見つかりません（「案件名」の列を確認してください）');

  const jobMap = {};  // 重複防止用

  const _addJob = (name, venue, price, category) => {
    if (!name || name === '') return;
    const key = `${name}__${venue}`;
    if (!jobMap[key]) {
      jobMap[key] = [
        String(name),
        String(venue || ''),
        Number(price) || 0,
        String(category || '')
      ];
    }
  };

  for (let i = headerRow + 1; i < srcData.length; i++) {
    const row = srcData[i];
    _addJob(row[13], row[14], row[15], row[16]); // 左側ブロック
    _addJob(row[18], row[19], row[20], row[21]); // 右側ブロック
  }

  const jobs = Object.values(jobMap);

  // ── マスタシートに書き込み ─────────────────────────────────────
  const dstSheet = masterSs.getSheetByName('案件マスタ');
  if (!dstSheet) throw new Error('マスタスプレッドシートに「案件マスタ」シートがありません');

  if (dstSheet.getLastRow() > 1) {
    dstSheet.getRange(2, 1, dstSheet.getLastRow() - 1, 4).clearContent();
  }
  if (jobs.length > 0) {
    dstSheet.getRange(2, 1, jobs.length, 4).setValues(jobs);
  }

  Logger.log(`案件マスタ: ${jobs.length}件 転記完了`);
  return jobs.length;
}
