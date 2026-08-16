// ============================================================
// GASエディタで手動実行するテスト関数
// ============================================================

function testToWareki() {
  const cases = [
    { input: new Date('2019-04-30'), expected: '平成31年4月' },
    { input: new Date('2019-05-01'), expected: '令和元年5月' },
    { input: new Date('2026-04-01'), expected: '令和8年4月' },
    { input: new Date('2026-01-01'), expected: '令和8年1月' },
  ];
  cases.forEach(({ input, expected }) => {
    const result = toWareki(input);
    if (result !== expected) {
      throw new Error(`toWareki(${input}): expected "${expected}", got "${result}"`);
    }
  });
  Logger.log('testToWareki: PASSED');
}

function testGetMonthlyFileName() {
  const result = getMonthlyFileName(new Date('2026-04-01'));
  const expected = '【ダンサー】令和8年4月分 外注連絡表';
  if (result !== expected) {
    throw new Error(`expected "${expected}", got "${result}"`);
  }
  Logger.log('testGetMonthlyFileName: PASSED');
}

function testDateOnlyHandling() {
  const cases = [
    { input: '2026/06/10', expectedDate: '2026/06/10', expectedWeekday: '水' },
    { input: '2026-06-11', expectedDate: '2026/06/11', expectedWeekday: '木' }
  ];

  cases.forEach(function(testCase) {
    const date = parseDate(testCase.input);
    const formatted = Utilities.formatDate(date, 'Asia/Tokyo', 'yyyy/MM/dd');
    const weekday = getWeekday_(testCase.input);
    if (formatted !== testCase.expectedDate || weekday !== testCase.expectedWeekday) {
      throw new Error(
        'date handling failed: ' + testCase.input
        + ' => ' + formatted + ' (' + weekday + ')'
      );
    }
  });

  Logger.log('testDateOnlyHandling: PASSED');
}

function testInputRowOccupancy() {
  const blankRow = ['', '', '', '', '', '', '', '', '', ''];
  const missingDateButPopulated = ['', '月', '', '', 'OWL TIP', '', '', '', 'テスト', 1];

  if (hasInputRowData_(blankRow)) {
    throw new Error('blank input row must not be treated as occupied');
  }
  if (!hasInputRowData_(missingDateButPopulated)) {
    throw new Error('a populated row with a blank date must be treated as occupied');
  }

  Logger.log('testInputRowOccupancy: PASSED');
}

function testSubmissionIdNormalization() {
  const valid = 'submission:20260816_test-123';
  if (normalizeSubmissionId_(valid) !== valid) {
    throw new Error('valid submission ID was changed');
  }
  if (normalizeSubmissionId_('') !== '') {
    throw new Error('empty submission ID must stay empty');
  }

  let rejected = false;
  try {
    normalizeSubmissionId_('invalid id with spaces');
  } catch (err) {
    rejected = true;
  }
  if (!rejected) {
    throw new Error('invalid submission ID must be rejected');
  }

  Logger.log('testSubmissionIdNormalization: PASSED');
}

function testLookupUnitPrice() {
  // MASTER_SPREADSHEET_IDが設定済みの状態でテスト
  const price = lookupUnitPrice('OWL TIP');
  if (price !== 50) {
    throw new Error(`lookupUnitPrice('OWL TIP'): expected 50, got ${price}`);
  }
  Logger.log('testLookupUnitPrice: PASSED');
}

function testGetMasterData() {
  const data = getMasterData();
  if (!data.jobList || data.jobList.length === 0) {
    throw new Error('getMasterData: jobList is empty');
  }
  if (!data.dancerNames || data.dancerNames.length === 0) {
    throw new Error('getMasterData: dancerNames is empty');
  }
  Logger.log('testGetMasterData: PASSED, jobs=' + data.jobList.length + ', dancers=' + data.dancerNames.length);
}

/**
 * THP環境への移管後に、必要なGoogle資産へアクセスできるかを確認する
 * データの書き込みやメール送信は行わない
 */
function verifyTHPMigrationAccess() {
  const props = PropertiesService.getScriptProperties();
  const templateId = props.getProperty('TEMPLATE_SPREADSHEET_ID');
  const folderId = props.getProperty('MONTHLY_FOLDER_ID');
  if (!templateId || !folderId) {
    throw new Error('月次テンプレートまたは月次保存フォルダの設定がありません');
  }

  const data = getMasterData();
  const templateName = DriveApp.getFileById(templateId).getName();
  const folderName = DriveApp.getFolderById(folderId).getName();
  const remainingMailQuota = MailApp.getRemainingDailyQuota();

  Logger.log(
    'THP移管確認: PASSED / 案件=' + data.jobList.length
    + '件 / 担当者=' + data.dancerNames.length
    + '件 / テンプレート=' + templateName
    + ' / 保存先=' + folderName
    + ' / メール残数=' + remainingMailQuota
  );
}

function testGetOrCreateMonthlySheet() {
  const ss = getOrCreateMonthlySpreadsheet(new Date());
  if (!ss) throw new Error('getOrCreateMonthlySpreadsheet returned null');
  Logger.log('testGetOrCreateMonthlySheet: PASSED, title=' + ss.getName());
}

function testAppendRows() {
  const ss = getOrCreateMonthlySpreadsheet(new Date());
  const inputSheet = ss.getSheetByName('2.入力表');
  const beforeCount = inputSheet.getLastRow();

  appendRowsToInputSheet(ss, [
    { date: '2026/04/01', jobName: 'OWL TIP', detail: 'テスト', name: '斉藤愛乃', qty: 3, unitPrice: 50 }
  ]);

  const afterCount = inputSheet.getLastRow();
  if (afterCount !== beforeCount + 1) {
    throw new Error(`appendRows: expected ${beforeCount + 1} rows, got ${afterCount}`);
  }
  Logger.log('testAppendRows: PASSED');
}

function testSendLineMessage() {
  // 実際のLINEユーザーIDに書き換えてテスト
  const testUserId = 'Uxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx';
  sendLineMessage(testUserId, 'テスト: GASからのpushメッセージ');
  Logger.log('testSendLineMessage: PASSED (check LINE app)');
}

function testDoPostSimulation() {
  const mockPayload = JSON.stringify({
    userId: 'Uxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx',
    rows: [
      { date: '2026/04/01', jobName: 'OWL TIP', detail: '', name: '斉藤愛乃', qty: 5 },
      { date: '2026/04/01', jobName: 'PCDL GUEST', detail: 'VIP対応', name: '伊藤悠亜', qty: 1 }
    ]
  });

  const mockEvent = { postData: { contents: mockPayload } };
  const result = handleSubmission(mockEvent);
  Logger.log('testDoPostSimulation: PASSED, result=' + JSON.stringify(result));
}

function runAllTests() {
  testToWareki();
  testGetMonthlyFileName();
  testDateOnlyHandling();
  testInputRowOccupancy();
  testSubmissionIdNormalization();
  Logger.log('=== Utils tests PASSED ===');
  // 以下はスプレッドシートIDが設定済みの場合のみ実行
  // testLookupUnitPrice();
  // testGetMasterData();
  // testGetOrCreateMonthlySheet();
  // testAppendRows();
  // testDoPostSimulation();
}
