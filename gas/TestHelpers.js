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
  if (!data.jobNames || data.jobNames.length === 0) {
    throw new Error('getMasterData: jobNames is empty');
  }
  if (!data.dancerNames || data.dancerNames.length === 0) {
    throw new Error('getMasterData: dancerNames is empty');
  }
  Logger.log('testGetMasterData: PASSED, jobs=' + data.jobNames.length + ', dancers=' + data.dancerNames.length);
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
  Logger.log('=== Utils tests PASSED ===');
  // 以下はスプレッドシートIDが設定済みの場合のみ実行
  // testLookupUnitPrice();
  // testGetMasterData();
  // testGetOrCreateMonthlySheet();
  // testAppendRows();
  // testDoPostSimulation();
}
