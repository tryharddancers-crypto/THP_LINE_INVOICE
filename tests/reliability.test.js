const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const test = require('node:test');
const vm = require('node:vm');

const ROOT = path.resolve(__dirname, '..');

function createBackendHarness(options = {}) {
  const properties = new Map();
  const counts = {
    append: 0,
    email: 0,
    line: 0,
    flush: 0,
    release: 0
  };

  const scriptProperties = {
    getProperty(key) {
      return properties.has(key) ? properties.get(key) : null;
    },
    setProperty(key, value) {
      properties.set(key, value);
      return this;
    },
    deleteProperty(key) {
      properties.delete(key);
      return this;
    },
    getProperties() {
      return Object.fromEntries(properties);
    }
  };

  const spreadsheet = {
    getId() {
      return 'monthly-sheet-id';
    }
  };

  const context = vm.createContext({
    console,
    Date,
    JSON,
    Math,
    Object,
    String,
    Number,
    RegExp,
    Error,
    Logger: { log() {} },
    LockService: {
      getScriptLock() {
        return {
          tryLock() {
            return options.lockAvailable !== false;
          },
          releaseLock() {
            counts.release += 1;
          }
        };
      }
    },
    PropertiesService: {
      getScriptProperties() {
        return scriptProperties;
      }
    },
    Utilities: {
      getUuid() {
        return 'fallback-' + (properties.size + 1);
      },
      sleep() {}
    },
    SpreadsheetApp: {
      flush() {
        counts.flush += 1;
      },
      openById() {
        return spreadsheet;
      }
    },
    parseDate(value) {
      return new Date(value.replace(/\//g, '-'));
    },
    getOrCreateMonthlySpreadsheet() {
      return spreadsheet;
    },
    appendRowsToInputSheet(_ss, rows) {
      counts.append += 1;
      return { rowNumbers: rows.map((_, index) => index + 8) };
    },
    sendInputRowsNotification_() {
      counts.email += 1;
      if (options.emailFails) throw new Error('mail failed');
    },
    sendLineMessage() {
      counts.line += 1;
      if (options.lineFails) throw new Error('line failed');
    },
    verifyLineIdToken_(token) {
      if (token === 'invalid') throw new Error('invalid token');
      return 'verified-line-user';
    },
    validateSubmissionRows_(rows) {
      return rows;
    },
    reportSystemError_() {
    },
    runOpportunisticMaintenance_() {
    }
  });

  for (const relativePath of ['gas/SubmissionGuard.js', 'gas/Code.js']) {
    vm.runInContext(
      fs.readFileSync(path.join(ROOT, relativePath), 'utf8'),
      context,
      { filename: relativePath }
    );
  }

  function submit(submissionId, overrides = {}) {
    const body = {
      submissionId,
      idToken: 'valid-token',
      userId: 'line-user',
      rows: [{
        date: '2026/08/16',
        jobName: 'TEST TIP',
        name: 'TEST USER',
        qty: 1,
        detail: '',
        unitPrice: 1
      }],
      ...overrides
    };
    return context.handleSubmission({ postData: { contents: JSON.stringify(body) } });
  }

  return { context, counts, properties, submit };
}

test('same submission ID writes and notifies only once', () => {
  const harness = createBackendHarness();
  const first = harness.submit('submission-1');
  const second = harness.submit('submission-1');

  assert.equal(first.duplicate, false);
  assert.equal(second.duplicate, true);
  assert.equal(harness.counts.append, 1);
  assert.equal(harness.counts.email, 1);
  assert.equal(harness.counts.line, 1);
});

test('different submission IDs remain separate submissions', () => {
  const harness = createBackendHarness();
  harness.submit('submission-1');
  harness.submit('submission-2');

  assert.equal(harness.counts.append, 2);
  assert.equal(harness.counts.email, 2);
  assert.equal(harness.counts.line, 2);
});

test('notification failures do not duplicate a saved row on retry', () => {
  const harness = createBackendHarness({ emailFails: true, lineFails: true });
  const first = harness.submit('submission-notification-failure');
  const second = harness.submit('submission-notification-failure');

  assert.equal(first.ok, true);
  assert.equal(second.duplicate, true);
  assert.equal(harness.counts.append, 1);
  assert.equal(harness.counts.email, 1);
  assert.equal(harness.counts.line, 1);
});

test('an unavailable lock rejects the request before writing', () => {
  const harness = createBackendHarness({ lockAvailable: false });

  assert.throws(
    () => harness.submit('submission-locked'),
    /現在ほかの送信を処理しています/
  );
  assert.equal(harness.counts.append, 0);
  assert.equal(harness.counts.release, 0);
});

test('an invalid submission ID is rejected before writing', () => {
  const harness = createBackendHarness();

  assert.throws(
    () => harness.submit('invalid id'),
    /submissionId is invalid/
  );
  assert.equal(harness.counts.append, 0);
});

test('an invalid LINE identity is rejected before writing', () => {
  const harness = createBackendHarness();
  assert.throws(
    () => harness.submit('submission-auth-failure', { idToken: 'invalid' }),
    /invalid token/
  );
  assert.equal(harness.counts.append, 0);
});

test('an expired submission record does not block a new write', () => {
  const harness = createBackendHarness();
  harness.properties.set(
    'submission.guard.expired-submission',
    JSON.stringify({
      savedAt: Date.now() - (8 * 24 * 60 * 60 * 1000),
      result: { ok: true, duplicate: false }
    })
  );

  const result = harness.submit('expired-submission');
  assert.equal(result.duplicate, false);
  assert.equal(harness.counts.append, 1);
});

test('a malformed submission record is removed and rewritten', () => {
  const harness = createBackendHarness();
  harness.properties.set('submission.guard.malformed-submission', '{bad json');

  const result = harness.submit('malformed-submission');
  assert.equal(result.duplicate, false);
  assert.equal(harness.counts.append, 1);
  assert.doesNotThrow(() => JSON.parse(
    harness.properties.get('submission.guard.malformed-submission')
  ));
});

test('automatic form submission has no PDF sender call', () => {
  const source = fs.readFileSync(path.join(ROOT, 'gas/Code.js'), 'utf8');
  assert.doesNotMatch(source, /sendPdfForPerson\s*\(/);
});

test('master loading failures return a safe error and alert the administrator', () => {
  const alerts = [];
  const context = vm.createContext({
    console, Date, JSON, Math, Object, String, Number, Error,
    Logger: { log() {} },
    ContentService: {
      MimeType: { JSON: 'JSON' },
      createTextOutput(text) {
        return {
          text,
          setMimeType() { return this; }
        };
      }
    },
    getMasterData() { throw new Error('master unavailable'); },
    reportSystemError_(label, err) { alerts.push({ label, message: err.message }); }
  });
  vm.runInContext(
    fs.readFileSync(path.join(ROOT, 'gas/Code.js'), 'utf8'),
    context,
    { filename: 'gas/Code.js' }
  );

  const response = JSON.parse(context.doGet({ parameter: { action: 'getMaster' } }).text);
  assert.equal(response.ok, false);
  assert.match(response.error, /入力候補を読み込めませんでした/);
  assert.deepEqual(alerts, [{ label: 'フォーム用マスタデータ取得', message: 'master unavailable' }]);
});

function extractInlineScript(htmlPath) {
  const html = fs.readFileSync(htmlPath, 'utf8');
  const scripts = [...html.matchAll(/<script(?:\s[^>]*)?>([\s\S]*?)<\/script>/gi)]
    .map(match => match[1])
    .filter(source => source.trim());
  return scripts[scripts.length - 1].replace(/\binit\(\);\s*$/, '');
}

function verifyFrontendLifecycle(relativePath) {
  const storage = new Map();
  let uuid = 0;
  const elements = new Map();
  const context = vm.createContext({
    console,
    JSON,
    Math,
    Date,
    Object,
    String,
    Number,
    setTimeout() {},
    fetch() {},
    liff: {},
    window: {
      crypto: {
        randomUUID() {
          uuid += 1;
          return 'generated-' + uuid;
        }
      },
      localStorage: {
        getItem(key) {
          return storage.has(key) ? storage.get(key) : null;
        },
        setItem(key, value) {
          storage.set(key, value);
        },
        removeItem(key) {
          storage.delete(key);
        }
      }
    },
    document: {
      getElementById(id) {
        if (!elements.has(id)) elements.set(id, { style: {}, textContent: '' });
        return elements.get(id);
      }
    }
  });

  vm.runInContext(
    extractInlineScript(path.join(ROOT, relativePath)),
    context,
    { filename: relativePath }
  );

  const rows = [{ date: '2026/08/16', name: 'TEST USER', qty: 1 }];
  vm.runInContext('prepareSubmissionId(' + JSON.stringify(rows) + ')', context);
  const firstId = vm.runInContext('pendingSubmissionId', context);

  vm.runInContext('backToForm()', context);
  vm.runInContext('prepareSubmissionId(' + JSON.stringify(rows) + ')', context);
  assert.equal(vm.runInContext('pendingSubmissionId', context), firstId);

  const changedRows = [{ date: '2026/08/16', name: 'TEST USER', qty: 2 }];
  vm.runInContext('prepareSubmissionId(' + JSON.stringify(changedRows) + ')', context);
  const changedId = vm.runInContext('pendingSubmissionId', context);
  assert.notEqual(changedId, firstId);

  // Returning to the original draft after an uncertain send must recover its ID.
  vm.runInContext('prepareSubmissionId(' + JSON.stringify(rows) + ')', context);
  assert.equal(vm.runInContext('pendingSubmissionId', context), firstId);

  vm.runInContext('clearPendingSubmission()', context);
  assert.equal(vm.runInContext('pendingSubmissionId', context), null);

  vm.runInContext('prepareSubmissionId(' + JSON.stringify(changedRows) + ')', context);
  assert.equal(vm.runInContext('pendingSubmissionId', context), changedId);
  vm.runInContext('clearPendingSubmission()', context);
  assert.equal(storage.size, 0);
}

for (const relativePath of ['gas/liff/index.html', 'liff-frontend/index.html']) {
  test(relativePath + ' preserves IDs for retries and replaces them after edits', () => {
    verifyFrontendLifecycle(relativePath);
  });
}

function createValidationHarness(options = {}) {
  const fetchCalls = [];
  const context = vm.createContext({
    Date,
    JSON,
    Number,
    String,
    Object,
    Array,
    RegExp,
    Error,
    PropertiesService: {
      getScriptProperties() {
        return {
          getProperty(key) {
            if (key === 'LIFF_ID') return '2009725727-3jvV9g52';
            return null;
          }
        };
      }
    },
    UrlFetchApp: {
      fetch(url, request) {
        fetchCalls.push({ url, request });
        return {
          getResponseCode() { return options.verifyCode || 200; },
          getContentText() {
            return JSON.stringify(options.verifyResponse || {
              sub: 'Uverified', aud: '2009725727', exp: Math.floor(Date.now() / 1000) + 300
            });
          }
        };
      }
    },
    getMasterData() {
      return {
        jobList: [
          { name: 'OWL TIP', billing: 'OWL', unitPrice: 50, category: 'TIP' },
          { name: 'WINX CA', billing: 'OWL', unitPrice: 1000, category: 'WINX' },
          { name: 'WINX CA', billing: 'KTN', unitPrice: 1200, category: 'WINX' }
        ],
        dancerNames: ['TEST USER']
      };
    },
    parseDateParts_(value) {
      const match = String(value || '').match(/^(\d{4})[\/-](\d{1,2})[\/-](\d{1,2})$/);
      if (!match) throw new Error('invalid date');
      return { year: Number(match[1]), month: Number(match[2]), day: Number(match[3]) };
    }
  });
  vm.runInContext(
    fs.readFileSync(path.join(ROOT, 'gas/SecurityValidation.js'), 'utf8'),
    context,
    { filename: 'gas/SecurityValidation.js' }
  );
  return { context, fetchCalls };
}

test('LINE ID token is verified against the configured LIFF channel', () => {
  const harness = createValidationHarness();
  assert.equal(harness.context.verifyLineIdToken_('token-value'), 'Uverified');
  assert.equal(harness.fetchCalls.length, 1);
  assert.equal(harness.fetchCalls[0].request.payload.client_id, '2009725727');
  assert.equal(harness.fetchCalls[0].request.payload.id_token, 'token-value');
});

test('invalid LINE ID token is rejected', () => {
  const harness = createValidationHarness({ verifyCode: 400, verifyResponse: { error: 'invalid' } });
  assert.throws(() => harness.context.verifyLineIdToken_('bad-token'), /LINEの認証期限/);
});

test('server validation replaces client prices with master prices', () => {
  const harness = createValidationHarness();
  const rows = harness.context.validateSubmissionRows_([{
    date: '2026/08/16', billing: 'OWL', name: 'TEST USER', jobName: 'OWL TIP', qty: 2, detail: '', unitPrice: 999999
  }]);
  assert.equal(rows[0].unitPrice, 50);
  assert.equal(rows[0].date, '2026/08/16');
  assert.equal(rows[0].billing, 'OWL');
  assert.equal(rows[0].category, 'TIP');
});

test('server validation resolves duplicated product names by store', () => {
  const harness = createValidationHarness();
  const rows = harness.context.validateSubmissionRows_([{
    date: '2026/08/16', billing: 'KTN', name: 'TEST USER', jobName: 'WINX CA', qty: 2
  }]);
  assert.equal(rows[0].billing, 'KTN');
  assert.equal(rows[0].category, 'WINX');
  assert.equal(rows[0].unitPrice, 1200);

  assert.throws(() => harness.context.validateSubmissionRows_([{
    date: '2026/08/16', name: 'TEST USER', jobName: 'WINX CA', qty: 1
  }]), /店舗/);
  assert.throws(() => harness.context.validateSubmissionRows_([{
    date: '2026/08/16', billing: 'BMB', name: 'TEST USER', jobName: 'WINX CA', qty: 1
  }]), /選択された店舗/);
});

test('server validation rejects mixed-month submissions', () => {
  const harness = createValidationHarness();
  assert.throws(() => harness.context.validateSubmissionRows_([
    { date: '2026/08/31', name: 'TEST USER', jobName: 'OWL TIP', qty: 1 },
    { date: '2026/09/01', name: 'TEST USER', jobName: 'OWL TIP', qty: 1 }
  ]), /異なる月/);
});

test('server validation rejects unknown people, jobs and invalid quantities', () => {
  const harness = createValidationHarness();
  assert.throws(() => harness.context.validateSubmissionRows_([
    { date: '2026/08/16', name: 'UNKNOWN', jobName: 'OWL TIP', qty: 1 }
  ]), /担当者/);
  assert.throws(() => harness.context.validateSubmissionRows_([
    { date: '2026/08/16', name: 'TEST USER', jobName: 'UNKNOWN', qty: 1 }
  ]), /商品/);
  assert.throws(() => harness.context.validateSubmissionRows_([
    { date: '2026/08/16', name: 'TEST USER', jobName: 'OWL TIP', qty: 0 }
  ]), /数量/);
  assert.throws(() => harness.context.validateSubmissionRows_([
    { date: '', name: 'TEST USER', jobName: 'OWL TIP', qty: 1 }
  ]), /日程/);
});

test('all frontend copies send a LINE ID token', () => {
  for (const relativePath of ['gas/liff/index.html', 'liff-frontend/index.html']) {
    const source = fs.readFileSync(path.join(ROOT, relativePath), 'utf8');
    assert.match(source, /liff\.getIDToken\(\)/);
    assert.match(source, /idToken,/);
    assert.match(source, /rows\.push\(\{ date, billing, jobName/);
    assert.doesNotMatch(source, /new Date\(\)\.toISOString\(\)\.split\('T'\)\[0\]/);
  }
});

test('all frontend copies time out stalled requests and reject failed master responses', () => {
  for (const relativePath of ['gas/liff/index.html', 'liff-frontend/index.html']) {
    const source = fs.readFileSync(path.join(ROOT, relativePath), 'utf8');
    assert.match(source, /const REQUEST_TIMEOUT_MS = 20000/);
    assert.match(source, /controller\.abort\(\)/);
    assert.match(source, /if \(!json\.ok\)/);
    assert.match(source, /Array\.isArray\(json\.data\.jobList\)/);
  }
});

test('date utilities reject impossible dates and preserve weekdays', () => {
  const context = vm.createContext({ Date, String, Number, Error });
  vm.runInContext(
    fs.readFileSync(path.join(ROOT, 'gas/Utils.js'), 'utf8'),
    context,
    { filename: 'gas/Utils.js' }
  );
  assert.equal(context.getWeekday_('2026/08/16'), '日');
  assert.equal(context.parseDateParts_('2024/02/29').day, 29);
  assert.throws(() => context.parseDateParts_('2026/02/29'), /存在しない日付/);
  assert.throws(() => context.parseDateParts_('2026/13/01'), /存在しない日付/);
});

test('job master reader finds a row-2 header and never exposes it as data', () => {
  const context = vm.createContext({ String, Number, Error });
  vm.runInContext(
    fs.readFileSync(path.join(ROOT, 'gas/MasterSheet.js'), 'utf8'),
    context,
    { filename: 'gas/MasterSheet.js' }
  );
  const sheet = {
    getDataRange() {
      return {
        getValues() {
          return [
            ['', '', '', ''],
            ['案件名', '現場名', '単価', '項目'],
            ['OWL TIP', 'OWL', 50, 'TIP'],
            ['WINX CA', 'KTN', 1200, 'WINX']
          ];
        }
      };
    }
  };
  const jobs = context.readJobMasterRows_(sheet);
  assert.equal(jobs.length, 2);
  assert.equal(jobs[0].name, 'OWL TIP');
  assert.equal(jobs[0].sourceRow, 3);
  assert.equal(jobs.some(job => job.name === '案件名'), false);
});

test('formula capacity protection extends rows and restores calculated columns', () => {
  let maxRows = 10;
  const formulas = new Map();
  const required = { 2: '=ROW()', 5: '=RC[2]', 6: '=RC[1]', 13: '=RC[-1]', 14: '=RC[-2]*RC[-1]' };
  Object.entries(required).forEach(([column, formula]) => formulas.set(`8:${column}`, formula));
  let copied = false;
  const sheet = {
    getMaxRows() { return maxRows; },
    insertRowsAfter(_after, count) { maxRows += count; },
    getRange(row, column, rowCount = 1, columnCount = 1) {
      return {
        getFormulasR1C1() {
          return Array.from({ length: rowCount }, (_, r) =>
            Array.from({ length: columnCount }, (_, c) => formulas.get(`${row + r}:${column + c}`) || '')
          );
        },
        copyTo() { copied = true; },
        setFormulaR1C1(formula) {
          formulas.set(`${row}:${column}`, formula);
          return this;
        }
      };
    }
  };
  const context = vm.createContext({
    console,
    Date,
    JSON,
    Math,
    Object,
    String,
    Number,
    Error,
    SpreadsheetApp: { CopyPasteType: { PASTE_FORMAT: 'PASTE_FORMAT' } }
  });
  vm.runInContext(
    fs.readFileSync(path.join(ROOT, 'gas/MonthlySheet.js'), 'utf8'),
    context,
    { filename: 'gas/MonthlySheet.js' }
  );
  context.ensureInputSheetCapacity_(sheet, 11, 12);
  assert.equal(maxRows, 32);
  assert.equal(copied, true);
  for (const column of [2, 5, 6, 13, 14]) {
    assert.ok(formulas.get(`11:${column}`));
    assert.ok(formulas.get(`12:${column}`));
  }
});

test('appended rows pin store, category and price from validated master data', () => {
  const values = new Map();
  const formulas = new Map();
  const required = { 2: '=ROW()', 5: '=RC[2]', 6: '=RC[1]', 13: '=RC[-1]', 14: '=RC[-2]*RC[-1]' };
  Object.entries(required).forEach(([column, formula]) => formulas.set(`8:${column}`, formula));

  const sheet = {
    getLastRow() { return 8; },
    getMaxRows() { return 20; },
    insertRowsAfter() {},
    getRange(row, column, rowCount = 1, columnCount = 1) {
      const range = {
        getValues() {
          return Array.from({ length: rowCount }, (_, r) =>
            Array.from({ length: columnCount }, (_, c) => values.get(`${row + r}:${column + c}`) ?? '')
          );
        },
        getFormulasR1C1() {
          return Array.from({ length: rowCount }, (_, r) =>
            Array.from({ length: columnCount }, (_, c) => formulas.get(`${row + r}:${column + c}`) || '')
          );
        },
        copyTo() { return range; },
        clearDataValidations() { return range; },
        setNumberFormat() { return range; },
        setFormulaR1C1(formula) { formulas.set(`${row}:${column}`, formula); return range; },
        setValue(value) { values.set(`${row}:${column}`, value); return range; }
      };
      return range;
    }
  };
  const ss = { getSheetByName(name) { return name === '2.入力表' ? sheet : null; } };
  const context = vm.createContext({
    console, Date, JSON, Math, Object, String, Number, Error,
    SpreadsheetApp: {
      CopyPasteType: { PASTE_FORMAT: 'PASTE_FORMAT' },
      flush() {}
    },
    parseDate() { return new Date(Date.UTC(2026, 7, 16, 12)); },
    getWeekday_() { return '日'; }
  });
  vm.runInContext(
    fs.readFileSync(path.join(ROOT, 'gas/MonthlySheet.js'), 'utf8'),
    context,
    { filename: 'gas/MonthlySheet.js' }
  );
  context.appendRowsToInputSheet(ss, [{
    date: '2026/08/16', billing: 'KTN', category: 'WINX', jobName: 'WINX CA',
    name: 'TEST USER', qty: 2, detail: '', unitPrice: 1200
  }]);
  assert.equal(values.get('8:5'), 'KTN');
  assert.equal(values.get('8:6'), 'WINX');
  assert.equal(values.get('8:13'), 1200);
  assert.ok(formulas.get('8:14'));
});

test('notification retry and health-check safeguards are present', () => {
  const queueSource = fs.readFileSync(path.join(ROOT, 'gas/NotificationQueue.js'), 'utf8');
  const healthSource = fs.readFileSync(path.join(ROOT, 'gas/HealthCheck.js'), 'utf8');
  const monitoringSource = fs.readFileSync(path.join(ROOT, 'gas/Monitoring.js'), 'utf8');
  const codeSource = fs.readFileSync(path.join(ROOT, 'gas/Code.js'), 'utf8');
  assert.match(queueSource, /function retryPendingNotifications\(/);
  assert.match(queueSource, /NOTIFICATION_MAX_ATTEMPTS_/);
  assert.match(healthSource, /function runSystemHealthCheck_\(/);
  assert.match(monitoringSource, /function runOpportunisticMaintenance_\(/);
  assert.match(monitoringSource, /function reconcileReliabilityTriggers_\(/);
  assert.match(monitoringSource, /DEFAULT_ADMIN_ALERT_EMAIL_ = 'tryharddancers@gmail.com'/);
  assert.match(monitoringSource, /DEFAULT_ADMIN_ALERT_CC_ = 'h\.fujimoto@vexum-ai\.com'/);
  assert.match(codeSource, /runOpportunisticMaintenance_\((?:true|false)\)/);
});

test('system alerts are copied to the VEXUM support address', () => {
  const sent = [];
  const context = vm.createContext({
    console, Date, JSON, Math, Object, String, Number, Error,
    Logger: { log() {} },
    PropertiesService: {
      getScriptProperties() {
        return {
          getProperty(key) {
            if (key === 'ADMIN_ALERT_EMAIL') return 'tryharddancers@gmail.com';
            return null;
          }
        };
      }
    },
    Session: { getEffectiveUser() { return { getEmail() { return ''; } }; } },
    Utilities: {
      DigestAlgorithm: { SHA_256: 'SHA_256' },
      computeDigest() { return [1, 2, 3, 4]; }
    },
    CacheService: {
      getScriptCache() {
        return { get() { return null; }, put() {} };
      }
    },
    GmailApp: {
      sendEmail(to, subject, body, options) {
        sent.push({ to, subject, body, options });
      }
    }
  });
  vm.runInContext(
    fs.readFileSync(path.join(ROOT, 'gas/Monitoring.js'), 'utf8'),
    context,
    { filename: 'gas/Monitoring.js' }
  );

  context.reportSystemError_('test context', new Error('test failure'));

  assert.equal(sent.length, 1);
  assert.equal(sent[0].to, 'tryharddancers@gmail.com');
  assert.equal(sent[0].options.cc, 'h.fujimoto@vexum-ai.com');
});

test('trigger reconciliation creates missing triggers and removes duplicates', () => {
  const created = [];
  const deleted = [];
  const properties = new Map();
  const existing = [
    { getHandlerFunction() { return 'retryPendingNotifications'; } },
    { getHandlerFunction() { return 'retryPendingNotifications'; } }
  ];
  const context = vm.createContext({
    console, Date, JSON, Math, Object, String, Number, Error,
    Logger: { log() {} },
    PropertiesService: {
      getScriptProperties() {
        return {
          getProperty(key) { return properties.get(key) || null; },
          setProperty(key, value) { properties.set(key, value); }
        };
      }
    },
    Session: { getEffectiveUser() { return { getEmail() { return ''; } }; } },
    ScriptApp: {
      getProjectTriggers() { return existing; },
      deleteTrigger(trigger) {
        deleted.push(trigger);
        const index = existing.indexOf(trigger);
        if (index >= 0) existing.splice(index, 1);
      },
      newTrigger(handler) {
        return {
          timeBased() { return this; },
          everyMinutes() { return this; },
          everyDays() { return this; },
          atHour() { return this; },
          create() {
            created.push(handler);
            const trigger = { getHandlerFunction() { return handler; } };
            existing.push(trigger);
            return trigger;
          }
        };
      }
    }
  });
  vm.runInContext(
    fs.readFileSync(path.join(ROOT, 'gas/Monitoring.js'), 'utf8'),
    context,
    { filename: 'gas/Monitoring.js' }
  );

  const status = context.setupReliabilityTriggers();
  assert.deepEqual(created, ['dailySystemHealthCheck']);
  assert.equal(deleted.length, 1);
  assert.equal(properties.get('ADMIN_ALERT_EMAIL'), 'tryharddancers@gmail.com');
  assert.equal(status.retryPendingNotifications, true);
  assert.equal(status.dailySystemHealthCheck, true);
});

test('opportunistic maintenance is throttled and does not repeat work per request', () => {
  const properties = new Map();
  const counts = { retry: 0, health: 0, release: 0 };
  const context = vm.createContext({
    console, Date, JSON, Math, Object, String, Number, Error,
    Logger: { log() {} },
    PropertiesService: {
      getScriptProperties() {
        return {
          getProperty(key) { return properties.get(key) || null; },
          setProperty(key, value) { properties.set(key, value); }
        };
      }
    },
    Session: { getEffectiveUser() { return { getEmail() { return ''; } }; } },
    LockService: {
      getScriptLock() {
        return {
          tryLock() { return true; },
          releaseLock() { counts.release += 1; }
        };
      }
    }
  });
  vm.runInContext(
    fs.readFileSync(path.join(ROOT, 'gas/Monitoring.js'), 'utf8'),
    context,
    { filename: 'gas/Monitoring.js' }
  );
  context.retryPendingNotifications_ = function() { counts.retry += 1; };
  context.dailySystemHealthCheck = function() { counts.health += 1; };
  context.runOpportunisticMaintenance_(true);
  context.runOpportunisticMaintenance_(true);

  assert.deepEqual(counts, { retry: 1, health: 1, release: 4 });
  assert.equal(properties.get('ADMIN_ALERT_EMAIL'), 'tryharddancers@gmail.com');
  assert.ok(properties.get('OPS_LAST_RETRY_FALLBACK_AT'));
  assert.ok(properties.get('OPS_LAST_HEALTH_FALLBACK_AT'));
  assert.equal(properties.has('OPS_LAST_TRIGGER_AUDIT_AT'), false);
});

test('lightweight form-load maintenance skips the daily health check', () => {
  const properties = new Map();
  const counts = { retry: 0, health: 0 };
  const context = vm.createContext({
    console, Date, JSON, Math, Object, String, Number, Error,
    Logger: { log() {} },
    PropertiesService: {
      getScriptProperties() {
        return {
          getProperty(key) { return properties.get(key) || null; },
          setProperty(key, value) { properties.set(key, value); }
        };
      }
    },
    Session: { getEffectiveUser() { return { getEmail() { return ''; } }; } },
    LockService: {
      getScriptLock() {
        return { tryLock() { return true; }, releaseLock() {} };
      }
    }
  });
  vm.runInContext(
    fs.readFileSync(path.join(ROOT, 'gas/Monitoring.js'), 'utf8'),
    context,
    { filename: 'gas/Monitoring.js' }
  );
  context.retryPendingNotifications_ = function() { counts.retry += 1; return {}; };
  context.dailySystemHealthCheck = function() { counts.health += 1; };
  context.runOpportunisticMaintenance_(false);

  assert.deepEqual(counts, { retry: 1, health: 0 });
  assert.equal(properties.has('OPS_LAST_HEALTH_FALLBACK_AT'), false);
});

test('a failed recipient is queued without resending successful recipients', () => {
  const queued = [];
  const sentTo = [];
  const context = vm.createContext({
    console,
    Date,
    JSON,
    Math,
    Object,
    String,
    Number,
    Error,
    Logger: { log() {} },
    GmailApp: {
      sendEmail(email) {
        sentTo.push(email);
        if (email === 'fail@example.com') throw new Error('mail unavailable');
      }
    },
    getOutsourceContactEmailMap_() {
      return { OK: 'ok@example.com', FAIL: 'fail@example.com' };
    },
    enqueueNotificationRetry_(submissionId, spreadsheetId, personName, rowNumbers) {
      queued.push({ submissionId, spreadsheetId, personName, rowNumbers });
    },
    reportSystemError_() {},
    Utilities: { formatDate() { return '2026/08/16'; } }
  });
  vm.runInContext(
    fs.readFileSync(path.join(ROOT, 'gas/MonthlySheet.js'), 'utf8'),
    context,
    { filename: 'gas/MonthlySheet.js' }
  );
  context.readInsertedInputRows_ = function() {
    const base = {
      date: '2026/08/16', weekday: '日', venue: 'TEST', category: 'TIP',
      jobName: 'TEST TIP', detail: '', qty: 1, unitPrice: 1, total: 1
    };
    return [
      { ...base, name: 'OK' },
      { ...base, name: 'FAIL' }
    ];
  };
  const result = context.sendInputRowsNotification_(
    { getId() { return 'monthly-id'; } },
    [8, 9],
    'submission-mail-test'
  );
  assert.deepEqual([...result.sent], ['OK']);
  assert.deepEqual([...result.queued], ['FAIL']);
  assert.deepEqual(sentTo, ['ok@example.com', 'fail@example.com']);
  assert.equal(queued.length, 1);
  assert.equal(queued[0].personName, 'FAIL');
});

test('all recipients are queued when the email address master cannot be read', () => {
  const queued = [];
  const alerts = [];
  const context = vm.createContext({
    console, Date, JSON, Math, Object, String, Number, Error,
    Logger: { log() {} },
    getOutsourceContactEmailMap_() { throw new Error('contact master unavailable'); },
    enqueueNotificationRetry_(submissionId, spreadsheetId, personName, rowNumbers) {
      queued.push({ submissionId, spreadsheetId, personName, rowNumbers });
    },
    reportSystemError_(label, err) { alerts.push({ label, message: err.message }); },
    Utilities: { formatDate() { return '2026/08/16'; } }
  });
  vm.runInContext(
    fs.readFileSync(path.join(ROOT, 'gas/MonthlySheet.js'), 'utf8'),
    context,
    { filename: 'gas/MonthlySheet.js' }
  );
  context.readInsertedInputRows_ = function() {
    return [
      { name: 'A' },
      { name: 'B' }
    ];
  };

  const result = context.sendInputRowsNotification_(
    { getId() { return 'monthly-id'; } },
    [8, 9],
    'submission-contact-master-failure'
  );
  assert.deepEqual([...result.sent], []);
  assert.deepEqual([...result.queued], ['A', 'B']);
  assert.deepEqual(queued.map(item => item.personName), ['A', 'B']);
  assert.equal(alerts.at(-1).label, 'メール宛先一覧の取得');
});
