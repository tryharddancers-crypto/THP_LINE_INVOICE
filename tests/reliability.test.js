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
          { name: 'OWL TIP', billing: 'OWL', unitPrice: 50 },
          { name: 'WINX CA', billing: 'OWL', unitPrice: 1000 },
          { name: 'WINX CA', billing: 'KTN', unitPrice: 1000 }
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
    date: '2026/08/16', name: 'TEST USER', jobName: 'OWL TIP', qty: 2, detail: '', unitPrice: 999999
  }]);
  assert.equal(rows[0].unitPrice, 50);
  assert.equal(rows[0].date, '2026/08/16');
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
  }
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

test('notification retry and health-check safeguards are present', () => {
  const queueSource = fs.readFileSync(path.join(ROOT, 'gas/NotificationQueue.js'), 'utf8');
  const healthSource = fs.readFileSync(path.join(ROOT, 'gas/HealthCheck.js'), 'utf8');
  assert.match(queueSource, /function retryPendingNotifications\(/);
  assert.match(queueSource, /NOTIFICATION_MAX_ATTEMPTS_/);
  assert.match(healthSource, /function runSystemHealthCheck_\(/);
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
