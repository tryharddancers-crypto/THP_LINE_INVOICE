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
