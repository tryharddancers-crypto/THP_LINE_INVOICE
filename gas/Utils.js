/**
 * 西暦DateオブジェクトをGENGO+年+月の文字列に変換する
 * @param {Date} date
 * @returns {string} 例: "令和8年4月"
 */
function toWareki(date) {
  const y = date.getFullYear();
  const m = date.getMonth() + 1;
  const reiwaStart = new Date('2019-05-01');
  const heiseiStart = new Date('1989-01-08');

  if (date >= reiwaStart) {
    const reiwaYear = y - 2018;
    const label = reiwaYear === 1 ? '令和元年' : `令和${reiwaYear}年`;
    return `${label}${m}月`;
  } else if (date >= heiseiStart) {
    const heiseiYear = y - 1988;
    const label = heiseiYear === 1 ? '平成元年' : `平成${heiseiYear}年`;
    return `${label}${m}月`;
  }
  return `${y}年${m}月`;
}

/**
 * 月次スプレッドシートのファイル名を生成する
 * @param {Date} date
 * @returns {string} 例: "【ダンサー】令和8年4月分 外注連絡表"
 */
function getMonthlyFileName(date) {
  return `【ダンサー】${toWareki(date)}分 外注連絡表`;
}

/**
 * YYYY/MM/DD または YYYY-MM-DD を日付要素に分解する
 * @param {string} dateStr
 * @returns {{year: number, month: number, day: number}}
 */
function parseDateParts_(dateStr) {
  const match = String(dateStr || '').trim().match(/^(\d{4})[\/-](\d{1,2})[\/-](\d{1,2})$/);
  if (!match) {
    throw new Error('日付の形式が正しくありません: ' + dateStr);
  }

  const year = Number(match[1]);
  const month = Number(match[2]);
  const day = Number(match[3]);
  const check = new Date(Date.UTC(year, month - 1, day));
  if (
    check.getUTCFullYear() !== year
    || check.getUTCMonth() !== month - 1
    || check.getUTCDate() !== day
  ) {
    throw new Error('存在しない日付です: ' + dateStr);
  }

  return { year: year, month: month, day: day };
}

/**
 * 日付文字列をタイムゾーンの影響を受けにくいDateに変換する
 * @param {string} dateStr
 * @returns {Date}
 */
function parseDate(dateStr) {
  const parts = parseDateParts_(dateStr);
  return new Date(Date.UTC(parts.year, parts.month - 1, parts.day, 12));
}

/**
 * 日付文字列から曜日を返す
 * @param {string} dateStr
 * @returns {string}
 */
function getWeekday_(dateStr) {
  const parts = parseDateParts_(dateStr);
  const dayNames = ['日', '月', '火', '水', '木', '金', '土'];
  return dayNames[new Date(Date.UTC(parts.year, parts.month - 1, parts.day)).getUTCDay()];
}

/**
 * DateをYYYY/MM/DD形式にフォーマットする
 * @param {Date} date
 * @returns {string}
 */
function formatDate(date) {
  const y = date.getFullYear();
  const m = String(date.getMonth() + 1).padStart(2, '0');
  const d = String(date.getDate()).padStart(2, '0');
  return `${y}/${m}/${d}`;
}

/**
 * 現在の日時をYYYY/MM/DD HH:MM:SS形式で返す
 * @returns {string}
 */
function nowString() {
  const now = new Date();
  const y = now.getFullYear();
  const mo = String(now.getMonth() + 1).padStart(2, '0');
  const d = String(now.getDate()).padStart(2, '0');
  const h = String(now.getHours()).padStart(2, '0');
  const mi = String(now.getMinutes()).padStart(2, '0');
  const s = String(now.getSeconds()).padStart(2, '0');
  return `${y}/${mo}/${d} ${h}:${mi}:${s}`;
}
