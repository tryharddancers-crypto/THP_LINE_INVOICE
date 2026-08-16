var MAX_SUBMISSION_ROWS_ = 100;
var MAX_DETAIL_LENGTH_ = 500;
var MAX_QUANTITY_ = 1000000;

/** Create an error that is safe to show to the form user. */
function createClientError_(message) {
  const err = new Error(message);
  err.isClientError = true;
  return err;
}

/**
 * Verify the LIFF ID token with LINE and return the authenticated LINE user ID.
 * The userId sent in the JSON body is deliberately ignored.
 */
function verifyLineIdToken_(idToken) {
  const token = String(idToken || '').trim();
  if (!token) {
    throw createClientError_('LINEの認証情報を確認できませんでした。フォームを開き直して再度お試しください。');
  }

  const props = PropertiesService.getScriptProperties();
  const liffId = String(props.getProperty('LIFF_ID') || '').trim();
  const channelId = String(props.getProperty('LINE_CHANNEL_ID') || liffId.split('-')[0] || '').trim();
  if (!/^\d+$/.test(channelId)) {
    throw new Error('LINE_CHANNEL_IDまたはLIFF_IDの設定が正しくありません');
  }

  const response = UrlFetchApp.fetch('https://api.line.me/oauth2/v2.1/verify', {
    method: 'post',
    contentType: 'application/x-www-form-urlencoded',
    payload: { id_token: token, client_id: channelId },
    muteHttpExceptions: true
  });
  const responseCode = response.getResponseCode();
  let verified = null;
  try {
    verified = JSON.parse(response.getContentText());
  } catch (err) {
    verified = null;
  }

  if (responseCode >= 500) {
    throw new Error('LINE認証サーバーへの接続に失敗しました: HTTP ' + responseCode);
  }
  if (
    responseCode !== 200
    || !verified
    || !verified.sub
    || String(verified.aud) !== channelId
    || (verified.exp && Number(verified.exp) * 1000 <= Date.now())
  ) {
    throw createClientError_('LINEの認証期限が切れています。フォームを開き直して再度お試しください。');
  }

  return String(verified.sub);
}

/**
 * Normalize and validate form rows against the current master data.
 * Unit prices are always replaced by the server-side master value.
 */
function validateSubmissionRows_(rows) {
  if (!Array.isArray(rows) || rows.length === 0) {
    throw createClientError_('入力内容がありません');
  }
  if (rows.length > MAX_SUBMISSION_ROWS_) {
    throw createClientError_('一度に送信できる入力は' + MAX_SUBMISSION_ROWS_ + '件までです');
  }

  const master = getMasterData();
  const dancerSet = {};
  (master.dancerNames || []).forEach(function(name) {
    dancerSet[String(name).trim()] = true;
  });

  const jobsByKey = {};
  const jobsByName = {};
  (master.jobList || []).forEach(function(job) {
    const name = String(job.name || '').trim();
    const billing = String(job.billing || '').trim();
    if (!name) return;
    if (!billing) {
      throw new Error('案件マスタの店舗名が空欄です: ' + name);
    }
    const price = Number(job.unitPrice);
    if (!Number.isFinite(price) || price < 0) {
      throw new Error('案件マスタの単価が不正です: ' + name);
    }
    const key = billing + '\u0000' + name;
    if (Object.prototype.hasOwnProperty.call(jobsByKey, key)) {
      throw new Error('案件マスタに同じ店舗・商品が重複しています: ' + billing + ' / ' + name);
    }
    const normalizedJob = {
      name: name,
      billing: billing,
      category: String(job.category || '').trim(),
      unitPrice: price
    };
    jobsByKey[key] = normalizedJob;
    if (!jobsByName[name]) jobsByName[name] = [];
    jobsByName[name].push(normalizedJob);
  });

  let targetMonth = '';
  return rows.map(function(row, index) {
    const itemNo = index + 1;
    let parts;
    try {
      parts = parseDateParts_(row && row.date);
    } catch (err) {
      throw createClientError_(itemNo + '件目の日程を正しく入力してください。');
    }
    const date = [
      String(parts.year).padStart(4, '0'),
      String(parts.month).padStart(2, '0'),
      String(parts.day).padStart(2, '0')
    ].join('/');
    const month = String(parts.year) + '-' + String(parts.month).padStart(2, '0');
    if (!targetMonth) targetMonth = month;
    if (month !== targetMonth) {
      throw createClientError_('異なる月の日付は一度に送信できません。月ごとに分けて送信してください。');
    }

    const name = String(row && row.name || '').trim();
    if (!name || !dancerSet[name]) {
      throw createClientError_(itemNo + '件目の担当者が現在のマスタにありません。フォームを開き直してください。');
    }

    const jobName = String(row && row.jobName || '').trim();
    const matchingJobs = jobsByName[jobName] || [];
    if (!jobName || matchingJobs.length === 0) {
      throw createClientError_(itemNo + '件目の商品が現在の案件マスタにありません。フォームを開き直してください。');
    }

    let billing = String(row && row.billing || '').trim();
    if (!billing) {
      if (matchingJobs.length === 1) {
        // Older cached forms did not send the store. Keep unambiguous items compatible.
        billing = matchingJobs[0].billing;
      } else {
        throw createClientError_(itemNo + '件目の店舗を確認できません。フォームを開き直して店舗から選び直してください。');
      }
    }
    const selectedJob = jobsByKey[billing + '\u0000' + jobName];
    if (!selectedJob) {
      throw createClientError_(itemNo + '件目の商品は選択された店舗の案件マスタにありません。フォームを開き直してください。');
    }

    const qty = Number(row && row.qty);
    if (!Number.isFinite(qty) || qty <= 0 || qty > MAX_QUANTITY_) {
      throw createClientError_(itemNo + '件目の数量を正しく入力してください。');
    }

    const detail = String(row && row.detail || '').trim();
    if (detail.length > MAX_DETAIL_LENGTH_) {
      throw createClientError_(itemNo + '件目の詳細は' + MAX_DETAIL_LENGTH_ + '文字以内で入力してください。');
    }

    return {
      date: date,
      jobName: jobName,
      billing: selectedJob.billing,
      category: selectedJob.category,
      name: name,
      qty: qty,
      detail: detail,
      unitPrice: selectedJob.unitPrice
    };
  });
}
