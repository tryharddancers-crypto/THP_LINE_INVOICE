/**
 * LINE Messaging APIでプッシュメッセージを送信する
 * @param {string} userId LINEユーザーID
 * @param {string} message 送信するテキスト
 */
function sendLineMessage(userId, message) {
  const token = PropertiesService.getScriptProperties().getProperty('LINE_CHANNEL_ACCESS_TOKEN');
  if (!token) throw new Error('LINE_CHANNEL_ACCESS_TOKEN is not set');

  const payload = {
    to: userId,
    messages: [{ type: 'text', text: message }]
  };

  const options = {
    method: 'post',
    contentType: 'application/json',
    headers: { Authorization: `Bearer ${token}` },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  const response = UrlFetchApp.fetch('https://api.line.me/v2/bot/message/push', options);
  const code = response.getResponseCode();
  if (code !== 200) {
    throw new Error(`LINE push failed: ${code} ${response.getContentText()}`);
  }
}
