/**
 * ChatWorkにメッセージを送信する共通関数
 */
function sendChatworkMessage(body) {
  const props = PropertiesService.getScriptProperties();
  const apiToken = props.getProperty('CHATWORK_API_KEY');  // スクリプト プロパティに値を設定して下さい
  const roomId = 40851424;

  if (!apiToken) {
    console.error('APIキーが設定されていません。');
    return;
  }

  const options = {
    method: 'post',
    headers: { 'X-ChatWorkToken': apiToken },
    payload: { body: body }
  };

  UrlFetchApp.fetch(`https://api.chatwork.com/v2/rooms/${roomId}/messages`, options);
}