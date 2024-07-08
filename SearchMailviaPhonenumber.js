function getEmailFromPhoneNumber() {
  const apiKey = 'AIzaSyD-aE8cpQrP5ydtnvmAyrqX135Lnq4ct4M'; // ここにGoogle APIキーを入力
  const cx = '8050577eb54c44127'; // ここにカスタム検索エンジンIDを入力
  const spreadsheetId = '1aiJNCCSxRqmgrnp6IaSzpz7gkSiw7HRlg3hZhBelLuo'; // ここにスプレッドシートIDを入力
  const sheetName = 'Sheet1'; // ここにシート名を入力

  const sheet = SpreadsheetApp.openById(spreadsheetId).getSheetByName(sheetName);
  const data = sheet.getDataRange().getValues();
  
  for (let i = 1; i < data.length; i++) { // 1行目はヘッダーのため2行目から開始
    const phoneNumber = data[i][0]; // 電話番号が1列目にあると仮定
    if (phoneNumber) {
      try {
        const email = searchEmailByPhoneNumber(apiKey, cx, phoneNumber);
        if (email) {
          sheet.getRange(i + 1, 2).setValue(email); // メールアドレスを2列目に設定
        } else {
          Logger.log(`メールアドレスが見つかりませんでした: ${phoneNumber}`);
        }
      } catch (error) {
        Logger.log(`エラーが発生しました: ${error.message} - 電話番号: ${phoneNumber}`);
      }
    }
  }
}

function searchEmailByPhoneNumber(apiKey, cx, phoneNumber) {
  const searchQuery = `${phoneNumber} メールアドレス`;
  const apiUrl = `https://www.googleapis.com/customsearch/v1?q=${encodeURIComponent(searchQuery)}&key=${apiKey}&cx=${cx}`;

  const response = UrlFetchApp.fetch(apiUrl, {muteHttpExceptions: true});
  const responseCode = response.getResponseCode();
  if (responseCode !== 200) {
    throw new Error(`Google検索APIのリクエストに失敗しました: ${responseCode} - ${response.getContentText()}`);
  }
  
  const results = JSON.parse(response.getContentText());
  if (results.items && results.items.length > 0) {
    for (let i = 0; i < Math.min(results.items.length, 3); i++) {
      const firstResultUrl = results.items[i].link;
      const pageContent = UrlFetchApp.fetch(firstResultUrl, {muteHttpExceptions: true}).getContentText();
      const emailRegex = /[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}/g;
      const emailMatches = pageContent.match(emailRegex);
      if (emailMatches && emailMatches.length > 0) {
        return emailMatches[0];
      }
    }
  }
  return null;
}
