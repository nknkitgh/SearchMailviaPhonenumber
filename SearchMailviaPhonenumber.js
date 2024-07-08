function getEmailFromPhoneNumber() {
  const apiKey = 'AIzaSyD-aE8cpQrP5ydtnvmAyrqX135Lnq4ct4M'; // ここにGoogle APIキーを入力
  const spreadsheetId = '1aiJNCCSxRqmgrnp6IaSzpz7gkSiw7HRlg3hZhBelLuo'; // ここにスプレッドシートIDを入力
  const sheetName = 'Sheet1'; // ここにシート名を入力

  const sheet = SpreadsheetApp.openById(spreadsheetId).getSheetByName(sheetName);
  const data = sheet.getDataRange().getValues();
  
  for (let i = 1; i < data.length; i++) { // 1行目はヘッダーのため2行目から開始
    const phoneNumber = data[i][0]; // 電話番号が1列目にあると仮定
    if (phoneNumber) {
      const email = searchEmailByPhoneNumber(apiKey, phoneNumber);
      if (email) {
        sheet.getRange(i + 1, 2).setValue(email); // メールアドレスを2列目に設定
      }
    }
  }
}

function searchEmailByPhoneNumber(apiKey, phoneNumber) {
  const searchQuery = `${phoneNumber} メールアドレス`;
  const apiUrl = `https://www.googleapis.com/customsearch/v1?q=${encodeURIComponent(searchQuery)}&key=${apiKey}&cx=YOUR_SEARCH_ENGINE_ID`;

  const response = UrlFetchApp.fetch(apiUrl);
  const results = JSON.parse(response.getContentText());

  if (results.items && results.items.length > 0) {
    const firstResultUrl = results.items[0].link;
    const pageContent = UrlFetchApp.fetch(firstResultUrl).getContentText();
    const emailRegex = /[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}/;
    const emailMatch = pageContent.match(emailRegex);
    return emailMatch ? emailMatch[0] : null;
  }
  return null;
}
