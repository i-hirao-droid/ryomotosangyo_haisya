// ▼▼▼ この部分にご自身のスプレッドシートIDを設定してください ▼▼▼
const SPREADSHEET_ID = '14YStto4gK0-OqsX_fjym6XovRlq2k0jahc1d5wBotrg';
// ▲▲▲ 設定はここまで ▲▲▲

/**
 * スプレッドシートが開かれたときに実行され、カスタムメニューを追加します。
 */
function onOpen(e) {
  SpreadsheetApp.getUi()
      .createMenu('📄 日報メニュー')
      .addItem('PDFダウンロード', 'showSidebar')
      .addToUi();
}

/**
 * HTMLファイルからサイドバーを生成して表示します。
 */
function showSidebar() {
  const html = HtmlService.createHtmlOutputFromFile('Sidebar')
      .setTitle('PDF ダウンロード')
      .setWidth(300);
  SpreadsheetApp.getUi().showSidebar(html);
}

/**
 * 指定されたシートをPDFとしてエクスポートします。
 * @param {string} sheetName - サイドバーから渡されるシート名
 */
function createPdf(sheetName) {
  try {
    const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = spreadsheet.getSheetByName(sheetName);

    if (!sheet) {
      throw new Error(`シート「${sheetName}」が見つかりませんでした。`);
    }

    const sheetId = sheet.getSheetId();
    
    // PDFエクスポート用のURL設定
    const url = `https://docs.google.com/spreadsheets/d/${SPREADSHEET_ID}/export?` +
      'format=pdf' +
      `&gid=${sheetId}` +
      '&size=A4' +
      '&portrait=false' + // 横向き
      '&fitw=true' +
      '&sheetnames=false' +
      '&printtitle=false' +
      '&gridlines=false';

    const token = ScriptApp.getOAuthToken();
    const response = UrlFetchApp.fetch(url, {
      headers: { 'Authorization': `Bearer ${token}` }
    });

    const blob = response.getBlob();
    const base64Data = Utilities.base64Encode(blob.getBytes());
    const fileName = `${sheetName}_${Utilities.formatDate(new Date(), "JST", "yyyy-MM-dd")}.pdf`;

    return {
      data: base64Data,
      filename: fileName
    };

  } catch (e) {
    return { error: e.message };
  }
}
