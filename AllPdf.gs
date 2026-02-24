/**
 * メニューを追加
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('日報作成')
    .addItem('日付選択と出力 (一括)', 'showDateSelectionDialog')
    .addSeparator()
    .addItem('PDFダウンロード (個別)', 'showSidebar')
    .addToUi();
}

/* ==============================================
   ▼ 機能1：日付選択による一括処理
   ============================================== */

/**
 * 日付選択ダイアログを表示
 */
function showDateSelectionDialog() {
  const html = HtmlService.createHtmlOutputFromFile('Dialog')
    .setWidth(350)
    .setHeight(350);
  SpreadsheetApp.getUi().showModalDialog(html, '処理する日付を選択してください');
}

/**
 * メイン処理：選択された日付のデータを処理してPDF(ZIP)を作成し、ドライブのURLを返す
 */
function processDailyReports(selectedDateStr) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // 日付の整形
  selectedDateStr = selectedDateStr.trim().replace(/-/g, '/');

  const scheduleSheet = ss.getSheetByName('配車予定');
  const workSheet = ss.getSheetByName('日報 作業用');
  const refSheet = ss.getSheetByName('日報 転記用');
  const cylinderSheet = ss.getSheetByName('シリンダー日報');
  const normalReportSheet = ss.getSheetByName('日報');

  if (!scheduleSheet || !workSheet || !refSheet) {
    throw new Error('必要なシートが見つかりません。シート名を確認してください。');
  }

  const lastRow = scheduleSheet.getLastRow();
  if (lastRow < 2) return { success: false, message: "配車予定にデータがありません" };
  
  const dataRange = scheduleSheet.getRange(2, 2, lastRow - 1, 3).getValues();

  let blobs = [];
  let processedCount = 0;
  let dateMatchCount = 0;

  for (let i = 0; i < dataRange.length; i++) {
    const rowDate = dataRange[i][0];
    const vehicleNo = dataRange[i][1];
    const driverId = dataRange[i][2];

    let rowDateStr = "";
    if (rowDate instanceof Date) {
      rowDateStr = Utilities.formatDate(rowDate, "JST", "yyyy/MM/dd");
    } else {
      rowDateStr = Utilities.formatDate(new Date(rowDate), "JST", "yyyy/MM/dd");
    }

    if (rowDateStr === selectedDateStr) {
      dateMatchCount++;

      // 転記
      workSheet.getRange('B3').setValue(new Date(selectedDateStr));
      workSheet.getRange('C3').setValue(vehicleNo);
      workSheet.getRange('E3').setValue(driverId);
      SpreadsheetApp.flush();

      const driverName = workSheet.getRange('D3').getValue();
      const sColumnValue = findSColumnValue(refSheet, vehicleNo); 

      const fileDate = selectedDateStr.replace(/\//g, ''); 
      const fileNameBase = `${fileDate}_${vehicleNo}_${driverName}`;
      
      let isCreated = false;

      // ▼▼▼ 修正箇所：両方とも false (横向き) に設定 ▼▼▼

      // 1. シリンダー日報 -> 横向き (false)
      if (sColumnValue.includes('シリンダー')) {
        blobs.push(createPdfBlob(ss, cylinderSheet, fileNameBase + '_シリンダー日報', false));
        isCreated = true;
      }

      // 2. 通常の日報 -> 横向き (false)
      if (sColumnValue.includes('スクラップ') || sColumnValue.includes('サーバー') || sColumnValue.includes('バルク')) {
        blobs.push(createPdfBlob(ss, normalReportSheet, fileNameBase + '_日報', false));
        isCreated = true;
      }

      if (isCreated) {
        processedCount++;
      }
    }
  }

  if (dateMatchCount === 0) {
    return { success: false, message: `日付「${selectedDateStr}」のデータが配車予定に見つかりませんでした。` };
  }

  if (blobs.length === 0) {
    return { success: false, message: `データは見つかりましたが、PDF出力条件（S列の値）に合致するものがありませんでした。` };
  }

  const zipName = `日報一括出力_${selectedDateStr.replace(/\//g, '')}.zip`;
  const zipBlob = Utilities.zip(blobs, zipName);
  
  // マイドライブに保存してURLを返す
  const zipFile = DriveApp.createFile(zipBlob);
  const driveUrl = zipFile.getUrl();
  
  return {
    success: true,
    message: `${processedCount}件のPDFを作成しました。`,
    driveUrl: driveUrl,
    filename: zipName
  };
}

/**
 * 補助関数: 車番(Q列)をもとにS列の値を取得する
 */
function findSColumnValue(sheet, key) {
  const startRow = 9; 
  const lastRow = sheet.getLastRow();
  
  if (lastRow < startRow) return "";
  
  const vehicleNoValues = sheet.getRange(startRow, 17, lastRow - startRow + 1, 1).getValues(); 
  const sValues = sheet.getRange(startRow, 19, lastRow - startRow + 1, 1).getValues();

  const strKey = String(key).trim();

  for (let i = 0; i < vehicleNoValues.length; i++) {
    if (String(vehicleNoValues[i][0]).trim() == strKey) {
      return String(sValues[i][0]);
    }
  }
  return "";
}

/**
 * PDF生成関数（向き指定対応）
 * @param {boolean} isPortrait - trueなら縦、falseなら横
 */
function createPdfBlob(ss, sheet, fileName, isPortrait) {
  const spreadSheetId = ss.getId();
  const sheetId = sheet.getSheetId();
  
  // 引数が省略された場合は縦(true)にする
  const portraitSetting = (isPortrait === undefined) ? true : isPortrait;

  const url = `https://docs.google.com/spreadsheets/d/${spreadSheetId}/export?` +
    'format=pdf&size=A4&fitw=true&sheetnames=false&printtitle=false&pagenumbers=false&gridlines=false&fzr=false' +
    `&portrait=${portraitSetting}` + 
    `&gid=${sheetId}`;
  
  const token = ScriptApp.getOAuthToken();
  const response = UrlFetchApp.fetch(url, {
    headers: { 'Authorization': 'Bearer ' + token }
  });
  
  return response.getBlob().setName(fileName + '.pdf');
}


/* ==============================================
   ▼ 機能2：サイドバーからの個別出力
   ============================================== */

function showSidebar() {
  const html = HtmlService.createHtmlOutputFromFile('Sidebar')
      .setTitle('PDF ダウンロード')
      .setWidth(300);
  SpreadsheetApp.getUi().showSidebar(html);
}

function createManualPdf(sheetName) {
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const spreadSheetId = spreadsheet.getId();
    const sheet = spreadsheet.getSheetByName(sheetName);

    if (!sheet) {
      throw new Error(`シート「${sheetName}」が見つかりませんでした。`);
    }

    const sheetId = sheet.getSheetId();
    
    const url = `https://docs.google.com/spreadsheets/d/${spreadSheetId}/export?` +
      'format=pdf' +
      `&gid=${sheetId}` +
      '&size=A4' +
      '&portrait=false' + 
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
