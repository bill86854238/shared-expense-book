/**
 * 檢查 SettleUp_transactions 試算表內容
 */
function inspectSettleUpData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const spreadsheetFile = DriveApp.getFileById(ss.getId());
  const parentFolders = spreadsheetFile.getParents();

  if (!parentFolders.hasNext()) {
    Logger.log('❌ 無法取得資料夾');
    return;
  }

  const folder = parentFolders.next();
  let settleUpSpreadsheet = null;
  const files = folder.getFilesByType(MimeType.GOOGLE_SHEETS);

  while (files.hasNext()) {
    const file = files.next();
    const fileName = file.getName();

    if (fileName.toLowerCase() === 'settleup_transactions') {
      settleUpSpreadsheet = SpreadsheetApp.openById(file.getId());
      break;
    }
  }

  if (!settleUpSpreadsheet) {
    Logger.log('❌ 找不到 SettleUp_transactions 試算表');
    return;
  }

  const sheet = settleUpSpreadsheet.getSheets()[0];
  const data = sheet.getDataRange().getValues();

  Logger.log('========== SettleUp 試算表分析 ==========');
  Logger.log('總行數: ' + data.length);
  Logger.log('');
  Logger.log('標題列: ' + data[0].join(' | '));
  Logger.log('');

  // 統計 Type 欄位
  let expenseCount = 0;
  let transferCount = 0;
  const typeIndex = 10; // Type 欄位在第 11 欄（索引 10）

  for (let i = 1; i < data.length; i++) {
    const type = String(data[i][typeIndex] || '').trim();
    if (type === 'transfer') {
      transferCount++;
      if (transferCount <= 5) {
        Logger.log('Transfer 範例 ' + transferCount + ': ' +
          data[i][0] + ' paid ' + data[i][1] + ' to ' + data[i][3] +
          ' (' + data[i][5] + ')');
      }
    } else if (type === 'expense') {
      expenseCount++;
    }
  }

  Logger.log('');
  Logger.log('支出記錄 (expense): ' + expenseCount);
  Logger.log('結算記錄 (transfer): ' + transferCount);
  Logger.log('');

  // 檢查前 5 筆支出記錄的分帳情況
  Logger.log('--- 前 5 筆支出記錄 ---');
  let count = 0;
  for (let i = 1; i < data.length && count < 5; i++) {
    const type = String(data[i][typeIndex] || '').trim();
    if (type === 'expense' || type === '') {
      count++;
      Logger.log('記錄 ' + count + ':');
      Logger.log('  Who paid: ' + data[i][0]);
      Logger.log('  Amount: ' + data[i][1]);
      Logger.log('  For whom: ' + data[i][3]);
      Logger.log('  Split amounts: ' + data[i][4]);
      Logger.log('  Purpose: ' + data[i][5]);
      Logger.log('  Type: ' + type);
      Logger.log('');
    }
  }

  Logger.log('========================================');
}

/**
 * 檢查匯入後的支出記錄工作表內容
 */
function inspectExpensesSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEET_NAMES.EXPENSES);

  if (!sheet) {
    Logger.log('❌ 找不到支出記錄工作表');
    return;
  }

  const data = sheet.getDataRange().getValues();
  const headers = data[0];

  Logger.log('========== 支出記錄工作表分析 ==========');
  Logger.log('總行數: ' + data.length);
  Logger.log('欄位數: ' + headers.length);
  Logger.log('標題列: ' + headers.join(' | '));
  Logger.log('');

  // 搜尋結算記錄
  let settlementCount = 0;
  const recordTypeIndex = headers.indexOf('記錄類型');
  const categoryIndex = headers.indexOf('分類');
  const itemIndex = headers.indexOf('項目');

  for (let i = 1; i < data.length; i++) {
    const recordType = data[i][recordTypeIndex];
    const category = data[i][categoryIndex];
    const item = String(data[i][itemIndex] || '');

    if (recordType === 'settlement' || category === '結算' || item.includes('💰結算')) {
      settlementCount++;
      if (settlementCount <= 5) {
        Logger.log('結算記錄 ' + settlementCount + ': ' + item +
          ' | 金額: ' + data[i][headers.indexOf('金額')] +
          ' | 記錄類型: ' + recordType);
      }
    }
  }

  Logger.log('');
  Logger.log('找到結算記錄: ' + settlementCount + ' 筆');
  Logger.log('');

  // 檢查前 5 筆支出記錄
  Logger.log('--- 前 5 筆支出記錄 ---');
  for (let i = 1; i <= 5 && i < data.length; i++) {
    Logger.log('記錄 ' + i + ':');
    Logger.log('  項目: ' + data[i][itemIndex]);
    Logger.log('  金額: ' + data[i][headers.indexOf('金額')]);
    Logger.log('  付款人: ' + data[i][headers.indexOf('付款人')]);
    Logger.log('  你的部分: ' + data[i][headers.indexOf('你的部分')]);
    Logger.log('  對方的部分: ' + data[i][headers.indexOf('對方的部分')]);
    Logger.log('  你實付: ' + data[i][headers.indexOf('你實付')]);
    Logger.log('  對方實付: ' + data[i][headers.indexOf('對方實付')]);
    Logger.log('  記錄類型: ' + data[i][recordTypeIndex]);
    Logger.log('');
  }

  Logger.log('========================================');
}
