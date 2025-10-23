/**
 * SettleUp 匯入工具
 *
 * ⚠️ 使用說明：
 * 1. 從 SettleUp APP 匯出 CSV
 * 2. 將 CSV 匯入 Google Sheets（名稱必須是 "SettleUp_transactions"）
 * 3. 執行本檔案中的 importSettleUp() 函式
 * 4. 匯入後，系統會自動計算分帳，但不會處理「結算/轉帳」記錄
 * 5. 請手動在 SettleUp 中查看正確的餘額，然後在本系統中「調整餘額」
 *
 * 注意事項：
 * - 只會匯入 expense 類型的記錄
 * - transfer（結算/轉帳）會被跳過，不影響餘額計算
 * - 如需記錄不等額墊付，請在 CSV 的 Amount 欄位用分號分隔（例如：50;1590）
 */

/**
 * 匯入 SettleUp 資料
 * 執行前請確認：
 * 1. 同一資料夾中有名為 "SettleUp_transactions" 的 Google Sheets
 * 2. 該試算表包含從 SettleUp 匯出的 CSV 資料
 */
function importSettleUp() {
  const myName = '零幻'; // ⚠️ 請修改為你在 SettleUp 中的名字

  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const expensesSheet = ss.getSheetByName(CONFIG.SHEET_NAMES.EXPENSES);

  if (!expensesSheet) {
    ui.alert('❌ 錯誤', '找不到「支出記錄」工作表。請先執行「初始化系統」。', ui.ButtonSet.OK);
    return;
  }

  try {
    // 尋找 SettleUp_transactions 試算表
    const spreadsheetFile = DriveApp.getFileById(ss.getId());
    const parentFolders = spreadsheetFile.getParents();

    if (!parentFolders.hasNext()) {
      ui.alert('❌ 錯誤', '無法取得試算表所在資料夾', ui.ButtonSet.OK);
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
      ui.alert('❌ 錯誤', '找不到名為「SettleUp_transactions」的試算表。\n\n請確認：\n1. 已將 CSV 匯入 Google Sheets\n2. 試算表名稱為「SettleUp_transactions」\n3. 試算表與本系統在同一資料夾', ui.ButtonSet.OK);
      return;
    }

    const sheet = settleUpSpreadsheet.getSheets()[0];
    const data = sheet.getDataRange().getValues();

    const expenses = [];
    let skippedTransfers = 0;
    let errors = [];

    // 從第 2 行開始（跳過標題列）
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (!row[0] && !row[1]) continue; // 跳過空行

      try {
        const expense = parseSettleUpSheetRow(row, i + 1, myName);

        if (expense.type === 'transfer') {
          // 跳過 transfer 記錄（不納入餘額計算）
          skippedTransfers++;
          continue;
        }

        expenses.push(expense);
      } catch (e) {
        if (errors.length < 10) {
          Logger.log('第 ' + (i + 1) + ' 行錯誤: ' + e.message);
        }
        errors.push(`第 ${i + 1} 行：${e.message}`);
      }
    }

    if (expenses.length === 0) {
      ui.alert('⚠️ 無資料', '沒有可匯入的支出記錄。', ui.ButtonSet.OK);
      return;
    }

    // 寫入試算表
    const dataToWrite = expenses.map(exp => [
      exp.date,
      exp.item,
      exp.amount,
      exp.payer,
      exp.actualPayer,
      exp.yourPart,
      exp.partnerPart,
      exp.yourActualPaid || 0,
      exp.partnerActualPaid || 0,
      exp.category,
      false, // isRecurring
      '', // recurringDay
      new Date().getTime() + Math.random(), // ID
      'expense' // recordType
    ]);

    const lastRow = expensesSheet.getLastRow();
    expensesSheet.getRange(lastRow + 1, 1, dataToWrite.length, 14).setValues(dataToWrite);

    // 顯示結果
    let message = `✅ 匯入完成！\n\n` +
                  `✓ 成功匯入：${expenses.length} 筆支出記錄\n` +
                  `✓ 已跳過：${skippedTransfers} 筆結算/轉帳記錄\n\n` +
                  `⚠️ 重要提醒：\n` +
                  `請至 SettleUp 查看正確的餘額，\n` +
                  `然後在本系統使用「調整餘額」功能\n` +
                  `以同步初始餘額。`;

    if (errors.length > 0) {
      message += `\n\n⚠️ 錯誤記錄（${errors.length} 筆）：\n` + errors.slice(0, 5).join('\n');
      if (errors.length > 5) {
        message += `\n... 還有 ${errors.length - 5} 筆錯誤`;
      }
    }

    ui.alert('📥 匯入結果', message, ui.ButtonSet.OK);

  } catch (e) {
    ui.alert('❌ 匯入失敗', '發生錯誤：' + e.message, ui.ButtonSet.OK);
    Logger.log('匯入錯誤：' + e.toString());
  }
}

/**
 * 資料檢查工具
 * 用於檢查匯入後的資料完整性
 */
function checkImportedData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEET_NAMES.EXPENSES);

  if (!sheet) {
    Logger.log('❌ 找不到工作表');
    return;
  }

  const data = sheet.getDataRange().getValues();
  const headers = data[0];

  const recordTypeIdx = headers.indexOf('記錄類型');
  const yourPartIdx = headers.indexOf('你的部分');
  const partnerPartIdx = headers.indexOf('對方的部分');
  const yourActualIdx = headers.indexOf('你實付');
  const partnerActualIdx = headers.indexOf('對方實付');

  let expenseCount = 0;
  let settlementCount = 0;
  let emptyPartCount = 0;
  let zeroPaidCount = 0;

  for (let i = 1; i < data.length; i++) {
    const recordType = data[i][recordTypeIdx];

    if (recordType === 'expense') {
      expenseCount++;

      const yourPart = data[i][yourPartIdx];
      const partnerPart = data[i][partnerPartIdx];
      const yourActual = data[i][yourActualIdx];
      const partnerActual = data[i][partnerActualIdx];

      if (yourPart === '' || yourPart === null || partnerPart === '' || partnerPart === null) {
        emptyPartCount++;
      }

      if (yourActual === 0 && partnerActual === 0) {
        zeroPaidCount++;
      }
    } else if (recordType === 'settlement') {
      settlementCount++;
    }
  }

  Logger.log('========== 資料檢查結果 ==========');
  Logger.log('總記錄數: ' + (data.length - 1));
  Logger.log('支出記錄: ' + expenseCount);
  Logger.log('結算記錄: ' + settlementCount);
  Logger.log('分帳為空: ' + emptyPartCount);
  Logger.log('實付為 0: ' + zeroPaidCount);
  Logger.log('=================================');
}
