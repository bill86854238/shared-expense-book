/**
 * 情侶共同記帳系統 - Google Apps Script
 * 完整版本
 */

// ==================== 設定區 ====================
const CONFIG = {
  SHEET_NAMES: {
    EXPENSES: '支出記錄',
    RECURRING: '週期設定',
    SETTINGS: '設定',
    CATEGORIES: '分類設定'
  },
  COLORS: {
    HEADER: '#8b5cf6',
    YOUR: '#dbeafe',
    PARTNER: '#fce7f3',
    BOTH: '#f3e8ff'
  }
};

// ==================== 初始化函數 ====================

function initializeSpreadsheet() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // 檢查是否已有支出記錄工作表
  const expensesSheet = ss.getSheetByName(CONFIG.SHEET_NAMES.EXPENSES);

  if (expensesSheet) {
    const dataCount = expensesSheet.getLastRow() - 1; // 扣掉標題列

    if (dataCount > 0) {
      // 已有資料，拒絕初始化
      ui.alert(
        '⚠️ 系統已存在資料',
        `目前有 ${dataCount} 筆支出記錄。\n\n` +
        '初始化功能僅供「首次使用」！\n\n' +
        '如果要升級資料結構（新增欄位），請使用：\n' +
        '📊 記帳系統 → 🔄 升級資料結構\n\n' +
        '如果要清空重置，請使用：\n' +
        '📊 記帳系統 → ⚠️ 重置系統（危險）',
        ui.ButtonSet.OK
      );
      return;
    }
  }

  // 首次初始化
  createExpensesSheet(ss);
  createRecurringSheet(ss);
  createSettingsSheet(ss);
  createCategoriesSheet(ss);
  setupTriggers();

  ui.alert('✅ 初始化完成！\n\n已建立：\n1. 支出記錄\n2. 週期設定\n3. 設定\n4. 分類設定\n\n並設定每日自動執行週期事件。');
}

function createExpensesSheet(ss) {
  let sheet = ss.getSheetByName(CONFIG.SHEET_NAMES.EXPENSES);
  if (!sheet) {
    // 只有不存在時才建立新的
    sheet = ss.insertSheet(CONFIG.SHEET_NAMES.EXPENSES);

    const headers = ['日期', '項目', '金額', '付款人', '實際付款人', '你的部分', '對方的部分', '你實付', '對方實付', '分類', '是否週期', '週期日期', 'ID', '記錄類型'];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

    sheet.getRange(1, 1, 1, headers.length)
      .setBackground(CONFIG.COLORS.HEADER)
      .setFontColor('#ffffff')
      .setFontWeight('bold')
      .setHorizontalAlignment('center');

    const widths = [100, 150, 100, 100, 100, 100, 100, 100, 100, 80, 80, 80, 120, 100];
    widths.forEach((width, i) => sheet.setColumnWidth(i + 1, width));

    sheet.setFrozenRows(1);
  }
  // 如果已存在，不做任何事（保護資料）
}

function createRecurringSheet(ss) {
  let sheet = ss.getSheetByName(CONFIG.SHEET_NAMES.RECURRING);
  if (!sheet) {
    // 只有不存在時才建立新的
    sheet = ss.insertSheet(CONFIG.SHEET_NAMES.RECURRING);

    const headers = ['啟用', '項目', '金額', '付款人', '你的部分', '對方的部分', '分類', '每月執行日', '備註'];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

    sheet.getRange(1, 1, 1, headers.length)
      .setBackground(CONFIG.COLORS.HEADER)
      .setFontColor('#ffffff')
      .setFontWeight('bold')
      .setHorizontalAlignment('center');

    sheet.setFrozenRows(1);

    const examples = [
      [true, '房租', 15000, '你', 15000, 0, '居住', 1, '每月 1 號自動扣款'],
      [true, '水電費', 2000, '兩人', 800, 1200, '居住', 10, '你付 800，對方付 1200'],
      [false, '網路費', 599, '你', 599, 0, '居住', 5, '已停用範例']
    ];

    sheet.getRange(2, 1, examples.length, examples[0].length).setValues(examples);
  }
  // 如果已存在，不做任何事（保護資料）
}

function createSettingsSheet(ss) {
  let sheet = ss.getSheetByName(CONFIG.SHEET_NAMES.SETTINGS);
  if (!sheet) {
    // 只有不存在時才建立新的
    sheet = ss.insertSheet(CONFIG.SHEET_NAMES.SETTINGS);

    const owner = ss.getOwner().getEmail();

    const settings = [
      ['設定項目', '值'],
      ['你的名字', '你'],
      ['對方的名字', '對方'],
      ['預設分類', '飲食,居住,交通,娛樂,寵物,服飾,其他'],
      ['週期事件最後執行日期', ''],
      ['允許存取的使用者', owner],
      ['記帳模式', '共同記帳'],
      ['介面配色', '紫色']
    ];

    sheet.getRange(1, 1, settings.length, 2).setValues(settings);
    sheet.getRange(1, 1, 1, 2)
      .setBackground(CONFIG.COLORS.HEADER)
      .setFontColor('#ffffff')
      .setFontWeight('bold');

    sheet.setColumnWidth(1, 200);
    sheet.setColumnWidth(2, 400);

    // 加入說明
    sheet.getRange('C6').setValue('多個使用者用逗號分隔，例如：user1@gmail.com, user2@gmail.com');
    sheet.getRange('C6').setFontSize(9).setFontColor('#999999');
    sheet.getRange('C7').setValue('個人記帳 / 共同記帳');
    sheet.getRange('C7').setFontSize(9).setFontColor('#999999');
    sheet.getRange('C8').setValue('紫色 / 藍色 / 綠色 / 粉色');
    sheet.getRange('C8').setFontSize(9).setFontColor('#999999');

    // 快速記帳按鈕設定
    const quickExpenseHeaders = ['表情符號', '項目', '金額', '分類'];
    const quickExpenseData = [
      ['🍳', '早餐', 50, '飲食'],
      ['🍱', '午餐', 100, '飲食'],
      ['🍽️', '晚餐', 150, '飲食'],
      ['☕', '咖啡', 60, '飲食'],
      ['🚇', '交通', 20, '交通'],
      ['🅿️', '停車', 50, '交通'],
      ['🍰', '點心', 80, '飲食'],
      ['🧋', '飲料', 50, '飲食']
    ];

    sheet.getRange(8, 1).setValue('快速記帳按鈕設定');
    sheet.getRange(8, 1).setFontWeight('bold').setFontSize(11);
    sheet.getRange(9, 1, 1, 4).setValues([quickExpenseHeaders]);
    sheet.getRange(9, 1, 1, 4)
      .setBackground(CONFIG.COLORS.HEADER)
      .setFontColor('#ffffff')
      .setFontWeight('bold')
      .setHorizontalAlignment('center');

    sheet.getRange(10, 1, quickExpenseData.length, 4).setValues(quickExpenseData);
    sheet.getRange(10, 1, quickExpenseData.length, 4).setHorizontalAlignment('center');

    // 設定欄位寬度
    sheet.setColumnWidth(1, 100);  // 表情符號
    sheet.setColumnWidth(2, 120);  // 項目
    sheet.setColumnWidth(3, 80);   // 金額
    sheet.setColumnWidth(4, 100);  // 分類

    // 加入說明
    sheet.getRange('A18').setValue('💡 提示：可以自由新增、修改或刪除快速記帳按鈕（最多 12 個）');
    sheet.getRange('A18').setFontSize(9).setFontColor('#999999');
  }
  // 如果已存在，不做任何事（保護資料）
}

function createCategoriesSheet(ss) {
  let sheet = ss.getSheetByName(CONFIG.SHEET_NAMES.CATEGORIES);
  if (!sheet) {
    // 只有不存在時才建立新的
    sheet = ss.insertSheet(CONFIG.SHEET_NAMES.CATEGORIES);

    // 橫向佈局：第1行是主分類，第2行往下是子分類
    const mainCategories = ['飲食', '居住', '交通', '娛樂', '寵物', '服飾', '其他'];
    const subCategories = {
      '飲食': ['早餐', '午餐', '晚餐', '宵夜', '飲料', '點心'],
      '居住': ['房租', '水電', '網路', '家具'],
      '交通': ['捷運', '公車', '計程車', '加油', '停車'],
      '娛樂': ['電影', '遊戲', '旅遊'],
      '寵物': ['飼料', '看醫生', '美容'],
      '服飾': [],
      '其他': []
    };

    // 設定第1行（主分類）
    sheet.getRange(1, 1, 1, mainCategories.length).setValues([mainCategories]);
    sheet.getRange(1, 1, 1, mainCategories.length)
      .setBackground(CONFIG.COLORS.HEADER)
      .setFontColor('#ffffff')
      .setFontWeight('bold')
      .setHorizontalAlignment('center')
      .setVerticalAlignment('middle');

    // 設定子分類（第2行往下）
    let maxRows = 1;
    mainCategories.forEach((cat, colIndex) => {
      const subs = subCategories[cat] || [];
      if (subs.length > 0) {
        const colData = subs.map(sub => [sub]);
        sheet.getRange(2, colIndex + 1, subs.length, 1).setValues(colData);
        sheet.getRange(2, colIndex + 1, subs.length, 1).setHorizontalAlignment('center');
        maxRows = Math.max(maxRows, subs.length + 1);
      }
    });

    // 設定欄位寬度
    for (let i = 1; i <= mainCategories.length; i++) {
      sheet.setColumnWidth(i, 120);
    }

    // 凍結第1行
    sheet.setFrozenRows(1);

    // 加入說明
    sheet.getRange(maxRows + 2, 1).setValue('💡 使用說明：');
    sheet.getRange(maxRows + 2, 1).setFontWeight('bold').setFontSize(11);
    sheet.getRange(maxRows + 3, 1, 1, 3).merge();
    sheet.getRange(maxRows + 3, 1).setValue(
      '• 第1行：主分類名稱\n' +
      '• 第2行往下：該主分類的子分類（選填）\n' +
      '• 要新增主分類：在右邊加新欄位\n' +
      '• 要新增子分類：在該欄下方加新行'
    );
    sheet.getRange(maxRows + 3, 1).setFontSize(9).setFontColor('#666666').setWrap(true);
  }
  // 如果已存在，不做任何事（保護資料）
}

// ==================== 核心功能 ====================

/**
 * 取得分類列表（從分類設定工作表讀取）
 */
function getCategories() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const categoriesSheet = ss.getSheetByName(CONFIG.SHEET_NAMES.CATEGORIES);

    if (!categoriesSheet) {
      // 找不到分類設定工作表，嘗試從舊的設定工作表讀取（向下相容）
      return getCategoriesFromSettings();
    }

    const data = categoriesSheet.getDataRange().getValues();
    if (data.length === 0) {
      return ['飲食', '居住', '交通', '娛樂', '寵物', '服飾', '其他'];
    }

    const categories = [];
    const mainCategories = data[0]; // 第1行是主分類

    // 遍歷每一欄（主分類）
    for (let col = 0; col < mainCategories.length; col++) {
      const mainCat = String(mainCategories[col]).trim();
      if (!mainCat) continue; // 跳過空的主分類

      // 加入主分類
      categories.push(mainCat);

      // 讀取該欄的子分類（第2行往下）
      for (let row = 1; row < data.length; row++) {
        const subCat = data[row][col];
        const subCatStr = String(subCat).trim();

        // 跳過空白、說明文字（包含「使用說明」、「提示」等）
        if (!subCatStr ||
            subCatStr.includes('使用說明') ||
            subCatStr.includes('提示') ||
            subCatStr.includes('💡') ||
            subCatStr.includes('•')) {
          continue;
        }

        categories.push(mainCat + '>' + subCatStr);
      }
    }

    return categories.length > 0 ? categories : ['飲食', '居住', '交通', '娛樂', '寵物', '服飾', '其他'];
  } catch (e) {
    Logger.log('讀取分類失敗: ' + e.toString());
    return ['飲食', '居住', '交通', '娛樂', '寵物', '服飾', '其他'];
  }
}

/**
 * 從舊的設定工作表讀取分類（向下相容）
 */
function getCategoriesFromSettings() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const settingsSheet = ss.getSheetByName(CONFIG.SHEET_NAMES.SETTINGS);

    if (!settingsSheet) {
      return ['飲食', '居住', '交通', '娛樂', '寵物', '服飾', '其他'];
    }

    // 讀取 B4 儲存格（預設分類）
    const categoriesStr = settingsSheet.getRange('B4').getValue();

    if (!categoriesStr) {
      return ['飲食', '居住', '交通', '娛樂', '寵物', '服飾', '其他'];
    }

    // 分割逗號並去除空白
    const categories = String(categoriesStr).split(',').map(c => c.trim()).filter(c => c.length > 0);

    return categories.length > 0 ? categories : ['飲食', '居住', '交通', '娛樂', '寵物', '服飾', '其他'];
  } catch (e) {
    Logger.log('從設定工作表讀取分類失敗: ' + e.toString());
    return ['飲食', '居住', '交通', '娛樂', '寵物', '服飾', '其他'];
  }
}

function addExpense(item, amount, payer, actualPayer, yourPart, partnerPart, category, isRecurring, recurringDay, yourActualPaid, partnerActualPaid, expenseDate, expenseTime) {
  // 檢查頻率限制
  checkRateLimit('addExpense');

  // 驗證輸入
  if (!validateText(item, 100)) {
    throw new Error('項目名稱無效（必須是 1-100 字元）');
  }

  if (!validateNumber(amount, 0.01, 9999999)) {
    throw new Error('金額無效（必須介於 0.01 到 9,999,999）');
  }

  if (!validateNumber(yourPart, 0, amount)) {
    throw new Error('你的部分金額無效');
  }

  if (!validateNumber(partnerPart, 0, amount)) {
    throw new Error('對方的部分金額無效');
  }

  // 檢查金額合理性
  if (Math.abs((yourPart + partnerPart) - amount) > 0.01) {
    throw new Error('分帳金額總和必須等於總金額');
  }

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEET_NAMES.EXPENSES);

  // 處理日期時間：如果有提供 expenseDate 和 expenseTime，使用它們；否則使用當下時間
  let date;
  if (expenseDate && expenseTime) {
    // 組合日期和時間字串 (YYYY-MM-DD HH:MM)
    date = new Date(expenseDate + ' ' + expenseTime);
  } else if (expenseDate) {
    // 只有日期，時間設為當下
    const now = new Date();
    date = new Date(expenseDate + ' ' + now.toTimeString().substring(0, 5));
  } else {
    // 都沒有，使用當下時間
    date = new Date();
  }

  const id = date.getTime();

  // 過濾和轉義輸入
  const safeItem = escapeHtml(item.trim());
  const safeCategory = escapeHtml(category);

  // 向下相容：如果沒有提供實際付款金額，根據 actualPayer 推算
  let finalYourActualPaid = yourActualPaid;
  let finalPartnerActualPaid = partnerActualPaid;

  if (yourActualPaid === null || yourActualPaid === undefined) {
    if (actualPayer === '你') {
      finalYourActualPaid = amount;
      finalPartnerActualPaid = 0;
    } else if (actualPayer === '對方') {
      finalYourActualPaid = 0;
      finalPartnerActualPaid = amount;
    } else if (actualPayer === '各自') {
      finalYourActualPaid = yourPart;
      finalPartnerActualPaid = partnerPart;
    } else {
      // 其他情況（例如舊資料）
      finalYourActualPaid = 0;
      finalPartnerActualPaid = 0;
    }
  }

  const row = [
    Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy-MM-dd'),
    safeItem,
    amount,
    payer,
    actualPayer || payer,  // 實際付款人，向下相容
    yourPart,
    partnerPart,
    finalYourActualPaid,  // 你實際付出的金額
    finalPartnerActualPaid,  // 對方實際付出的金額
    safeCategory,
    isRecurring || false,
    recurringDay || '',
    id,
    'expense'  // 記錄類型：支出
  ];

  sheet.appendRow(row);

  const lastRow = sheet.getLastRow();
  sheet.getRange(lastRow, 1, 1, 14).setHorizontalAlignment('center');

  let color = CONFIG.COLORS.BOTH;
  if (payer === '你') color = CONFIG.COLORS.YOUR;
  else if (payer === '對方') color = CONFIG.COLORS.PARTNER;

  sheet.getRange(lastRow, 1, 1, 13).setBackground(color);

  // 記錄日誌
  logAction('新增支出', `項目: ${safeItem}, 金額: ${amount}, 付款人: ${payer}`);

  return id;
}

function executeRecurringExpenses() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const recurringSheet = ss.getSheetByName(CONFIG.SHEET_NAMES.RECURRING);
  const today = new Date();
  const currentDay = today.getDate();

  const data = recurringSheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    const [enabled, item, amount, payer, yourPart, partnerPart, category, executeDay, note] = data[i];

    if (enabled === true && executeDay === currentDay) {
      addExpense(item, amount, payer, yourPart, partnerPart, category, true, executeDay);
      Logger.log(`已自動新增週期支出：${item} - ${amount}`);
    }
  }

  const settingsSheet = ss.getSheetByName(CONFIG.SHEET_NAMES.SETTINGS);
  settingsSheet.getRange('B5').setValue(Utilities.formatDate(today, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss'));
}

function getStatistics(startDate, endDate) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEET_NAMES.EXPENSES);
  const data = sheet.getDataRange().getValues();

  let yourTotal = 0;
  let partnerTotal = 0;
  const categoryStats = {};

  for (let i = 1; i < data.length; i++) {
    const [date, item, amount, payer, actualPayer, yourPart, partnerPart, category] = data[i];

    if (startDate && new Date(date) < new Date(startDate)) continue;
    if (endDate && new Date(date) > new Date(endDate)) continue;

    yourTotal += Number(yourPart) || 0;
    partnerTotal += Number(partnerPart) || 0;

    if (!categoryStats[category]) {
      categoryStats[category] = 0;
    }
    categoryStats[category] += Number(amount) || 0;
  }

  return {
    yourTotal: yourTotal,
    partnerTotal: partnerTotal,
    difference: yourTotal - partnerTotal,
    categoryStats: categoryStats,
    total: yourTotal + partnerTotal
  };
}

// ==================== 登入與權限管理 ====================

/**
 * 檢查使用者是否有權限
 */
function checkUserPermission() {
  const user = Session.getActiveUser().getEmail();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const settingsSheet = ss.getSheetByName(CONFIG.SHEET_NAMES.SETTINGS);

  // 讀取允許的使用者清單（從設定工作表）
  const allowedUsers = settingsSheet.getRange('B6').getValue();

  if (!allowedUsers) {
    // 如果沒設定，預設允許試算表擁有者
    const owner = ss.getOwner().getEmail();
    return {
      allowed: user === owner,
      user: user,
      name: user.split('@')[0]
    };
  }

  const userList = allowedUsers.split(',').map(u => u.trim());

  return {
    allowed: userList.includes(user),
    user: user,
    name: user.split('@')[0]
  };
}

/**
 * 取得當前使用者資訊（含 Google 頭像）
 */
function getCurrentUser() {
  const user = Session.getActiveUser().getEmail();
  const permission = checkUserPermission();

  // 取得使用者照片和名稱
  let photoUrl = '';
  let displayName = user.split('@')[0]; // 預設使用 email 前綴

  try {
    // 使用 People API 取得使用者資訊
    const userInfo = People.People.get('people/me', {
      personFields: 'photos,names'
    });

    if (userInfo.photos && userInfo.photos.length > 0) {
      photoUrl = userInfo.photos[0].url;
    }

    // 取得使用者的顯示名稱
    if (userInfo.names && userInfo.names.length > 0) {
      displayName = userInfo.names[0].displayName || user.split('@')[0];
    }
  } catch (e) {
    Logger.log('無法取得使用者資訊: ' + e.toString());
    // 使用預設 Google 帳號圖示
    photoUrl = 'https://www.gstatic.com/images/branding/product/1x/avatar_circle_blue_512dp.png';
  }

  // 取得「對方的名字」設定
  let partnerName = '對方';
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const settingsSheet = ss.getSheetByName(CONFIG.SHEET_NAMES.SETTINGS);
    if (settingsSheet) {
      const partnerNameValue = settingsSheet.getRange('B3').getValue();
      if (partnerNameValue && partnerNameValue.trim()) {
        partnerName = String(partnerNameValue).trim();
      }
    }
  } catch (e) {
    Logger.log('無法取得對方名字設定: ' + e.toString());
  }

  return {
    email: user,
    name: displayName,
    partnerName: partnerName,
    photoUrl: photoUrl,
    allowed: permission.allowed
  };
}

/**
 * 取得應用程式設定（記帳模式、介面配色）
 */
function getAppSettings() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const settingsSheet = ss.getSheetByName(CONFIG.SHEET_NAMES.SETTINGS);

    if (!settingsSheet) {
      Logger.log('設定工作表不存在');
      return {
        mode: '共同記帳',
        theme: '紫色'
      };
    }

    // 讀取記帳模式（B7）和介面配色（B8）
    const mode = settingsSheet.getRange('B7').getValue() || '共同記帳';
    const theme = settingsSheet.getRange('B8').getValue() || '紫色';

    return {
      mode: String(mode).trim(),
      theme: String(theme).trim()
    };
  } catch (e) {
    Logger.log('讀取應用程式設定失敗: ' + e.toString());
    return {
      mode: '共同記帳',
      theme: '紫色'
    };
  }
}

/**
 * 取得快速記帳按鈕設定
 */
function getQuickExpenseButtons() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const settingsSheet = ss.getSheetByName(CONFIG.SHEET_NAMES.SETTINGS);

    if (!settingsSheet) {
      Logger.log('設定工作表不存在');
      return [];
    }

    // 讀取第 10 行開始的快速記帳設定（最多 12 個）
    const data = settingsSheet.getRange(10, 1, 12, 4).getValues();
    const buttons = [];

    for (let i = 0; i < data.length; i++) {
      const [emoji, item, amount, category] = data[i];

      // 如果項目和金額都有值，則加入按鈕清單
      if (item && amount) {
        buttons.push({
          emoji: emoji || '📝',
          item: String(item).trim(),
          amount: Number(amount) || 0,
          category: String(category).trim() || '其他'
        });
      }
    }

    Logger.log('載入了 ' + buttons.length + ' 個快速記帳按鈕');
    return buttons;
  } catch (e) {
    Logger.log('讀取快速記帳設定失敗: ' + e.toString());
    return [];
  }
}

// ==================== Web App API ====================

function doGet() {
  // 檢查權限
  const permission = checkUserPermission();

  if (!permission.allowed) {
    // 無權限時顯示錯誤頁面
    return HtmlService.createHtmlOutput(`
      <html>
        <head>
          <meta charset="UTF-8">
          <meta name="viewport" content="width=device-width, initial-scale=1.0">
          <style>
            body {
              font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', sans-serif;
              background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
              min-height: 100vh;
              display: flex;
              align-items: center;
              justify-content: center;
              margin: 0;
            }
            .error-box {
              background: white;
              padding: 40px;
              border-radius: 20px;
              text-align: center;
              box-shadow: 0 20px 60px rgba(0, 0, 0, 0.3);
              max-width: 400px;
            }
            .error-icon {
              font-size: 4em;
              margin-bottom: 20px;
            }
            h1 {
              color: #ef4444;
              margin-bottom: 10px;
            }
            p {
              color: #666;
              line-height: 1.6;
            }
            .user-info {
              background: #f3f4f6;
              padding: 15px;
              border-radius: 10px;
              margin-top: 20px;
              color: #374151;
            }
          </style>
        </head>
        <body>
          <div class="error-box">
            <div class="error-icon">🔒</div>
            <h1>無法存取</h1>
            <p>抱歉，您沒有權限使用這個記帳系統。</p>
            <div class="user-info">
              <strong>您的帳號：</strong><br>
              ${permission.user}
            </div>
            <p style="font-size: 0.9em; margin-top: 20px;">
              請聯絡系統管理員將您的 Email 加入允許清單。
            </p>
          </div>
        </body>
      </html>
    `).setTitle('無法存取');
  }

  // 有權限則顯示主頁面
  const htmlOutput = HtmlService.createHtmlOutputFromFile('index')
    .setTitle('共同記帳')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.DEFAULT); // 防止點擊劫持

  // 強制設定 viewport - 禁止縮放避免 iPhone 輸入時自動縮放
  htmlOutput.addMetaTag('viewport', 'width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no');

  return htmlOutput;
}

function getExpenses(filters) {
  // 權限驗證
  const permission = checkUserPermission();
  if (!permission.allowed) {
    throw new Error('無權限訪問');
  }

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEET_NAMES.EXPENSES);

  // 檢查工作表是否存在
  if (!sheet) {
    Logger.log('支出記錄工作表不存在');
    return { expenses: [], total: 0, hasMore: false };
  }

  const data = sheet.getDataRange().getValues();

  // 如果只有標題列，返回空陣列
  if (data.length <= 1) {
    Logger.log('沒有支出記錄');
    return { expenses: [], total: 0, hasMore: false };
  }

  // 解析分頁參數
  const offset = (filters && filters.offset) ? Number(filters.offset) : 0;
  const limit = (filters && filters.limit) ? Number(filters.limit) : 50;

  const allExpenses = [];
  for (let i = 1; i < data.length; i++) {
    // 跳過空白列
    if (!data[i][1]) {
      continue;
    }

    // 格式化日期（處理 Date 物件或字串）
    let dateStr = data[i][0];
    if (dateStr instanceof Date) {
      dateStr = Utilities.formatDate(dateStr, Session.getScriptTimeZone(), 'yyyy-MM-dd');
    } else if (typeof dateStr === 'string') {
      // 已經是字串，保持原樣
    } else {
      // 其他情況，使用當前日期
      dateStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
    }

    allExpenses.push({
      date: dateStr,
      item: String(data[i][1] || ''),
      amount: Number(data[i][2]) || 0,
      payer: String(data[i][3] || ''),
      actualPayer: String(data[i][4] || data[i][3] || ''),  // 實際付款人，向下相容
      yourPart: Number(data[i][5]) || 0,
      partnerPart: Number(data[i][6]) || 0,
      yourActualPaid: Number(data[i][7]) >= 0 ? Number(data[i][7]) : null,  // 你實際付出的金額，向下相容
      partnerActualPaid: Number(data[i][8]) >= 0 ? Number(data[i][8]) : null,  // 對方實際付出的金額，向下相容
      category: String(data[i][9] || '其他'),
      isRecurring: Boolean(data[i][10]),
      recurringDay: data[i][11] || '',
      id: String(data[i][12] || '')
    });
  }

  // 按日期排序（新到舊）
  allExpenses.sort(function(a, b) {
    return b.date.localeCompare(a.date);
  });

  const total = allExpenses.length;
  const expenses = allExpenses.slice(offset, offset + limit);
  const hasMore = (offset + limit) < total;

  Logger.log('成功載入 ' + expenses.length + ' 筆支出記錄（共 ' + total + ' 筆，offset: ' + offset + '）');

  return {
    expenses: expenses,
    total: total,
    hasMore: hasMore
  };
}

/**
 * 取得所有支出記錄（不分頁，用於儀表板和統計）
 */
function getAllExpenses() {
  // 權限驗證
  const permission = checkUserPermission();
  if (!permission.allowed) {
    throw new Error('無權限訪問');
  }

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEET_NAMES.EXPENSES);

  // 檢查工作表是否存在
  if (!sheet) {
    Logger.log('支出記錄工作表不存在');
    return [];
  }

  const data = sheet.getDataRange().getValues();

  // 如果只有標題列，返回空陣列
  if (data.length <= 1) {
    Logger.log('沒有支出記錄');
    return [];
  }

  const expenses = [];
  for (let i = 1; i < data.length; i++) {
    // 跳過空白列
    if (!data[i][1]) {
      continue;
    }

    // 格式化日期（處理 Date 物件或字串）
    let dateStr = data[i][0];
    if (dateStr instanceof Date) {
      dateStr = Utilities.formatDate(dateStr, Session.getScriptTimeZone(), 'yyyy-MM-dd');
    } else if (typeof dateStr === 'string') {
      // 已經是字串，保持原樣
    } else {
      // 其他情況，使用當前日期
      dateStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
    }

    expenses.push({
      date: dateStr,
      item: String(data[i][1] || ''),
      amount: Number(data[i][2]) || 0,
      payer: String(data[i][3] || ''),
      actualPayer: String(data[i][4] || data[i][3] || ''),  // 實際付款人，向下相容
      yourPart: Number(data[i][5]) || 0,
      partnerPart: Number(data[i][6]) || 0,
      yourActualPaid: Number(data[i][7]) >= 0 ? Number(data[i][7]) : null,  // 你實際付出的金額，向下相容
      partnerActualPaid: Number(data[i][8]) >= 0 ? Number(data[i][8]) : null,  // 對方實際付出的金額，向下相容
      category: String(data[i][9] || '其他'),
      isRecurring: Boolean(data[i][10]),
      recurringDay: data[i][11] || '',
      id: String(data[i][12] || '')
    });
  }

  Logger.log('成功載入所有 ' + expenses.length + ' 筆支出記錄');
  return expenses;
}

function addExpenseFromWeb(expenseData) {
  return addExpense(
    expenseData.item,
    expenseData.amount,
    expenseData.payer,
    expenseData.actualPayer || expenseData.payer,  // 實際付款人，向下相容
    expenseData.yourPart,
    expenseData.partnerPart,
    expenseData.category,
    expenseData.isRecurring,
    expenseData.recurringDay,
    expenseData.yourActualPaid || null,  // 你實際付出的金額
    expenseData.partnerActualPaid || null,  // 對方實際付出的金額
    expenseData.expenseDate || null,  // 支出日期
    expenseData.expenseTime || null   // 支出時間
  );
}

function getStatisticsFromWeb(startDate, endDate) {
  return getStatistics(startDate, endDate);
}

/**
 * 供網頁呼叫 - 更新支出記錄
 */
function updateExpenseById(updatedData) {
  // 檢查頻率限制
  checkRateLimit('updateExpense');

  // 權限檢查
  const currentUser = Session.getActiveUser().getEmail();
  const permission = checkUserPermission();

  if (!permission.allowed) {
    throw new Error('無權限操作');
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const owner = ss.getOwner().getEmail();
  const sheet = ss.getSheetByName(CONFIG.SHEET_NAMES.EXPENSES);
  const data = sheet.getDataRange().getValues();

  // 只有管理員可以編輯記錄
  if (currentUser !== owner) {
    logAction('更新支出失敗', `非管理員嘗試更新 ID: ${updatedData.id}`);
    throw new Error('只有管理員可以編輯記錄');
  }

  // 驗證輸入
  if (!validateText(updatedData.item, 100)) {
    throw new Error('項目名稱無效（必須是 1-100 字元）');
  }

  if (!validateNumber(updatedData.amount, 0.01, 9999999)) {
    throw new Error('金額無效（必須介於 0.01 到 9,999,999）');
  }

  if (!validateNumber(updatedData.yourPart, 0, updatedData.amount)) {
    throw new Error('你的部分金額無效');
  }

  if (!validateNumber(updatedData.partnerPart, 0, updatedData.amount)) {
    throw new Error('對方的部分金額無效');
  }

  // 檢查金額合理性
  if (Math.abs((updatedData.yourPart + updatedData.partnerPart) - updatedData.amount) > 0.01) {
    throw new Error('分帳金額總和必須等於總金額');
  }

  // 找到 ID 欄位（第 13 欄）並更新
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][12]) === String(updatedData.id)) {
      const oldItem = data[i][1];
      const oldAmount = data[i][2];

      // 過濾和轉義輸入
      const safeItem = escapeHtml(updatedData.item.trim());
      const safeCategory = escapeHtml(updatedData.category);

      // 更新資料（保留原有的日期和 ID）
      sheet.getRange(i + 1, 2).setValue(safeItem);           // 項目
      sheet.getRange(i + 1, 3).setValue(updatedData.amount); // 金額
      sheet.getRange(i + 1, 4).setValue(updatedData.payer);  // 付款人
      // 第 5 欄是「實際付款人」，編輯功能暫不更新
      sheet.getRange(i + 1, 6).setValue(updatedData.yourPart);     // 你的部分
      sheet.getRange(i + 1, 7).setValue(updatedData.partnerPart);  // 對方的部分
      // 第 8, 9 欄是「你實付」、「對方實付」，編輯功能暫不更新
      sheet.getRange(i + 1, 10).setValue(safeCategory);       // 分類

      // 更新背景顏色
      let color = CONFIG.COLORS.BOTH;
      if (updatedData.payer === '你') color = CONFIG.COLORS.YOUR;
      else if (updatedData.payer === '對方') color = CONFIG.COLORS.PARTNER;
      sheet.getRange(i + 1, 1, 1, 13).setBackground(color);

      // 記錄日誌
      logAction('更新支出', `ID: ${updatedData.id}, 原: ${oldItem}($${oldAmount}) → 新: ${safeItem}($${updatedData.amount})`);

      Logger.log(`已更新記錄 ID: ${updatedData.id}`);
      return true;
    }
  }

  throw new Error('找不到該記錄');
}

/**
 * 供網頁呼叫 - 刪除支出記錄
 */
function deleteExpenseById(id) {
  // 檢查頻率限制
  checkRateLimit('deleteExpense');

  // 權限檢查
  const currentUser = Session.getActiveUser().getEmail();
  const permission = checkUserPermission();

  if (!permission.allowed) {
    throw new Error('無權限操作');
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const owner = ss.getOwner().getEmail();
  const sheet = ss.getSheetByName(CONFIG.SHEET_NAMES.EXPENSES);
  const data = sheet.getDataRange().getValues();

  // 只有管理員可以刪除記錄
  if (currentUser !== owner) {
    logAction('刪除支出失敗', `非管理員嘗試刪除 ID: ${id}`);
    throw new Error('只有管理員可以刪除記錄');
  }

  // 找到 ID 欄位（第 13 欄）
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][12]) === String(id)) {
      const item = data[i][1];
      const amount = data[i][2];

      sheet.deleteRow(i + 1);

      // 記錄日誌
      logAction('刪除支出', `項目: ${item}, 金額: ${amount}, ID: ${id}`);

      Logger.log(`已刪除記錄 ID: ${id}`);
      return true;
    }
  }

  throw new Error('找不到該記錄');
}

// ==================== 安全性與驗證 ====================

/**
 * 輸入驗證 - Email
 */
function validateEmail(email) {
  const emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
  return emailRegex.test(email);
}

/**
 * 輸入驗證 - 數字
 */
function validateNumber(value, min = 0, max = 9999999) {
  const num = Number(value);
  return !isNaN(num) && num >= min && num <= max;
}

/**
 * 輸入驗證 - 文字
 */
function validateText(text, maxLength = 100) {
  if (!text || typeof text !== 'string') return false;
  text = text.trim();
  return text.length > 0 && text.length <= maxLength;
}

/**
 * HTML 編碼防止 XSS
 */
function escapeHtml(text) {
  if (!text) return '';
  return String(text)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#039;');
}

/**
 * 記錄操作日誌（敏感資訊脫敏）
 */
function logAction(action, details) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let logSheet = ss.getSheetByName('操作日誌');

  // 如果沒有日誌工作表，創建一個
  if (!logSheet) {
    logSheet = ss.insertSheet('操作日誌');
    logSheet.getRange(1, 1, 1, 5).setValues([['時間', '使用者', '動作', '詳細資訊', 'IP/裝置']]);
    logSheet.getRange(1, 1, 1, 5)
      .setBackground(CONFIG.COLORS.HEADER)
      .setFontColor('#ffffff')
      .setFontWeight('bold');

    // 保護日誌工作表，只有擁有者可以編輯
    const protection = logSheet.protect().setDescription('操作日誌保護');
    protection.removeEditors(protection.getEditors());
    protection.addEditor(ss.getOwner().getEmail());
  }

  const user = Session.getActiveUser().getEmail();
  const timestamp = new Date();
  const userAgent = Session.getTemporaryActiveUserKey(); // 簡易裝置識別

  // 轉義日誌內容防止注入
  const safeDetails = escapeHtml(String(details));

  logSheet.appendRow([
    Utilities.formatDate(timestamp, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss'),
    user,
    action,
    safeDetails,
    userAgent
  ]);
}

/**
 * 檢查請求頻率限制（防止濫用）
 */
function checkRateLimit(action) {
  const userCache = CacheService.getUserCache();
  const key = 'ratelimit_' + action;
  const count = userCache.get(key);

  if (count && Number(count) > 50) { // 每分鐘最多 50 次
    throw new Error('操作過於頻繁，請稍後再試');
  }

  userCache.put(key, Number(count || 0) + 1, 60); // 60 秒過期
}

// ==================== 成員管理 ====================

/**
 * 取得所有成員列表
 */
function getMembers() {
  try {
    Logger.log('getMembers: 開始執行');
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    Logger.log('getMembers: 已取得試算表');

    let owner;
    try {
      owner = ss.getOwner().getEmail();
      Logger.log('getMembers: 擁有者 = ' + owner);
    } catch (ownerError) {
      Logger.log('getMembers: 無法取得擁有者，使用當前用戶');
      owner = Session.getActiveUser().getEmail();
    }

    // 檢查設定工作表是否存在
    let settingsSheet = ss.getSheetByName(CONFIG.SHEET_NAMES.SETTINGS);
    Logger.log('getMembers: 設定工作表存在 = ' + (settingsSheet != null));

    if (!settingsSheet) {
      Logger.log('getMembers: 設定工作表不存在，返回擁有者');
      return [{
        email: owner,
        isOwner: true,
        name: owner.split('@')[0]
      }];
    }

    // 讀取允許的使用者清單
    let allowedUsers;
    try {
      allowedUsers = settingsSheet.getRange('B6').getValue();
      Logger.log('getMembers: 允許的使用者 = ' + allowedUsers);
    } catch (rangeError) {
      Logger.log('getMembers: 讀取 B6 失敗: ' + rangeError.toString());
      return [{
        email: owner,
        isOwner: true,
        name: owner.split('@')[0]
      }];
    }

    if (!allowedUsers || allowedUsers.toString().trim() === '') {
      Logger.log('getMembers: 白名單為空，返回擁有者');
      return [{
        email: owner,
        isOwner: true,
        name: owner.split('@')[0]
      }];
    }

    const userList = allowedUsers.toString().split(',').map(function(u) {
      return u.trim();
    }).filter(function(u) {
      return u;
    });

    Logger.log('getMembers: 用戶列表 = ' + userList.join(', '));

    // 確保擁有者在列表中
    const members = userList.map(function(email) {
      return {
        email: email,
        isOwner: email === owner,
        name: email.split('@')[0]
      };
    });

    // 如果擁有者不在列表中，加入擁有者
    const ownerExists = members.some(function(m) {
      return m.email === owner;
    });

    if (!ownerExists) {
      Logger.log('getMembers: 擁有者不在列表中，添加');
      members.unshift({
        email: owner,
        isOwner: true,
        name: owner.split('@')[0]
      });
    }

    Logger.log('getMembers: 返回 ' + members.length + ' 個成員');
    return members;
  } catch (error) {
    Logger.log('getMembers error: ' + error.toString());
    Logger.log('getMembers error stack: ' + error.stack);

    // 發生錯誤時返回空陣列並拋出錯誤，讓前端知道
    throw new Error('無法載入成員列表：' + error.message);
  }
}

/**
 * 邀請新成員
 */
function inviteMember(email) {
  // 檢查頻率限制
  checkRateLimit('inviteMember');

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const settingsSheet = ss.getSheetByName(CONFIG.SHEET_NAMES.SETTINGS);
  const currentUser = Session.getActiveUser().getEmail();
  const owner = ss.getOwner().getEmail();

  // 只有擁有者可以邀請成員
  if (currentUser !== owner) {
    logAction('邀請成員失敗', `非管理員嘗試邀請: ${email}`);
    throw new Error('只有系統管理員可以邀請成員');
  }

  // 驗證 email 格式
  if (!validateEmail(email)) {
    throw new Error('請輸入有效的 Email 地址');
  }

  email = email.trim().toLowerCase();

  // 讀取現有成員
  const allowedUsers = settingsSheet.getRange('B6').getValue();
  let userList = [];

  if (allowedUsers) {
    userList = allowedUsers.split(',').map(u => u.trim().toLowerCase());
  }

  // 檢查是否已存在
  if (userList.includes(email)) {
    throw new Error('此成員已在列表中');
  }

  // 新增成員
  userList.push(email);
  settingsSheet.getRange('B6').setValue(userList.join(', '));

  // 發送邀請郵件
  try {
    const appUrl = ScriptApp.getService().getUrl();

    // 驗證 URL 安全性（確保是 HTTPS）
    if (!appUrl.startsWith('https://')) {
      throw new Error('不安全的應用 URL');
    }

    // 轉義郵件中的變數以防止 XSS
    const safeUser = escapeHtml(currentUser);
    const safeAppUrl = escapeHtml(appUrl);

    MailApp.sendEmail({
      to: email,
      subject: '【共同記帳】邀請您加入記帳系統',
      htmlBody: `
        <div style="font-family: Arial, sans-serif; max-width: 600px; margin: 0 auto;">
          <h2 style="color: #667eea;">💑 共同記帳邀請</h2>
          <p>您好！</p>
          <p><strong>${safeUser}</strong> 邀請您一起使用共同記帳系統。</p>
          <p>點擊下方連結即可開始使用：</p>
          <p style="text-align: center; margin: 30px 0;">
            <a href="${safeAppUrl}" style="background: linear-gradient(135deg, #667eea 0%, #764ba2 100%); color: white; padding: 15px 30px; text-decoration: none; border-radius: 10px; display: inline-block;">
              開始使用
            </a>
          </p>
          <p style="color: #999; font-size: 0.9em;">如果按鈕無法點擊，請複製此連結：<br>${safeAppUrl}</p>
        </div>
      `
    });
  } catch (e) {
    Logger.log('發送郵件失敗: ' + e.toString());
    // 即使郵件發送失敗，仍然添加成員
  }

  // 記錄日誌
  logAction('邀請成員', `已邀請: ${email}`);

  return {
    success: true,
    message: '已成功邀請 ' + email
  };
}

/**
 * 移除成員
 */
function removeMember(email) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const settingsSheet = ss.getSheetByName(CONFIG.SHEET_NAMES.SETTINGS);
  const currentUser = Session.getActiveUser().getEmail();
  const owner = ss.getOwner().getEmail();

  // 只有擁有者可以移除成員
  if (currentUser !== owner) {
    throw new Error('只有系統管理員可以移除成員');
  }

  // 不能移除擁有者自己
  if (email === owner) {
    throw new Error('無法移除系統管理員');
  }

  email = email.trim().toLowerCase();

  // 讀取現有成員
  const allowedUsers = settingsSheet.getRange('B6').getValue();

  if (!allowedUsers) {
    throw new Error('成員列表為空');
  }

  let userList = allowedUsers.split(',').map(u => u.trim().toLowerCase());

  // 移除成員
  userList = userList.filter(u => u !== email);

  settingsSheet.getRange('B6').setValue(userList.join(', '));

  // 記錄日誌
  logAction('移除成員', `已移除: ${email}`);

  return {
    success: true,
    message: '已移除 ' + email
  };
}

// ==================== 觸發器 ====================

function setupTriggers() {
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => ScriptApp.deleteTrigger(trigger));

  ScriptApp.newTrigger('executeRecurringExpenses')
    .timeBased()
    .atHour(9)
    .everyDays(1)
    .create();

  Logger.log('已設定每日觸發器');
}

function manualExecuteRecurring() {
  executeRecurringExpenses();
  SpreadsheetApp.getUi().alert('✅ 週期事件執行完成！\n\n請檢查「支出記錄」工作表。');
}

/**
 * 升級資料結構 - 向下相容地新增欄位
 */
function upgradeDataStructure() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEET_NAMES.EXPENSES);

  if (!sheet) {
    ui.alert('❌ 錯誤', '找不到「支出記錄」工作表。\n\n請先執行「初始化系統」。', ui.ButtonSet.OK);
    return;
  }

  // 檢查是否已有「實際付款人」欄位
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const actualPayerIndex = headers.indexOf('實際付款人');

  if (actualPayerIndex !== -1) {
    ui.alert('✅ 資料結構已是最新版本', '無需升級。', ui.ButtonSet.OK);
    return;
  }

  // 確認升級
  const response = ui.alert(
    '🔄 升級資料結構',
    '即將在「付款人」欄位後方新增「實際付款人」欄位。\n\n' +
    '這是為了支援墊付功能（例如：我幫對方墊付）。\n\n' +
    '升級過程不會刪除任何資料，舊資料會自動相容。\n\n' +
    '確定要升級嗎？',
    ui.ButtonSet.YES_NO
  );

  if (response !== ui.Button.YES) {
    ui.alert('✅ 已取消升級');
    return;
  }

  // 執行升級：在第 5 欄（付款人後）插入新欄位
  sheet.insertColumnAfter(4); // 在第 4 欄後插入

  // 設定標題
  sheet.getRange(1, 5).setValue('實際付款人');
  sheet.getRange(1, 5)
    .setBackground(CONFIG.COLORS.HEADER)
    .setFontColor('#ffffff')
    .setFontWeight('bold')
    .setHorizontalAlignment('center');

  // 設定欄寬
  sheet.setColumnWidth(5, 100);

  // 自動填入舊資料：實際付款人 = 付款人
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    const payerData = sheet.getRange(2, 4, lastRow - 1, 1).getValues(); // 第 4 欄是付款人
    sheet.getRange(2, 5, lastRow - 1, 1).setValues(payerData); // 複製到第 5 欄
  }

  ui.alert('✅ 升級完成！\n\n已新增「實際付款人」欄位，\n舊資料已自動設定為與「付款人」相同。');
}

/**
 * 升級快速記帳設定 - 在設定工作表新增快速記帳按鈕區域
 */
function addQuickExpenseSettings() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const settingsSheet = ss.getSheetByName(CONFIG.SHEET_NAMES.SETTINGS);

  if (!settingsSheet) {
    ui.alert('❌ 錯誤', '找不到「設定」工作表。\n\n請先執行「初始化系統」。', ui.ButtonSet.OK);
    return;
  }

  // 檢查第 8 行是否已有快速記帳設定
  const cell8 = settingsSheet.getRange('A8').getValue();
  if (cell8 === '快速記帳按鈕設定') {
    ui.alert('✅ 快速記帳設定已存在', '無需重複新增。', ui.ButtonSet.OK);
    return;
  }

  // 確認升級
  const response = ui.alert(
    '🔄 新增快速記帳設定',
    '即將在「設定」工作表新增快速記帳按鈕設定區域。\n\n' +
    '你可以在試算表直接修改快速記帳按鈕的項目、金額和分類。\n\n' +
    '這不會影響任何現有資料。\n\n' +
    '確定要新增嗎？',
    ui.ButtonSet.YES_NO
  );

  if (response !== ui.Button.YES) {
    ui.alert('✅ 已取消');
    return;
  }

  // 新增快速記帳按鈕設定
  const quickExpenseHeaders = ['表情符號', '項目', '金額', '分類'];
  const quickExpenseData = [
    ['🍳', '早餐', 50, '飲食'],
    ['🍱', '午餐', 100, '飲食'],
    ['🍽️', '晚餐', 150, '飲食'],
    ['☕', '咖啡', 60, '飲食'],
    ['🚇', '交通', 20, '交通'],
    ['🅿️', '停車', 50, '交通'],
    ['🍰', '點心', 80, '飲食'],
    ['🧋', '飲料', 50, '飲食']
  ];

  settingsSheet.getRange(8, 1).setValue('快速記帳按鈕設定');
  settingsSheet.getRange(8, 1).setFontWeight('bold').setFontSize(11);
  settingsSheet.getRange(9, 1, 1, 4).setValues([quickExpenseHeaders]);
  settingsSheet.getRange(9, 1, 1, 4)
    .setBackground(CONFIG.COLORS.HEADER)
    .setFontColor('#ffffff')
    .setFontWeight('bold')
    .setHorizontalAlignment('center');

  settingsSheet.getRange(10, 1, quickExpenseData.length, 4).setValues(quickExpenseData);
  settingsSheet.getRange(10, 1, quickExpenseData.length, 4).setHorizontalAlignment('center');

  // 設定欄位寬度
  settingsSheet.setColumnWidth(1, 100);  // 表情符號
  settingsSheet.setColumnWidth(2, 120);  // 項目
  settingsSheet.setColumnWidth(3, 80);   // 金額
  settingsSheet.setColumnWidth(4, 100);  // 分類

  // 加入說明
  settingsSheet.getRange('A18').setValue('💡 提示：可以自由新增、修改或刪除快速記帳按鈕（最多 12 個）');
  settingsSheet.getRange('A18').setFontSize(9).setFontColor('#999999');

  ui.alert('✅ 新增完成！\n\n已在「設定」工作表新增快速記帳按鈕設定區域。\n\n你現在可以直接在試算表編輯按鈕設定，重新整理網頁後就會生效！');
}

/**
 * 一鍵升級到最新版本 - 自動執行所有可用的升級
 */
function upgradeToLatest() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // 檢查是否已初始化
  const expensesSheet = ss.getSheetByName(CONFIG.SHEET_NAMES.EXPENSES);
  if (!expensesSheet) {
    ui.alert('❌ 錯誤', '找不到「支出記錄」工作表。\n\n請先執行「初始化系統」。', ui.ButtonSet.OK);
    return;
  }

  // 確認升級
  const response = ui.alert(
    '🔄 升級到最新版本',
    '即將檢查並執行所有可用的升級項目：\n\n' +
    '• v2.4 - 墊付功能（實際付款人欄位）\n' +
    '• v2.5 - 快速記帳按鈕設定\n' +
    '• v2.8 - 結算功能（記錄類型欄位）\n\n' +
    '已完成的升級會自動跳過，不會重複執行。\n\n' +
    '確定要開始升級嗎？',
    ui.ButtonSet.YES_NO
  );

  if (response !== ui.Button.YES) {
    ui.alert('✅ 已取消升級');
    return;
  }

  const upgrades = [];
  let hasUpgrade = false;

  // === 檢查 v2.4：墊付功能 ===
  const headers = expensesSheet.getRange(1, 1, 1, expensesSheet.getLastColumn()).getValues()[0];
  const actualPayerIndex = headers.indexOf('實際付款人');

  if (actualPayerIndex === -1) {
    // 需要升級 v2.4
    try {
      // 執行升級：在第 5 欄（付款人後）插入新欄位
      expensesSheet.insertColumnAfter(4);
      expensesSheet.getRange(1, 5).setValue('實際付款人');
      expensesSheet.getRange(1, 5)
        .setBackground(CONFIG.COLORS.HEADER)
        .setFontColor('#ffffff')
        .setFontWeight('bold')
        .setHorizontalAlignment('center');
      expensesSheet.setColumnWidth(5, 100);

      // 自動填入舊資料
      const lastRow = expensesSheet.getLastRow();
      if (lastRow > 1) {
        const payerData = expensesSheet.getRange(2, 4, lastRow - 1, 1).getValues();
        expensesSheet.getRange(2, 5, lastRow - 1, 1).setValues(payerData);
      }

      upgrades.push('✓ v2.4 墊付功能');
      hasUpgrade = true;
    } catch (e) {
      upgrades.push('✗ v2.4 墊付功能失敗：' + e.toString());
    }
  } else {
    upgrades.push('- v2.4 墊付功能（已安裝）');
  }

  // === 檢查 v2.5：快速記帳設定 ===
  const settingsSheet = ss.getSheetByName(CONFIG.SHEET_NAMES.SETTINGS);
  if (!settingsSheet) {
    upgrades.push('✗ v2.5 快速記帳設定失敗：找不到設定工作表');
  } else {
    const cell8 = settingsSheet.getRange('A8').getValue();
    if (cell8 !== '快速記帳按鈕設定') {
      // 需要升級 v2.5
      try {
        const quickExpenseHeaders = ['表情符號', '項目', '金額', '分類'];
        const quickExpenseData = [
          ['🍳', '早餐', 50, '飲食'],
          ['🍱', '午餐', 100, '飲食'],
          ['🍽️', '晚餐', 150, '飲食'],
          ['☕', '咖啡', 60, '飲食'],
          ['🚇', '交通', 20, '交通'],
          ['🅿️', '停車', 50, '交通'],
          ['🍰', '點心', 80, '飲食'],
          ['🧋', '飲料', 50, '飲食']
        ];

        settingsSheet.getRange(8, 1).setValue('快速記帳按鈕設定');
        settingsSheet.getRange(8, 1).setFontWeight('bold').setFontSize(11);
        settingsSheet.getRange(9, 1, 1, 4).setValues([quickExpenseHeaders]);
        settingsSheet.getRange(9, 1, 1, 4)
          .setBackground(CONFIG.COLORS.HEADER)
          .setFontColor('#ffffff')
          .setFontWeight('bold')
          .setHorizontalAlignment('center');
        settingsSheet.getRange(10, 1, quickExpenseData.length, 4).setValues(quickExpenseData);
        settingsSheet.getRange(10, 1, quickExpenseData.length, 4).setHorizontalAlignment('center');
        settingsSheet.setColumnWidth(1, 100);
        settingsSheet.setColumnWidth(2, 120);
        settingsSheet.setColumnWidth(3, 80);
        settingsSheet.setColumnWidth(4, 100);
        settingsSheet.getRange('A18').setValue('💡 提示：可以自由新增、修改或刪除快速記帳按鈕（最多 12 個）');
        settingsSheet.getRange('A18').setFontSize(9).setFontColor('#999999');

        upgrades.push('✓ v2.5 快速記帳設定');
        hasUpgrade = true;
      } catch (e) {
        upgrades.push('✗ v2.5 快速記帳設定失敗：' + e.toString());
      }
    } else {
      upgrades.push('- v2.5 快速記帳設定（已安裝）');
    }
  }

  // === 檢查 v2.8：結算功能（記錄類型欄位） ===
  const recordTypeIndex = headers.indexOf('記錄類型');

  if (recordTypeIndex === -1) {
    // 需要升級 v2.8
    try {
      // 在最後一欄新增「記錄類型」欄位
      const lastCol = expensesSheet.getLastColumn();
      expensesSheet.getRange(1, lastCol + 1).setValue('記錄類型');
      expensesSheet.getRange(1, lastCol + 1)
        .setBackground(CONFIG.COLORS.HEADER)
        .setFontColor('#ffffff')
        .setFontWeight('bold')
        .setHorizontalAlignment('center');
      expensesSheet.setColumnWidth(lastCol + 1, 100);

      // 自動填入現有資料為 'expense'
      const lastRow = expensesSheet.getLastRow();
      if (lastRow > 1) {
        const recordTypes = [];
        for (let i = 0; i < lastRow - 1; i++) {
          recordTypes.push(['expense']);
        }
        expensesSheet.getRange(2, lastCol + 1, lastRow - 1, 1).setValues(recordTypes);
      }

      upgrades.push('✓ v2.8 結算功能');
      hasUpgrade = true;
    } catch (e) {
      upgrades.push('✗ v2.8 結算功能失敗：' + e.toString());
    }
  } else {
    upgrades.push('- v2.8 結算功能（已安裝）');
  }

  // 顯示結果
  const message = upgrades.join('\n');
  if (hasUpgrade) {
    ui.alert('✅ 升級完成！\n\n' + message + '\n\n系統已升級到最新版本！');
  } else {
    ui.alert('✅ 已是最新版本\n\n' + message + '\n\n無需升級。');
  }
}

/**
 * 重置系統 - 清空所有資料（危險操作）
 */
function resetSystem() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const expensesSheet = ss.getSheetByName(CONFIG.SHEET_NAMES.EXPENSES);

  if (!expensesSheet) {
    ui.alert('❌ 錯誤', '找不到「支出記錄」工作表。', ui.ButtonSet.OK);
    return;
  }

  const dataCount = expensesSheet.getLastRow() - 1;

  // 第一次確認
  const response1 = ui.alert(
    '⚠️ 警告：即將重置系統',
    `目前有 ${dataCount} 筆支出記錄。\n\n` +
    '重置將會「永久刪除所有資料」！\n\n' +
    '強烈建議：\n' +
    '1. 先使用網頁版「匯出 CSV」備份\n' +
    '2. 或使用「檔案 → 建立副本」備份整個試算表\n\n' +
    '確定要繼續嗎？',
    ui.ButtonSet.YES_NO
  );

  if (response1 !== ui.Button.YES) {
    ui.alert('✅ 已取消重置');
    return;
  }

  // 第二次確認（最後機會）
  const response2 = ui.alert(
    '⚠️ 最後確認',
    '這是最後一次確認！\n\n' +
    `即將刪除 ${dataCount} 筆記錄，無法復原！\n\n` +
    '真的要繼續嗎？',
    ui.ButtonSet.YES_NO
  );

  if (response2 !== ui.Button.YES) {
    ui.alert('✅ 已取消重置');
    return;
  }

  // 執行重置
  expensesSheet.clear();
  const headers = ['日期', '項目', '金額', '付款人', '實際付款人', '你的部分', '對方的部分', '你實付', '對方實付', '分類', '是否週期', '週期日期', 'ID', '記錄類型'];
  expensesSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  expensesSheet.getRange(1, 1, 1, headers.length)
    .setBackground(CONFIG.COLORS.HEADER)
    .setFontColor('#ffffff')
    .setFontWeight('bold')
    .setHorizontalAlignment('center');

  const widths = [100, 150, 100, 100, 100, 100, 100, 100, 100, 80, 80, 80, 120, 100];
  widths.forEach((width, i) => expensesSheet.setColumnWidth(i + 1, width));

  expensesSheet.setFrozenRows(1);

  // 同樣重置週期設定和設定工作表
  const recurringSheet = ss.getSheetByName(CONFIG.SHEET_NAMES.RECURRING);
  if (recurringSheet) {
    recurringSheet.clear();
    const headers = ['啟用', '項目', '金額', '付款人', '你的部分', '對方的部分', '分類', '每月執行日', '備註'];
    recurringSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    recurringSheet.getRange(1, 1, 1, headers.length)
      .setBackground(CONFIG.COLORS.HEADER)
      .setFontColor('#ffffff')
      .setFontWeight('bold')
      .setHorizontalAlignment('center');
    recurringSheet.setFrozenRows(1);
  }

  const settingsSheet = ss.getSheetByName(CONFIG.SHEET_NAMES.SETTINGS);
  if (settingsSheet) {
    const owner = ss.getOwner().getEmail();
    settingsSheet.clear();
    const settings = [
      ['設定項目', '值'],
      ['你的名字', '你'],
      ['對方的名字', '對方'],
      ['預設分類', '飲食,居住,交通,娛樂,寵物,服飾,其他'],
      ['週期事件最後執行日期', ''],
      ['允許存取的使用者', owner],
      ['記帳模式', '共同記帳'],
      ['介面配色', '紫色']
    ];
    settingsSheet.getRange(1, 1, settings.length, 2).setValues(settings);
    settingsSheet.getRange(1, 1, 1, 2)
      .setBackground(CONFIG.COLORS.HEADER)
      .setFontColor('#ffffff')
      .setFontWeight('bold');
    settingsSheet.setColumnWidth(1, 200);
    settingsSheet.setColumnWidth(2, 400);
    settingsSheet.getRange('C6').setValue('多個使用者用逗號分隔，例如：user1@gmail.com, user2@gmail.com');
    settingsSheet.getRange('C6').setFontSize(9).setFontColor('#999999');
    settingsSheet.getRange('C7').setValue('個人記帳 / 共同記帳');
    settingsSheet.getRange('C7').setFontSize(9).setFontColor('#999999');
    settingsSheet.getRange('C8').setValue('紫色 / 藍色 / 綠色 / 粉色');
    settingsSheet.getRange('C8').setFontSize(9).setFontColor('#999999');

    // 快速記帳按鈕設定
    const quickExpenseHeaders = ['表情符號', '項目', '金額', '分類'];
    const quickExpenseData = [
      ['🍳', '早餐', 50, '飲食'],
      ['🍱', '午餐', 100, '飲食'],
      ['🍽️', '晚餐', 150, '飲食'],
      ['☕', '咖啡', 60, '飲食'],
      ['🚇', '交通', 20, '交通'],
      ['🅿️', '停車', 50, '交通'],
      ['🍰', '點心', 80, '飲食'],
      ['🧋', '飲料', 50, '飲食']
    ];

    settingsSheet.getRange(8, 1).setValue('快速記帳按鈕設定');
    settingsSheet.getRange(8, 1).setFontWeight('bold').setFontSize(11);
    settingsSheet.getRange(9, 1, 1, 4).setValues([quickExpenseHeaders]);
    settingsSheet.getRange(9, 1, 1, 4)
      .setBackground(CONFIG.COLORS.HEADER)
      .setFontColor('#ffffff')
      .setFontWeight('bold')
      .setHorizontalAlignment('center');

    settingsSheet.getRange(10, 1, quickExpenseData.length, 4).setValues(quickExpenseData);
    settingsSheet.getRange(10, 1, quickExpenseData.length, 4).setHorizontalAlignment('center');

    // 設定欄位寬度
    settingsSheet.setColumnWidth(1, 100);  // 表情符號
    settingsSheet.setColumnWidth(2, 120);  // 項目
    settingsSheet.setColumnWidth(3, 80);   // 金額
    settingsSheet.setColumnWidth(4, 100);  // 分類

    // 加入說明
    settingsSheet.getRange('A18').setValue('💡 提示：可以自由新增、修改或刪除快速記帳按鈕（最多 12 個）');
    settingsSheet.getRange('A18').setFontSize(9).setFontColor('#999999');
  }

  // 重置分類設定工作表
  const categoriesSheet = ss.getSheetByName(CONFIG.SHEET_NAMES.CATEGORIES);
  if (categoriesSheet) {
    ss.deleteSheet(categoriesSheet);
  }
  createCategoriesSheet(ss);

  ui.alert('✅ 重置完成！\n\n所有資料已清空，系統已重新初始化。');
}

// ==================== 選單 ====================

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('📊 記帳系統')
    .addItem('🚀 初始化系統（僅首次）', 'initializeSpreadsheet')
    .addItem('📱 開啟網頁版', 'openWebApp')
    .addSeparator()
    .addItem('🔄 升級到最新版本', 'upgradeToLatest')
    .addItem('📥 匯入 SettleUp CSV', 'importSettleUpCSV')
    .addSeparator()
    .addItem('🔄 手動執行週期事件', 'manualExecuteRecurring')
    .addItem('📈 查看統計資料', 'showStatistics')
    .addSeparator()
    .addItem('⚙️ 設定觸發器', 'setupTriggers')
    .addItem('⚠️ 重置系統（危險）', 'resetSystem')
    .addToUi();
}

function openWebApp() {
  const url = ScriptApp.getService().getUrl();
  const safeUrl = escapeHtml(url);

  const html = HtmlService.createHtmlOutput(
    `<p>複製以下連結在瀏覽器開啟：</p>
     <input type="text" value="${safeUrl}" style="width:100%;padding:10px;" onclick="this.select()">
     <p><small>點擊輸入框即可選取全部文字</small></p>`
  ).setWidth(500).setHeight(150);

  SpreadsheetApp.getUi().showModalDialog(html, '網頁版連結');
}

function showStatistics() {
  const stats = getStatistics(null, null);

  let categoryList = '';
  for (const [category, amount] of Object.entries(stats.categoryStats)) {
    categoryList += `\n${category}: ${amount.toLocaleString()}`;
  }

  const message = `📊 統計資料\n\n💰 總支出: ${stats.total.toLocaleString()}\n\n👤 你付了: ${stats.yourTotal.toLocaleString()}\n👤 對方付了: ${stats.partnerTotal.toLocaleString()}\n\n${stats.difference > 0 ? `✅ 換對方付: ${Math.abs(stats.difference).toLocaleString()}` : stats.difference < 0 ? `⚠️ 換你付: ${Math.abs(stats.difference).toLocaleString()}` : `✅ 已結清`}\n\n📈 分類統計:${categoryList}`;

  SpreadsheetApp.getUi().alert(message);
}

// ==================== SettleUp CSV 匯入功能 ====================

/**
 * 匯入 SettleUp CSV 檔案
 */
function importSettleUpCSV() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // 檢查是否已初始化
  const expensesSheet = ss.getSheetByName(CONFIG.SHEET_NAMES.EXPENSES);
  if (!expensesSheet) {
    ui.alert('❌ 錯誤', '請先執行「初始化系統」！', ui.ButtonSet.OK);
    return;
  }

  try {
    // 取得試算表所在資料夾
    const spreadsheetFile = DriveApp.getFileById(ss.getId());
    const parentFolders = spreadsheetFile.getParents();

    if (!parentFolders.hasNext()) {
      ui.alert('❌ 錯誤', '無法取得試算表所在資料夾', ui.ButtonSet.OK);
      return;
    }

    const folder = parentFolders.next();

    // 尋找 SettleUp_transactions 試算表
    let settleUpSpreadsheet = null;
    const files = folder.getFilesByType(MimeType.GOOGLE_SHEETS);
    const fileList = [];

    while (files.hasNext()) {
      const file = files.next();
      const fileName = file.getName();
      fileList.push(fileName);

      if (fileName.toLowerCase() === 'settleup_transactions') {
        settleUpSpreadsheet = SpreadsheetApp.openById(file.getId());
        break;
      }
    }

    if (!settleUpSpreadsheet) {
      ui.alert(
        '❌ 找不到試算表',
        '在此資料夾中找不到「SettleUp_transactions」試算表。\n\n' +
        '資料夾中的試算表：\n' + fileList.slice(0, 10).join('\n') +
        (fileList.length > 10 ? '\n... 還有 ' + (fileList.length - 10) + ' 個試算表' : '') +
        '\n\n請確認：\n' +
        '1. 已上傳 CSV 並轉換為 Google 試算表\n' +
        '2. 試算表名稱為「SettleUp_transactions」（不區分大小寫）',
        ui.ButtonSet.OK
      );
      return;
    }

    // 讀取試算表並提取所有名字
    const sheet = settleUpSpreadsheet.getSheets()[0];
    const data = sheet.getDataRange().getValues();

    if (data.length < 2) {
      ui.alert('❌ 錯誤', '試算表中沒有資料', ui.ButtonSet.OK);
      return;
    }

    // 從 "Who paid" 和 "For whom" 欄位提取所有名字
    const namesSet = new Set();

    // 過濾函數：判斷是否為有效的人名
    function isValidName(name) {
      if (!name || name.length === 0) return false;

      // 排除標題列
      if (name === 'Who paid' || name === 'For whom') return false;

      // 排除包含分號的多人選項（例如：「佩樺;零幻」）
      if (name.includes(';')) return false;

      // 排除包含特殊字符的項目（商品名、商家名等）
      const invalidPatterns = [
        /【.*】/,           // 包含【】的商品名
        /\[NT\$.*\]/,      // 包含價格標記
        /商家:/,           // 商家前綴
        /\$/,              // 包含金錢符號
        /http/,            // 包含網址
        /\d{3,}/,          // 包含3位以上連續數字
        /x\s*\d+/i,        // 包含 x1, x2 等數量標記
        /材料價|現作|皇家-/, // 商品相關關鍵字
        /森林|商店|超市|市場|企業|公司|店面/ // 商家相關關鍵字
      ];

      for (const pattern of invalidPatterns) {
        if (pattern.test(name)) return false;
      }

      // 名字長度應該合理（1-10個字）
      if (name.length > 10) return false;

      return true;
    }

    for (let i = 1; i < data.length; i++) {
      const whoPaid = String(data[i][0] || '').trim();
      const forWhom = String(data[i][3] || '').trim();

      if (whoPaid && isValidName(whoPaid)) {
        namesSet.add(whoPaid);
      }

      if (forWhom) {
        // "For whom" 可能包含多個名字（分號分隔）
        const names = forWhom.split(';').map(n => n.trim()).filter(n => n);
        names.forEach(name => {
          if (isValidName(name)) {
            namesSet.add(name);
          }
        });
      }
    }

    const names = Array.from(namesSet).sort();

    if (names.length === 0) {
      ui.alert('❌ 錯誤', '試算表中找不到任何名字', ui.ButtonSet.OK);
      return;
    }

    // 將名字和試算表ID儲存到快取，供 HTML 對話框使用
    // 不儲存完整資料，避免超過大小限制
    const cache = CacheService.getUserCache();
    cache.put('importNames', JSON.stringify(names), 300); // 5分鐘有效期
    cache.put('importSpreadsheetId', settleUpSpreadsheet.getId(), 300);

    // 顯示 HTML 對話框讓使用者選擇名字
    const template = HtmlService.createTemplateFromFile('nameSelector');
    template.names = names;
    const html = template.evaluate()
      .setWidth(400)
      .setHeight(280);

    ui.showModalDialog(html, '選擇你的名字');

  } catch (e) {
    ui.alert('❌ 錯誤', '匯入失敗：' + e.message, ui.ButtonSet.OK);
    Logger.log('匯入錯誤：' + e.toString());
  }
}

/**
 * 使用者從對話框選擇名字後，處理實際匯入
 * 由 nameSelector.html 呼叫
 */
function processImportWithName(myName) {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const expensesSheet = ss.getSheetByName(CONFIG.SHEET_NAMES.EXPENSES);

  try {
    // 從快取讀取試算表ID
    const cache = CacheService.getUserCache();
    const spreadsheetId = cache.get('importSpreadsheetId');

    if (!spreadsheetId) {
      throw new Error('快取資料已過期，請重新執行匯入');
    }

    // 重新讀取試算表資料
    const settleUpSpreadsheet = SpreadsheetApp.openById(spreadsheetId);
    const sheet = settleUpSpreadsheet.getSheets()[0];
    const data = sheet.getDataRange().getValues();

    Logger.log('開始匯入，使用者名字：' + myName);
    Logger.log('找到試算表，共 ' + data.length + ' 行');

    const expenses = [];
    let skippedTransfers = 0;
    let errors = [];

    // 從第 2 行開始（跳過標題列，索引 0 是標題）
    for (let i = 1; i < data.length; i++) {
      const row = data[i];

      // 跳過空行
      if (!row[0] && !row[1]) continue;

      try {
        const expense = parseSettleUpSheetRow(row, i + 1, myName);

        // transfer 類型（債務結算）轉為結算記錄
        if (expense.type === 'transfer') {
          skippedTransfers++;
          // 轉換為結算記錄格式
          expense.recordType = 'settlement';
          expense.category = '結算';
          expense.item = '[💰結算] ' + expense.item;
          // 結算記錄的分帳金額設為 0
          expense.yourPart = 0;
          expense.partnerPart = 0;
          expense.splitType = '';
        }

        expenses.push(expense);
      } catch (e) {
        // 只記錄前 10 個錯誤的詳細資訊
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
      exp.yourActualPaid || 0,  // 你實付
      exp.partnerActualPaid || 0,  // 對方實付
      exp.category,
      false, // isRecurring
      '', // recurringDay
      new Date().getTime() + Math.random(), // ID
      exp.recordType || 'expense' // 記錄類型
    ]);

    const lastRow = expensesSheet.getLastRow();
    expensesSheet.getRange(lastRow + 1, 1, dataToWrite.length, 14).setValues(dataToWrite);

    // 顯示結果
    let message = `✅ 匯入完成！\n\n` +
                  `✓ 成功匯入：${expenses.length} 筆記錄\n` +
                  `- 其中結算記錄：${skippedTransfers} 筆\n`;

    if (errors.length > 0) {
      message += `\n⚠️ 錯誤記錄（${errors.length} 筆）：\n` + errors.slice(0, 5).join('\n');
      if (errors.length > 5) {
        message += `\n... 還有 ${errors.length - 5} 筆錯誤`;
      }
    }

    ui.alert('📥 匯入結果', message, ui.ButtonSet.OK);

    // 清除快取
    cache.removeAll(['importNames', 'importSpreadsheetId']);

  } catch (e) {
    ui.alert('❌ 錯誤', '匯入失敗：' + e.message, ui.ButtonSet.OK);
    Logger.log('匯入錯誤：' + e.toString());
    throw e; // 回傳錯誤給 HTML 對話框
  }
}

/**
 * 解析 SettleUp 試算表的一行資料
 * @param {Array} row - 試算表的一行（陣列格式）
 * @param {number} rowNum - 行號（用於錯誤訊息）
 * @param {string} myName - 使用者在 SettleUp 中的名字
 */
function parseSettleUpSheetRow(row, rowNum, myName) {
  // 試算表格式：Who paid, Amount, Currency, For whom, Split amounts, Purpose, Category, Date & time, Exchange rate, Converted amount, Type, Receipt
  // 索引：        0          1        2         3         4              5        6         7            8              9                10     11

  if (row.length < 11) {
    throw new Error('欄位數量不足');
  }

  const whoPaid = String(row[0] || '').trim();
  const amountRaw = String(row[1] || '').trim();
  const forWhom = String(row[3] || '').trim();
  const splitAmounts = String(row[4] || '').trim();
  const purpose = String(row[5] || '支出').trim();
  const category = String(row[6] || '').trim() || autoDetectCategory(purpose);
  const dateTime = row[7]; // 可能是 Date 物件或字串
  const type = String(row[10] || '').trim();

  // 解析金額：可能是單一數字或分號分隔的多個數字
  let amount = 0;
  let actualPayments = {}; // 記錄每個人實際付了多少

  if (amountRaw.includes(';')) {
    // Amount 包含分號：例如 "178;230"
    const amounts = amountRaw.split(';').map(a => parseFloat(a.trim()) || 0);
    amount = amounts.reduce((sum, a) => sum + a, 0); // 總金額

    // 對應到 Who paid 的位置
    if (whoPaid.includes(';')) {
      const payers = whoPaid.split(';').map(p => p.trim());
      for (let i = 0; i < payers.length && i < amounts.length; i++) {
        actualPayments[payers[i]] = amounts[i];
      }
    }
  } else {
    // 單一金額
    amount = parseFloat(amountRaw) || 0;
  }

  // 判斷付款人
  let payer = '你';
  let isSplitPayment = false;

  if (Object.keys(actualPayments).length > 0) {
    // 有 actualPayments 資料：表示多人分別付款
    isSplitPayment = true;

    // 檢查你付了多少
    const yourPaid = actualPayments[myName] || 0;
    const totalPaid = Object.values(actualPayments).reduce((sum, a) => sum + a, 0);
    const partnerPaid = totalPaid - yourPaid;

    if (yourPaid > 0 && partnerPaid > 0) {
      payer = '共同';
    } else if (yourPaid > 0) {
      payer = '你';
    } else {
      payer = '對方';
    }
  } else if (whoPaid.includes(';')) {
    // Who paid 包含多人但 Amount 沒有分號：取第一個人
    const firstPayer = whoPaid.split(';')[0].trim();
    if (type === 'transfer') {
      if (firstPayer !== myName) {
        payer = '對方';
      }
    } else {
      if (firstPayer !== myName) {
        payer = '對方';
      }
    }
  } else if (type === 'transfer') {
    // Transfer: whoPaid 墊付給 forWhom
    if (whoPaid !== myName) {
      payer = '對方';
    }
  } else {
    // 一般支出：單人付款
    if (whoPaid !== myName) {
      payer = '對方';
    }
  }

  // 解析分帳方式
  const splitInfo = parseSplitInfo(forWhom, splitAmounts, amount, whoPaid, myName, isSplitPayment, actualPayments);

  // 轉換日期格式
  let date;
  if (dateTime instanceof Date) {
    // 如果是 Date 物件，直接格式化
    date = Utilities.formatDate(dateTime, Session.getScriptTimeZone(), 'yyyy/M/d');
  } else {
    // 如果是字串，解析後格式化（2021-07-01 13:39:55 → 2021/7/1）
    const dateStr = String(dateTime).split(' ')[0];
    date = dateStr.replace(/-/g, '/').replace(/^(\d{4})\/0?(\d+)\/0?(\d+)$/, '$1/$2/$3');
  }

  return {
    date: date,
    item: purpose,
    amount: amount,
    category: category,
    payer: payer,
    splitType: splitInfo.splitType,
    yourPart: splitInfo.yourPart,
    partnerPart: splitInfo.partnerPart,
    yourRatio: splitInfo.yourRatio,
    partnerRatio: splitInfo.partnerRatio,
    actualPayer: payer,
    yourActualPaid: splitInfo.yourActualPaid || 0,
    partnerActualPaid: splitInfo.partnerActualPaid || 0,
    type: type
  };
}

/**
 * 解析分帳資訊
 * @param {string} myName - 使用者在 SettleUp 中的名字
 * @param {boolean} isSplitPayment - 是否為分開付款（Who paid 包含多人）
 * @param {object} actualPayments - 實際付款金額對照表 {名字: 金額}
 */
function parseSplitInfo(forWhom, splitAmounts, totalAmount, whoPaid, myName, isSplitPayment, actualPayments) {
  const people = forWhom.split(';').map(p => p.trim());
  const amounts = splitAmounts.split(';').map(a => parseFloat(a) || 0);

  // 判斷誰實際付款
  let yourActualPaid = 0;
  let partnerActualPaid = 0;

  if (actualPayments && Object.keys(actualPayments).length > 0) {
    // 有實際付款資料：使用 actualPayments（Amount 包含分號的情況）
    yourActualPaid = actualPayments[myName] || 0;

    // 計算對方付的總額（所有不是你的人）
    for (const [name, paid] of Object.entries(actualPayments)) {
      if (name !== myName) {
        partnerActualPaid += paid;
      }
    }
  } else if (whoPaid && whoPaid.includes(';')) {
    // Who paid 包含多人，但沒有 actualPayments
    // 假設：各自付款 = 各自應付（按比例分攤）
    yourActualPaid = null;
    partnerActualPaid = null;
  } else if (whoPaid) {
    // 單人付款
    yourActualPaid = (whoPaid === myName) ? totalAmount : 0;
    partnerActualPaid = (whoPaid !== myName) ? totalAmount : 0;
  } else {
    // 沒有 whoPaid 資訊，預設為你付款
    yourActualPaid = totalAmount;
    partnerActualPaid = 0;
  }

  // 只有一個人 → 100% 金額分帳
  if (people.length === 1) {
    if (people[0] === myName) {
      const yPart = totalAmount;
      const pPart = 0;
      return {
        splitType: '金額',
        yourPart: yPart,
        partnerPart: pPart,
        yourRatio: '',
        partnerRatio: '',
        yourActualPaid: yourActualPaid !== null ? yourActualPaid : yPart,
        partnerActualPaid: partnerActualPaid !== null ? partnerActualPaid : pPart
      };
    } else {
      const yPart = 0;
      const pPart = totalAmount;
      return {
        splitType: '金額',
        yourPart: yPart,
        partnerPart: pPart,
        yourRatio: '',
        partnerRatio: '',
        yourActualPaid: yourActualPaid !== null ? yourActualPaid : yPart,
        partnerActualPaid: partnerActualPaid !== null ? partnerActualPaid : pPart
      };
    }
  }

  // 兩個人 - 檢查是否均分
  if (people.length === 2 && amounts.length === 2) {
    const diff = Math.abs(amounts[0] - amounts[1]);

    // 均分（差距小於 0.1）
    if (diff < 0.1) {
      // 找出你應分和對方應分的金額
      const yourIndex = people.indexOf(myName);
      const yPart = yourIndex >= 0 ? amounts[yourIndex] : totalAmount / 2;
      const pPart = totalAmount - yPart;

      return {
        splitType: '自動均分',
        yourPart: yPart,
        partnerPart: pPart,
        yourRatio: '',
        partnerRatio: '',
        yourActualPaid: yourActualPaid !== null ? yourActualPaid : yPart,
        partnerActualPaid: partnerActualPaid !== null ? partnerActualPaid : pPart
      };
    }

    // 不等額 - 金額分帳
    const yourIndex = people.indexOf(myName);
    let partnerIndex = -1;

    // 找到對方的索引（不是自己的那個）
    for (let i = 0; i < people.length; i++) {
      if (people[i] !== myName) {
        partnerIndex = i;
        break;
      }
    }

    const yPart = yourIndex >= 0 ? amounts[yourIndex] : 0;
    const pPart = partnerIndex >= 0 ? amounts[partnerIndex] : 0;

    return {
      splitType: '金額',
      yourPart: yPart,
      partnerPart: pPart,
      yourRatio: '',
      partnerRatio: '',
      yourActualPaid: yourActualPaid !== null ? yourActualPaid : yPart,
      partnerActualPaid: partnerActualPaid !== null ? partnerActualPaid : pPart
    };
  }

  // 預設均分
  const halfAmount = totalAmount / 2;
  return {
    splitType: '自動均分',
    yourPart: halfAmount,
    partnerPart: halfAmount,
    yourRatio: '',
    partnerRatio: '',
    yourActualPaid: yourActualPaid !== null ? yourActualPaid : halfAmount,
    partnerActualPaid: partnerActualPaid !== null ? partnerActualPaid : halfAmount
  };
}

/**
 * CSV 解析器（處理引號和空欄位）
 */
function parseCSVLine(line) {
  const result = [];
  let current = '';
  let inQuotes = false;

  for (let i = 0; i < line.length; i++) {
    const char = line[i];

    if (char === '"') {
      inQuotes = !inQuotes;
    } else if (char === ',' && !inQuotes) {
      // 保留空字串，不要 trim 掉空欄位
      result.push(current.replace(/^"/, '').replace(/"$/, '').trim());
      current = '';
    } else {
      current += char;
    }
  }

  // 最後一個欄位
  result.push(current.replace(/^"/, '').replace(/"$/, '').trim());

  return result;
}

/**
 * 根據項目名稱自動偵測分類
 */
function autoDetectCategory(item) {
  const keywords = {
    '飲食': ['早餐', '午餐', '晚餐', '點心', '飲料', '咖啡', '麥當勞', '全聯', '吐司', '泡泡冰', '水煎包'],
    '交通': ['車票', '交通', '停車', '台鐵', '高鐵', '捷運', 'Uber', '油錢'],
    '居住': ['房租', '水費', '電費', '瓦斯', '網路', '日用品', '衛生紙'],
    '娛樂': ['書籍', '課程', 'APP', 'Hahow', '旅遊', '電影', '遊戲'],
    '服飾': ['衣服', '鞋子', '包包', '美髮', '美容', '化妝', '保養'],
    '其他': ['醫療', '保險', '稅']
  };

  for (const [category, words] of Object.entries(keywords)) {
    for (const word of words) {
      if (item.includes(word)) {
        return category;
      }
    }
  }

  return '其他';
}

// ==================== 結算功能 ====================

/**
 * 新增結算記錄
 * @param {string} direction - 結算方向：'partner_pay_me' 或 'i_pay_partner'
 * @param {number} amount - 結算金額
 * @param {string} date - 結算日期 (yyyy-mm-dd)
 * @param {string} note - 備註（選填）
 */
function addSettlement(direction, amount, date, note) {
  // 檢查頻率限制
  checkRateLimit('addSettlement');

  // 驗證輸入
  if (!validateNumber(amount, 0.01, 9999999)) {
    throw new Error('金額無效（必須介於 0.01 到 9,999,999）');
  }

  if (!['partner_pay_me', 'i_pay_partner'].includes(direction)) {
    throw new Error('結算方向無效');
  }

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEET_NAMES.EXPENSES);
  const id = new Date().getTime();

  // 根據方向設定項目名稱
  let item = '';
  if (direction === 'partner_pay_me') {
    item = '[💰結算] 對方還款';
  } else {
    item = '[💰結算] 我還款';
  }

  if (note) {
    item += ' - ' + escapeHtml(note.trim());
  }

  // 結算記錄的欄位
  const row = [
    date || Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd'),
    item,
    amount,
    direction === 'partner_pay_me' ? '對方' : '你',  // 付款人（誰給錢）
    direction === 'partner_pay_me' ? '對方' : '你',  // 實際付款人
    0,  // 你的部分
    0,  // 對方的部分
    0,  // 你實付
    0,  // 對方實付
    '結算',  // 分類
    false,  // 是否週期
    '',  // 週期日期
    id,
    'settlement'  // 記錄類型：結算
  ];

  sheet.appendRow(row);

  const lastRow = sheet.getLastRow();
  sheet.getRange(lastRow, 1, 1, 14).setHorizontalAlignment('center');

  // 設定特殊背景色（淺綠色）
  sheet.getRange(lastRow, 1, 1, 14).setBackground('#d1fae5');

  return {
    success: true,
    message: '結算記錄已新增'
  };
}

/**
 * 清空所有支出記錄（保留標題列）
 * ⚠️ 警告：此操作會刪除所有記錄，無法復原！
 */
function clearAllExpenses() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEET_NAMES.EXPENSES);

  if (!sheet) {
    Logger.log('找不到支出記錄工作表');
    return;
  }

  const lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    // 清除內容而不是刪除行（保留格式和公式）
    const numCols = sheet.getLastColumn();
    sheet.getRange(2, 1, lastRow - 1, numCols).clearContent();
    Logger.log('✅ 已清空 ' + (lastRow - 1) + ' 筆記錄的內容');
  } else {
    Logger.log('⚠️ 沒有記錄可以清空');
  }
}

/**
 * 診斷函數：分析支出記錄的分帳狀況
 * 在 Apps Script 編輯器中執行此函數可以看到詳細統計
 */
function diagnoseExpenseData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEET_NAMES.EXPENSES);

  if (!sheet) {
    Logger.log('找不到支出記錄工作表');
    return;
  }

  const data = sheet.getDataRange().getValues();
  const headers = data[0];

  // 找出欄位索引
  const colIndex = {
    item: headers.indexOf('項目'),
    amount: headers.indexOf('金額'),
    yourPart: headers.indexOf('你的部分'),
    partnerPart: headers.indexOf('對方的部分'),
    yourActualPaid: headers.indexOf('你實付'),
    partnerActualPaid: headers.indexOf('對方實付'),
    recordType: headers.indexOf('記錄類型')
  };

  let totalExpenses = 0;
  let totalSettlements = 0;
  let yourPartSum = 0;
  let partnerPartSum = 0;
  let yourActualSum = 0;
  let partnerActualSum = 0;
  let emptyPartCount = 0; // yourPart 和 partnerPart 都是空的記錄數

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const recordType = row[colIndex.recordType];

    if (recordType === 'settlement') {
      totalSettlements++;
      continue;
    }

    totalExpenses++;

    const yourPart = row[colIndex.yourPart];
    const partnerPart = row[colIndex.partnerPart];
    const yourActual = Number(row[colIndex.yourActualPaid]) || 0;
    const partnerActual = Number(row[colIndex.partnerActualPaid]) || 0;

    yourPartSum += Number(yourPart) || 0;
    partnerPartSum += Number(partnerPart) || 0;
    yourActualSum += yourActual;
    partnerActualSum += partnerActual;

    // 檢查是否為空
    if ((yourPart === '' || yourPart === null || yourPart === undefined) &&
        (partnerPart === '' || partnerPart === null || partnerPart === undefined)) {
      emptyPartCount++;
      if (i < 5) { // 顯示前幾筆空記錄
        Logger.log('空分帳記錄範例 ' + i + ': ' + row[colIndex.item] + ', 金額: ' + row[colIndex.amount]);
      }
    }
  }

  Logger.log('========== 診斷結果 ==========');
  Logger.log('支出記錄總數: ' + totalExpenses);
  Logger.log('結算記錄總數: ' + totalSettlements);
  Logger.log('');
  Logger.log('你應付總額: ' + yourPartSum.toFixed(2));
  Logger.log('對方應付總額: ' + partnerPartSum.toFixed(2));
  Logger.log('應付總和: ' + (yourPartSum + partnerPartSum).toFixed(2));
  Logger.log('');
  Logger.log('你實付總額: ' + yourActualSum.toFixed(2));
  Logger.log('對方實付總額: ' + partnerActualSum.toFixed(2));
  Logger.log('實付總和: ' + (yourActualSum + partnerActualSum).toFixed(2));
  Logger.log('');
  Logger.log('分帳為空的記錄數: ' + emptyPartCount + ' (' + (emptyPartCount / totalExpenses * 100).toFixed(1) + '%)');
  Logger.log('============================');

  return {
    totalExpenses: totalExpenses,
    totalSettlements: totalSettlements,
    yourPartSum: yourPartSum,
    partnerPartSum: partnerPartSum,
    yourActualSum: yourActualSum,
    partnerActualSum: partnerActualSum,
    emptyPartCount: emptyPartCount
  };
}
