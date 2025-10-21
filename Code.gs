/**
 * 情侶共同記帳系統 - Google Apps Script
 * 完整版本
 */

// ==================== 設定區 ====================
const CONFIG = {
  SHEET_NAMES: {
    EXPENSES: '支出記錄',
    RECURRING: '週期設定',
    SETTINGS: '設定'
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
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  createExpensesSheet(ss);
  createRecurringSheet(ss);
  createSettingsSheet(ss);
  setupTriggers();

  SpreadsheetApp.getUi().alert('✅ 初始化完成！\n\n已建立：\n1. 支出記錄\n2. 週期設定\n3. 設定\n\n並設定每日自動執行週期事件。');
}

function createExpensesSheet(ss) {
  let sheet = ss.getSheetByName(CONFIG.SHEET_NAMES.EXPENSES);
  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.SHEET_NAMES.EXPENSES);
  } else {
    sheet.clear();
  }

  const headers = ['日期', '項目', '金額', '付款人', '實際付款人', '你的部分', '對方的部分', '分類', '是否週期', '週期日期', 'ID'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

  sheet.getRange(1, 1, 1, headers.length)
    .setBackground(CONFIG.COLORS.HEADER)
    .setFontColor('#ffffff')
    .setFontWeight('bold')
    .setHorizontalAlignment('center');

  const widths = [100, 150, 100, 100, 100, 100, 100, 80, 80, 120];
  widths.forEach((width, i) => sheet.setColumnWidth(i + 1, width));

  sheet.setFrozenRows(1);
}

function createRecurringSheet(ss) {
  let sheet = ss.getSheetByName(CONFIG.SHEET_NAMES.RECURRING);
  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.SHEET_NAMES.RECURRING);
  } else {
    sheet.clear();
  }

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

function createSettingsSheet(ss) {
  let sheet = ss.getSheetByName(CONFIG.SHEET_NAMES.SETTINGS);
  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.SHEET_NAMES.SETTINGS);
  } else {
    sheet.clear();
  }

  const owner = ss.getOwner().getEmail();

  const settings = [
    ['設定項目', '值'],
    ['你的名字', '你'],
    ['對方的名字', '對方'],
    ['預設分類', '飲食,居住,交通,娛樂,其他'],
    ['週期事件最後執行日期', ''],
    ['允許存取的使用者', owner]
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
}

// ==================== 核心功能 ====================

function addExpense(item, amount, payer, actualPayer, yourPart, partnerPart, category, isRecurring, recurringDay) {
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
  const date = new Date();
  const id = date.getTime();

  // 過濾和轉義輸入
  const safeItem = escapeHtml(item.trim());
  const safeCategory = escapeHtml(category);

  const row = [
    Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy-MM-dd'),
    safeItem,
    amount,
    payer,
    actualPayer || payer,  // 實際付款人，向下相容
    yourPart,
    partnerPart,
    safeCategory,
    isRecurring || false,
    recurringDay || '',
    id
  ];

  sheet.appendRow(row);

  const lastRow = sheet.getLastRow();
  sheet.getRange(lastRow, 1, 1, 11).setHorizontalAlignment('center');

  let color = CONFIG.COLORS.BOTH;
  if (payer === '你') color = CONFIG.COLORS.YOUR;
  else if (payer === '對方') color = CONFIG.COLORS.PARTNER;

  sheet.getRange(lastRow, 1, 1, 11).setBackground(color);

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

  // 取得使用者照片 URL
  let photoUrl = '';
  try {
    // 使用 People API 取得使用者資訊
    const userInfo = People.People.get('people/me', {
      personFields: 'photos'
    });

    if (userInfo.photos && userInfo.photos.length > 0) {
      photoUrl = userInfo.photos[0].url;
    }
  } catch (e) {
    Logger.log('無法取得使用者照片: ' + e.toString());
    // 使用預設 Google 帳號圖示
    photoUrl = 'https://www.gstatic.com/images/branding/product/1x/avatar_circle_blue_512dp.png';
  }

  return {
    email: user,
    name: user.split('@')[0],
    photoUrl: photoUrl,
    allowed: permission.allowed
  };
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
      actualPayer: String(data[i][4] || data[i][3] || ''),  // 新增：實際付款人，向下相容
      yourPart: Number(data[i][5]) || 0,
      partnerPart: Number(data[i][6]) || 0,
      category: String(data[i][7] || '其他'),
      isRecurring: Boolean(data[i][8]),
      recurringDay: data[i][9] || '',
      id: String(data[i][10] || '')
    });
  }

  Logger.log('成功載入 ' + expenses.length + ' 筆支出記錄');
  return expenses;
}

function addExpenseFromWeb(expenseData) {
  return addExpense(
    expenseData.item,
    expenseData.amount,
    expenseData.payer,
    expenseData.actualPayer || expenseData.payer,  // 新增：實際付款人，向下相容
    expenseData.yourPart,
    expenseData.partnerPart,
    expenseData.category,
    expenseData.isRecurring,
    expenseData.recurringDay
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

  // 找到 ID 欄位（第 10 欄）並更新
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][9]) === String(updatedData.id)) {
      const oldItem = data[i][1];
      const oldAmount = data[i][2];

      // 過濾和轉義輸入
      const safeItem = escapeHtml(updatedData.item.trim());
      const safeCategory = escapeHtml(updatedData.category);

      // 更新資料（保留原有的日期和 ID）
      sheet.getRange(i + 1, 2).setValue(safeItem);           // 項目
      sheet.getRange(i + 1, 3).setValue(updatedData.amount); // 金額
      sheet.getRange(i + 1, 4).setValue(updatedData.payer);  // 付款人
      sheet.getRange(i + 1, 5).setValue(updatedData.yourPart);     // 你的部分
      sheet.getRange(i + 1, 6).setValue(updatedData.partnerPart);  // 對方的部分
      sheet.getRange(i + 1, 7).setValue(safeCategory);       // 分類

      // 更新背景顏色
      let color = CONFIG.COLORS.BOTH;
      if (updatedData.payer === '你') color = CONFIG.COLORS.YOUR;
      else if (updatedData.payer === '對方') color = CONFIG.COLORS.PARTNER;
      sheet.getRange(i + 1, 1, 1, 10).setBackground(color);

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

  // 找到 ID 欄位（第 10 欄）
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][9]) === String(id)) {
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

// ==================== 選單 ====================

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('📊 記帳系統')
    .addItem('🚀 初始化系統', 'initializeSpreadsheet')
    .addItem('📱 開啟網頁版', 'openWebApp')
    .addSeparator()
    .addItem('🔄 手動執行週期事件', 'manualExecuteRecurring')
    .addItem('📈 查看統計資料', 'showStatistics')
    .addSeparator()
    .addItem('⚙️ 設定觸發器', 'setupTriggers')
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
