const puppeteer = require('puppeteer-core');
const fs = require('fs');
const path = require('path');

const bookAssetsDir = path.join(__dirname, 'book', 'assets');
const screenshotsDir = path.join(__dirname, 'screenshots');
if (!fs.existsSync(bookAssetsDir)) fs.mkdirSync(bookAssetsDir, { recursive: true });

async function main() {
  const browser = await puppeteer.launch({
    executablePath: '/Applications/Brave Browser.app/Contents/MacOS/Brave Browser',
    headless: true,
    args: ['--no-sandbox', '--disable-setuid-sandbox', '--hide-scrollbars']
  });

  const page = await browser.newPage();

  // Helper to render HTML content and screenshot
  async function renderHtmlAndSave(html, width, height, outputPath) {
    await page.setViewport({ width, height, deviceScaleFactor: 2 });
    await page.setContent(html, { waitUntil: 'domcontentloaded' });
    await new Promise(r => setTimeout(r, 300));
    await page.screenshot({ path: outputPath });
    console.log('Generated:', path.basename(outputPath));
  }

  // =========================================================================
  // Asset 00: 00_deploy_step.png (Google Apps Script Deployment Dialog)
  // =========================================================================
  const html00 = `
  <!DOCTYPE html>
  <html>
  <head>
    <meta charset="utf-8">
    <style>
      * { box-sizing: border-box; font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, "Helvetica Neue", sans-serif; }
      body { margin: 0; padding: 40px; background: #e8eaed; display: flex; justify-content: center; align-items: center; min-height: 100vh; }
      .dialog {
        width: 680px; background: #fff; border-radius: 8px;
        box-shadow: 0 8px 24px rgba(0,0,0,0.18), 0 2px 6px rgba(0,0,0,0.1);
        overflow: hidden; border: 1px solid #dadce0;
      }
      .header {
        padding: 20px 24px 16px; border-bottom: 1px solid #e0e0e0;
        display: flex; justify-content: space-between; align-items: center;
      }
      .header h2 { margin: 0; font-size: 20px; font-weight: 500; color: #202124; }
      .header-close { color: #5f6368; font-size: 24px; cursor: pointer; line-height: 1; }
      .content { display: flex; min-height: 380px; }
      .sidebar {
        width: 220px; border-right: 1px solid #e0e0e0; background: #f8f9fa; padding: 16px;
      }
      .sidebar-label { font-size: 11px; font-weight: 600; color: #5f6368; text-transform: uppercase; margin-bottom: 12px; letter-spacing: 0.5px; }
      .sidebar-item {
        display: flex; align-items: center; gap: 10px; padding: 10px 12px; border-radius: 6px;
        background: #e8f0fe; color: #1967d2; font-size: 14px; font-weight: 500;
      }
      .form-pane { flex: 1; padding: 24px 28px; }
      .form-group { margin-bottom: 22px; }
      .form-label { display: block; font-size: 13px; font-weight: 600; color: #3c4043; margin-bottom: 6px; }
      .form-input {
        width: 100%; padding: 10px 12px; border: 1px solid #dadce0; border-radius: 4px;
        font-size: 14px; color: #202124; outline: none; background: #fff;
      }
      .form-select {
        width: 100%; padding: 10px 12px; border: 1px solid #dadce0; border-radius: 4px;
        font-size: 14px; color: #202124; background: #fff; cursor: pointer; outline: none;
      }
      .highlight-box {
        position: relative; border: 2px solid #ea4335 !important; border-radius: 6px; padding: 2px;
        background: #fdf2f2;
      }
      .badge-note {
        position: absolute; right: 8px; top: -11px; background: #ea4335; color: white;
        font-size: 11px; padding: 2px 8px; border-radius: 10px; font-weight: 600;
      }
      .footer {
        padding: 16px 24px; border-top: 1px solid #e0e0e0; display: flex; justify-content: flex-end; gap: 12px;
        background: #f8f9fa;
      }
      .btn { padding: 9px 24px; border-radius: 4px; font-size: 14px; font-weight: 500; cursor: pointer; border: none; }
      .btn-cancel { background: transparent; color: #1a73e8; border: 1px solid #dadce0; }
      .btn-primary { background: #1a73e8; color: #fff; box-shadow: 0 1px 2px rgba(0,0,0,0.15); }
    </style>
  </head>
  <body>
    <div class="dialog">
      <div class="header">
        <h2>新部署 (New deployment)</h2>
        <span class="header-close">×</span>
      </div>
      <div class="content">
        <div class="sidebar">
          <div class="sidebar-label">選取類型</div>
          <div class="sidebar-item">
            <span>⚙️</span> 網頁應用程式
          </div>
        </div>
        <div class="form-pane">
          <div class="form-group">
            <label class="form-label">說明 (Description)</label>
            <input type="text" class="form-input" value="正式版 v1.0" readonly>
          </div>
          <div class="form-group">
            <label class="form-label">網頁應用程式 (Web App) 設定</label>
            <div style="font-size: 12px; color: #5f6368; margin-bottom: 12px;">指定誰能以何種身分存取此記帳服務</div>
          </div>
          
          <div class="form-group">
            <label class="form-label">執行身分 (Execute as)：</label>
            <div class="highlight-box">
              <span class="badge-note">【關鍵設定 1】請選「我」</span>
              <select class="form-select" style="border: none;">
                <option selected>我 (your-email@gmail.com)</option>
              </select>
            </div>
            <div style="font-size: 11px; color: #5f6368; margin-top: 4px;">由您的帳號權限安全讀寫 Google 試算表</div>
          </div>

          <div class="form-group">
            <label class="form-label">具有存取權限的人 (Who has access)：</label>
            <div class="highlight-box">
              <span class="badge-note">【關鍵設定 2】請選「所有人」</span>
              <select class="form-select" style="border: none;">
                <option selected>所有人 (Anyone)</option>
              </select>
            </div>
            <div style="font-size: 11px; color: #5f6368; margin-top: 4px;">方便共同記帳成員無痛透過手機網址存取</div>
          </div>
        </div>
      </div>
      <div class="footer">
        <button class="btn btn-cancel">取消</button>
        <button class="btn btn-primary">部署 (Deploy)</button>
      </div>
    </div>
  </body>
  </html>
  `;
  await renderHtmlAndSave(html00, 780, 640, path.join(bookAssetsDir, '00_deploy_step.png'));

  // =========================================================================
  // Asset 01: 01_mobile_interface.png (Direct copy of mobile shared overview)
  // =========================================================================
  const mobileOverviewSrc = path.join(screenshotsDir, '01_Mobile_Shared_Overview.png');
  if (fs.existsSync(mobileOverviewSrc)) {
    fs.copyFileSync(mobileOverviewSrc, path.join(bookAssetsDir, '01_mobile_interface.png'));
    console.log('Generated: 01_mobile_interface.png');
  }

  // =========================================================================
  // Asset 02: 02_theme_styles.png (4 Themes Comparison Grid)
  // =========================================================================
  const html02 = `
  <!DOCTYPE html>
  <html>
  <head>
    <meta charset="utf-8">
    <style>
      * { box-sizing: border-box; font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, "PingFang TC", "Microsoft JhengHei", sans-serif; }
      body { margin: 0; padding: 40px; background: #f3f4f6; }
      .title-banner { text-align: center; margin-bottom: 30px; }
      .title-banner h1 { margin: 0 0 8px 0; font-size: 26px; color: #111827; }
      .title-banner p { margin: 0; color: #6b7280; font-size: 14px; }
      .grid { display: grid; grid-template-columns: repeat(2, 1fr); gap: 24px; max-width: 900px; margin: 0 auto; }
      .theme-card {
        border-radius: 16px; overflow: hidden; box-shadow: 0 10px 25px rgba(0,0,0,0.08);
        border: 1px solid rgba(0,0,0,0.06); background: white;
      }
      .card-top { padding: 22px 20px; color: white; }
      .theme-purple .card-top { background: linear-gradient(135deg, #667eea 0%, #764ba2 100%); }
      .theme-green .card-top { background: linear-gradient(135deg, #10b981 0%, #059669 100%); }
      .theme-pink .card-top { background: linear-gradient(135deg, #ec4899 0%, #db2777 100%); }
      .theme-blue .card-top { background: linear-gradient(135deg, #3b82f6 0%, #1e40af 100%); }
      
      .mini-header { display: flex; justify-content: space-between; align-items: center; margin-bottom: 14px; }
      .mini-title { font-size: 16px; font-weight: 700; display: flex; align-items: center; gap: 6px; }
      .mini-badge { font-size: 11px; padding: 3px 8px; border-radius: 20px; background: rgba(255,255,255,0.25); font-weight: 500; }
      
      .mini-balance {
        background: rgba(255,255,255,0.95); border-radius: 12px; padding: 14px 16px; color: #1f2937;
        box-shadow: 0 4px 12px rgba(0,0,0,0.06); text-align: center;
      }
      .mini-balance-title { font-size: 11px; color: #6b7280; margin-bottom: 3px; font-weight: 500; }
      .mini-balance-val { font-size: 19px; font-weight: 800; color: #059669; }
      
      .card-body { padding: 18px 20px; background: #fff; }
      .quick-tags { display: flex; gap: 8px; flex-wrap: wrap; margin-bottom: 14px; }
      .tag {
        font-size: 12px; padding: 6px 12px; border-radius: 8px; font-weight: 600; color: white;
      }
      .theme-purple .tag { background: #667eea; }
      .theme-green .tag { background: #10b981; }
      .theme-pink .tag { background: #ec4899; }
      .theme-blue .tag { background: #3b82f6; }
      
      .theme-label-bar {
        display: flex; justify-content: space-between; align-items: center; border-top: 1px solid #f3f4f6;
        padding-top: 12px; font-size: 13px; font-weight: 600;
      }
      .theme-name { color: #374151; display: flex; align-items: center; gap: 6px; }
      .theme-desc { color: #9ca3af; font-size: 12px; font-weight: 400; }
    </style>
  </head>
  <body>
    <div class="title-banner">
      <h1>🎨 四款精心調配風格主題</h1>
      <p>點擊右上角設定即可即時切換個人裝置專屬配色，隨心打造最舒適的記帳視覺體驗</p>
    </div>
    <div class="grid">
      <!-- 1. 紫色 -->
      <div class="theme-card theme-purple">
        <div class="card-top">
          <div class="mini-header">
            <div class="mini-title">💑 共同記帳</div>
            <div class="mini-badge">經典預設</div>
          </div>
          <div class="mini-balance">
            <div class="mini-balance-title">💰 結算狀態</div>
            <div class="mini-balance-val">小涵 應付你 $2,074</div>
          </div>
        </div>
        <div class="card-body">
          <div class="quick-tags">
            <div class="tag">🍳 早餐 $80</div>
            <div class="tag">☕ 咖啡 $120</div>
            <div class="tag">🍱 午餐 $150</div>
          </div>
          <div class="theme-label-bar">
            <span class="theme-name">🟣 優雅紫 (Purple)</span>
            <span class="theme-desc">優雅溫暖 · 科技質感</span>
          </div>
        </div>
      </div>

      <!-- 2. 綠色 -->
      <div class="theme-card theme-green">
        <div class="card-top">
          <div class="mini-header">
            <div class="mini-title">💑 共同記帳</div>
            <div class="mini-badge">護眼推薦</div>
          </div>
          <div class="mini-balance">
            <div class="mini-balance-title">💰 結算狀態</div>
            <div class="mini-balance-val">小涵 應付你 $2,074</div>
          </div>
        </div>
        <div class="card-body">
          <div class="quick-tags">
            <div class="tag">🍳 早餐 $80</div>
            <div class="tag">☕ 咖啡 $120</div>
            <div class="tag">🍱 午餐 $150</div>
          </div>
          <div class="theme-label-bar">
            <span class="theme-name">🟢 清新綠 (Green)</span>
            <span class="theme-desc">自然護眼 · 舒暢明亮</span>
          </div>
        </div>
      </div>

      <!-- 3. 粉色 -->
      <div class="theme-card theme-pink">
        <div class="card-top">
          <div class="mini-header">
            <div class="mini-title">💑 共同記帳</div>
            <div class="mini-badge">活力熱情</div>
          </div>
          <div class="mini-balance">
            <div class="mini-balance-title">💰 結算狀態</div>
            <div class="mini-balance-val">小涵 應付你 $2,074</div>
          </div>
        </div>
        <div class="card-body">
          <div class="quick-tags">
            <div class="tag">🍳 早餐 $80</div>
            <div class="tag">☕ 咖啡 $120</div>
            <div class="tag">🍱 午餐 $150</div>
          </div>
          <div class="theme-label-bar">
            <span class="theme-name">💗 活力粉紅 (Pink)</span>
            <span class="theme-desc">熱情甜美 · 溫暖陪伴</span>
          </div>
        </div>
      </div>

      <!-- 4. 藍色 -->
      <div class="theme-card theme-blue">
        <div class="card-top">
          <div class="mini-header">
            <div class="mini-title">💑 共同記帳</div>
            <div class="mini-badge">俐落沉穩</div>
          </div>
          <div class="mini-balance">
            <div class="mini-balance-title">💰 結算狀態</div>
            <div class="mini-balance-val">小涵 應付你 $2,074</div>
          </div>
        </div>
        <div class="card-body">
          <div class="quick-tags">
            <div class="tag">🍳 早餐 $80</div>
            <div class="tag">☕ 咖啡 $120</div>
            <div class="tag">🍱 午餐 $150</div>
          </div>
          <div class="theme-label-bar">
            <span class="theme-name">🔵 商務藍 (Blue)</span>
            <span class="theme-desc">沉穩專業 · 簡約純粹</span>
          </div>
        </div>
      </div>
    </div>
  </body>
  </html>
  `;
  await renderHtmlAndSave(html02, 980, 720, path.join(bookAssetsDir, '02_theme_styles.png'));

  // =========================================================================
  // Asset 03: 03_currency_project.png (Travel Foreign Currency & Project Tag)
  // =========================================================================
  const html03 = `
  <!DOCTYPE html>
  <html>
  <head>
    <meta charset="utf-8">
    <style>
      * { box-sizing: border-box; font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, "PingFang TC", sans-serif; }
      body { margin: 0; padding: 40px; background: #e0e7ff; display: flex; justify-content: center; align-items: center; min-height: 100vh; }
      .phone-frame {
        width: 420px; background: #ffffff; border-radius: 28px;
        box-shadow: 0 16px 36px rgba(79, 70, 229, 0.2), 0 2px 8px rgba(0,0,0,0.06);
        overflow: hidden; border: 4px solid #fff;
      }
      .phone-header {
        background: linear-gradient(135deg, #4f46e5 0%, #7c3aed 100%); padding: 24px 20px 20px; color: white;
      }
      .top-nav { display: flex; justify-content: space-between; align-items: center; margin-bottom: 12px; }
      .app-title { font-size: 17px; font-weight: 700; display: flex; align-items: center; gap: 8px; }
      .project-badge {
        display: inline-flex; align-items: center; gap: 6px; padding: 6px 14px;
        background: rgba(255,255,255,0.22); border-radius: 20px; font-size: 13px; font-weight: 600;
        border: 1px solid rgba(255,255,255,0.35);
      }
      .form-container { padding: 22px 20px; }
      .form-title { font-size: 16px; font-weight: 700; color: #1e1b4b; margin-bottom: 16px; display: flex; align-items: center; gap: 6px; }
      .field-group { margin-bottom: 16px; }
      .field-label { display: block; font-size: 12px; font-weight: 600; color: #4b5563; margin-bottom: 6px; }
      .input-box {
        width: 100%; padding: 11px 14px; border: 1.5px solid #e5e7eb; border-radius: 10px;
        font-size: 14px; color: #111827; background: #f9fafb; font-weight: 500;
      }
      .currency-row { display: flex; gap: 10px; }
      .currency-select {
        width: 140px; padding: 11px; border: 1.5px solid #4f46e5; border-radius: 10px;
        background: #eef2ff; color: #4338ca; font-size: 14px; font-weight: 700;
      }
      .exchange-hint {
        margin-top: 8px; padding: 10px 14px; background: #f5f3ff; border: 1.5px dashed #8b5cf6;
        border-radius: 10px; font-size: 13px; color: #6d28d9; display: flex; justify-content: space-between; align-items: center;
      }
      .split-card {
        padding: 12px 14px; background: #f0fdf4; border: 1px solid #bbf7d0; border-radius: 10px;
        display: flex; justify-content: space-between; align-items: center; font-size: 13px; color: #166534; font-weight: 600;
      }
      .submit-btn {
        width: 100%; padding: 14px; background: linear-gradient(135deg, #4f46e5 0%, #7c3aed 100%);
        color: white; border: none; border-radius: 12px; font-size: 15px; font-weight: 700;
        cursor: pointer; box-shadow: 0 4px 14px rgba(79, 70, 229, 0.35); margin-top: 10px;
      }
    </style>
  </head>
  <body>
    <div class="phone-frame">
      <div class="phone-header">
        <div class="top-nav">
          <div class="app-title">✈️ 出國外幣記帳模式</div>
          <span style="font-size: 12px; opacity: 0.8;">2026/09/03</span>
        </div>
        <div class="project-badge">
          🏷️ 專案標籤：2026東京自由行
        </div>
      </div>
      
      <div class="form-container">
        <div class="form-title">
          <span>💴 新增出國消費</span>
        </div>

        <div class="field-group">
          <label class="field-label">消費項目</label>
          <div class="input-box">淺草寺雷門御守與精選手工紀念品</div>
        </div>

        <div class="field-group">
          <label class="field-label">幣別與外幣金額</label>
          <div class="currency-row">
            <div class="currency-select">🇯🇵 JPY 日圓</div>
            <div class="input-box" style="flex: 1; font-weight: 700; font-size: 16px;">12,000</div>
          </div>
          
          <div class="exchange-hint">
            <span>💱 實時匯率 <strong>0.2100</strong></span>
            <span style="font-weight: 700;">折合 NT$ 2,520</span>
          </div>
        </div>

        <div class="field-group">
          <label class="field-label">付款帳戶與分帳方式</label>
          <div class="input-box" style="margin-bottom: 8px;">💳 玉山富利卡（海外刷卡 3% 回饋）</div>
          <div class="split-card">
            <span>👥 雙人均分 (50% / 50%)</span>
            <span>各負擔 NT$ 1,260</span>
          </div>
        </div>

        <button class="submit-btn">✔ 儲存至「2026東京自由行」帳本</button>
      </div>
    </div>
  </body>
  </html>
  `;
  await renderHtmlAndSave(html03, 520, 680, path.join(bookAssetsDir, '03_currency_project.png'));

  // =========================================================================
  // Asset 04: 04_recurring_settings.png (Recurring Expenses Spreadsheet View)
  // =========================================================================
  const html04 = `
  <!DOCTYPE html>
  <html>
  <head>
    <meta charset="utf-8">
    <style>
      * { box-sizing: border-box; font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, "PingFang TC", sans-serif; }
      body { margin: 0; padding: 30px; background: #f8f9fa; }
      .sheets-container {
        max-width: 960px; margin: 0 auto; background: #fff; border-radius: 8px;
        box-shadow: 0 4px 20px rgba(0,0,0,0.08); border: 1px solid #dadce0; overflow: hidden;
      }
      .sheets-header {
        background: #107c41; color: white; padding: 12px 18px; display: flex; align-items: center; justify-content: space-between;
      }
      .sheets-title { font-size: 16px; font-weight: 600; display: flex; align-items: center; gap: 8px; }
      .table-wrapper { overflow-x: auto; }
      table { width: 100%; border-collapse: collapse; font-size: 13px; text-align: left; }
      th {
        background: #f1f3f4; color: #3c4043; padding: 10px 14px; font-weight: 600;
        border: 1px solid #dadce0; text-align: center;
      }
      td { padding: 11px 14px; border: 1px solid #e8eaed; color: #202124; }
      tr:nth-child(even) td { background: #fdfdfd; }
      tr:hover td { background: #e8f0fe; }
      .check-cell { text-align: center; color: #107c41; font-weight: bold; font-size: 15px; }
      .money-cell { text-align: right; font-family: "SF Mono", Menlo, Consolas, monospace; font-weight: 600; }
      .center-cell { text-align: center; }
      .badge-cat {
        display: inline-block; padding: 3px 8px; border-radius: 12px; font-size: 11px; font-weight: 500;
      }
      .cat-live { background: #e0f2fe; color: #0369a1; }
      .cat-play { background: #fce7f3; color: #be185d; }
      .sheets-footer-tabs {
        background: #f8f9fa; border-top: 1px solid #dadce0; display: flex; padding: 0 10px;
      }
      .tab {
        padding: 9px 18px; font-size: 13px; color: #5f6368; border-bottom: 3px solid transparent;
        display: flex; align-items: center; gap: 6px; font-weight: 500;
      }
      .tab.active {
        color: #107c41; border-bottom: 3px solid #107c41; background: #fff; font-weight: 600;
      }
    </style>
  </head>
  <body>
    <div class="sheets-container">
      <div class="sheets-header">
        <div class="sheets-title">
          <span>📊</span> 共同記帳本 — 「週期設定」工作表
        </div>
        <span style="font-size: 12px; opacity: 0.9;">自動於每月指定日期觸發記帳</span>
      </div>

      <div class="table-wrapper">
        <table>
          <thead>
            <tr>
              <th style="width: 60px;">啟用</th>
              <th>項目名稱</th>
              <th>總金額</th>
              <th style="width: 80px;">代墊人</th>
              <th>你應負擔</th>
              <th>對方負擔</th>
              <th>分類</th>
              <th style="width: 100px;">扣款日</th>
            </tr>
          </thead>
          <tbody>
            <tr>
              <td class="check-cell">✔ TRUE</td>
              <td style="font-weight: 600;">每月房租與大樓管理費</td>
              <td class="money-cell">$24,000</td>
              <td class="center-cell">你</td>
              <td class="money-cell">$12,000</td>
              <td class="money-cell">$12,000</td>
              <td class="center-cell"><span class="badge-cat cat-live">🏠 居住>家具</span></td>
              <td class="center-cell" style="font-weight: 600; color: #1a73e8;">每月 1 號</td>
            </tr>
            <tr>
              <td class="check-cell">✔ TRUE</td>
              <td style="font-weight: 600;">居家光纖寬頻網路</td>
              <td class="money-cell">$699</td>
              <td class="center-cell">你</td>
              <td class="money-cell">$349.5</td>
              <td class="money-cell">$349.5</td>
              <td class="center-cell"><span class="badge-cat cat-live">🏠 居住>網路</span></td>
              <td class="center-cell" style="font-weight: 600; color: #1a73e8;">每月 10 號</td>
            </tr>
            <tr>
              <td class="check-cell">✔ TRUE</td>
              <td style="font-weight: 600;">Netflix 家庭高級方案</td>
              <td class="money-cell">$390</td>
              <td class="center-cell">對方</td>
              <td class="money-cell">$195</td>
              <td class="money-cell">$195</td>
              <td class="center-cell"><span class="badge-cat cat-play">🎮 娛樂>遊戲</span></td>
              <td class="center-cell" style="font-weight: 600; color: #1a73e8;">每月 15 號</td>
            </tr>
            <tr>
              <td class="check-cell">✔ TRUE</td>
              <td style="font-weight: 600;">天然氣瓦斯定期基本費</td>
              <td class="money-cell">$420</td>
              <td class="center-cell">對方</td>
              <td class="money-cell">$210</td>
              <td class="money-cell">$210</td>
              <td class="center-cell"><span class="badge-cat cat-live">🏠 居住>水電</span></td>
              <td class="center-cell" style="font-weight: 600; color: #1a73e8;">每月 20 號</td>
            </tr>
          </tbody>
        </table>
      </div>

      <div class="sheets-footer-tabs">
        <div class="tab">📋 支出記錄</div>
        <div class="tab active">🔄 週期設定</div>
        <div class="tab">⚙️ 設定</div>
        <div class="tab">🏷️ 分類設定</div>
        <div class="tab">💳 付款帳戶</div>
      </div>
    </div>
  </body>
  </html>
  `;
  await renderHtmlAndSave(html04, 1020, 390, path.join(bookAssetsDir, '04_recurring_settings.png'));

  // =========================================================================
  // Asset 05: 05_charts_report.png (Direct copy of Mobile Chart Modal)
  // =========================================================================
  const chartModalSrc = path.join(screenshotsDir, '04_Mobile_Chart_Modal_Category.png');
  if (fs.existsSync(chartModalSrc)) {
    fs.copyFileSync(chartModalSrc, path.join(bookAssetsDir, '05_charts_report.png'));
    console.log('Generated: 05_charts_report.png');
  }

  // =========================================================================
  // Asset 06: 06_sheets_backend.png (Google Sheets Database Backend Panorama)
  // =========================================================================
  const html06 = `
  <!DOCTYPE html>
  <html>
  <head>
    <meta charset="utf-8">
    <style>
      * { box-sizing: border-box; font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, "PingFang TC", sans-serif; }
      body { margin: 0; padding: 25px; background: #e8eaed; }
      .window {
        max-width: 1080px; margin: 0 auto; background: #fff; border-radius: 8px;
        box-shadow: 0 10px 30px rgba(0,0,0,0.15); border: 1px solid #dadce0; overflow: hidden;
      }
      .top-menu-bar {
        background: #f8f9fa; border-bottom: 1px solid #dadce0; padding: 8px 16px;
        display: flex; align-items: center; gap: 16px; font-size: 13px; color: #3c4043;
      }
      .sheet-title-row {
        display: flex; align-items: center; gap: 10px; font-weight: 600; font-size: 15px; color: #107c41;
      }
      .menu-items { display: flex; gap: 14px; font-size: 13px; position: relative; }
      .menu-item { cursor: pointer; padding: 4px 8px; border-radius: 4px; }
      .menu-item.active { background: #e8f0fe; color: #1a73e8; font-weight: 600; }
      
      .dropdown-menu {
        position: absolute; top: 32px; left: 0; width: 230px; background: white; border: 1px solid #dadce0;
        border-radius: 6px; box-shadow: 0 4px 16px rgba(0,0,0,0.15); padding: 6px 0; z-index: 10;
      }
      .drop-item {
        padding: 8px 16px; font-size: 13px; color: #202124; display: flex; justify-content: space-between; align-items: center;
      }
      .drop-item.highlight {
        background: #e8f0fe; color: #1a73e8; font-weight: 600; border-left: 3px solid #1a73e8;
      }
      
      table { width: 100%; border-collapse: collapse; font-size: 12px; }
      th {
        background: #f1f3f4; color: #5f6368; padding: 8px 12px; border: 1px solid #dadce0; text-align: center;
        font-weight: 600;
      }
      td { padding: 9px 12px; border: 1px solid #e8eaed; color: #202124; }
      tr:nth-child(even) td { background: #fdfdfd; }
      .settle-row td { background: #ecfdf5 !important; font-weight: 600; color: #065f46; }
      
      .bottom-tabs {
        background: #f8f9fa; border-top: 1px solid #dadce0; display: flex; padding: 0 12px;
      }
      .tab {
        padding: 9px 18px; font-size: 13px; color: #5f6368; border-bottom: 3px solid transparent;
        display: flex; align-items: center; gap: 6px; font-weight: 500;
      }
      .tab.active {
        color: #107c41; border-bottom: 3px solid #107c41; background: #fff; font-weight: 600;
      }
    </style>
  </head>
  <body>
    <div class="window">
      <!-- 頂部選單列 -->
      <div class="top-menu-bar">
        <div class="sheet-title-row">
          <span>📗</span> 共同記帳系統 (Shared Expense Database)
        </div>
        <div class="menu-items">
          <div class="menu-item active">檔案 (File)</div>
          <div class="menu-item">編輯</div>
          <div class="menu-item">檢視</div>
          <div class="menu-item">插入</div>
          <div class="menu-item">資料</div>
          <div class="menu-item">擴充功能</div>
          
          <!-- 模擬開啟下載選單 -->
          <div class="dropdown-menu">
            <div class="drop-item">共用與共編...</div>
            <div class="drop-item">建立副本</div>
            <div style="border-top: 1px solid #e0e0e0; margin: 4px 0;"></div>
            <div class="drop-item highlight">
              <span>下載 (Download) ➔</span>
            </div>
            <div style="padding-left: 24px; font-size: 12px; color: #1a73e8; padding-top: 4px; padding-bottom: 4px;">
              • Microsoft Excel (.xlsx)<br>
              • 逗號分隔值 (.csv)
            </div>
            <div style="border-top: 1px solid #e0e0e0; margin: 4px 0;"></div>
            <div class="drop-item">版本歷程記錄</div>
          </div>
        </div>
      </div>

      <!-- 資料表格 -->
      <table>
        <thead>
          <tr>
            <th style="width: 90px;">日期</th>
            <th>消費項目</th>
            <th style="width: 80px;">金額</th>
            <th style="width: 70px;">付款人</th>
            <th style="width: 80px;">你實付</th>
            <th style="width: 80px;">對方實付</th>
            <th style="width: 80px;">你負擔</th>
            <th style="width: 80px;">對方負擔</th>
            <th>分類</th>
            <th>付款帳戶</th>
          </tr>
        </thead>
        <tbody>
          <tr>
            <td style="text-align: center;">2026-09-03</td>
            <td>週末手作早午餐食材</td>
            <td style="text-align: right; font-weight: 600;">$480</td>
            <td style="text-align: center;">你</td>
            <td style="text-align: right;">$480</td>
            <td style="text-align: right;">$0</td>
            <td style="text-align: right;">$240</td>
            <td style="text-align: right;">$240</td>
            <td>🍜 飲食>早餐</td>
            <td>💵 日常現金</td>
          </tr>
          <tr>
            <td style="text-align: center;">2026-09-02</td>
            <td>日系超市生鮮與日常補給</td>
            <td style="text-align: right; font-weight: 600;">$2,180</td>
            <td style="text-align: center;">對方</td>
            <td style="text-align: right;">$0</td>
            <td style="text-align: right;">$2,180</td>
            <td style="text-align: right;">$1,090</td>
            <td style="text-align: right;">$1,090</td>
            <td>🍜 飲食>晚餐</td>
            <td>💳 玉山富利卡</td>
          </tr>
          <tr>
            <td style="text-align: center;">2026-09-01</td>
            <td>米其林推薦日式定食</td>
            <td style="text-align: right; font-weight: 600;">$1,280</td>
            <td style="text-align: center;">兩人</td>
            <td style="text-align: right;">$800</td>
            <td style="text-align: right;">$480</td>
            <td style="text-align: right;">$640</td>
            <td style="text-align: right;">$640</td>
            <td>🍜 飲食>午餐</td>
            <td>💳 國泰信用卡</td>
          </tr>
          <tr class="settle-row">
            <td style="text-align: center;">2026-08-25</td>
            <td>[💰結算] 對方還款結清</td>
            <td style="text-align: right;">$1,500</td>
            <td style="text-align: center;">對方</td>
            <td style="text-align: right;">$0</td>
            <td style="text-align: right;">$1,500</td>
            <td style="text-align: right;">$0</td>
            <td style="text-align: right;">$0</td>
            <td>💰 結算</td>
            <td>📱 LINE Pay</td>
          </tr>
        </tbody>
      </table>

      <!-- 底部 5 大工作表標籤 -->
      <div class="bottom-tabs">
        <div class="tab active">📋 支出記錄</div>
        <div class="tab">🔄 週期設定</div>
        <div class="tab">⚙️ 設定</div>
        <div class="tab">🏷️ 分類設定</div>
        <div class="tab">💳 付款帳戶</div>
      </div>
    </div>
  </body>
  </html>
  `;
  await renderHtmlAndSave(html06, 1100, 480, path.join(bookAssetsDir, '06_sheets_backend.png'));

  await browser.close();
  console.log('🎉 All 7 book assets generated successfully!');
}

main().catch(console.error);
