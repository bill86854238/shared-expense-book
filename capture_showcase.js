/**
 * 自動截圖產生工具 (App Showcase Screenshot Generator)
 * 執行方式：node capture_showcase.js
 */

const puppeteer = require('puppeteer-core');
const http = require('http');
const fs = require('fs');
const path = require('path');

const screenshotsDir = path.join(__dirname, 'screenshots');
if (!fs.existsSync(screenshotsDir)) fs.mkdirSync(screenshotsDir, { recursive: true });

// 本地輕量 HTTP 伺服器
const server = http.createServer((req, res) => {
  const filePath = path.join(__dirname, 'index.html');
  const content = fs.readFileSync(filePath);
  res.writeHead(200, { 'Content-Type': 'text/html; charset=utf-8' });
  res.end(content);
});

server.listen(0, async () => {
  const port = server.address().port;
  console.log(`🌐 預覽伺服器已啟動於連接埠: ${port}`);

  // 自動尋找系統中的 Chromium 核心瀏覽器
  const browserPaths = [
    '/Applications/Brave Browser.app/Contents/MacOS/Brave Browser',
    '/Applications/Google Chrome.app/Contents/MacOS/Google Chrome',
    '/Applications/Microsoft Edge.app/Contents/MacOS/Microsoft Edge',
    '/Applications/Chromium.app/Contents/MacOS/Chromium'
  ];

  let executablePath = browserPaths.find(p => fs.existsSync(p));
  if (!executablePath) {
    console.error('❌ 未找到合適的 Chromium 瀏覽器');
    server.close();
    process.exit(1);
  }

  console.log(`🚀 啟動瀏覽器: ${executablePath}`);
  const browser = await puppeteer.launch({
    executablePath,
    headless: true,
    args: ['--no-sandbox', '--disable-setuid-sandbox', '--hide-scrollbars']
  });

  const page = await browser.newPage();

  async function cleanAndSave(filename) {
    await page.evaluate(() => {
      document.querySelectorAll('.toast').forEach(t => t.remove());
      const demoBtn = document.getElementById('demoToggleBtn');
      if (demoBtn) demoBtn.style.display = 'none';
    });
    await new Promise(r => setTimeout(r, 200));
    const localPath = path.join(screenshotsDir, filename);
    await page.screenshot({ path: localPath });
    console.log(`  📸 已儲存: screenshots/${filename}`);
  }

  console.log('📱 開始擷取行動端 (Mobile 390x844 Retina) 畫面...');
  await page.setViewport({ width: 390, height: 844, deviceScaleFactor: 2 });
  await page.goto(`http://localhost:${port}?demo=1`, { waitUntil: 'networkidle0' });
  await page.waitForSelector('.stat-card.balance');
  await new Promise(r => setTimeout(r, 800));

  // 1. 行動端 - 共同記帳首頁
  await cleanAndSave('01_mobile_shared_overview.png');

  // 2. 行動端 - 統計分析儀表板
  await page.evaluate(() => {
    const el = document.getElementById('dashboardPeriod').closest('div[style*="margin-bottom: 10px"]') || document.getElementById('dashboardPeriod');
    const y = el.getBoundingClientRect().top + window.pageYOffset - 65;
    window.scrollTo({ top: Math.max(0, y) });
  });
  await new Promise(r => setTimeout(r, 600));
  await cleanAndSave('02_mobile_dashboard.png');

  // 3. 行動端 - 歷史收支明細清單
  await page.evaluate(() => {
    const el = document.querySelector('.filter-section') || document.getElementById('expensesList');
    const y = el.getBoundingClientRect().top + window.pageYOffset - 14;
    window.scrollTo({ top: Math.max(0, y) });
  });
  await new Promise(r => setTimeout(r, 600));
  await cleanAndSave('03_mobile_expenses_list.png');

  // 4. 行動端 - 圖表分析彈窗 (分類佔比圓環圖)
  await page.evaluate(() => {
    window.scrollTo(0, 0);
    const modal = document.getElementById('chartModal');
    if (modal) {
      modal.style.background = 'rgba(15, 23, 42, 0.88)';
      modal.style.backdropFilter = 'blur(8px)';
    }
    showCharts();
  });
  await new Promise(r => setTimeout(r, 1200));
  await cleanAndSave('04_mobile_chart_modal_category.png');

  // 5. 行動端 - 圖表分析彈窗 (每日與月度趨勢)
  await page.evaluate(() => {
    switchChartTab(2);
  });
  await new Promise(r => setTimeout(r, 1000));
  await cleanAndSave('05_mobile_chart_modal_trend.png');

  // 6. 行動端 - 個人記帳模式首頁
  await page.evaluate(() => {
    closeChartModal();
    selectMode('個人記帳');
    window.scrollTo(0, 0);
  });
  await new Promise(r => setTimeout(r, 800));
  await cleanAndSave('06_mobile_personal_mode.png');

  // 7. 行動端 - 個人記帳統計分析
  await page.evaluate(() => {
    const el = document.getElementById('dashboardPeriod').closest('div[style*="margin-bottom: 10px"]') || document.getElementById('dashboardPeriod');
    const y = el.getBoundingClientRect().top + window.pageYOffset - 65;
    window.scrollTo({ top: Math.max(0, y) });
  });
  await new Promise(r => setTimeout(r, 600));
  await cleanAndSave('07_mobile_personal_dashboard.png');

  console.log('💻 開始擷取桌面端 (Desktop 1200x850 Retina) 畫面...');
  await page.setViewport({ width: 1200, height: 850, deviceScaleFactor: 2 });
  await page.evaluate(() => {
    selectMode('共同記帳');
    window.scrollTo(0, 0);
  });
  await new Promise(r => setTimeout(r, 800));
  await cleanAndSave('08_desktop_shared_overview.png');

  // 9. 桌面端 - 統計儀表板全景
  await page.evaluate(() => {
    const el = document.getElementById('dashboardPeriod').closest('div[style*="margin-bottom: 10px"]') || document.getElementById('dashboardPeriod');
    const y = el.getBoundingClientRect().top + window.pageYOffset - 50;
    window.scrollTo({ top: Math.max(0, y) });
  });
  await new Promise(r => setTimeout(r, 600));
  await cleanAndSave('09_desktop_dashboard_analytics.png');

  await browser.close();
  server.close();
  console.log('✨ 全套 9 張展示截圖已擷取完畢！');
});
