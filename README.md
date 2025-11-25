# 💑 共同記帳系統

> 基於 Google Apps Script 的免費共享記帳工具，適合情侶、室友、家庭等多人使用。  
> 本版本已維持在穩定狀態，後續新功能（例如付款方式、更進階的自訂）將提供於另一個新版專案。

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)
[![Google Apps Script](https://img.shields.io/badge/Google%20Apps%20Script-v8-blue.svg)](https://developers.google.com/apps-script)
[![Version](https://img.shields.io/badge/Version-v2.56-blue.svg)](https://github.com/bill86854238/shared-expense-book/releases/tag/v2.56)
[![Security: 5/5](https://img.shields.io/badge/Security-⭐⭐⭐⭐⭐-brightgreen.svg)](./最終安全報告.md)

## ✨ 特色

- 💰 **免費使用**：基於 Apps Script，不需伺服器費用  
- 🔒 **安全保護**：具備 XSS、注入攻擊等完整防護  
- ☁️ **雲端同步**：資料儲存在個人 Google 試算表  
- 📱 **多裝置支援**：手機、平板、電腦皆可使用  
- 🎭 **雙模式記帳**：個人 / 共同記帳自由切換  
- 💵 **完整個人記帳**：支援收入 + 支出紀錄、自動計算淨收入  
- 💑 **共享記帳功能**：支援不等額付款、自動結算  
- 📥 **資料匯入**：相容 SettleUp、AndroMoney  
- 🎨 **主題選擇**：四種配色，自動記憶  
- 📊 **圖表化分析**：Google Charts 儀表板  
- 👥 **成員管理**：Google 登入、Email 邀請、權限控制  
- ⚡ **快速記帳**：常用支出一鍵帶入  
- 📈 **儀表板**：支援多時間範圍比較  
- 🏷️ **彈性分類**：主/子分類階層化管理，自訂分類

## 🎯 適用場景

- 情侶共同記帳  
- 室友生活費分帳  
- 朋友旅遊分費  
- 家庭日常開銷管理  
- 小型團隊共用費用  
- 個人財務紀錄

## 📸 主要功能

### 💰 記帳功能

**共同記帳模式：**

- 新增支出（付款人可選我／對方／各付各的）  
- 分帳方式支援：  
  - 自動均分  
  - 自訂金額分帳  
  - 百分比分帳  
- 不等額墊付、自動計算負擔額  
- 智能結算提醒  
- 週期性支出（房租、訂閱等）

**個人記帳模式：**

- 收入記錄（薪水、獎金、投資等 9 類）  
- 支出記錄  
- 自動計算淨收入  
- 收入/支出分開管理

**通用功能：**

- 編輯、刪除記錄  
- 多條件查詢  
- CSV 匯出  
- 資料匯入  
  - SettleUp  
  - AndroMoney

### 👥 成員管理

- Google 登入  
- 顯示頭像與名稱  
- Email 邀請成員  
- 成員列表管理  
- 移除成員  
- 白名單權限控制

### 📊 數據分析

**共同記帳：**

- 儀表板（本月、上月、近 7 天 / 30 天、月比較）  
- 支出分類分析  
- 趨勢圖  
- 付款方分析

**個人記帳：**

- 收支統計  
- 淨收入計算  
- 收支趨勢圖  
- 類別分析圖

### ⚡ 其他快速功能

- 快速記帳按鈕（早餐、咖啡、交通等）  
- 週期性自動記帳  
- 每日觸發器

### 🛡️ 安全防護

- 前後端輸入驗證  
- XSS 攻擊防護  
- 請求頻率限制  
- 操作日誌  
- 白名單權限  
- HTTPS

> 更多請見：[完整功能清單.md](./完整功能清單.md)

---

## 🚀 快速開始

1. 建立 Google 試算表  
2. 將 Code.gs、index.html、nameSelector.html 複製到 Apps Script  
3. 啟用 People API  
4. 部署為網頁應用程式  
5. 初始化並設定白名單  

如需更詳細教學，可參考 FAQ.md（38 個常見問題）或使用互動式部署工具包。

---

## 🎁 部署助手工具包（付費版）

若希望快速部署，可使用以下工具包：
- 互動式部署助手（4 步驟）  
- 白名單設定自動產生器  
- 主題客製化指南  
- 自訂分類教學  
- 30 天 Email 技術支援  

[前往 Gumroad 購買 →](https://billions65.gumroad.com/l/kwvhy)

---

> 核心功能完全開源免費；付費工具包提供更友善的部署體驗。
