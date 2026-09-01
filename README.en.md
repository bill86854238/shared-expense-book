# 💑 Shared Expense Book

🌐 **Language / 語言 / 言語**: [繁體中文](README.md) | [English](README.en.md) | [日本語](README.ja.md)

> A private, self-hosted shared expense tracker built on Google Apps Script and Google Sheets.  
> Features **$0 running cost, forever ad-free, and 100% data ownership**, crafted for couples, roommates, families, and personal budgeting.

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](./LICENSE)
[![Google Apps Script](https://img.shields.io/badge/Google%20Apps%20Script-v8-blue.svg)](https://developers.google.com/apps-script)
[![Version](https://img.shields.io/badge/Version-v2.60-blue.svg)](https://github.com/bill86854238/shared-expense-book/releases/tag/v2.60)
[![Security](https://img.shields.io/badge/Security-5%20Stars-brightgreen.svg)]()

---

## ✨ Key Features

- 💰 **100% Free & No Ads**: Powered entirely by Google Workspace. No servers or external databases needed.
- 🔒 **Complete Data Privacy**: Your data lives exclusively in your own Google Drive and is never sent to third-party servers.
- 🎭 **Dual Mode (Personal & Shared)**:
  - **Personal Mode**: Track income and expenses separately with automatic net balance calculation.
  - **Shared Mode**: Flexible bill splitting (50/50, custom amounts, percentages, or full advance payments) with smart debt balance calculation.
- ⚡ **1-Second Quick Log**: Customizable quick buttons (Breakfast, Coffee, Transport) for instant one-tap entry.
- ✈️ **Travel & Foreign Currencies**: Real-time multi-currency conversion with dedicated project tags for vacation budgeting.
- ⏰ **Recurring Auto-Expenses**: Schedule monthly recurring bills (Rent, Utilities, Subscriptions) automatically.
- 🎨 **4 Elegant Color Themes**: Purple, Green, Coral, and Navy, with automatic preference saving.
- 📥 **Seamless Data Import & Export**: Compatible with SettleUp and AndroMoney data; export full backups to Excel (.xlsx) / CSV anytime.
- 🛡️ **Production-Grade Security**: Protection against XSS, CSRF/clickjacking, rate limiting, and Google OAuth whitelist authorization.

---

## 🎯 Use Cases

- **Couples & Spouses**: Effortlessly split household bills, groceries, and dining expenses.
- **Roommates**: Transparent shared rent, utility splitting, and hassle-free settlements.
- **Family Budgets**: Centralized family spending management and recurring bills.
- **Travel & Vacations**: Project-based travel expense tracking with multi-currency conversion.
- **Personal Finances**: Full monthly income, expenses, and savings monitoring.

---

## 🚀 Quick Start (Copy & Play)

Deployment takes just **2 simple steps**:

1. **Make a Copy of the Template**:  
   👉 [Click here to copy the Google Sheets template](https://docs.google.com/spreadsheets/d/1ZEXa0R0MGEMyG9W-Rh2t7bj1HCeM-gJ0chN7qqDw054/copy)
2. **Initialize & Deploy**:
   - Open your copied sheet, click top menu **`📊 記帳系統 (Expense Book)` → `1️⃣ 第一步：初始化系統 (Initialize)`** (automatically creates tabs and schedules triggers).
   - Go to **`Extensions` → `Apps Script` → Click top-right `Deploy` → `New deployment`** (Select Web app, Execute as: Me, Who has access: Anyone), and copy the generated Web App URL to open on your phone!

> 💡 **Note**: The core architecture has been streamlined—**no need to manually enable People API or configure triggers**. Ready to use out of the box.

For common troubleshooting, check [FAQ.md](./FAQ.md).

---

## 🎁 Deployment Toolkit (Premium Pack)

For an enhanced experience, an optional toolkit is available:
- **Profile Avatars**: Automatic Google avatar syncing (People API integration)
- **Smart Generators**: Whitelist configuration helper
- **Customization Guides**: In-depth color theme and category styling guide
- **Dedicated Support**: 30 days of email technical assistance

👉 [Support this project on Gumroad ($3 USD) →](https://billions65.gumroad.com/l/kwvhy)

---

## ❓ FAQ

- **Q: Can other people see my financial data if they copy the template?**  
  **A:** No. Each user creates an independent copy stored entirely in their own Google Drive.
- **Q: How do I invite my partner or roommate?**  
  **A:** In the `設定 (Settings)` tab of your sheet, add their Gmail address in cell `B6`, then share your Web App URL with them.
- **Q: Can I backup my data?**  
  **A:** Yes, anytime via `File → Download → Microsoft Excel (.xlsx)` or `CSV`.

---

## 📝 License

Distributed under the [MIT License](./LICENSE). Free for personal and commercial use.
