# 💑 Shared Expense Book

🌐 **Language / 語言 / 言語**: [繁體中文](README.md) | [English](README.en.md) | [日本語](README.ja.md)

> A free, private, and self-hosted shared expense tracker built on Google Apps Script and Google Sheets.  
> **$0 hosting costs, zero ads, and 100% data ownership.** Designed for couples, roommates, families, and personal budgeting.

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](./LICENSE)
[![Google Apps Script](https://img.shields.io/badge/Google%20Apps%20Script-v8-blue.svg)](https://developers.google.com/apps-script)
[![Version](https://img.shields.io/badge/Version-v2.60-blue.svg)](https://github.com/bill86854238/shared-expense-book/releases/tag/v2.60)
[![Security](https://img.shields.io/badge/Security-5%20Stars-brightgreen.svg)]()

---

## ✨ Key Features

- 💰 **100% Free & Ad-Free**: Runs entirely on your Google account. No subscriptions, server fees, or in-app ads.
- 🔒 **Complete Data Privacy**: All records live securely in your own Google Drive. No third-party servers ever touch your finances.
- 🎭 **Dual Mode (Personal vs. Shared)**:
  - **Personal Mode**: Track income and personal spending separately, with automatic monthly net savings calculation.
  - **Shared Mode**: Split bills evenly (50/50), by custom amounts, by percentages, or record 100% advance payments (reimbursements). Live balance calculation instantly shows who owes whom.
- ⚡ **One-Tap Quick Logging**: Preset buttons for routine purchases (Breakfast, Coffee, Transit, Groceries) to log expenses in 1 second.
- ✈️ **Travel & Foreign Currencies**: Real-time multi-currency exchange conversion with project tags for vacations.
- ⏰ **Automated Recurring Expenses**: Automatically logs recurring monthly bills (Rent, Utilities, Subscriptions) on schedule.
- 🎨 **4 Clean Themes**: Purple, Green, Coral, and Navy with persistent user preference memory.
- 📥 **Easy Data Import & Export**: Import history from SettleUp or AndroMoney, and export full backups to Excel (.xlsx) / CSV anytime.
- 🛡️ **Hardened Security**: Protected against XSS, clickjacking, and unauthorized access via Google OAuth email whitelist.

---

## 🎯 Use Cases

- **Couples & Spouses**: Transparent household budget sharing, groceries, and dining expenses.
- **Roommates**: Clean rent & utility splitting with hassle-free monthly settlements.
- **Family Budgets**: Centralized household finances and automated bill tracking.
- **Vacation Budgets**: Project-based travel expense tracking with live currency conversion.
- **Personal Finances**: Full monthly income, expenses, and savings monitoring.

---

## 🚀 Quick Start (Copy & Play)

Get started in **just 2 steps**:

1. **Copy the Template**:  
   👉 [Click here to make a copy of the Google Sheets template](https://docs.google.com/spreadsheets/d/1ZEXa0R0MGEMyG9W-Rh2t7bj1HCeM-gJ0chN7qqDw054/copy)
2. **Initialize & Deploy**:
   - Open your copied sheet, click the top menu **`📊 記帳系統` → `1️⃣ 第一步：初始化系統`** (automatically creates tabs and schedules daily triggers).
   - Go to **`Extensions` → `Apps Script` → Click top-right `Deploy` → `New deployment`** (Select Web app, Execute as: Me, Who has access: Anyone), and copy the generated Web App URL to open on your phone!

> 💡 **Note**: No need to manually enable People API or configure triggers. Everything is automated.

For troubleshooting and common questions, see [FAQ.md](./FAQ.md).

---

## 🎁 Deployment Toolkit (Optional Premium Pack)

For extra conveniences, an optional toolkit is available:
- **Profile Avatars**: Automatic Google avatar syncing (People API integration)
- **Smart Generators**: Email whitelist configuration generator
- **Customization Guides**: Step-by-step color theme and category styling guide
- **Dedicated Support**: 30 days of email technical assistance

👉 [Support this project on Gumroad ($3 USD) →](https://billions65.gumroad.com/l/kwvhy)

---

## ❓ FAQ

- **Q: Can other people see my financial data if they copy the template?**  
  **A:** No. Each user creates an independent copy stored entirely in their own private Google Drive.
- **Q: How do I invite my partner or roommate?**  
  **A:** In the `設定 (Settings)` tab of your sheet, add their Gmail address in cell `B6`, then share your Web App URL with them.
- **Q: Can I backup my data?**  
  **A:** Yes, anytime via `File → Download → Microsoft Excel (.xlsx)` or `CSV`.

---

## 📝 License

Distributed under the [MIT License](./LICENSE). Free for personal and commercial use.
