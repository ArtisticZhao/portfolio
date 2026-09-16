# Repository instructions

- This repository is synchronized with the Google Apps Script project bound to the `Portfolio` Google Sheets workbook through the user's deployment plugin.
- Treat the repository files as the source of truth for Apps Script code review and implementation.
- Make Apps Script code changes in this repository only. The user will deploy or synchronize those changes to Google Apps Script.
- Do not edit or deploy the bound Google Apps Script project through the browser unless the user explicitly asks for that action.
- Preserve compatibility with the existing `Portfolio` workbook, especially the `Trades`, `Positions`, `CashLedger`, `NAV_Daily`, and `PnL_View` sheets.
