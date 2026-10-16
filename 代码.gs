// =================================================================
// 主菜单与主流程控制
// =================================================================

function onOpen() {
  SpreadsheetApp.getUi()
      .createMenu('投资工具')
      .addItem('【一键执行】每日更新全流程', 'runDailyPortfolioUpdate')
      .addItem('【初始化】收益视图与曲线数据', 'setupPerformanceView')
      .addToUi();
}

function runDailyPortfolioUpdate() {
  const doc = getSpreadsheet_();
  ensurePerformanceSheets_(doc);
  fillTradeDerivedFields_(doc);
  fillPositionsFromTrades_(doc);
  const symbolInfoMap = updatePositionsPrices_(doc);
  fillTradesNameFromSymbol_(doc, symbolInfoMap);
  sortPositionsRows_(doc);

  // Fetch live FX once per run (throws rather than silently defaulting to
  // 1:1 — see getFxToCnyMap_). The sync below detects trades recorded late
  // and recomputes the affected CashLedger/NAV_Daily history on its own.
  const fxMap = getFxToCnyMap_(doc);
  syncCashAndNavHistory_(doc, fxMap);
}
