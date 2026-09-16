// =================================================================
// 主菜单与主流程控制
// =================================================================

function onOpen() {
  SpreadsheetApp.getUi()
      .createMenu('投资工具')
      .addItem('【一键执行】每日更新全流程', 'runDailyPortfolioUpdate')
      .addItem('【初始化】收益视图与曲线数据', 'setupPerformanceView')
      .addSeparator()
      .addItem('⚠️【全量重建】现金与NAV历史', 'rebuildCashAndNavHistory')
      .addToUi();
}

function runDailyPortfolioUpdate() {
  const doc = getSpreadsheet_();
  ensurePerformanceSheets_(doc);
  fillTradeDerivedFields_(doc);
  fillPositionsFromTrades_(doc);
  const symbolInfoMap = updatePositionsPrices_(doc);
  fillTradesNameFromSymbol_(doc, symbolInfoMap);

  // Fetch FX once per run and reuse it everywhere below, instead of each
  // consumer independently re-fetching (and each retry now throws rather
  // than silently defaulting to 1:1 — see getFxToCnyMap_).
  const fxMap = getFxToCnyMap_(doc);
  const fxLookup = currency => {
    if (currency === "CNY") return 1;
    if (fxMap[currency]) return fxMap[currency];
    Logger.log(`Unknown currency "${currency}", defaulting FX to 1.`);
    return 1;
  };

  fillCashFlowDerivedFields_(doc, fxLookup);
  appendTodayCashLedgerRow_(doc, fxMap);
  appendDailyNavSnapshot_(doc, fxMap);
  appendTodayToHistoryCaches_(doc, fxMap);
}
