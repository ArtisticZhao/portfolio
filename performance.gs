/**
 * Ensure performance-related sheets exist with expected headers.
 * Sheets:
 * - CashLedger: generated cash ledger from Trades + CashFlow (external flows)
 * - NAV_Daily: daily account snapshot in CNY
 * - PnL_View: dashboard formulas
 * - FxHistory / PriceHistory: cache sheets for historical rebuilds
 */
function ensurePerformanceSheets_(spreadsheet) {
  ensureSheetWithHeader_(spreadsheet, "CashLedger", [
    "Date",
    "TradeCashFlowCNY",
    "ExternalCashFlowCNY",
    "CashBalanceCNY",
    "Note"
  ]);

  ensureSheetWithHeader_(spreadsheet, "NAV_Daily", [
    "Date",
    "EquityValueCNY",
    "CashBalanceCNY",
    "ExternalCashFlowCNY",
    "TotalAssetCNY",
    "PrevTotalAssetCNY",
    "TradingPnL",
    "DailyReturn",
    "DataQuality"
  ]);

  ensureSheetWithHeader_(spreadsheet, "PnL_View", [
    "Metric",
    "Value",
    "Notes"
  ]);

  ensureSheetWithHeader_(spreadsheet, "FxHistory", [
    "Date",
    "USDCNY",
    "HKDCNY",
    "Source"
  ], true);

  ensureSheetWithHeader_(spreadsheet, "PriceHistory", [
    "Date",
    "Symbol",
    "Market",
    "ClosePrice",
    "Currency",
    "Source"
  ], true);

  setupPnLViewFormula_(spreadsheet);
}

function ensureSheetWithHeader_(spreadsheet, name, headers, hidden) {
  let sheet = spreadsheet.getSheetByName(name);
  let isNew = false;
  if (!sheet) {
    sheet = spreadsheet.insertSheet(name);
    isNew = true;
  }

  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  const current = headerRange.getValues()[0];
  let needsWrite = false;

  for (let i = 0; i < headers.length; i++) {
    if (current[i] !== headers[i]) {
      needsWrite = true;
      break;
    }
  }

  if (needsWrite) {
    headerRange.setValues([headers]);
  }

  if (isNew && hidden) sheet.hideSheet();

  return sheet;
}

/**
 * Capture one daily snapshot into NAV_Daily. Append-only: only ever
 * touches "today"'s row, historical rows are never rewritten here.
 * Snapshot date is today in spreadsheet timezone.
 * @param {Spreadsheet} spreadsheet
 * @param {Object} fxMap CNY->1, USD->rate, HKD->rate (see getFxToCnyMap_)
 */
function appendDailyNavSnapshot_(spreadsheet, fxMap) {
  const navSheet = spreadsheet.getSheetByName("NAV_Daily");
  const posSheet = spreadsheet.getSheetByName("Positions");
  if (!navSheet || !posSheet) {
    throw new Error('Missing sheet: "NAV_Daily" or "Positions"');
  }

  if (!isWeekdayTradingDay_(spreadsheet)) {
    Logger.log("Today is weekend, skip NAV snapshot for trading-day view.");
    return;
  }

  const today = getTodayDateOnly_(spreadsheet);
  const todayKey = formatDateKey_(today);

  const equityValueCNY = computeEquityValueCNY_(spreadsheet, fxMap);
  const cashBalanceCNY = getCashBalanceCNYForDate_(spreadsheet, today);
  const externalCashFlowCNY = getExternalCashFlowCNYForDate_(spreadsheet, today);
  const totalAssetCNY = equityValueCNY + cashBalanceCNY;

  // Upsert by date to avoid duplicate rows when running repeatedly on same day.
  const lastRow = navSheet.getLastRow();
  const rows = lastRow >= 2 ? navSheet.getRange(2, 1, lastRow - 1, 1).getValues() : [];

  let targetRow = -1;
  for (let i = 0; i < rows.length; i++) {
    const d = rows[i][0];
    if (!d) continue;
    if (formatDateKey_(d) === todayKey) {
      targetRow = i + 2;
      break;
    }
  }

  if (targetRow === -1) {
    targetRow = navSheet.getLastRow() + 1;
  }

  navSheet.getRange(targetRow, 1, 1, 5).setValues([[
    today,
    equityValueCNY,
    cashBalanceCNY,
    externalCashFlowCNY,
    totalAssetCNY
  ]]);

  // Fill derived columns F:G:H.
  navSheet.getRange(targetRow, 6).setFormula(
    `=IF(A${targetRow}="",,IF(ROW()=2,,INDEX(E$2:E,ROW()-2)))`
  );
  navSheet.getRange(targetRow, 7).setFormula(
    `=IF(A${targetRow}="",,IF(F${targetRow}="",,E${targetRow}-F${targetRow}-D${targetRow}))`
  );
  navSheet.getRange(targetRow, 8).setFormula(
    `=IF(A${targetRow}="",,IF(F${targetRow}="",,IF(F${targetRow}+D${targetRow}=0,,G${targetRow}/(F${targetRow}+D${targetRow}))))`
  );
  navSheet.getRange(targetRow, 9).setValue("LIVE");

  Logger.log(`NAV snapshot updated for ${todayKey}`);
}

/**
 * Append/upsert only "today"'s CashLedger row from the running balance of
 * the last existing row. Historical rows are never recomputed here — use
 * rebuildCashLedgerHistory_ for a deliberate full rebuild.
 * @return {number} new cash balance after today's flows
 */
function appendTodayCashLedgerRow_(spreadsheet, fxMap) {
  const sheet = spreadsheet.getSheetByName("CashLedger");
  if (!sheet) throw new Error('Sheet "CashLedger" not found');

  const today = getTodayDateOnly_(spreadsheet);
  const todayKey = formatDateKey_(today);
  const fxLookup = currency => {
    if (currency === "CNY") return 1;
    if (fxMap[currency]) return fxMap[currency];
    Logger.log(`Unknown currency "${currency}", defaulting FX to 1.`);
    return 1;
  };

  const tradeFlows = getDailyTradeCashFlowsCNY_(spreadsheet, fxLookup);
  const externalFlows = getDailyExternalCashFlowsCNY_(spreadsheet, fxLookup);

  const tradeFlowToday = roundMoney_(tradeFlows[todayKey] || 0);
  const externalFlowToday = roundMoney_(externalFlows[todayKey] || 0);

  const lastRow = sheet.getLastRow();
  const priorRows = lastRow >= 2 ? sheet.getRange(2, 1, lastRow - 1, 5).getValues() : [];

  let targetRow = -1;
  let priorBalance = 0;
  priorRows.forEach((row, idx) => {
    if (!row[0]) return;
    const rowKey = formatDateKey_(row[0]);
    if (rowKey === todayKey) {
      targetRow = idx + 2;
    } else if (rowKey < todayKey) {
      priorBalance = Number(row[3]) || 0;
    }
  });

  const newBalance = roundMoney_(priorBalance + tradeFlowToday + externalFlowToday);
  if (targetRow === -1) targetRow = lastRow + 1;

  sheet.getRange(targetRow, 1, 1, 5).setValues([[
    today,
    tradeFlowToday,
    externalFlowToday,
    newBalance,
    "Trade + external cash flow"
  ]]);

  Logger.log(`CashLedger row upserted for ${todayKey}: balance=${newBalance}`);
  return newBalance;
}

/**
 * Pure helper: walk-forward cumulative cash balance from per-date trade and
 * external cash flows. No AutoDeposit floor — balance is allowed to go
 * negative, which is a real signal (missing a deposit record), not
 * something to paper over.
 * @return {Array<{date, tradeFlow, externalFlow, cashBalance}>} sorted by date
 */
function walkCashBalances_(tradeFlowsByDate, externalFlowsByDate) {
  const allKeys = Array.from(
    new Set(Object.keys(tradeFlowsByDate).concat(Object.keys(externalFlowsByDate)))
  ).sort();

  let cash = 0;
  return allKeys.map(key => {
    const tradeFlow = roundMoney_(tradeFlowsByDate[key] || 0);
    const externalFlow = roundMoney_(externalFlowsByDate[key] || 0);
    cash = roundMoney_(cash + tradeFlow + externalFlow);
    return { date: key, tradeFlow, externalFlow, cashBalance: cash };
  });
}

/**
 * Deliberate full CashLedger rebuild using cached historical FX (not
 * "today"'s live rate applied to every past trade). Idempotent: same
 * Trades/CashFlow/FxHistory input always produces the same output.
 */
function rebuildCashLedgerHistory_(spreadsheet) {
  const sheet = spreadsheet.getSheetByName("CashLedger");
  if (!sheet) throw new Error('Sheet "CashLedger" not found');

  const fxLookup = makeHistoricalFxLookup_(loadFxHistoryMap_(spreadsheet));
  const tradeFlows = getDailyTradeCashFlowsCNY_(spreadsheet, fxLookup);
  const externalFlows = getDailyExternalCashFlowsCNY_(spreadsheet, fxLookup);

  const walked = walkCashBalances_(tradeFlows, externalFlows);

  const rows = walked.map(w => [
    new Date(w.date + "T00:00:00"),
    w.tradeFlow,
    w.externalFlow,
    w.cashBalance,
    "Rebuilt from history"
  ]);

  const lastRow = sheet.getLastRow();
  if (lastRow > 1) sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).clearContent();
  if (rows.length > 0) sheet.getRange(2, 1, rows.length, 5).setValues(rows);

  Logger.log(`CashLedger history rebuilt: ${rows.length} rows.`);
}

/**
 * Per-date trade cash flow in CNY, from Trades.
 * @param {Function} fxLookup (currency, dateKey) -> rate, or (currency) -> rate
 * @return {Object} dateKey("yyyy-MM-dd") -> tradeCashFlowCNY
 */
function getDailyTradeCashFlowsCNY_(spreadsheet, fxLookup) {
  const sheet = spreadsheet.getSheetByName("Trades");
  if (!sheet) return {};

  const values = sheet.getDataRange().getValues();
  if (values.length < 2) return {};

  const header = values[0];
  const col = name => header.indexOf(name);
  const iDate = col("Date");
  const iSide = col("Side");
  const iQty = col("Quantity");
  const iPrice = col("Price");
  const iCurrency = col("Currency");
  const iFee = col("Fee");
  const iTax = col("Tax");
  const iOtherCost = col("OtherCost");

  if ([iDate, iSide, iQty, iPrice, iCurrency].some(i => i < 0)) {
    throw new Error("Trades sheet must contain Date, Side, Quantity, Price, and Currency columns");
  }

  const daily = {};

  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    if (!row[iDate]) continue;

    const side = String(row[iSide] || "").trim().toUpperCase();
    const gross = (Number(row[iQty]) || 0) * (Number(row[iPrice]) || 0);
    if (!side || gross <= 0) continue;

    const fee = iFee >= 0 ? (Number(row[iFee]) || 0) : 0;
    const tax = iTax >= 0 ? (Number(row[iTax]) || 0) : 0;
    const otherCost = iOtherCost >= 0 ? (Number(row[iOtherCost]) || 0) : 0;
    const currency = String(row[iCurrency] || "CNY").trim().toUpperCase();
    const key = formatDateKey_(row[iDate]);
    const fx = fxLookup(currency, key);

    if (!daily[key]) daily[key] = 0;

    if (side === "BUY") {
      daily[key] -= (gross + fee + tax + otherCost) * fx;
    } else if (side === "SELL" || side === "DIVIDEND") {
      daily[key] += (gross - fee - tax - otherCost) * fx;
    }
  }

  return daily;
}

/**
 * Per-date external cash flow in CNY, from CashFlow (Deposit/Withdrawal).
 * This is the ONLY source of external cash flow now — no more AutoDeposit.
 * @param {Function} fxLookup (currency, dateKey) -> rate, or (currency) -> rate
 * @return {Object} dateKey("yyyy-MM-dd") -> externalCashFlowCNY
 */
function getDailyExternalCashFlowsCNY_(spreadsheet, fxLookup) {
  const sheet = spreadsheet.getSheetByName("CashFlow");
  if (!sheet) return {};

  const values = sheet.getDataRange().getValues();
  if (values.length < 2) return {};

  const header = values[0];
  const col = name => header.indexOf(name);
  const iDate = col("Date");
  const iType = col("Type");
  const iAmount = col("Amount");
  const iCurrency = col("Currency");

  if ([iDate, iType, iAmount, iCurrency].some(i => i < 0)) {
    throw new Error("CashFlow sheet must contain Date, Type, Amount, and Currency columns");
  }

  const daily = {};

  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    if (!row[iDate]) continue;

    const type = String(row[iType] || "").trim().toUpperCase();
    const amount = Number(row[iAmount]) || 0;
    if (!type || amount === 0) continue;

    const currency = String(row[iCurrency] || "CNY").trim().toUpperCase();
    const key = formatDateKey_(row[iDate]);
    const fx = fxLookup(currency, key);
    const signed = type === "WITHDRAWAL" ? -Math.abs(amount) * fx : Math.abs(amount) * fx;

    daily[key] = roundMoney_((daily[key] || 0) + signed);
  }

  return daily;
}

/**
 * Recompute CashFlow.AmountCNY from Amount/Currency using the unified FX
 * policy, every run (not just when blank) — keeps this column's CNY basis
 * consistent with everything else instead of trusting a manually entered
 * or stale value.
 * @param {Function} fxLookup (currency, dateKey) -> rate, or (currency) -> rate
 */
function fillCashFlowDerivedFields_(spreadsheet, fxLookup) {
  const sheet = spreadsheet.getSheetByName("CashFlow");
  if (!sheet) return;

  const range = sheet.getDataRange();
  const values = range.getValues();
  if (values.length < 2) return;

  const header = values[0];
  const col = name => header.indexOf(name);
  const iDate = col("Date");
  const iAmount = col("Amount");
  const iCurrency = col("Currency");
  const iAmountCNY = col("AmountCNY");

  if ([iDate, iAmount, iCurrency, iAmountCNY].some(i => i < 0)) {
    throw new Error("CashFlow sheet must contain Date, Amount, Currency, and AmountCNY columns");
  }

  const rows = values.slice(1);
  rows.forEach(row => {
    if (!row[iDate]) return;
    const currency = String(row[iCurrency] || "CNY").trim().toUpperCase();
    const key = formatDateKey_(row[iDate]);
    const fx = fxLookup(currency, key);
    row[iAmountCNY] = roundMoney_((Number(row[iAmount]) || 0) * fx);
  });

  range.setValues([header].concat(rows));
}

function getCashLedgerRows_(spreadsheet) {
  const sheet = spreadsheet.getSheetByName("CashLedger");
  if (!sheet) return [];

  const values = sheet.getDataRange().getValues();
  if (values.length < 2) return [];

  const header = values[0];
  const col = name => header.indexOf(name);
  const iDate = col("Date");
  const iExternalFlow = col("ExternalCashFlowCNY");
  const iBalance = col("CashBalanceCNY");
  if ([iDate, iExternalFlow, iBalance].some(i => i < 0)) {
    throw new Error("CashLedger sheet must contain Date, ExternalCashFlowCNY, and CashBalanceCNY columns");
  }

  return values.slice(1)
    .filter(row => row[iDate])
    .map(row => ({
      date: row[iDate],
      externalFlow: Number(row[iExternalFlow]) || 0,
      balance: Number(row[iBalance]) || 0
    }));
}

function getCashBalanceCNYForDate_(spreadsheet, dateObj) {
  const targetMs = dateOnlyMs_(dateObj);
  let bestMs = -Infinity;
  let balance = 0;

  getCashLedgerRows_(spreadsheet).forEach(row => {
    const ms = dateOnlyMs_(row.date);
    if (ms <= targetMs && ms > bestMs) {
      bestMs = ms;
      balance = row.balance;
    }
  });

  return balance;
}

function getExternalCashFlowCNYForDate_(spreadsheet, dateObj) {
  const target = formatDateKey_(dateObj);
  return getCashLedgerRows_(spreadsheet).reduce((sum, row) => {
    return formatDateKey_(row.date) === target ? sum + row.externalFlow : sum;
  }, 0);
}

function dateOnlyMs_(d) {
  const date = (d instanceof Date) ? d : new Date(d);
  return new Date(date.getFullYear(), date.getMonth(), date.getDate()).getTime();
}

/**
 * Reconstruct positions as of a given date by replaying Trades up to (and
 * including) that date. Reuses calculatePositionsFromTradeRows_ (trades.gs)
 * unchanged — just pre-filters the rows.
 */
function buildPositionsAsOfDate_(spreadsheet, asOfDate) {
  const tradesSheet = spreadsheet.getSheetByName("Trades");
  if (!tradesSheet) throw new Error('Sheet "Trades" not found');

  const values = tradesSheet.getDataRange().getValues();
  if (values.length < 2) return {};

  const header = values[0];
  const iDate = header.indexOf("Date");
  const asOfMs = dateOnlyMs_(asOfDate);

  const rows = values.slice(1).filter(row => row[iDate] && dateOnlyMs_(row[iDate]) <= asOfMs);
  return calculatePositionsFromTradeRows_(header, rows);
}

/**
 * Current (live) equity value in CNY, from Positions' current LastPrice.
 * @param {Object} [fxMap] result of getFxToCnyMap_; fetched if omitted.
 */
function computeEquityValueCNY_(spreadsheet, fxMap) {
  const sheet = spreadsheet.getSheetByName("Positions");
  if (!sheet) throw new Error('Sheet "Positions" not found');

  const values = sheet.getDataRange().getValues();
  if (values.length < 2) return 0;

  const header = values[0];
  const col = name => header.indexOf(name);

  const iMarket = col("Market");
  const iQty = col("PositionQty");
  const iLastPrice = col("LastPrice");
  const iCurrency = col("Currency");

  if ([iMarket, iQty, iLastPrice].some(i => i < 0)) {
    throw new Error("Positions sheet must contain Market, PositionQty, LastPrice");
  }

  const map = fxMap || getFxToCnyMap_(spreadsheet);
  let total = 0;

  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    const qty = Number(row[iQty]) || 0;
    const px = Number(row[iLastPrice]) || 0;
    if (qty <= 0 || px <= 0) continue;

    const market = String(row[iMarket] || "").trim().toUpperCase();
    const currencyRaw = iCurrency >= 0 ? String(row[iCurrency] || "") : "";
    const currency = currencyRaw ? currencyRaw.toUpperCase() : defaultCurrencyByMarket_(market);
    const fx = map[currency] || 1;

    total += qty * px * fx;
  }

  return total;
}

function defaultCurrencyByMarket_(market) {
  if (market === "CN") return "CNY";
  if (market === "HK") return "HKD";
  if (market === "US") return "USD";
  return "CNY";
}

/**
 * Live FX rates to CNY. Retries a few times; throws (does NOT silently
 * fall back to 1:1) if GOOGLEFINANCE never resolves — writing a wrong
 * cash/equity number is worse than failing the run loudly.
 */
function getFxToCnyMap_(spreadsheet) {
  const map = { CNY: 1 };
  const targets = [
    { currency: "USD", symbol: "CURRENCY:USDCNY" },
    { currency: "HKD", symbol: "CURRENCY:HKDCNY" }
  ];

  const maxAttempts = 3;
  let lastResult = {};
  for (let attempt = 1; attempt <= maxAttempts; attempt++) {
    lastResult = fetchFromGoogle_(targets.map(t => ({ symbol: t.symbol })));
    const missing = targets.filter(t => !(lastResult[t.symbol] && lastResult[t.symbol].price > 0));
    if (missing.length === 0) break;
    if (attempt < maxAttempts) Utilities.sleep(4000);
  }

  targets.forEach(t => {
    const item = lastResult[t.symbol];
    if (item && item.price > 0) {
      map[t.currency] = item.price;
    } else {
      throw new Error(
        `Failed to fetch FX rate for ${t.currency} after ${maxAttempts} attempts ` +
        `(GOOGLEFINANCE ${t.symbol} unavailable). Aborting rather than writing wrong CNY figures.`
      );
    }
  });

  return map;
}

/**
 * @return {Object} dateKey -> { USD, HKD }
 */
function loadFxHistoryMap_(spreadsheet) {
  const map = {};
  const sheet = spreadsheet.getSheetByName("FxHistory");
  if (!sheet) return map;

  const values = sheet.getDataRange().getValues();
  if (values.length < 2) return map;

  const header = values[0];
  const iDate = header.indexOf("Date");
  const iUsd = header.indexOf("USDCNY");
  const iHkd = header.indexOf("HKDCNY");

  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    if (!row[iDate]) continue;
    map[formatDateKey_(row[iDate])] = {
      USD: Number(row[iUsd]) || 0,
      HKD: Number(row[iHkd]) || 0
    };
  }
  return map;
}

/**
 * @return {Function} (currency, dateKey) -> rate, walking back to the
 *   latest cached date <= dateKey (handles weekends/holidays).
 */
function makeHistoricalFxLookup_(fxHistoryMap) {
  const sortedKeys = Object.keys(fxHistoryMap).sort();
  return function(currency, dateKey) {
    if (currency === "CNY") return 1;
    for (let i = sortedKeys.length - 1; i >= 0; i--) {
      if (sortedKeys[i] <= dateKey) {
        const entry = fxHistoryMap[sortedKeys[i]];
        const rate = currency === "USD" ? entry.USD : currency === "HKD" ? entry.HKD : null;
        if (rate) return rate;
      }
    }
    throw new Error(`No historical FX rate available for ${currency} on/before ${dateKey}`);
  };
}

/**
 * @return {Object} symbol -> { dateKey: closePrice }
 */
function loadPriceHistoryMap_(spreadsheet) {
  const map = {};
  const sheet = spreadsheet.getSheetByName("PriceHistory");
  if (!sheet) return map;

  const values = sheet.getDataRange().getValues();
  if (values.length < 2) return map;

  const header = values[0];
  const iDate = header.indexOf("Date");
  const iSymbol = header.indexOf("Symbol");
  const iClose = header.indexOf("ClosePrice");

  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    if (!row[iDate] || !row[iSymbol]) continue;
    const symbol = normalizeSymbol_(row[iSymbol]);
    const price = Number(row[iClose]) || 0;
    if (price <= 0) continue;
    if (!map[symbol]) map[symbol] = {};
    map[symbol][formatDateKey_(row[iDate])] = price;
  }
  return map;
}

/**
 * @return {Function} (symbol, dateKey) -> price or null, walking back to
 *   the latest cached date <= dateKey.
 */
function makeHistoricalPriceLookup_(priceHistoryMap) {
  const sortedKeysCache = {};
  return function(symbol, dateKey) {
    const bySymbol = priceHistoryMap[symbol];
    if (!bySymbol) return null;
    if (!sortedKeysCache[symbol]) sortedKeysCache[symbol] = Object.keys(bySymbol).sort();
    const keys = sortedKeysCache[symbol];
    for (let i = keys.length - 1; i >= 0; i--) {
      if (keys[i] <= dateKey) return bySymbol[keys[i]];
    }
    return null;
  };
}

function enumerateWeekdayDateKeys_(startDateKey, endDateKey) {
  const keys = [];
  let cur = new Date(startDateKey + "T00:00:00");
  const end = new Date(endDateKey + "T00:00:00");
  while (cur.getTime() <= end.getTime()) {
    const day = cur.getDay();
    if (day !== 0 && day !== 6) keys.push(formatDateKey_(cur));
    cur = new Date(cur.getTime() + 86400000);
  }
  return keys;
}

/**
 * Ensure FxHistory covers every weekday in [startDateKey, endDateKey].
 * Cheap to call repeatedly — no-ops once the range is cached.
 */
function ensureFxHistoryCached_(spreadsheet, startDateKey, endDateKey) {
  const sheet = spreadsheet.getSheetByName("FxHistory");
  if (!sheet) throw new Error('Sheet "FxHistory" not found');

  const existing = loadFxHistoryMap_(spreadsheet);
  const weekdayKeys = enumerateWeekdayDateKeys_(startDateKey, endDateKey);
  const isComplete = weekdayKeys.every(k => existing[k] && existing[k].USD && existing[k].HKD);
  if (isComplete) {
    Logger.log("FxHistory already covers requested range.");
    return;
  }

  const fetched = fetchGoogleFinanceHistoryBatch_([
    { symbol: "USD", googleSymbol: "CURRENCY:USDCNY" },
    { symbol: "HKD", googleSymbol: "CURRENCY:HKDCNY" }
  ], startDateKey, endDateKey);

  const usdMap = fetched.USD || {};
  const hkdMap = fetched.HKD || {};
  const allKeys = Array.from(new Set(
    Object.keys(usdMap).concat(Object.keys(hkdMap)).concat(Object.keys(existing))
  )).sort();

  const rows = allKeys.map(key => [
    new Date(key + "T00:00:00"),
    usdMap[key] || (existing[key] && existing[key].USD) || "",
    hkdMap[key] || (existing[key] && existing[key].HKD) || "",
    "GoogleFinance"
  ]);

  const lastRow = sheet.getLastRow();
  if (lastRow > 1) sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).clearContent();
  if (rows.length > 0) sheet.getRange(2, 1, rows.length, 4).setValues(rows);

  Logger.log(`FxHistory rebuilt: ${rows.length} rows.`);
}

/**
 * Ensure PriceHistory covers every weekday in range for each symbol.
 * CN -> Tushare daily; HK -> Tushare hk_daily; anything left (US, or a
 * Tushare failure) -> one batched GOOGLEFINANCE historical call.
 * Only fetches what's missing; safe to call repeatedly.
 */
function ensurePriceHistoryCached_(spreadsheet, symbols, startDateKey, endDateKey) {
  const sheet = spreadsheet.getSheetByName("PriceHistory");
  if (!sheet) throw new Error('Sheet "PriceHistory" not found');

  const existing = loadPriceHistoryMap_(spreadsheet);
  const weekdayKeys = enumerateWeekdayDateKeys_(startDateKey, endDateKey);

  const needsFetch = symbols.filter(symbol => {
    const already = existing[symbol] || {};
    return !weekdayKeys.every(k => already[k] !== undefined);
  });

  if (needsFetch.length === 0) {
    Logger.log("PriceHistory already covers requested range for all symbols.");
    return;
  }

  const perSymbolResult = {};
  const googleFallbackSymbols = [];

  needsFetch.forEach(symbol => {
    const market = inferMarketFromSymbol_(symbol);
    let fetchedMap = null;

    if (market === "CN") {
      const tsCode = symbolToTsCode_(symbol);
      try {
        fetchedMap = fetchTushareDaily_(tsCode, startDateKey, endDateKey);
        if (Object.keys(fetchedMap).length > 0) perSymbolResult[symbol] = { map: fetchedMap, source: "Tushare" };
      } catch (e) {
        Logger.log(`Tushare daily failed for ${symbol}: ${e.message}`);
      }
    } else if (market === "HK") {
      const tsCode = hkSymbolToTsCode_(symbol);
      try {
        fetchedMap = fetchTushareHkDaily_(tsCode, startDateKey, endDateKey);
        if (Object.keys(fetchedMap).length > 0) perSymbolResult[symbol] = { map: fetchedMap, source: "Tushare" };
      } catch (e) {
        Logger.log(`Tushare hk_daily failed for ${symbol}: ${e.message}`);
      }
    }

    if (!perSymbolResult[symbol]) googleFallbackSymbols.push(symbol);
  });

  if (googleFallbackSymbols.length > 0) {
    const googleResults = fetchGoogleFinanceHistoryBatch_(
      googleFallbackSymbols.map(s => ({ symbol: s })),
      startDateKey,
      endDateKey
    );
    googleFallbackSymbols.forEach(symbol => {
      perSymbolResult[symbol] = { map: googleResults[symbol] || {}, source: "GoogleFinance" };
    });
  }

  const rowsToAppend = [];
  needsFetch.forEach(symbol => {
    const market = inferMarketFromSymbol_(symbol);
    const currency = inferCurrencyFromSymbol_(symbol);
    const already = existing[symbol] || {};
    const fetchedMap = perSymbolResult[symbol].map;
    const source = perSymbolResult[symbol].source;

    weekdayKeys.forEach(key => {
      if (already[key] !== undefined) return;
      const price = fetchedMap[key];
      rowsToAppend.push([
        new Date(key + "T00:00:00"),
        symbol,
        market,
        price || "",
        currency,
        price ? source : "ESTIMATED_MISSING"
      ]);
    });
  });

  if (rowsToAppend.length > 0) {
    const lastRow = sheet.getLastRow();
    sheet.getRange(lastRow + 1, 1, rowsToAppend.length, 6).setValues(rowsToAppend);
    Logger.log(`PriceHistory appended ${rowsToAppend.length} rows.`);
  }
}

function getEarliestDataDateKey_(spreadsheet) {
  const dates = [];

  ["Trades", "CashFlow"].forEach(name => {
    const sheet = spreadsheet.getSheetByName(name);
    if (!sheet) return;
    const values = sheet.getDataRange().getValues();
    if (values.length < 2) return;
    const iDate = values[0].indexOf("Date");
    if (iDate < 0) return;
    values.slice(1).forEach(r => { if (r[iDate]) dates.push(dateOnlyMs_(r[iDate])); });
  });

  if (dates.length === 0) {
    throw new Error("No Trades or CashFlow data found to determine rebuild start date.");
  }
  return formatDateKey_(new Date(Math.min(...dates)));
}

/**
 * Deliberate, explicit full historical rebuild of CashLedger and NAV_Daily
 * from the earliest Trades/CashFlow date to today, using cached historical
 * FX/prices (not "today"'s live rate/price). Idempotent and safe to
 * re-run — it never touches Trades or CashFlow (the originals).
 */
function rebuildCashAndNavHistory() {
  const spreadsheet = getSpreadsheet_();
  ensurePerformanceSheets_(spreadsheet);

  const startKey = getEarliestDataDateKey_(spreadsheet);
  const todayKey = formatDateKey_(getTodayDateOnly_(spreadsheet));
  Logger.log(`Rebuilding cash/NAV history from ${startKey} to ${todayKey}...`);

  ensureFxHistoryCached_(spreadsheet, startKey, todayKey);

  const fxLookup = makeHistoricalFxLookup_(loadFxHistoryMap_(spreadsheet));
  fillCashFlowDerivedFields_(spreadsheet, fxLookup);

  const symbols = Object.keys(buildPositionsFromTrades_(spreadsheet));
  ensurePriceHistoryCached_(spreadsheet, symbols, startKey, todayKey);

  rebuildCashLedgerHistory_(spreadsheet);
  rebuildNavDailyHistory_(spreadsheet, startKey, todayKey);

  Logger.log("Cash/NAV history rebuild complete.");
}

/**
 * Full NAV_Daily rebuild: for every weekday in range, replay Trades up to
 * that date (buildPositionsAsOfDate_) and price it with that day's cached
 * historical price/FX — not today's live Positions/LastPrice.
 */
function rebuildNavDailyHistory_(spreadsheet, startKey, endKey) {
  const sheet = spreadsheet.getSheetByName("NAV_Daily");
  if (!sheet) throw new Error('Sheet "NAV_Daily" not found');

  const fxLookup = makeHistoricalFxLookup_(loadFxHistoryMap_(spreadsheet));
  const priceLookup = makeHistoricalPriceLookup_(loadPriceHistoryMap_(spreadsheet));

  const tradeFlows = getDailyTradeCashFlowsCNY_(spreadsheet, fxLookup);
  const externalFlows = getDailyExternalCashFlowsCNY_(spreadsheet, fxLookup);
  const walked = walkCashBalances_(tradeFlows, externalFlows);

  const cashByDate = {};
  walked.forEach(w => { cashByDate[w.date] = w.cashBalance; });
  const cashKeys = walked.map(w => w.date);

  const weekdayKeys = enumerateWeekdayDateKeys_(startKey, endKey);

  const rows = [];
  const dataQualityFlags = [];

  weekdayKeys.forEach(key => {
    const asOfDate = new Date(key + "T00:00:00");
    const positions = buildPositionsAsOfDate_(spreadsheet, asOfDate);

    let equity = 0;
    let estimated = false;
    Object.keys(positions).forEach(symbol => {
      const p = positions[symbol];
      if (p.quantity <= 0) return;
      const price = priceLookup(symbol, key);
      if (price === null) { estimated = true; return; }
      const fx = p.currency === "CNY" ? 1 : fxLookup(p.currency, key);
      equity += p.quantity * price * fx;
    });

    let cashBalance = 0;
    for (let i = cashKeys.length - 1; i >= 0; i--) {
      if (cashKeys[i] <= key) { cashBalance = cashByDate[cashKeys[i]]; break; }
    }

    const externalFlow = roundMoney_(externalFlows[key] || 0);
    const totalAsset = roundMoney_(equity + cashBalance);

    rows.push([asOfDate, roundMoney_(equity), cashBalance, externalFlow, totalAsset]);
    dataQualityFlags.push(estimated ? "ESTIMATED" : "");
  });

  const lastRow = sheet.getLastRow();
  if (lastRow > 1) sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).clearContent();
  if (rows.length === 0) return;

  sheet.getRange(2, 1, rows.length, 5).setValues(rows);

  for (let i = 0; i < rows.length; i++) {
    const r = i + 2;
    if (i === 0) {
      sheet.getRange(r, 6, 1, 3).setValues([["", "", ""]]);
    } else {
      sheet.getRange(r, 6).setFormula(`=E${r - 1}`);
      sheet.getRange(r, 7).setFormula(`=E${r}-F${r}-D${r}`);
      sheet.getRange(r, 8).setFormula(`=IF(F${r}+D${r}=0,,G${r}/(F${r}+D${r}))`);
    }
  }

  sheet.getRange(2, 9, dataQualityFlags.length, 1).setValues(dataQualityFlags.map(v => [v]));

  Logger.log(`NAV_Daily history rebuilt: ${rows.length} rows.`);
}

/**
 * After a successful daily run, cache today's live price/FX into
 * FxHistory/PriceHistory so a future rebuild never needs to re-fetch them.
 */
function appendTodayToHistoryCaches_(spreadsheet, fxMap) {
  const todayKey = formatDateKey_(getTodayDateOnly_(spreadsheet));

  const fxSheet = spreadsheet.getSheetByName("FxHistory");
  if (fxSheet) {
    const existing = loadFxHistoryMap_(spreadsheet);
    if (!existing[todayKey]) {
      const lastRow = fxSheet.getLastRow();
      fxSheet.getRange(lastRow + 1, 1, 1, 4).setValues([[
        new Date(todayKey + "T00:00:00"), fxMap.USD || "", fxMap.HKD || "", "Live"
      ]]);
    }
  }

  const posSheet = spreadsheet.getSheetByName("Positions");
  const priceSheet = spreadsheet.getSheetByName("PriceHistory");
  if (!posSheet || !priceSheet) return;

  const data = posSheet.getDataRange().getValues();
  if (data.length <= 2) return;

  const header = data[0];
  const col = name => header.indexOf(name);
  const iMarket = col("Market");
  const iSymbol = col("Symbol");
  const iPrice = col("LastPrice");
  const iCurrency = col("Currency");

  const existingPrices = loadPriceHistoryMap_(spreadsheet);
  const rowsToAppend = [];

  for (let i = 2; i < data.length; i++) {
    const symbol = normalizeSymbol_(data[i][iSymbol]);
    const price = Number(data[i][iPrice]) || 0;
    if (!symbol || price <= 0) continue;
    if (existingPrices[symbol] && existingPrices[symbol][todayKey] !== undefined) continue;

    rowsToAppend.push([
      new Date(todayKey + "T00:00:00"),
      symbol,
      data[i][iMarket] || "",
      price,
      data[i][iCurrency] || "",
      "Live"
    ]);
  }

  if (rowsToAppend.length > 0) {
    const lastRow = priceSheet.getLastRow();
    priceSheet.getRange(lastRow + 1, 1, rowsToAppend.length, 6).setValues(rowsToAppend);
  }
}

function isWeekdayTradingDay_(spreadsheet) {
  const tz = spreadsheet.getSpreadsheetTimeZone() || Session.getScriptTimeZone();
  const today = new Date();
  const day = Number(Utilities.formatDate(today, tz, "u")); // 1..7 (Mon..Sun)
  return day <= 5;
}

function getTodayDateOnly_(spreadsheet) {
  const tz = spreadsheet.getSpreadsheetTimeZone() || Session.getScriptTimeZone();
  const key = Utilities.formatDate(new Date(), tz, "yyyy-MM-dd");
  return new Date(key + "T00:00:00");
}

function formatDateKey_(d) {
  const date = (d instanceof Date) ? d : new Date(d);
  return Utilities.formatDate(date, Session.getScriptTimeZone(), "yyyy-MM-dd");
}

function setupPnLViewFormula_(spreadsheet) {
  const sheet = spreadsheet.getSheetByName("PnL_View");
  if (!sheet) return;

  // Reset content area (keep headers)
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    sheet.getRange(2, 1, lastRow - 1, 3).clearContent();
  }

  const rows = [
    ["Latest Trading Date", '=IFERROR(INDEX(FILTER(NAV_Daily!A2:A,NAV_Daily!A2:A<>""),ROWS(FILTER(NAV_Daily!A2:A,NAV_Daily!A2:A<>""))),"")', "NAV_Daily latest date"],
    ["Total Asset (CNY)", '=IFERROR(INDEX(FILTER(NAV_Daily!E2:E,NAV_Daily!A2:A<>""),ROWS(FILTER(NAV_Daily!E2:E,NAV_Daily!A2:A<>""))),0)', "Latest equity plus cash"],
    ["Cash Balance (CNY)", '=IFERROR(INDEX(FILTER(NAV_Daily!C2:C,NAV_Daily!A2:A<>""),ROWS(FILTER(NAV_Daily!C2:C,NAV_Daily!A2:A<>""))),0)', "Latest generated cash balance"],
    ["Today PnL", '=IFERROR(INDEX(FILTER(NAV_Daily!G2:G,NAV_Daily!A2:A<>""),ROWS(FILTER(NAV_Daily!G2:G,NAV_Daily!A2:A<>""))),0)', "Trading PnL after external cash flow"],
    ["Today Return", '=IFERROR(INDEX(FILTER(NAV_Daily!H2:H,NAV_Daily!A2:A<>""),ROWS(FILTER(NAV_Daily!H2:H,NAV_Daily!A2:A<>""))),0)', "Trading return after external cash flow"],
    ["MTD PnL", '=IFERROR(SUM(FILTER(NAV_Daily!G2:G,NAV_Daily!A2:A>=EOMONTH(TODAY(),-1)+1,NAV_Daily!A2:A<=TODAY())),0)', "Month-to-date trading PnL"],
    ["YTD PnL", '=IFERROR(SUM(FILTER(NAV_Daily!G2:G,NAV_Daily!A2:A>=DATE(YEAR(TODAY()),1,1),NAV_Daily!A2:A<=TODAY())),0)', "Year-to-date trading PnL"],
    ["Optional Since Inception XIRR", '=IFERROR(XIRR({FILTER(CashLedger!C2:C,CashLedger!A2:A<>"",CashLedger!C2:C>0);-INDEX(FILTER(NAV_Daily!E2:E,NAV_Daily!A2:A<>""),ROWS(FILTER(NAV_Daily!E2:E,NAV_Daily!A2:A<>"")))},{FILTER(CashLedger!A2:A,CashLedger!A2:A<>"",CashLedger!C2:C>0);INDEX(FILTER(NAV_Daily!A2:A,NAV_Daily!A2:A<>""),ROWS(FILTER(NAV_Daily!A2:A,NAV_Daily!A2:A<>"")))}),"")', "Optional money-weighted reference"]
  ];

  sheet.getRange(2, 1, rows.length, 3).setValues(rows);
}

/**
 * One-click setup for initial sheet creation.
 */
function setupPerformanceView() {
  const ss = getSpreadsheet_();
  ensurePerformanceSheets_(ss);
}
