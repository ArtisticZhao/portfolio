const NAV_COLUMN_COUNT_ = 10;
const NAV_QUALITY_LIVE_ = "LIVE";
const NAV_QUALITY_PENDING_ = "PENDING";
const PRICE_NO_DATA_SOURCE_ = "NO_DATA";
// A weekday without a close is only recorded as a market holiday once it is
// this many days old; before that the close may simply not be published yet.
const PRICE_SETTLE_DAYS_ = 3;
const HISTORY_LOOKBACK_DAYS_ = 10;

/**
 * Ensure performance-related sheets exist with expected headers.
 * Sheets:
 * - CashLedger: generated cash ledger from Trades + CashFlow (external flows)
 * - NAV_Daily: daily account snapshot in CNY; InputSig records which
 *   Trades/CashFlow rows each row was computed from (see syncCashAndNavHistory_)
 * - PnL_View: dashboard formulas
 * - FxHistory / PriceHistory: cache sheets for historical closes
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
    "DataQuality",
    "InputSig"
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
 * Rewrite CashLedger from walked balances. It is a pure function of
 * Trades/CashFlow and cached historical FX, so rewriting it on every sync
 * is reproducible (unlike the old "today's live FX on every past trade").
 */
function writeCashLedger_(spreadsheet, walked) {
  const sheet = spreadsheet.getSheetByName("CashLedger");
  if (!sheet) throw new Error('Sheet "CashLedger" not found');

  const rows = walked.map(w => [
    new Date(w.date + "T00:00:00"),
    w.tradeFlow,
    w.externalFlow,
    w.cashBalance,
    "Trade + external cash flow"
  ]);

  const lastRow = sheet.getLastRow();
  if (lastRow > 1) sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).clearContent();
  if (rows.length > 0) sheet.getRange(2, 1, rows.length, 5).setValues(rows);

  Logger.log(`CashLedger rewritten: ${rows.length} rows.`);
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

function dateOnlyMs_(d) {
  const date = (d instanceof Date) ? d : new Date(d);
  return new Date(date.getFullYear(), date.getMonth(), date.getDate()).getTime();
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
 * @return {Object} symbol -> { dateKey: closePrice }. A value of 0 is a
 *   "NO_DATA" marker: that weekday was already queried and the market was
 *   closed, so it must not be fetched again.
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
  const iSource = header.indexOf("Source");

  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    if (!row[iDate] || !row[iSymbol]) continue;
    // Rows written before Symbol was forced to text came back as numbers
    // ("002317" -> 2317); restore the 6-digit A-share code.
    const symbol = typeof row[iSymbol] === "number"
      ? String(row[iSymbol]).padStart(6, "0")
      : normalizeSymbol_(row[iSymbol]);
    const key = formatDateKey_(row[iDate]);
    const price = Number(row[iClose]) || 0;
    if (!map[symbol]) map[symbol] = {};

    if (price > 0) {
      map[symbol][key] = price;
    } else if (row[iSource] === PRICE_NO_DATA_SOURCE_ && map[symbol][key] === undefined) {
      map[symbol][key] = 0;
    }
  }
  return map;
}

/**
 * @return {Function} (symbol, dateKey) -> { price, settled }.
 *   price: close on dateKey, else the latest earlier close, else null.
 *   settled: false when dateKey itself has neither a price nor a NO_DATA
 *   marker yet, i.e. the answer may still change once the cache fills in.
 */
function makeHistoricalPriceLookup_(priceHistoryMap) {
  const sortedKeysCache = {};
  return function(symbol, dateKey) {
    const bySymbol = priceHistoryMap[symbol];
    if (!bySymbol) return { price: null, settled: false };
    if (bySymbol[dateKey] > 0) return { price: bySymbol[dateKey], settled: true };

    const settled = bySymbol[dateKey] === 0;
    if (!sortedKeysCache[symbol]) sortedKeysCache[symbol] = Object.keys(bySymbol).sort();
    const keys = sortedKeysCache[symbol];
    for (let i = keys.length - 1; i >= 0; i--) {
      if (keys[i] < dateKey && bySymbol[keys[i]] > 0) {
        return { price: bySymbol[keys[i]], settled: settled };
      }
    }
    return { price: null, settled: settled };
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

function addDaysToDateKey_(dateKey, days) {
  const d = new Date(dateKey + "T00:00:00");
  d.setDate(d.getDate() + days);
  return formatDateKey_(d);
}

/**
 * Ensure FxHistory spans [startDateKey, endDateKey]: some cached date on or
 * before the start (lookups walk back) and one on or after the end.
 * Individual weekdays may be missing; no-ops once the span is cached.
 */
function ensureFxHistoryCached_(spreadsheet, startDateKey, endDateKey) {
  const sheet = spreadsheet.getSheetByName("FxHistory");
  if (!sheet) throw new Error('Sheet "FxHistory" not found');

  const existing = loadFxHistoryMap_(spreadsheet);
  const cachedKeys = Object.keys(existing)
    .filter(k => existing[k].USD && existing[k].HKD)
    .sort();
  const isCovered = cachedKeys.length > 0 &&
    cachedKeys[0] <= startDateKey &&
    cachedKeys[cachedKeys.length - 1] >= endDateKey;
  if (isCovered) {
    Logger.log("FxHistory already covers requested range.");
    return;
  }

  const fetched = fetchGoogleFinanceHistoryBatch_([
    { symbol: "USD", googleSymbol: "CURRENCY:USDCNY" },
    { symbol: "HKD", googleSymbol: "CURRENCY:HKDCNY" }
  ], addDaysToDateKey_(startDateKey, -HISTORY_LOOKBACK_DAYS_), endDateKey);

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
 * Ensure PriceHistory has an entry (close or NO_DATA marker) for every
 * requested symbol/date. CN stocks -> Tushare daily, then Sina; CN
 * funds/ETFs -> Sina; HK and US -> one batched GOOGLEFINANCE historical
 * call. Only fetches what's missing; safe to call repeatedly.
 *
 * @param {Object} neededBySymbol symbol -> [dateKey] the symbol was held on
 * @param {string} todayKey NO_DATA markers are only written for dates older
 *   than PRICE_SETTLE_DAYS_, so a close that is merely not published yet is
 *   retried on the next run instead of being recorded as a holiday.
 */
function ensurePriceHistoryCached_(spreadsheet, neededBySymbol, todayKey) {
  const sheet = spreadsheet.getSheetByName("PriceHistory");
  if (!sheet) throw new Error('Sheet "PriceHistory" not found');

  const existing = loadPriceHistoryMap_(spreadsheet);
  const missingKeys = [];
  const needsFetch = Object.keys(neededBySymbol).filter(symbol => {
    const already = existing[symbol] || {};
    const missing = neededBySymbol[symbol].filter(k => already[k] === undefined);
    missing.forEach(k => missingKeys.push(k));
    return missing.length > 0;
  });

  if (needsFetch.length === 0) {
    Logger.log("PriceHistory already covers requested dates for all symbols.");
    return;
  }

  missingKeys.sort();
  // Look back a little so a symbol first needed on a market holiday still
  // has an earlier close to fall back to.
  const startDateKey = addDaysToDateKey_(missingKeys[0], -HISTORY_LOOKBACK_DAYS_);
  const endDateKey = missingKeys[missingKeys.length - 1];
  const weekdayKeys = enumerateWeekdayDateKeys_(startDateKey, endDateKey);
  const settleCutoffKey = addDaysToDateKey_(todayKey, -PRICE_SETTLE_DAYS_);

  const perSymbolResult = {};
  const googleFallbackSymbols = [];

  needsFetch.forEach(symbol => {
    const market = inferMarketFromSymbol_(symbol);
    let fetchedMap = null;

    if (market === "CN") {
      // Tushare "daily" only covers stocks, so funds/ETFs (e.g. 513300) go
      // straight to Sina; stocks fall back to it when Tushare has nothing.
      if (!isCnFundOrEtfSymbol_(symbol)) {
        try {
          fetchedMap = fetchTushareDaily_(symbolToTsCode_(symbol), startDateKey, endDateKey);
          if (Object.keys(fetchedMap).length > 0) perSymbolResult[symbol] = { map: fetchedMap, source: "Tushare" };
        } catch (e) {
          Logger.log(`Tushare daily failed for ${symbol}: ${e.message}`);
        }
      }
      if (!perSymbolResult[symbol]) {
        try {
          fetchedMap = fetchSinaDailyHistory_(symbol, startDateKey, endDateKey);
        } catch (e) {
          Logger.log(`Sina history failed for ${symbol}: ${e.message}`);
          fetchedMap = {};
        }
        // GOOGLEFINANCE has no A-share history, so there is no further fallback.
        perSymbolResult[symbol] = { map: fetchedMap, source: "Sina" };
      }
    }

    // HK and US go to GOOGLEFINANCE (Tushare hk_daily is limited to one
    // call per minute on this account, so it is not used).
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

    // An empty result means the source failed, not that every day was a
    // holiday: write nothing so the next run retries.
    if (Object.keys(fetchedMap).length === 0) {
      Logger.log(`No historical prices returned for ${symbol} (${startDateKey}..${endDateKey}).`);
      return;
    }

    weekdayKeys.forEach(key => {
      if (already[key] !== undefined) return;
      const price = fetchedMap[key];
      if (!price && key > settleCutoffKey) return;
      rowsToAppend.push([
        new Date(key + "T00:00:00"),
        symbol,
        market,
        price || "",
        currency,
        price ? source : PRICE_NO_DATA_SOURCE_
      ]);
    });
  });

  if (rowsToAppend.length > 0) {
    const lastRow = sheet.getLastRow();
    // Keep Symbol as text, otherwise Sheets turns "002317" into 2317.
    sheet.getRange(lastRow + 1, 2, rowsToAppend.length, 1).setNumberFormat("@");
    sheet.getRange(lastRow + 1, 1, rowsToAppend.length, 6).setValues(rowsToAppend);
    Logger.log(`PriceHistory appended ${rowsToAppend.length} rows.`);
  }
}

function isWeekdayDateKey_(dateKey) {
  const day = new Date(dateKey + "T00:00:00").getDay();
  return day !== 0 && day !== 6;
}

/**
 * NAV_Daily only has weekday rows, so anything dated on a weekend is
 * accounted for on the following Monday.
 */
function toNavDateKey_(dateKey) {
  let key = dateKey;
  while (!isWeekdayDateKey_(key)) key = addDaysToDateKey_(key, 1);
  return key;
}

function shortHash_(text) {
  const bytes = Utilities.computeDigest(
    Utilities.DigestAlgorithm.MD5,
    text,
    Utilities.Charset.UTF_8
  );
  return bytes.slice(0, 4).map(b => (b & 0xff).toString(16).padStart(2, "0")).join("");
}

/**
 * Fingerprint of everything the user entered for each NAV date, stored in
 * NAV_Daily.InputSig. A NAV row whose stored signature differs from the
 * current one was computed from different Trades/CashFlow rows (a trade was
 * back-filled, edited, or deleted) and must be recomputed.
 *
 * Format: "<trade rows>+<cash flow rows>:<hash>", e.g. "2+0:a1b2c3d4".
 * Row order within a day does not matter. Dates without any input have no
 * entry (stored as an empty cell).
 *
 * @param {Array<Array>} tradeValues Trades sheet values including header
 * @param {Array<Array>} cashFlowValues CashFlow sheet values including header
 * @return {Object} NAV dateKey -> signature
 */
function buildDailyInputSignatures_(tradeValues, cashFlowValues) {
  const byDate = {};
  const add = (dateValue, kind, fields) => {
    const key = toNavDateKey_(formatDateKey_(dateValue));
    if (!byDate[key]) byDate[key] = { trades: 0, flows: 0, lines: [] };
    byDate[key][kind]++;
    byDate[key].lines.push(kind + "|" + formatDateKey_(dateValue) + "|" + fields.join("|"));
  };
  const text = v => String(v === null || v === undefined ? "" : v).trim().toUpperCase();
  const num = v => Number(v) || 0;

  if (tradeValues && tradeValues.length > 1) {
    const header = tradeValues[0];
    const col = name => header.indexOf(name);
    const iDate = col("Date");
    const iSymbol = col("Symbol");
    const textCols = ["Symbol", "Side", "Currency"].map(col);
    const numCols = ["Quantity", "Price", "Fee", "Tax", "OtherCost"].map(col);

    tradeValues.slice(1).forEach(row => {
      if (!row[iDate] || !normalizeSymbol_(row[iSymbol])) return;
      add(row[iDate], "trades", textCols.map(i => i >= 0 ? text(row[i]) : "")
        .concat(numCols.map(i => i >= 0 ? num(row[i]) : 0)));
    });
  }

  if (cashFlowValues && cashFlowValues.length > 1) {
    const header = cashFlowValues[0];
    const col = name => header.indexOf(name);
    const iDate = col("Date");
    const iType = col("Type");
    const iAmount = col("Amount");
    const iCurrency = col("Currency");

    cashFlowValues.slice(1).forEach(row => {
      if (iDate < 0 || !row[iDate]) return;
      add(row[iDate], "flows", [text(row[iType]), num(row[iAmount]), text(row[iCurrency])]);
    });
  }

  const signatures = {};
  Object.keys(byDate).forEach(key => {
    const d = byDate[key];
    signatures[key] = `${d.trades}+${d.flows}:${shortHash_(d.lines.sort().join("\n"))}`;
  });
  return signatures;
}

/**
 * First NAV row that can no longer be trusted. Everything from that row on
 * is recomputed, because positions and cash are cumulative: a trade
 * back-filled on day D changes D and every day after it.
 *
 * @param {Array<string>} weekdayKeys expected NAV dates, ascending
 * @param {Array<Object>} navRows existing rows in sheet order: { dateKey, quality, sig }
 * @param {Object} signatures see buildDailyInputSignatures_
 * @return {number} index into weekdayKeys (== number of rows to keep), or
 *   -1 when NAV_Daily is fully up to date
 */
function findFirstDirtyIndex_(weekdayKeys, navRows, signatures, todayKey) {
  for (let i = 0; i < weekdayKeys.length; i++) {
    const key = weekdayKeys[i];
    const nav = navRows[i];
    // Missing row, gap, duplicate or out-of-order date.
    if (!nav || nav.dateKey !== key) return i;
    // Today is always re-priced with live quotes.
    if (key >= todayKey) return i;
    // Snapshot taken with live/incomplete prices: finalize with the close.
    if (nav.quality === NAV_QUALITY_LIVE_ || nav.quality === NAV_QUALITY_PENDING_) return i;
    if (nav.sig !== (signatures[key] || "")) return i;
  }
  return navRows.length > weekdayKeys.length ? weekdayKeys.length : -1;
}

function readNavRows_(navSheet) {
  const lastRow = navSheet.getLastRow();
  if (lastRow < 2) return [];

  const rows = navSheet.getRange(2, 1, lastRow - 1, NAV_COLUMN_COUNT_).getValues().map(row => ({
    dateKey: row[0] ? formatDateKey_(row[0]) : "",
    quality: String(row[8] || "").trim(),
    sig: String(row[9] || "").trim()
  }));
  while (rows.length > 0 && !rows[rows.length - 1].dateKey) rows.pop();
  return rows;
}

/**
 * Open holdings at the close of each requested date, replaying Trades in
 * memory with the same rules as Positions (calculatePositionsFromTradeRows_).
 * @return {Object} dateKey -> { symbol: { quantity, currency } }
 */
function buildHoldingsByDate_(tradeValues, dateKeys) {
  const result = {};
  const header = tradeValues && tradeValues.length > 0 ? tradeValues[0] : [];
  const iDate = header.indexOf("Date");

  const dated = (tradeValues || []).slice(1)
    .filter(row => row[iDate])
    .map((row, idx) => ({ row, idx, key: formatDateKey_(row[iDate]) }))
    .sort((a, b) => a.key === b.key ? a.idx - b.idx : (a.key < b.key ? -1 : 1));

  let count = 0;
  dateKeys.forEach(key => {
    while (count < dated.length && dated[count].key <= key) count++;
    const positions = count > 0
      ? calculatePositionsFromTradeRows_(header, dated.slice(0, count).map(d => d.row))
      : {};

    const holdings = {};
    Object.keys(positions).forEach(symbol => {
      const p = positions[symbol];
      if (p.quantity > 0) holdings[symbol] = { quantity: p.quantity, currency: p.currency };
    });
    result[key] = holdings;
  });
  return result;
}

/**
 * FX for the sync: cached historical rate for past dates, this run's live
 * rate for today (its close is not published yet).
 */
function makeSyncFxLookup_(fxHistoryMap, liveFxMap, todayKey) {
  const historical = makeHistoricalFxLookup_(fxHistoryMap);
  return function(currency, dateKey) {
    if (currency === "CNY") return 1;
    if (dateKey >= todayKey && liveFxMap[currency]) return liveFxMap[currency];
    return historical(currency, dateKey);
  };
}

function hasForeignCurrency_(sheetValues) {
  if (!sheetValues || sheetValues.length < 2) return false;
  const iCurrency = sheetValues[0].indexOf("Currency");
  if (iCurrency < 0) return false;
  return sheetValues.slice(1).some(row => {
    const currency = String(row[iCurrency] || "").trim().toUpperCase();
    return currency && currency !== "CNY";
  });
}

function getEarliestDateKey_(sheetValuesList) {
  let earliest = null;
  sheetValuesList.forEach(values => {
    if (!values || values.length < 2) return;
    const iDate = values[0].indexOf("Date");
    if (iDate < 0) return;
    values.slice(1).forEach(row => {
      if (!row[iDate]) return;
      const key = formatDateKey_(row[iDate]);
      if (earliest === null || key < earliest) earliest = key;
    });
  });
  return earliest;
}

/**
 * Bring CashLedger and NAV_Daily in line with Trades/CashFlow.
 *
 * There is no separate "full rebuild": every run compares each NAV row's
 * stored InputSig with the current Trades/CashFlow rows for that date and
 * recomputes from the first mismatch through today. That covers trades
 * recorded late, edited trades, days the script did not run, and the very
 * first run (no rows yet). To force a rebuild from some date, delete the
 * NAV_Daily rows from that date down.
 *
 * Past dates are priced with cached historical closes/FX; today uses the
 * live Positions prices and is marked LIVE so the next run finalizes it.
 *
 * @param {Object} fxMap live rates from getFxToCnyMap_
 */
function syncCashAndNavHistory_(spreadsheet, fxMap) {
  const navSheet = spreadsheet.getSheetByName("NAV_Daily");
  const tradesSheet = spreadsheet.getSheetByName("Trades");
  if (!navSheet || !tradesSheet) {
    throw new Error('Missing sheet: "NAV_Daily" or "Trades"');
  }

  const cashFlowSheet = spreadsheet.getSheetByName("CashFlow");
  const tradeValues = tradesSheet.getDataRange().getValues();
  const cashFlowValues = cashFlowSheet ? cashFlowSheet.getDataRange().getValues() : [];

  const startKey = getEarliestDateKey_([tradeValues, cashFlowValues]);
  if (!startKey) {
    Logger.log("No Trades or CashFlow data, skip cash/NAV sync.");
    return;
  }

  const todayKey = formatDateKey_(getTodayDateOnly_(spreadsheet));
  const weekdayKeys = enumerateWeekdayDateKeys_(startKey, todayKey);
  const signatures = buildDailyInputSignatures_(tradeValues, cashFlowValues);
  const firstDirty = findFirstDirtyIndex_(weekdayKeys, readNavRows_(navSheet), signatures, todayKey);
  if (firstDirty === -1) {
    Logger.log("NAV_Daily is up to date, nothing to recompute.");
    return;
  }

  const dirtyKeys = weekdayKeys.slice(firstDirty);
  const pastDirtyKeys = dirtyKeys.filter(k => k < todayKey);
  Logger.log(dirtyKeys.length > 0
    ? `Recomputing NAV_Daily from ${dirtyKeys[0]} (${dirtyKeys.length} rows).`
    : "Trimming extra NAV_Daily rows.");

  const pastWeekdayKeys = weekdayKeys.filter(k => k < todayKey);
  if (pastWeekdayKeys.length > 0 &&
      (hasForeignCurrency_(tradeValues) || hasForeignCurrency_(cashFlowValues))) {
    ensureFxHistoryCached_(spreadsheet, startKey, pastWeekdayKeys[pastWeekdayKeys.length - 1]);
  }
  const fxLookup = makeSyncFxLookup_(loadFxHistoryMap_(spreadsheet), fxMap, todayKey);

  fillCashFlowDerivedFields_(spreadsheet, fxLookup);
  const tradeFlows = getDailyTradeCashFlowsCNY_(spreadsheet, fxLookup);
  const externalFlows = getDailyExternalCashFlowsCNY_(spreadsheet, fxLookup);
  const walked = walkCashBalances_(tradeFlows, externalFlows);
  writeCashLedger_(spreadsheet, walked);

  const holdingsByDate = buildHoldingsByDate_(tradeValues, dirtyKeys);
  const neededBySymbol = {};
  pastDirtyKeys.forEach(key => {
    Object.keys(holdingsByDate[key]).forEach(symbol => {
      if (!neededBySymbol[symbol]) neededBySymbol[symbol] = [];
      neededBySymbol[symbol].push(key);
    });
  });
  ensurePriceHistoryCached_(spreadsheet, neededBySymbol, todayKey);
  const priceLookup = makeHistoricalPriceLookup_(loadPriceHistoryMap_(spreadsheet));

  const liveEquity = dirtyKeys.indexOf(todayKey) >= 0
    ? computeEquityValueCNY_(spreadsheet, fxMap)
    : 0;

  const rows = computeNavRows_({
    dateKeys: dirtyKeys,
    prevDateKey: firstDirty > 0 ? weekdayKeys[firstDirty - 1] : "",
    todayKey: todayKey,
    liveEquity: liveEquity,
    holdingsByDate: holdingsByDate,
    priceLookup: priceLookup,
    fxLookup: fxLookup,
    walked: walked,
    externalFlows: externalFlows,
    signatures: signatures
  });

  writeNavRows_(navSheet, firstDirty + 2, rows);
  Logger.log(`NAV_Daily synced: ${rows.length} rows written.`);
}

/**
 * Pure NAV row computation for consecutive weekday dates.
 * @return {Array<Object>} { dateKey, equity, cashBalance, externalFlow, totalAsset, quality, sig }
 */
function computeNavRows_(input) {
  const cashKeys = input.walked.map(w => w.date);
  const flowKeys = Object.keys(input.externalFlows).sort();
  let cashIdx = -1;
  let prevKey = input.prevDateKey;

  return input.dateKeys.map(key => {
    let equity = 0;
    let quality = "";

    if (key >= input.todayKey) {
      equity = input.liveEquity;
      quality = NAV_QUALITY_LIVE_;
    } else {
      let estimated = false;
      let pending = false;
      const holdings = input.holdingsByDate[key];
      Object.keys(holdings).forEach(symbol => {
        const h = holdings[symbol];
        const quote = input.priceLookup(symbol, key);
        if (!quote.settled) pending = true;
        if (quote.price === null) { estimated = true; return; }
        equity += h.quantity * quote.price * input.fxLookup(h.currency, key);
      });
      quality = pending ? NAV_QUALITY_PENDING_ : (estimated ? "ESTIMATED" : "");
    }

    while (cashIdx + 1 < cashKeys.length && cashKeys[cashIdx + 1] <= key) cashIdx++;
    const cashBalance = cashIdx >= 0 ? input.walked[cashIdx].cashBalance : 0;

    // Includes flows dated on the weekend/gap since the previous NAV row.
    const externalFlow = roundMoney_(flowKeys
      .filter(k => k > prevKey && k <= key)
      .reduce((sum, k) => sum + input.externalFlows[k], 0));
    prevKey = key;

    return {
      dateKey: key,
      equity: roundMoney_(equity),
      cashBalance: cashBalance,
      externalFlow: externalFlow,
      totalAsset: roundMoney_(equity + cashBalance),
      quality: quality,
      sig: input.signatures[key] || ""
    };
  });
}

/**
 * Overwrite NAV_Daily from startRow down with the given rows and clear
 * whatever was below. Rows above startRow are left untouched.
 */
function writeNavRows_(navSheet, startRow, rows) {
  const lastRow = navSheet.getLastRow();
  if (lastRow >= startRow) {
    navSheet.getRange(startRow, 1, lastRow - startRow + 1, NAV_COLUMN_COUNT_).clearContent();
  }
  if (rows.length === 0) return;

  navSheet.getRange(startRow, 1, rows.length, 5).setValues(rows.map(r => [
    new Date(r.dateKey + "T00:00:00"),
    r.equity,
    r.cashBalance,
    r.externalFlow,
    r.totalAsset
  ]));

  // Derived columns F:G:H. The first NAV row has no previous day.
  navSheet.getRange(startRow, 6, rows.length, 3).setFormulas(rows.map((_, i) => {
    const r = startRow + i;
    if (r === 2) return ["", "", ""];
    return [
      `=E${r - 1}`,
      `=E${r}-F${r}-D${r}`,
      `=IF(F${r}+D${r}=0,,G${r}/(F${r}+D${r}))`
    ];
  }));

  navSheet.getRange(startRow, 9, rows.length, 2).setValues(rows.map(r => [r.quality, r.sig]));
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
