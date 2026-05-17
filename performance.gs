/**
 * Ensure performance-related sheets exist with expected headers.
 * Sheets:
 * - CashFlow: manual external deposit/withdraw records
 * - CashBalance: computed cash retained inside investment accounts
 * - NAV_Daily: daily account snapshot in CNY
 * - PnL_View: dashboard formulas
 *
 * CashFlow sign rule:
 * - Positive amount = cash into investment account (deposit)
 * - Negative amount = cash out of investment account (withdraw)
 */
function ensurePerformanceSheets_(spreadsheet) {
  ensureSheetWithHeader_(spreadsheet, "CashFlow", [
    "Date",
    "Type",
    "Amount",
    "Currency",
    "Note",
    "AmountCNY"
  ]);

  ensureSheetWithHeader_(spreadsheet, "CashBalance", [
    "Date",
    "CashBalanceCNY",
    "Currency",
    "Note"
  ]);

  ensureSheetWithHeader_(spreadsheet, "NAV_Daily", [
    "Date",
    "EquityValueCNY",
    "NetFlowCNY",
    "TotalAssetCNY",
    "PrevTotalAssetCNY",
    "TradingPnL",
    "DailyReturn"
  ]);

  ensureSheetWithHeader_(spreadsheet, "PnL_View", [
    "Metric",
    "Value",
    "Notes"
  ]);

  setupCashFlowFormula_(spreadsheet);
  setupPnLViewFormula_(spreadsheet);
}

function ensureSheetWithHeader_(spreadsheet, name, headers) {
  let sheet = spreadsheet.getSheetByName(name);
  if (!sheet) sheet = spreadsheet.insertSheet(name);

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
}

function setupCashFlowFormula_(spreadsheet) {
  const sheet = spreadsheet.getSheetByName("CashFlow");
  if (!sheet) return;

  // AmountCNY = Amount * FX(Currency->CNY)
  const formula = [
    '=ARRAYFORMULA(IF(A2:A="",,IF(C2:C="",,C2:C*IF(D2:D="CNY",1,IF(D2:D="USD",IFERROR(GOOGLEFINANCE("CURRENCY:USDCNY"),),IF(D2:D="HKD",IFERROR(GOOGLEFINANCE("CURRENCY:HKDCNY"),),))))))'
  ];

  sheet.getRange("F2").setFormula(formula[0]);
}

/**
 * Capture one daily snapshot into NAV_Daily.
 * Snapshot date is today in spreadsheet timezone.
 */
function appendDailyNavSnapshot_(spreadsheet) {
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

  const equityValueCNY = computeEquityValueCNY_(spreadsheet);
  const netFlowCNY = getNetFlowCNYForDate_(spreadsheet, today);
  const cashBalanceCNY = updateCashBalanceForDate_(spreadsheet, today);
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

  navSheet.getRange(targetRow, 1, 1, 4).setValues([[
    today,
    equityValueCNY,
    netFlowCNY,
    totalAssetCNY
  ]]);

  // Fill derived columns E:F:G.
  navSheet.getRange(targetRow, 5).setFormula(
    `=IFERROR(INDEX(D$2:D,ROW()-2),)`
  );
  navSheet.getRange(targetRow, 6).setFormula(
    `=IF(E${targetRow}="",,D${targetRow}-E${targetRow}-C${targetRow})`
  );
  navSheet.getRange(targetRow, 7).setFormula(
    `=IF(E${targetRow}="",,IF(E${targetRow}=0,,F${targetRow}/E${targetRow}))`
  );

  Logger.log(`NAV snapshot updated for ${todayKey}`);
}

function updateCashBalanceForDate_(spreadsheet, dateObj) {
  const sheet = spreadsheet.getSheetByName("CashBalance");
  if (!sheet) throw new Error('Sheet "CashBalance" not found');

  const previous = getLatestCashBalanceBeforeDate_(spreadsheet, dateObj);
  const cashFlowCNY = getCashFlowCNYBetweenDates_(
    spreadsheet,
    previous.date,
    dateObj
  );
  const tradeFlowCNY = getTradeCashFlowCNYBetweenDates_(
    spreadsheet,
    previous.date,
    dateObj
  );
  const balance = previous.balance + cashFlowCNY + tradeFlowCNY;

  upsertCashBalance_(spreadsheet, dateObj, balance, [
    previous.date ? `Prev cash ${formatDateKey_(previous.date)}: ${roundMoney_(previous.balance)}` : "No previous cash balance",
    `External flow: ${roundMoney_(cashFlowCNY)}`,
    `Trade flow: ${roundMoney_(tradeFlowCNY)}`
  ].join("; "));

  return balance;
}

function getLatestCashBalanceBeforeDate_(spreadsheet, dateObj) {
  const sheet = spreadsheet.getSheetByName("CashBalance");
  if (!sheet) return { date: null, balance: 0 };

  const values = sheet.getDataRange().getValues();
  if (values.length < 2) return { date: null, balance: 0 };

  const header = values[0];
  const col = name => header.indexOf(name);
  const iDate = col("Date");
  const iBalance = col("CashBalanceCNY");
  if (iDate < 0 || iBalance < 0) {
    throw new Error("CashBalance sheet must contain Date and CashBalanceCNY columns");
  }

  const targetMs = dateOnlyMs_(dateObj);
  let bestDate = null;
  let bestBalance = 0;
  let bestMs = -Infinity;

  for (let i = 1; i < values.length; i++) {
    const d = values[i][iDate];
    if (!d) continue;
    const ms = dateOnlyMs_(d);
    if (ms < targetMs && ms > bestMs) {
      bestMs = ms;
      bestDate = d;
      bestBalance = Number(values[i][iBalance]) || 0;
    }
  }

  return { date: bestDate, balance: bestBalance };
}

function upsertCashBalance_(spreadsheet, dateObj, balance, note) {
  const sheet = spreadsheet.getSheetByName("CashBalance");
  const targetKey = formatDateKey_(dateObj);
  const lastRow = sheet.getLastRow();
  const rows = lastRow >= 2 ? sheet.getRange(2, 1, lastRow - 1, 1).getValues() : [];

  let targetRow = -1;
  for (let i = 0; i < rows.length; i++) {
    if (!rows[i][0]) continue;
    if (formatDateKey_(rows[i][0]) === targetKey) {
      targetRow = i + 2;
      break;
    }
  }

  if (targetRow === -1) targetRow = lastRow + 1;

  sheet.getRange(targetRow, 1, 1, 4).setValues([[
    dateObj,
    balance,
    "CNY",
    note
  ]]);
}

function getCashFlowCNYBetweenDates_(spreadsheet, startDateExclusive, endDateInclusive) {
  const sheet = spreadsheet.getSheetByName("CashFlow");
  if (!sheet) return 0;

  const values = sheet.getDataRange().getValues();
  if (values.length < 2) return 0;

  const header = values[0];
  const col = name => header.indexOf(name);
  const iDate = col("Date");
  const iAmountCNY = col("AmountCNY");
  if (iDate < 0 || iAmountCNY < 0) {
    throw new Error("CashFlow sheet must contain Date and AmountCNY columns");
  }

  let sum = 0;
  for (let i = 1; i < values.length; i++) {
    const d = values[i][iDate];
    if (!d || !isDateInWindow_(d, startDateExclusive, endDateInclusive)) continue;
    sum += Number(values[i][iAmountCNY]) || 0;
  }
  return sum;
}

function getTradeCashFlowCNYBetweenDates_(spreadsheet, startDateExclusive, endDateInclusive) {
  const sheet = spreadsheet.getSheetByName("Trades");
  if (!sheet) return 0;

  const values = sheet.getDataRange().getValues();
  if (values.length < 2) return 0;

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

  const fxMap = getFxToCnyMap_(spreadsheet);
  let sum = 0;

  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    if (!row[iDate] || !isDateInWindow_(row[iDate], startDateExclusive, endDateInclusive)) continue;

    const side = String(row[iSide] || "").trim().toUpperCase();
    const gross = (Number(row[iQty]) || 0) * (Number(row[iPrice]) || 0);
    if (!side || gross <= 0) continue;

    const fee = iFee >= 0 ? (Number(row[iFee]) || 0) : 0;
    const tax = iTax >= 0 ? (Number(row[iTax]) || 0) : 0;
    const otherCost = iOtherCost >= 0 ? (Number(row[iOtherCost]) || 0) : 0;
    const currency = String(row[iCurrency] || "CNY").trim().toUpperCase();
    const fx = fxMap[currency] || 1;

    if (side === "BUY") {
      sum -= (gross + fee + tax + otherCost) * fx;
    } else if (side === "SELL") {
      sum += (gross - fee - tax - otherCost) * fx;
    }
  }

  return sum;
}

function isDateInWindow_(dateValue, startDateExclusive, endDateInclusive) {
  const ms = dateOnlyMs_(dateValue);
  const startMs = startDateExclusive ? dateOnlyMs_(startDateExclusive) : -Infinity;
  const endMs = dateOnlyMs_(endDateInclusive);
  return ms > startMs && ms <= endMs;
}

function dateOnlyMs_(d) {
  const date = (d instanceof Date) ? d : new Date(d);
  return new Date(date.getFullYear(), date.getMonth(), date.getDate()).getTime();
}

function computeEquityValueCNY_(spreadsheet) {
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

  const fxMap = getFxToCnyMap_(spreadsheet);
  let total = 0;

  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    const qty = Number(row[iQty]) || 0;
    const px = Number(row[iLastPrice]) || 0;
    if (qty <= 0 || px <= 0) continue;

    const market = String(row[iMarket] || "").trim().toUpperCase();
    const currencyRaw = iCurrency >= 0 ? String(row[iCurrency] || "") : "";
    const currency = currencyRaw ? currencyRaw.toUpperCase() : defaultCurrencyByMarket_(market);
    const fx = fxMap[currency] || 1;

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

function getFxToCnyMap_(spreadsheet) {
  const map = { CNY: 1, USD: 1, HKD: 1 };
  const fx = fetchFromGoogle_([
    { symbol: "CURRENCY:USDCNY" },
    { symbol: "CURRENCY:HKDCNY" }
  ]);

  if (fx["CURRENCY:USDCNY"] && fx["CURRENCY:USDCNY"].price > 0) {
    map.USD = fx["CURRENCY:USDCNY"].price;
  }
  if (fx["CURRENCY:HKDCNY"] && fx["CURRENCY:HKDCNY"].price > 0) {
    map.HKD = fx["CURRENCY:HKDCNY"].price;
  }

  return map;
}

function getNetFlowCNYForDate_(spreadsheet, dateObj) {
  const sheet = spreadsheet.getSheetByName("CashFlow");
  if (!sheet) return 0;

  const values = sheet.getDataRange().getValues();
  if (values.length < 2) return 0;

  const header = values[0];
  const col = name => header.indexOf(name);

  const iDate = col("Date");
  const iAmountCNY = col("AmountCNY");

  if (iDate < 0 || iAmountCNY < 0) {
    throw new Error("CashFlow sheet must contain Date and AmountCNY columns");
  }

  const target = formatDateKey_(dateObj);
  let sum = 0;

  for (let i = 1; i < values.length; i++) {
    const d = values[i][iDate];
    if (!d) continue;
    if (formatDateKey_(d) !== target) continue;
    sum += Number(values[i][iAmountCNY]) || 0;
  }

  return sum;
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
    ["Latest Trading Date", '=IFERROR(LOOKUP(2,1/(NAV_Daily!A2:A<>""),NAV_Daily!A2:A),"")', "NAV_Daily latest date"],
    ["Total Asset (CNY)", '=IFERROR(LOOKUP(2,1/(NAV_Daily!D2:D<>""),NAV_Daily!D2:D),0)', "Latest total asset"],
    ["Today PnL", '=IFERROR(LOOKUP(2,1/(NAV_Daily!F2:F<>""),NAV_Daily!F2:F),0)', "Trading-day PnL"],
    ["Today Return", '=IFERROR(LOOKUP(2,1/(NAV_Daily!G2:G<>""),NAV_Daily!G2:G),0)', "Trading-day return"],
    ["MTD PnL", '=IFERROR(SUM(FILTER(NAV_Daily!F2:F,NAV_Daily!A2:A>=EOMONTH(TODAY(),-1)+1,NAV_Daily!A2:A<=TODAY())),0)', "Month-to-date"],
    ["YTD PnL", '=IFERROR(SUM(FILTER(NAV_Daily!F2:F,NAV_Daily!A2:A>=DATE(YEAR(TODAY()),1,1),NAV_Daily!A2:A<=TODAY())),0)', "Year-to-date"],
    ["Since Inception XIRR", '=IFERROR(XIRR({FILTER(CashFlow!F2:F,CashFlow!A2:A<>"",CashFlow!F2:F<>0);-LOOKUP(2,1/(NAV_Daily!D2:D<>""),NAV_Daily!D2:D)},{FILTER(CashFlow!A2:A,CashFlow!A2:A<>"",CashFlow!F2:F<>0);LOOKUP(2,1/(NAV_Daily!A2:A<>""),NAV_Daily!A2:A)}),"")', "Money-weighted return"],
    ["MTD XIRR", '=IFERROR(XIRR({FILTER(CashFlow!F2:F,CashFlow!A2:A>=EOMONTH(TODAY(),-1)+1,CashFlow!A2:A<=TODAY(),CashFlow!F2:F<>0);-XLOOKUP(MAX(FILTER(NAV_Daily!A2:A,NAV_Daily!A2:A<=TODAY(),NAV_Daily!D2:D<>"")),NAV_Daily!A2:A,NAV_Daily!D2:D)},{FILTER(CashFlow!A2:A,CashFlow!A2:A>=EOMONTH(TODAY(),-1)+1,CashFlow!A2:A<=TODAY(),CashFlow!F2:F<>0);XLOOKUP(MAX(FILTER(NAV_Daily!A2:A,NAV_Daily!A2:A<=TODAY(),NAV_Daily!D2:D<>"")),NAV_Daily!A2:A,NAV_Daily!A2:A)}),"")', "Money-weighted monthly"],
    ["YTD XIRR", '=IFERROR(XIRR({FILTER(CashFlow!F2:F,CashFlow!A2:A>=DATE(YEAR(TODAY()),1,1),CashFlow!A2:A<=TODAY(),CashFlow!F2:F<>0);-XLOOKUP(MAX(FILTER(NAV_Daily!A2:A,NAV_Daily!A2:A<=TODAY(),NAV_Daily!D2:D<>"")),NAV_Daily!A2:A,NAV_Daily!D2:D)},{FILTER(CashFlow!A2:A,CashFlow!A2:A>=DATE(YEAR(TODAY()),1,1),CashFlow!A2:A<=TODAY(),CashFlow!F2:F<>0);XLOOKUP(MAX(FILTER(NAV_Daily!A2:A,NAV_Daily!A2:A<=TODAY(),NAV_Daily!D2:D<>"")),NAV_Daily!A2:A,NAV_Daily!A2:A)}),"")', "Money-weighted yearly"]
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
