/**
 * Fetch latest prices for a list of assets.
 * @param {Array<Object>} assets
 *   [{ symbol, market, currency }]
 * @return {Object} symbol -> price
 */
function fetchLatestPrices_(assets) {
  const result = {};

  const sinaAssets = [];
  const googleAssets = [];

  assets.forEach(a => {
    if (a.market === "CN") {
      sinaAssets.push(a);
    } else {
      googleAssets.push(a);
    }
  });

  Object.assign(result, fetchFromSina_(sinaAssets));
  Object.assign(result, fetchFromGoogle_(googleAssets));

  return result;
}

/**
 * Fetch latest prices and names using GOOGLEFINANCE.
 * Assumes symbols are already in valid Google Finance format.
 *
 * @param {Array<Object>} assets [{ symbol }]
 * @return {Object} symbol -> { name, price }
 */
function fetchFromGoogle_(assets) {
  const result = {};
  if (!assets || assets.length === 0) return result;

  // Create a temporary spreadsheet
  const tempSs = SpreadsheetApp.create("TempGoogleFinanceFetcher");
  const tempSheet = tempSs.getSheets()[0];

  try {
    // Write formulas
    assets.forEach((a, index) => {
      const row = index + 1;

      // Name
      tempSheet
        .getRange(row, 1)
        .setFormula(`=IFERROR(GOOGLEFINANCE("${a.symbol}", "name"), "")`);

      // Price (fallback to last close). Currency pairs only resolve without
      // an attribute ("price" yields #N/A), and a "close" query returns a
      // Date/Close table, so the value has to be picked out with INDEX.
      const live = /^CURRENCY:/i.test(a.symbol)
        ? `GOOGLEFINANCE("${a.symbol}")`
        : `GOOGLEFINANCE("${a.symbol}", "price")`;
      tempSheet
        .getRange(row, 2)
        .setFormula(
          `=IFERROR(` +
          `${live},` +
          `IFERROR(INDEX(GOOGLEFINANCE("${a.symbol}", "close", WORKDAY(TODAY()-1,-1)), 2, 2), "")` +
          `)`
        );
    });

    // Wait for formulas to evaluate
    SpreadsheetApp.flush();
    Utilities.sleep(8000);

    const values = tempSheet
      .getRange(1, 1, assets.length, 2)
      .getValues();

    values.forEach((row, idx) => {
      const name = row[0];
      const price = row[1];
      const symbol = assets[idx].symbol;

      if (typeof price === "number" && price > 0) {
        result[symbol] = {
          name: name || "",
          price: price
        };
      }
    });

  } finally {
    // Clean up temp spreadsheet
    DriveApp.getFileById(tempSs.getId()).setTrashed(true);
  }

  return result;
}

/**
 * Fetch historical daily close prices for a batch of symbols in one shot,
 * using GOOGLEFINANCE's date-range query. One temp spreadsheet, one wait,
 * one read for the whole batch (unlike fetchFromGoogle_'s live-price calls,
 * this returns a Date/Close table per symbol that needs to spill, so each
 * symbol gets its own 2-column slot).
 *
 * @param {Array<Object>} assets [{ symbol, googleSymbol? }] googleSymbol
 *   overrides what's passed to GOOGLEFINANCE (e.g. "CURRENCY:USDCNY"),
 *   while the returned map is still keyed by `symbol`.
 * @param {string} startDateKey "yyyy-MM-dd"
 * @param {string} endDateKey "yyyy-MM-dd"
 * @return {Object} symbol -> { "yyyy-MM-dd": close }
 */
function fetchGoogleFinanceHistoryBatch_(assets, startDateKey, endDateKey) {
  const result = {};
  if (!assets || assets.length === 0) return result;

  const start = new Date(startDateKey + "T00:00:00");
  const end = new Date(endDateKey + "T00:00:00");
  const startExpr = `DATE(${start.getFullYear()},${start.getMonth() + 1},${start.getDate()})`;
  const endExpr = `DATE(${end.getFullYear()},${end.getMonth() + 1},${end.getDate()})`;

  const tempSs = SpreadsheetApp.create("TempGoogleFinanceHistoryFetcher");
  const tempSheet = tempSs.getSheets()[0];

  try {
    // No IFERROR wrapper: it strips the date format off the spilled table,
    // so the Date column would come back as serial numbers. A failed lookup
    // just leaves "#N/A" in the first cell, which is skipped below.
    assets.forEach((a, index) => {
      const col = index * 2 + 1;
      const sym = a.googleSymbol || a.symbol;
      tempSheet
        .getRange(1, col)
        .setFormula(`=GOOGLEFINANCE("${sym}","close",${startExpr},${endExpr})`);
    });

    SpreadsheetApp.flush();
    Utilities.sleep(10000);

    // The table's timestamps are in the temp spreadsheet's timezone, which
    // need not match the script's.
    const tz = tempSs.getSpreadsheetTimeZone();
    const rowCount = tempSheet.getMaxRows();

    assets.forEach((a, index) => {
      const col = index * 2 + 1;
      const rows = tempSheet.getRange(1, col, rowCount, 2).getValues();
      const map = {};
      rows.forEach(r => {
        if (!(r[0] instanceof Date)) return;
        const close = Number(r[1]);
        if (!isNaN(close) && close > 0) {
          map[Utilities.formatDate(r[0], tz, "yyyy-MM-dd")] = close;
        }
      });
      result[a.symbol] = map;
    });
  } finally {
    DriveApp.getFileById(tempSs.getId()).setTrashed(true);
  }

  return result;
}

/**
 * Fetch latest prices and names for A-shares from Sina.
 * @param {Array<Object>} assets [{ symbol }]
 * @return {Object} symbol -> { name, price }
 */
function fetchFromSina_(assets) {
  const result = {};
  if (!assets || assets.length === 0) return result;

  const sinaSymbols = [];
  const reverseMap = {}; // sinaSymbol -> original symbol

  assets.forEach(a => {
    const sina = convertToSinaSymbol_(a.symbol);
    if (sina) {
      sinaSymbols.push(sina);
      reverseMap[sina] = a.symbol;
    }
  });

  if (sinaSymbols.length === 0) return result;

  const url = `https://hq.sinajs.cn/list=${sinaSymbols.join(",")}`;
  const params = {
    muteHttpExceptions: true,
    headers: {
      Referer: "https://finance.sina.com.cn"
    }
  };

  let responseText;
  try {
    responseText = UrlFetchApp.fetch(url, params).getContentText("GBK");
  } catch (e) {
    Logger.log(`Sina fetch failed: ${e.message}`);
    return result;
  }

  const lines = responseText.split(";");
  lines.forEach(line => {
    if (!line) return;

    const match = line.match(/hq_str_(\w+)=["](.*)["]/);
    if (!match) return;

    const sinaSymbol = match[1];
    const dataStr = match[2];
    const fields = dataStr.split(",");

    if (fields.length < 4) return;

    const name = fields[0];
    const price = parseFloat(fields[3]);

    if (
      reverseMap[sinaSymbol] &&
      !isNaN(price) &&
      price > 0
    ) {
      const symbol = reverseMap[sinaSymbol];
      result[symbol] = {
        name: name,
        price: price
      };
    }
  });

  return result;
}

/**
 * Fetch historical daily closes (unadjusted) for one A-share stock or ETF
 * from Sina's K-line service. The service returns the most recent N trading
 * days (N <= 1023, about four years), so anything older is not available.
 *
 * @param {string} symbol e.g. "513300"
 * @param {string} startDateKey "yyyy-MM-dd"
 * @param {string} endDateKey "yyyy-MM-dd"
 * @return {Object} "yyyy-MM-dd" -> close
 */
function fetchSinaDailyHistory_(symbol, startDateKey, endDateKey) {
  const sina = convertToSinaSymbol_(symbol);
  if (!sina) throw new Error(`Cannot convert "${symbol}" to a Sina symbol`);

  const url = "https://quotes.sina.cn/cn/api/json_v2.php/CN_MarketDataService.getKLineData" +
    `?symbol=${sina}&scale=240&ma=no&datalen=1023`;
  const response = UrlFetchApp.fetch(url, {
    muteHttpExceptions: true,
    headers: {
      Referer: "https://finance.sina.com.cn"
    }
  });

  const text = response.getContentText();
  let rows;
  try {
    rows = JSON.parse(text);
  } catch (e) {
    rows = null;
  }
  if (!Array.isArray(rows)) {
    throw new Error(`Unexpected Sina response (HTTP ${response.getResponseCode()}): ${text.slice(0, 200)}`);
  }

  const result = {};
  rows.forEach(row => {
    const close = parseFloat(row.close);
    if (row.day >= startDateKey && row.day <= endDateKey && close > 0) {
      result[row.day] = close;
    }
  });
  return result;
}

/**
 * Convert CN listed security symbol to Sina format.
 * Supports:
 * - SSE stocks: 6xxxxx -> sh6xxxxx
 * - SZSE stocks: 0xxxxx / 3xxxxx -> sz0xxxxx / sz3xxxxx
 * - SSE ETFs: 5xxxxx -> sh5xxxxx
 * - SZSE ETFs: 1xxxxx -> sz1xxxxx
 *
 * @param {string} symbol
 * @return {string|null}
 */
function convertToSinaSymbol_(symbol) {
  const s = String(symbol).trim();

  // SSE: stocks and ETFs
  if (/^[65]\d{5}$/.test(s)) return "sh" + s;

  // SZSE: stocks and ETFs
  if (/^[013]\d{5}$/.test(s)) return "sz" + s;

  return null;
}
