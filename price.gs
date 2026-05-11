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

      // Price (fallback to last close)
      tempSheet
        .getRange(row, 2)
        .setFormula(
          `=IFERROR(` +
          `GOOGLEFINANCE("${a.symbol}", "price"),` +
          `IFERROR(GOOGLEFINANCE("${a.symbol}", "close", WORKDAY(TODAY()-1,-1)), "")` +
          `)`
        );
    });

    // Wait for formulas to evaluate
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
