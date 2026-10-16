/**
 * Update LastPrice in Positions sheet by fetching latest prices.
 * @param {Spreadsheet} spreadsheet
 */
function updatePositionsPrices_(spreadsheet) {
  const sheet = spreadsheet.getSheetByName("Positions");
  if (!sheet) {
    throw new Error('Sheet "Positions" not found');
  }

  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return;

  const header = data[0];
  const col = name => header.indexOf(name);

  const iMarket = col("Market");
  const iName = col("Name");
  const iSymbol = col("Symbol");
  const iStatus = col("Status");
  const iQty = col("PositionQty");
  const iLastPrice = col("LastPrice");

  const cnAssets = [];
  const googleAssets = [];

  for (let i = 1; i < data.length; i++) {
    const symbol = data[i][iSymbol];
    const market = data[i][iMarket];
    const status = iStatus >= 0 ? String(data[i][iStatus] || "").trim().toUpperCase() : "OPEN";
    const qty = iQty >= 0 ? (Number(data[i][iQty]) || 0) : 0;
    if (!symbol || !market) continue;
    if (status === "CLOSED" || qty <= 0) continue;

    if (market === "CN") {
      cnAssets.push({ symbol: symbol, row: i + 1 });
    } else if (market === "US" || market === "HK") {
      googleAssets.push({ symbol: symbol, row: i + 1 });
    }
  }

  // Fetch prices
  const cnPrices = fetchFromSina_(cnAssets);
  const googlePrices = fetchFromGoogle_(googleAssets);

  // Write back prices
  cnAssets.forEach(a => {
    const item = cnPrices[a.symbol];
    if (item && item.price > 0) {
      sheet.getRange(a.row, iLastPrice + 1).setValue(item.price);
    }
    if (item.name) {
      sheet.getRange(a.row, iName + 1).setValue(item.name);
    }
  });

  googleAssets.forEach(a => {
    const item = googlePrices[a.symbol];
    if (item && item.price > 0) {
      sheet.getRange(a.row, iLastPrice + 1).setValue(item.price);
    }
    if (item.name) {
      sheet.getRange(a.row, iName + 1).setValue(item.name);
    }
  });

  const symbolInfoMap = {};
  Object.assign(symbolInfoMap, cnPrices);
  Object.assign(symbolInfoMap, googlePrices);
  return symbolInfoMap;
}

// Positions columns the script writes; every other column is formula-driven
// and left empty, as in fillPositionsFromTrades_.
const POSITIONS_SCRIPT_COLUMNS_ = [
  "Name", "Symbol", "Status", "PositionQty", "AvgCost", "TotalCost", "LastPrice", "RealizedPnL"
];

/**
 * Reorder Positions rows: CLOSED first, then open positions by Weight
 * descending. Run after updatePositionsPrices_ so Weight reflects today's
 * prices.
 * @param {Spreadsheet} spreadsheet
 */
function sortPositionsRows_(spreadsheet) {
  const sheet = spreadsheet.getSheetByName("Positions");
  if (!sheet) {
    throw new Error('Sheet "Positions" not found');
  }

  // Same layout as fillPositionsFromTrades_: header in row 1, data from row 3.
  const data = sheet.getDataRange().getValues();
  if (data.length < 4) return;

  const header = data[0];
  const col = name => header.indexOf(name);
  const iSymbol = col("Symbol");
  const iStatus = col("Status");
  const iQty = col("PositionQty");
  const iWeight = col("Weight") >= 0 ? col("Weight") : col("MarketValue");
  const keep = POSITIONS_SCRIPT_COLUMNS_.map(col).filter(i => i >= 0);

  const isClosed = row =>
    String(row[iStatus] || "").trim().toUpperCase() === "CLOSED" || !(Number(row[iQty]) > 0);
  const weight = row => iWeight >= 0 ? (Number(row[iWeight]) || 0) : 0;

  const rows = data.slice(2).filter(row => row[iSymbol]);
  if (rows.length < 2) return;
  rows.sort((a, b) => {
    const ca = isClosed(a), cb = isClosed(b);
    if (ca !== cb) return ca ? -1 : 1;
    if (!ca && weight(a) !== weight(b)) return weight(b) - weight(a);
    return String(a[iSymbol]) < String(b[iSymbol]) ? -1 : 1;
  });

  const out = rows.map(row => header.map((_, i) => keep.indexOf(i) >= 0 ? row[i] : ""));
  sheet.getRange(3, 1, out.length, header.length).setValues(out);
}
