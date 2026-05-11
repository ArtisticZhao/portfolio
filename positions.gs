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
  const iLastPrice = col("LastPrice");

  const cnAssets = [];
  const googleAssets = [];

  for (let i = 1; i < data.length; i++) {
    const symbol = data[i][iSymbol];
    const market = data[i][iMarket];
    if (!symbol || !market) continue;

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
