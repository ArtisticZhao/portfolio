/**
 * Get the active spreadsheet.
 * This is the default entry point for all portfolio-related functions.
 * @return {Spreadsheet}
 */
function getSpreadsheet_() {
  return SpreadsheetApp.getActiveSpreadsheet();
}

/**
 * Build positions from Trades sheet using weighted average cost.
 * BUY increases quantity and cost.
 * SELL reduces quantity and cost at average cost, including sell fees.
 * @param {Spreadsheet} spreadsheet
 * @return {Object} Map: symbol -> { quantity, cost }
 */
function buildPositionsFromTrades_(spreadsheet) {
  const tradesSheet = spreadsheet.getSheetByName("Trades");
  if (!tradesSheet) {
    throw new Error('Sheet "Trades" not found');
  }

  const values = tradesSheet.getDataRange().getValues();
  if (values.length < 2) return {};

  const header = values[0];
  const rows = values.slice(1);

  const col = (name) => {
    const idx = header.indexOf(name);
    if (idx === -1) throw new Error(`Missing column in Trades: ${name}`);
    return idx;
  };

  const iSymbol = col("Symbol");
  const iSide   = col("Side");
  const iQty    = col("Quantity");
  const iPrice  = col("Price");
  const iFee    = header.includes("Fee") ? col("Fee") : -1;
  const iTax    = header.includes("Tax") ? col("Tax") : -1;

  const positions = {};

  for (let rowIdx = 0; rowIdx < rows.length; rowIdx++) {
    const r = rows[rowIdx];
    const symbol = r[iSymbol];
    if (!symbol) continue;

    const side  = String(r[iSide]).trim().toUpperCase();
    const qty   = Number(r[iQty]) || 0;
    const price = Number(r[iPrice]) || 0;
    const fee   = iFee >= 0 ? (Number(r[iFee]) || 0) : 0;
    const tax   = iTax >= 0 ? (Number(r[iTax]) || 0) : 0;

    if (!positions[symbol]) {
      positions[symbol] = { quantity: 0, cost: 0 };
    }

    const p = positions[symbol];

    if (side === "BUY") {
      p.quantity += qty;
      p.cost += qty * price + fee + tax;

    } else if (side === "SELL") {
      if (qty > p.quantity) {
        throw new Error(
          `Sell quantity exceeds position for ${symbol} at row ${rowIdx + 2}`
        );
      }

      const avgCost = p.quantity > 0 ? p.cost / p.quantity : 0;
      p.quantity -= qty;
      p.cost -= avgCost * qty;
      p.cost -= (fee + tax); // reduce remaining cost by sell expenses
    }
  }

  return positions;
}


/**
 * Fill Positions sheet from Trades.
 * This function rebuilds all position rows to match the Positions header.
 * Market price related fields are left empty for formulas or later updates.
 * @param {Spreadsheet} spreadsheet
 */
function fillPositionsFromTrades_(spreadsheet) {
  const sheet = spreadsheet.getSheetByName("Positions");
  if (!sheet) {
    throw new Error('Sheet "Positions" not found');
  }

  const positions = buildPositionsFromTrades_(spreadsheet);
  const symbols = Object.keys(positions).filter(
    s => positions[s].quantity > 0
  );

  // Clear existing data but keep header row
  const lastRow = sheet.getLastRow();
  if (lastRow > 2) {
    sheet
      .getRange(3, 1, lastRow - 2, sheet.getLastColumn())
      .clearContent();
  }

  if (symbols.length === 0) {
    Logger.log("No open positions.");
    return;
  }

  const rows = symbols.map(symbol => {
    const p = positions[symbol];
    const avgCost = p.quantity > 0 ? p.cost / p.quantity : 0;

    return [
      "",          // Market
      "",          // Name
      symbol,      // Symbol
      "",          // Currency
      p.quantity,  // PositionQty
      avgCost,     // AvgCost
      p.cost,      // TotalCost
      "",          // LastPrice (filled by formula or later)
      "",          // MarketValue (formula)
      "",          // UnrealizedPnL (formula)
      "",          // UnrealizedPnL% (formula)
      ""           // Weight (formula)
    ];
  });

  sheet
    .getRange(3, 1, rows.length, rows[0].length)
    .setValues(rows);

  Logger.log(`Positions rebuilt: ${rows.length} symbols.`);
}

/**
 * Fill missing Name values in Trades sheet based on symbolInfoMap.
 * Only empty Name cells will be updated.
 *
 * @param {Spreadsheet} spreadsheet
 * @param {Object} symbolInfoMap  symbol -> { name, price }
 */
function fillTradesNameFromSymbol_(spreadsheet, symbolInfoMap) {
  const sheet = spreadsheet.getSheetByName("Trades");
  if (!sheet) {
    throw new Error('Sheet "Trades" not found');
  }

  if (!symbolInfoMap || typeof symbolInfoMap !== "object") {
    Logger.log("No symbol info provided, skip filling trade names.");
    return;
  }

  const range = sheet.getDataRange();
  const values = range.getValues();
  if (values.length < 2) return;

  const header = values[0];
  const rows = values.slice(1);

  const col = name => header.indexOf(name);

  const iSymbol = col("Symbol");
  const iName   = col("Name");

  if (iSymbol < 0 || iName < 0) {
    throw new Error("Trades sheet must contain Symbol and Name columns");
  }

  let updated = false;

  for (let i = 0; i < rows.length; i++) {
    const symbol = rows[i][iSymbol];
    const name   = rows[i][iName];

    // Only fill when Name is empty
    if (!symbol || name) continue;

    const info = symbolInfoMap[symbol];
    if (info && info.name) {
      rows[i][iName] = info.name;
      updated = true;
    }
  }

  if (updated) {
    // Write back only once
    range.setValues([header].concat(rows));
    Logger.log("Trades Name column updated successfully.");
  } else {
    Logger.log("No missing Names to update in Trades.");
  }
}
