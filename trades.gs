/**
 * Get the active spreadsheet.
 * This is the default entry point for all portfolio-related functions.
 * @return {Spreadsheet}
 */
function getSpreadsheet_() {
  return SpreadsheetApp.getActiveSpreadsheet();
}

const TRADE_FEE_RULES_ = {
  CN_COMMISSION_RATE: 0.00025,
  CN_MIN_FEE: 5,
  CN_STAMP_DUTY_SELL_RATE: 0.0005,
  HK_TRADING_FEE_RATE: 0.0000565,
  HK_SFC_LEVY_RATE: 0.000027,
  HK_AFRC_LEVY_RATE: 0.0000015,
  HK_CCASS_RATE: 0.000042,
  HK_STAMP_DUTY_RATE: 0.001
};

function normalizeSymbol_(symbol) {
  return String(symbol || "").trim().toUpperCase();
}

function inferMarketFromSymbol_(symbol) {
  const s = normalizeSymbol_(symbol);
  if (/^HKG:\d{4}$/.test(s)) return "HK";
  if (/^\d{6}$/.test(s)) return "CN";
  if (s) return "US";
  return "";
}

function inferCurrencyFromSymbol_(symbol) {
  const market = inferMarketFromSymbol_(symbol);
  if (market === "CN") return "CNY";
  if (market === "HK") return "HKD";
  if (market === "US") return "USD";
  return "";
}

function isCnFundOrEtfSymbol_(symbol) {
  const s = normalizeSymbol_(symbol);
  return /^(5\d{5}|1[56]\d{4})$/.test(s);
}

function roundMoney_(value) {
  return Math.round((Number(value) || 0) * 100) / 100;
}

function calculateTradeCosts_(symbol, side, quantity, price) {
  const market = inferMarketFromSymbol_(symbol);
  const normalizedSide = String(side || "").trim().toUpperCase();
  // Cash dividends use actual withholding tax and fees entered by the user.
  if (normalizedSide === "DIVIDEND") return { fee: 0, tax: 0, otherCost: 0 };
  const gross = (Number(quantity) || 0) * (Number(price) || 0);
  if (!market || gross <= 0) return { fee: 0, tax: 0, otherCost: 0 };

  if (market === "CN") {
    const fee = Math.max(
      TRADE_FEE_RULES_.CN_MIN_FEE,
      gross * TRADE_FEE_RULES_.CN_COMMISSION_RATE
    );
    const tax = normalizedSide === "SELL" && !isCnFundOrEtfSymbol_(symbol)
      ? gross * TRADE_FEE_RULES_.CN_STAMP_DUTY_SELL_RATE
      : 0;
    return {
      fee: roundMoney_(fee),
      tax: roundMoney_(tax),
      otherCost: 0
    };
  }

  if (market === "HK") {
    const feeRate =
      TRADE_FEE_RULES_.HK_TRADING_FEE_RATE +
      TRADE_FEE_RULES_.HK_SFC_LEVY_RATE +
      TRADE_FEE_RULES_.HK_AFRC_LEVY_RATE +
      TRADE_FEE_RULES_.HK_CCASS_RATE;
    return {
      fee: roundMoney_(gross * feeRate),
      tax: Math.ceil(gross * TRADE_FEE_RULES_.HK_STAMP_DUTY_RATE),
      otherCost: 0
    };
  }

  return { fee: 0, tax: 0, otherCost: 0 };
}

function collectKnownTradeNames_(rows, iSymbol, iName) {
  const names = {};
  rows.forEach(row => {
    const symbol = normalizeSymbol_(row[iSymbol]);
    const name = String(row[iName] || "").trim();
    if (symbol && name && !names[symbol]) names[symbol] = name;
  });
  return names;
}

/**
 * Fill derived trade fields so new rows only need Date, Symbol, Side,
 * Quantity, and Price. Existing non-empty cells are preserved.
 */
function fillTradeDerivedFields_(spreadsheet) {
  const sheet = spreadsheet.getSheetByName("Trades");
  if (!sheet) {
    throw new Error('Sheet "Trades" not found');
  }

  const range = sheet.getDataRange();
  const values = range.getValues();
  if (values.length < 2) return {};

  const header = values[0];
  const rows = values.slice(1);
  const col = name => header.indexOf(name);

  const iName = col("Name");
  const iSymbol = col("Symbol");
  const iSide = col("Side");
  const iQty = col("Quantity");
  const iPrice = col("Price");
  const iCurrency = col("Currency");
  const iFee = col("Fee");
  const iTax = col("Tax");
  const iOtherCost = col("OtherCost");

  if ([iName, iSymbol, iSide, iQty, iPrice, iCurrency, iFee, iTax, iOtherCost].some(i => i < 0)) {
    throw new Error("Trades sheet must contain Name, Symbol, Side, Quantity, Price, Currency, Fee, Tax, and OtherCost columns");
  }

  const knownNames = collectKnownTradeNames_(rows, iSymbol, iName);
  let updated = false;

  rows.forEach(row => {
    const symbol = normalizeSymbol_(row[iSymbol]);
    if (!symbol) return;

    if (!row[iName] && knownNames[symbol]) {
      row[iName] = knownNames[symbol];
      updated = true;
    }

    if (!row[iCurrency]) {
      row[iCurrency] = inferCurrencyFromSymbol_(symbol);
      updated = true;
    }

    const costs = calculateTradeCosts_(
      symbol,
      row[iSide],
      row[iQty],
      row[iPrice]
    );

    if (row[iFee] === "" || row[iFee] === null) {
      row[iFee] = costs.fee;
      updated = true;
    }
    if (row[iTax] === "" || row[iTax] === null) {
      row[iTax] = costs.tax;
      updated = true;
    }
    if (row[iOtherCost] === "" || row[iOtherCost] === null) {
      row[iOtherCost] = costs.otherCost;
      updated = true;
    }
  });

  if (updated) {
    // Keep symbols like "002317" as text. Skip if the column is a typed
    // Table column (format is fixed there and setNumberFormat throws).
    try {
      sheet.getRange(2, iSymbol + 1, rows.length, 1).setNumberFormat("@");
    } catch (e) {
      Logger.log("Skip Trades Symbol number format: " + e.message);
    }
    range.setValues([header].concat(rows));
    Logger.log("Trades derived fields updated.");
  }

  return knownNames;
}

function calculatePositionsFromTradeRows_(header, rows) {
  const col = (name) => {
    const idx = header.indexOf(name);
    if (idx === -1) throw new Error(`Missing column in Trades: ${name}`);
    return idx;
  };

  const iDate   = col("Date");
  const iName   = header.includes("Name") ? col("Name") : -1;
  const iSymbol = col("Symbol");
  const iSide   = col("Side");
  const iQty    = col("Quantity");
  const iPrice  = col("Price");
  const iCurrency  = header.includes("Currency") ? col("Currency") : -1;
  const iFee       = header.includes("Fee") ? col("Fee") : -1;
  const iTax       = header.includes("Tax") ? col("Tax") : -1;
  const iOtherCost = header.includes("OtherCost") ? col("OtherCost") : -1;

  const positions = {};

  const sortedRows = rows
    .map((row, idx) => ({ row, idx }))
    .filter(item => normalizeSymbol_(item.row[iSymbol]))
    .sort((a, b) => {
      const aMs = a.row[iDate] ? dateOnlyMs_(a.row[iDate]) : 0;
      const bMs = b.row[iDate] ? dateOnlyMs_(b.row[iDate]) : 0;
      return aMs === bMs ? a.idx - b.idx : aMs - bMs;
    });

  sortedRows.forEach(item => {
    const r = item.row;
    const symbol = normalizeSymbol_(r[iSymbol]);

    const side  = String(r[iSide]).trim().toUpperCase();
    const qty   = Number(r[iQty]) || 0;
    const price = Number(r[iPrice]) || 0;
    const fee       = iFee >= 0 ? (Number(r[iFee]) || 0) : 0;
    const tax       = iTax >= 0 ? (Number(r[iTax]) || 0) : 0;
    const otherCost = iOtherCost >= 0 ? (Number(r[iOtherCost]) || 0) : 0;

    if (!positions[symbol]) {
      positions[symbol] = {
        name: "",
        market: inferMarketFromSymbol_(symbol),
        currency: inferCurrencyFromSymbol_(symbol),
        quantity: 0,
        cost: 0,
        realizedPnL: 0
      };
    }

    const p = positions[symbol];
    if (iName >= 0 && r[iName]) p.name = r[iName];
    if (iCurrency >= 0 && r[iCurrency]) {
      p.currency = String(r[iCurrency]).trim().toUpperCase();
    }

    if (side === "BUY") {
      p.quantity += qty;
      p.cost += qty * price + fee + tax + otherCost;

    } else if (side === "SELL") {
      if (qty > p.quantity) {
        throw new Error(
          `Sell quantity exceeds position for ${symbol} at row ${item.idx + 2}`
        );
      }

      const avgCost = p.quantity > 0 ? p.cost / p.quantity : 0;
      const releasedCost = avgCost * qty;
      const proceeds = qty * price - fee - tax - otherCost;
      p.realizedPnL += proceeds - releasedCost;
      p.quantity -= qty;
      p.cost -= releasedCost;

      if (Math.abs(p.quantity) < 1e-9) {
        p.quantity = 0;
        p.cost = 0;
      }
    } else if (side === "DIVIDEND") {
      // Quantity is the eligible share count; Price is the gross dividend/share.
      // Payment may arrive after a sale. Do not alter shares or acquisition cost.
      p.realizedPnL += qty * price - fee - tax - otherCost;
    }
  });

  return positions;
}

/**
 * Build positions from Trades sheet using weighted average cost.
 * BUY increases quantity and cost. SELL releases average cost and
 * records realized PnL. DIVIDEND adds net cash income to realized PnL only.
 * @param {Spreadsheet} spreadsheet
 * @return {Object} Map: symbol -> position
 */
function buildPositionsFromTrades_(spreadsheet) {
  const tradesSheet = spreadsheet.getSheetByName("Trades");
  if (!tradesSheet) {
    throw new Error('Sheet "Trades" not found');
  }

  const values = tradesSheet.getDataRange().getValues();
  if (values.length < 2) return {};

  return calculatePositionsFromTradeRows_(values[0], values.slice(1));
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
  const symbols = Object.keys(positions).sort();

  // Clear existing data but keep header row
  const lastRow = sheet.getLastRow();
  if (lastRow > 2) {
    sheet
      .getRange(3, 1, lastRow - 2, sheet.getLastColumn())
      .clearContent();
  }

  if (symbols.length === 0) {
    Logger.log("No trade-derived positions.");
    return;
  }

  // Positions Symbol column is a Table column fixed to plain text, so no
  // setNumberFormat here (it throws on typed Table columns).
  const rows = symbols.map(symbol => {
    const p = positions[symbol];
    const avgCost = p.quantity > 0 ? p.cost / p.quantity : 0;
    const status = p.quantity > 0 ? "OPEN" : "CLOSED";

    return [
      "",          // Market
      p.name || "",// Name
      symbol,      // Symbol
      "",          // Currency
      status,      // Status
      p.quantity,  // PositionQty
      avgCost,     // AvgCost
      p.cost,      // TotalCost
      "",          // LastPrice (filled by formula or later)
      "",          // MarketValue (formula)
      "",          // UnrealizedPnL (formula)
      "",          // UnrealizedPnL% (formula)
      p.realizedPnL,// RealizedPnL
      "",          // TotalPnL (formula)
      "",          // TotalPnL% (formula)
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
    const symbol = normalizeSymbol_(rows[i][iSymbol]);
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
