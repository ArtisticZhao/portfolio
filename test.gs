/**
 * Test fetchFromGoogle_ without interacting with portfolio sheets.
 */
function testFetchFromGoogle_() {
  const testAssets = [
    { symbol: "AAPL" },
    { symbol: "MSFT" },
    { symbol: "HKG:0700" },
    { symbol: "HKG:9988" }
  ];

  const result = fetchFromGoogle_(testAssets);

  Logger.log("=== Google Finance Fetch Result ===");
  Object.keys(result).forEach(symbol => {
    const item = result[symbol];
    Logger.log(`${symbol} | ${item.name} | ${item.price}`);
  });
}

/**
 * Test fetchFromSina_ with name and price.
 */
function testFetchFromSinaWithName_() {
  const testAssets = [
    { symbol: "600519" },
    { symbol: "000001" },
    { symbol: "513520" }
  ];

  const result = fetchFromSina_(testAssets);

  Logger.log("=== Sina Fetch Result (Name + Price) ===");
  Object.keys(result).forEach(symbol => {
    const item = result[symbol];
    Logger.log(`${symbol} | ${item.name} | ${item.price}`);
  });
}

function assertAlmostEqual_(actual, expected, label) {
  if (Math.abs((Number(actual) || 0) - expected) > 0.000001) {
    throw new Error(`${label}: expected ${expected}, got ${actual}`);
  }
}

function testBuildPositionsFromTradesWeightedCost_() {
  const header = ["Date", "Name", "Symbol", "Side", "Quantity", "Price", "Currency", "Fee", "Tax", "OtherCost"];
  const rows = [
    [new Date("2026-01-01"), "Alpha", "AAA", "BUY", 100, 10, "USD", 1, 0, 0],
    [new Date("2026-01-02"), "Alpha", "AAA", "BUY", 100, 20, "USD", 1, 0, 0],
    [new Date("2026-01-03"), "Alpha", "AAA", "SELL", 50, 30, "USD", 2, 0, 0],
    [new Date("2026-01-04"), "Beta", "BBB", "BUY", 10, 5, "USD", 0, 0, 0],
    [new Date("2026-01-05"), "Beta", "BBB", "SELL", 10, 7, "USD", 1, 0, 0]
  ];

  const positions = calculatePositionsFromTradeRows_(header, rows);
  assertAlmostEqual_(positions.AAA.quantity, 150, "AAA remaining quantity");
  assertAlmostEqual_(positions.AAA.cost, 2251.5, "AAA remaining cost");
  assertAlmostEqual_(positions.AAA.realizedPnL, 747.5, "AAA realized PnL");
  assertAlmostEqual_(positions.BBB.quantity, 0, "BBB closed quantity");
  assertAlmostEqual_(positions.BBB.cost, 0, "BBB closed cost");
  assertAlmostEqual_(positions.BBB.realizedPnL, 19, "BBB realized PnL");
}

/**
 * No more AutoDeposit floor: cash balance is a plain running sum of trade
 * + external flows, and is allowed to go negative (a real signal, not
 * something to paper over with a synthetic deposit).
 */
function testWalkCashBalancesNoAutoDeposit_() {
  const tradeFlows = {
    "2026-01-01": -100,
    "2026-01-02": 40,
    "2026-01-03": -60
  };
  const externalFlows = {
    "2026-01-01": 50
  };

  const rows = walkCashBalances_(tradeFlows, externalFlows);
  assertAlmostEqual_(rows[0].cashBalance, -50, "day 1 balance (no auto deposit, negative allowed)");
  assertAlmostEqual_(rows[1].cashBalance, -10, "day 2 balance");
  assertAlmostEqual_(rows[2].cashBalance, -70, "day 3 balance");
}
