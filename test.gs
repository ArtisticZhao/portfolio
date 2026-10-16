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
 * Buys are paid from cash on hand (earlier proceeds) first; only the rest
 * is new money. Cash never goes negative.
 */
function testWalkCashBalancesAutoFunding_() {
  const rows = walkCashBalances_({
    "2026-01-01": { buy: 100, proceeds: 0 },
    "2026-01-02": { buy: 0, proceeds: 40 },
    "2026-01-03": { buy: 60, proceeds: 0 },
    "2026-01-05": { buy: 30, proceeds: 50 }
  });
  assertAlmostEqual_(rows[0].newMoney, 100, "first buy is all new money");
  assertAlmostEqual_(rows[0].reinvested, 0, "nothing to reinvest yet");
  assertAlmostEqual_(rows[0].cashBalance, 0, "day 1 balance");
  assertAlmostEqual_(rows[1].cashBalance, 40, "sale proceeds stay in the account");
  assertAlmostEqual_(rows[1].tradeFlow, 40, "net trade flow");
  assertAlmostEqual_(rows[2].reinvested, 40, "buy uses cash on hand first");
  assertAlmostEqual_(rows[2].newMoney, 20, "only the shortfall is new money");
  assertAlmostEqual_(rows[2].cashBalance, 0, "day 3 balance");
  assertAlmostEqual_(rows[3].reinvested, 30, "same-day sale funds the buy");
  assertAlmostEqual_(rows[3].newMoney, 0, "no new money needed");
  assertAlmostEqual_(rows[3].cashBalance, 20, "day 5 balance");
  assertAlmostEqual_(rows[3].totalInvested, 120, "total real money put in");
}

function assertEqual_(actual, expected, label) {
  if (actual !== expected) {
    throw new Error(`${label}: expected ${expected}, got ${actual}`);
  }
}

const TEST_TRADE_HEADER_ = ["Date", "Name", "Symbol", "Side", "Quantity", "Price", "Currency", "Fee", "Tax", "OtherCost"];

function testTradeRow_(dateKey, symbol, side, qty, price) {
  return [new Date(dateKey + "T00:00:00"), "", symbol, side, qty, price, "CNY", 5, 0, 0];
}

/**
 * InputSig must change when a trade is back-filled or edited, and must not
 * change when rows are merely reordered.
 */
function testDailyInputSignatures_() {
  const buy = testTradeRow_("2026-10-08", "600519", "BUY", 100, 10);
  const sell = testTradeRow_("2026-10-09", "600519", "SELL", 100, 11);
  const late = testTradeRow_("2026-10-08", "000001", "BUY", 200, 5);

  const base = buildDailyInputSignatures_([TEST_TRADE_HEADER_, buy, sell]);
  assertEqual_(base["2026-10-08"].split(":")[0], "1", "one trade counted");
  assertEqual_(base["2026-10-07"], undefined, "no signature without input");

  const backfilled = buildDailyInputSignatures_([TEST_TRADE_HEADER_, buy, sell, late]);
  assertEqual_(backfilled["2026-10-08"].split(":")[0], "2", "back-filled trade counted");
  assertEqual_(backfilled["2026-10-09"], base["2026-10-09"], "other day untouched");

  const reordered = buildDailyInputSignatures_([TEST_TRADE_HEADER_, late, sell, buy]);
  assertEqual_(reordered["2026-10-08"], backfilled["2026-10-08"], "row order ignored");

  const edited = buildDailyInputSignatures_(
    [TEST_TRADE_HEADER_, testTradeRow_("2026-10-08", "600519", "BUY", 100, 10.5), sell]
  );
  assertEqual_(edited["2026-10-08"].split(":")[0], "1", "edit keeps the count");
  assertEqual_(edited["2026-10-08"] !== base["2026-10-08"], true, "edited price changes signature");

  // 2026-10-10 is a Saturday: accounted for on Monday's NAV row.
  const weekend = buildDailyInputSignatures_(
    [TEST_TRADE_HEADER_, testTradeRow_("2026-10-10", "600519", "BUY", 100, 10)]
  );
  assertEqual_(weekend["2026-10-12"].split(":")[0], "1", "weekend trade rolls to Monday");
}

function testFindFirstDirtyIndex_() {
  const keys = ["2026-10-08", "2026-10-09", "2026-10-12", "2026-10-13"];
  const sigs = { "2026-10-08": "1:aaaaaaaa", "2026-10-12": "2:bbbbbbbb" };
  const nav = () => [
    { dateKey: "2026-10-08", quality: "", sig: "1:aaaaaaaa" },
    { dateKey: "2026-10-09", quality: "", sig: "" },
    { dateKey: "2026-10-12", quality: "", sig: "2:bbbbbbbb" }
  ];

  // Weekend run with nothing new: fully up to date.
  assertEqual_(findFirstDirtyIndex_(keys.slice(0, 3), nav(), sigs, "2026-10-17"), -1, "clean");
  // Weekday run: only today's row is rewritten.
  assertEqual_(findFirstDirtyIndex_(keys, nav(), sigs, "2026-10-13"), 3, "today only");

  // A trade back-filled on 10-09 after that row was computed.
  const backfilled = Object.assign({}, sigs, { "2026-10-09": "1:cccccccc" });
  assertEqual_(findFirstDirtyIndex_(keys, nav(), backfilled, "2026-10-13"), 1, "back-fill");

  // The only trade of a day was deleted.
  const deleted = { "2026-10-12": "2:bbbbbbbb" };
  assertEqual_(findFirstDirtyIndex_(keys, nav(), deleted, "2026-10-13"), 0, "deletion");

  // Script did not run on 10-09: the row is missing.
  const gap = nav();
  gap.splice(1, 1);
  assertEqual_(findFirstDirtyIndex_(keys, gap, sigs, "2026-10-13"), 1, "missing row");

  // Yesterday's live snapshot gets finalized with the close.
  const live = nav();
  live[2].quality = "LIVE";
  assertEqual_(findFirstDirtyIndex_(keys, live, sigs, "2026-10-13"), 2, "live finalize");

  assertEqual_(findFirstDirtyIndex_(keys, [], sigs, "2026-10-13"), 0, "first run");
}

/**
 * A trade back-filled on day 1 changes holdings, and so NAV, on every later
 * day as well.
 */
function testComputeNavRowsAfterBackfill_() {
  const keys = ["2026-10-08", "2026-10-09"];
  const trades = [
    TEST_TRADE_HEADER_,
    testTradeRow_("2026-10-09", "600519", "SELL", 50, 12),
    testTradeRow_("2026-10-08", "600519", "BUY", 100, 10)
  ];
  const holdings = buildHoldingsByDate_(trades, keys);
  assertAlmostEqual_(holdings["2026-10-08"]["600519"].quantity, 100, "day 1 holding");
  assertAlmostEqual_(holdings["2026-10-09"]["600519"].quantity, 50, "day 2 holding");

  const prices = { "2026-10-08": 11, "2026-10-09": 12 };
  const tradeFlows = { "2026-10-08": { buy: 1005, proceeds: 0 }, "2026-10-09": { buy: 0, proceeds: 595 } };

  const rows = computeNavRows_({
    dateKeys: keys,
    prevDateKey: "",
    todayKey: "2026-10-12",
    liveEquity: 0,
    holdingsByDate: holdings,
    priceLookup: (symbol, key) => ({ price: prices[key], settled: true }),
    fxLookup: () => 1,
    walked: walkCashBalances_(tradeFlows),
    signatures: { "2026-10-08": "1:aaaaaaaa" }
  });

  assertAlmostEqual_(rows[0].equity, 1100, "day 1 equity");
  assertAlmostEqual_(rows[0].cashBalance, 0, "day 1 cash");
  assertAlmostEqual_(rows[0].externalFlow, 1005, "day 1 external flow");
  assertAlmostEqual_(rows[1].equity, 600, "day 2 equity");
  assertAlmostEqual_(rows[1].totalAsset, 1195, "day 2 total asset");
  assertEqual_(rows[0].sig, "1:aaaaaaaa", "signature stored");
  assertEqual_(rows[1].quality, "", "settled close is final");
}

/**
 * Auto funding through NAV: a buy's new money is external flow, so it does
 * not show up as trading PnL. Starting mid-history only
 * counts flows after prevDateKey.
 */
function testComputeNavRowsAutoFunding_() {
  const keys = ["2026-10-08", "2026-10-09", "2026-10-12"];
  const trades = [
    TEST_TRADE_HEADER_,
    testTradeRow_("2026-10-08", "600519", "BUY", 100, 10),
    testTradeRow_("2026-10-09", "600519", "SELL", 50, 12),
    testTradeRow_("2026-10-12", "600519", "BUY", 100, 12)
  ];
  const holdings = buildHoldingsByDate_(trades, keys);
  const prices = { "2026-10-08": 11, "2026-10-09": 12, "2026-10-12": 12 };
  const walked = walkCashBalances_({
    "2026-10-08": { buy: 1005, proceeds: 0 },
    "2026-10-09": { buy: 0, proceeds: 595 },
    "2026-10-12": { buy: 1205, proceeds: 0 }
  });
  const input = {
    dateKeys: keys,
    prevDateKey: "",
    todayKey: "2026-10-13",
    liveEquity: 0,
    holdingsByDate: holdings,
    priceLookup: (symbol, key) => ({ price: prices[key], settled: true }),
    fxLookup: () => 1,
    walked: walked,
    signatures: {}
  };

  const rows = computeNavRows_(input);
  assertAlmostEqual_(rows[0].externalFlow, 1005, "buy auto-funded");
  assertAlmostEqual_(rows[0].cashBalance, 0, "no idle cash after auto funding");
  assertAlmostEqual_(rows[1].externalFlow, 0, "sale needs no funding");
  assertAlmostEqual_(rows[1].cashBalance, 595, "sale proceeds kept");
  assertAlmostEqual_(rows[2].externalFlow, 610, "only the shortfall is funded");
  assertAlmostEqual_(rows[2].cashBalance, 0, "day 3 cash");
  // TradingPnL = total - prevTotal - externalFlow (NAV_Daily column G).
  assertAlmostEqual_(rows[0].totalAsset - 0 - rows[0].externalFlow, 95, "day 1 PnL: +100 price, -5 fee");

  const tail = computeNavRows_(Object.assign({}, input, { dateKeys: keys.slice(2), prevDateKey: keys[1] }));
  assertAlmostEqual_(tail[0].externalFlow, 610, "mid-history start skips earlier flows");
}
