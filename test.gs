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
