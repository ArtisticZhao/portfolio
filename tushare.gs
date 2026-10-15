/**
 * Tushare Pro HTTP client for historical price data.
 * Token lives in Script Properties (Project Settings > Script Properties),
 * never in this repository.
 */
function getTushareToken_() {
  const token = PropertiesService.getScriptProperties().getProperty("TUSHARE_TOKEN");
  if (!token) {
    throw new Error(
      'Missing TUSHARE_TOKEN. Add it in Apps Script "Project Settings > Script Properties".'
    );
  }
  return token;
}

function callTushare_(apiName, params, fields) {
  const payload = {
    api_name: apiName,
    token: getTushareToken_(),
    params: params,
    fields: fields
  };

  const response = UrlFetchApp.fetch("https://api.tushare.pro", {
    method: "post",
    contentType: "application/json",
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });

  const json = JSON.parse(response.getContentText());
  if (json.code !== 0) {
    throw new Error(`Tushare ${apiName} error ${json.code}: ${json.msg}`);
  }
  return json.data;
}

/**
 * @return {Object} dateKey("yyyy-MM-dd") -> close price
 */
function parseTushareDateCloseRows_(data) {
  const dateIdx = data.fields.indexOf("trade_date");
  const closeIdx = data.fields.indexOf("close");
  const result = {};
  data.items.forEach(row => {
    const raw = String(row[dateIdx]);
    const key = `${raw.slice(0, 4)}-${raw.slice(4, 6)}-${raw.slice(6, 8)}`;
    result[key] = Number(row[closeIdx]);
  });
  return result;
}

function fetchTushareDaily_(tsCode, startDateKey, endDateKey) {
  const data = callTushare_(
    "daily",
    { ts_code: tsCode, start_date: startDateKey.replace(/-/g, ""), end_date: endDateKey.replace(/-/g, "") },
    "trade_date,close"
  );
  return parseTushareDateCloseRows_(data);
}

/**
 * "002317" -> "002317.SZ" / "600519" -> "600519.SH". Mirrors convertToSinaSymbol_ (price.gs).
 */
function symbolToTsCode_(symbol) {
  const s = String(symbol).trim();
  if (/^[65]\d{5}$/.test(s)) return s + ".SH";
  if (/^[013]\d{5}$/.test(s)) return s + ".SZ";
  return null;
}

