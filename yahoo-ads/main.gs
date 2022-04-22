// Yahoo!広告の認証
const URL_TOKEN = "https://biz-oauth.yahoo.co.jp/oauth/v1/token";
const YSS_URL_API = "https://ads-search.yahooapis.jp/api/v7";
const YDN_URL_API = "https://ads-display.yahooapis.jp/api/v7";
const YAHOO_CLIENT_ID = "";
const YAHOO_CLIENT_SECRET = "";
const YAHOO_REFRESH_TOKEN = "";
const YSS_ACCOUNT_ID = "";
const YDN_ACCOUNT_ID = "";

function yahooAdsMain() {
  const accessToken = getAccessToken();
  yahooAdsYssMain(YSS_URL_API, accessToken, YSS_ACCOUNT_ID);
  yahooAdsYdnMain(YDN_URL_API, accessToken, YDN_ACCOUNT_ID);
}

//アクセストークン再取得
function getAccessToken() {
  const payload = {
    grant_type: "refresh_token",
    client_id: YAHOO_CLIENT_ID,
    client_secret: YAHOO_CLIENT_SECRET,
    refresh_token: YAHOO_REFRESH_TOKEN,
  };
  const params = {
    payload: payload,
  };

  const response = UrlFetchApp.fetch(URL_TOKEN, params);
  const responseBody = JSON.parse(response.getContentText());

  return responseBody["access_token"];
}

function yahooAdsYssMain(url, accessToken, accountId) {
  const yahooAdsYssWriteSheet =
    SpreadsheetApp.openById("**書き込み用のスプシid**").getSheetByName(
      "YahooAdsYss"
    );

  const jobId = getYssReportJobId(url, accessToken, accountId);
  Utilities.sleep(10000);
  const yssData = downloadYssReport(url, accessToken, accountId, jobId);

  writeYahooAdsYssWriteSheet(yahooAdsYssWriteSheet, yssData);
}

//YSSの日ごとのレポートを作成してJOBIDを取得
function getYssReportJobId(url, accessToken, accountId) {
  const headers = {
    Authorization: "Bearer " + accessToken,
  };
  const payload = {
    accountId: accountId,
    operand: [
      {
        accountId: accountId,
        fields: [
          // "DAY",
          "CAMPAIGN_NAME", // キャンペーン名
          "CAMPAIGN_ID", // キャンペーンID
          "CONVERSIONS", // コンバージョン数
          "COST", // 費用
          "IMPS", // インプレッション数
          "CLICKS", // クリック数
          // "CLICK_RATE",          // クリック率
          // "AVG_CPC",             // 平均クリック単価
          // "CONV_RATE",           // コンバージョン率
          // "CONV_VALUE",          // コンバージョン値
        ],
        reportCompressType: "NONE",
        reportDateRangeType: "YESTERDAY",
        reportDownloadEncode: "UTF8",
        reportDownloadFormat: "CSV",
        reportIncludeDeleted: "TRUE",
        reportLanguage: "JA",
        reportName: "yssCampaignDayReport",
        reportType: "CAMPAIGN",
        sortFields: [
          {
            field: "CAMPAIGN_ID",
            reportSortType: "ASC",
          },
        ],
      },
    ],
  };

  const params = {
    method: "post",
    // muteHttpExceptions : true,
    contentType: "application/json",
    headers: headers,
    payload: JSON.stringify(payload),
  };

  const response = UrlFetchApp.fetch(
    url + "/ReportDefinitionService/add",
    params
  );
  const responseBody = JSON.parse(response.getContentText());
  const reportJobId =
    responseBody["rval"]["values"][0]["reportDefinition"]["reportJobId"];
  return Number(reportJobId);
}

//レポートをダウンロード
function downloadYssReport(url, accessToken, accountId, jobId) {
  const headers = {
    Authorization: "Bearer " + accessToken,
  };
  const payload = {
    accountId: accountId,
    reportJobId: jobId,
  };
  const params = {
    method: "post",
    contentType: "application/json",
    headers: headers,
    payload: JSON.stringify(payload),
  };

  const httpResponse = UrlFetchApp.fetch(
    url + "/ReportDefinitionService/download",
    params
  );
  const status = httpResponse.getResponseCode();
  if (status !== 200) {
    throw "HttpRequestError";
  }
  const response = httpResponse.getContentText();
  const lines = response.split("\n");
  if (lines.length === 3) {
    return [];
  }

  return lines.slice(1, -1).map((line) => {
    const data = line.split(",");
    return {
      campaignName: String(data[0]),
      conversions: Number(data[2]),
      cost: Math.round(data[3]),
      imps: String(data[4]),
      clicks: Number(data[5]),
    };
  });
}

function writeYahooAdsYssWriteSheet(sheet, yssData) {
  const column = sheet.getLastColumn() + 1; //書き込み列を取得

  // headerに日付を記入
  const date = new Date();
  const dateFormat = Utilities.formatDate(date, "JST", "YYYY/MM/dd");
  sheet.getRange(1, column).setValues([[dateFormat]]);

  for (let i = 0; i < yssData.length; i++) {
    if (!yssData[i].campaignName) continue;
    sheet.getRange(i + 2, 1).setValues([[yssData[i].campaignName]]);
    sheet.getRange(i + 2, column).setValues([[yssData[i].cost]]);
  }
}

function yahooAdsYdnMain(url, accessToken, accountId) {
  const yahooAdsYdnWriteSheet =
    SpreadsheetApp.openById("**書き込み用スプシid**").getSheetByName(
      "YahooAdsYdn"
    );

  const reportJobId = getYdnReportJobId(url, accessToken, accountId);
  Utilities.sleep(10000);
  const ydnData = downloadYdnReport(url, accessToken, accountId, reportJobId);

  writeYahooAdsYdnWriteSheet(yahooAdsYdnWriteSheet, ydnData);
}

//YDNのレポートJOBIDを取得
function getYdnReportJobId(url, accessToken, accountId) {
  const headers = {
    Authorization: "Bearer " + accessToken,
  };
  const payload = {
    accountId: accountId,
    operand: [
      {
        accountId: accountId,
        fields: [
          // "DAY",
          "CAMPAIGN_NAME", // キャンペーン名
          "CONVERSIONS", // コンバージョン数
          "COST", // 費用
          "IMPS", // インプレッション数
          "CAMPAIGN_ID", // キャンペーンID
          // "CLICKS",              // クリック数
          // "CLICK_RATE",          // クリック率
          // "AVG_CPC",             // 平均クリック単価
          // "CONV_RATE",           // コンバージョン率
          // "CONV_VALUE",          // コンバージョン値
        ],
        reportCompressType: "NONE",
        reportDateRangeType: "YESTERDAY",
        reportDownloadEncode: "UTF8",
        reportDownloadFormat: "CSV",
        reportLanguage: "JA",
        reportName: "ydnCampaignDayReport",
        // "reportType": "CAMPAIGN",
        sortFields: [
          {
            field: "CAMPAIGN_ID",
            reportSortType: "ASC",
          },
        ],
      },
    ],
  };
  const params = {
    method: "post",
    muteHttpExceptions: true,
    contentType: "application/json",
    headers: headers,
    payload: JSON.stringify(payload),
  };

  const response = UrlFetchApp.fetch(
    url + "/ReportDefinitionService/add",
    params
  );
  const responseBody = JSON.parse(response);
  const reportJobId =
    responseBody["rval"]["values"][0]["reportDefinition"]["reportJobId"];
  return Number(reportJobId);
}

//レポートをダウンロード
function downloadYdnReport(url, accessToken, accountId, jobId) {
  const headers = {
    Authorization: "Bearer " + accessToken,
  };
  const payload = {
    accountId: accountId,
    reportJobId: jobId,
  };
  const params = {
    method: "post",
    contentType: "application/json",
    headers: headers,
    payload: JSON.stringify(payload),
  };

  const httpResponse = UrlFetchApp.fetch(
    url + "/ReportDefinitionService/download",
    params
  );
  const status = httpResponse.getResponseCode();
  if (status !== 200) {
    throw "HttpRequestError";
  }
  const response = httpResponse.getContentText();
  const lines = response.split("\n");
  if (lines.length === 3) {
    return [];
  }

  return lines.slice(1, -1).map((line) => {
    const data = line.split(",");
    return {
      campaignName: String(data[0]),
      conversions: Number(data[1]),
      cost: Math.round(data[2]),
      imps: String(data[3]),
    };
  });
}

function writeYahooAdsYdnWriteSheet(sheet, ydnData) {
  const column = sheet.getLastColumn() + 1; //書き込み列を取得

  // headerに日付を記入
  const date = new Date();
  const dateFormat = Utilities.formatDate(date, "JST", "YYYY/MM/dd");
  sheet.getRange(1, column).setValues([[dateFormat]]);

  for (let i = 0; i < ydnData.length; i++) {
    if (ydnData[i].campaignName == "Total") continue;
    sheet.getRange(i + 2, 1).setValues([[ydnData[i].campaignName]]);
    sheet.getRange(i + 2, column).setValues([[ydnData[i].cost]]);
  }
}
