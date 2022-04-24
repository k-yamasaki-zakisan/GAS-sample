// 参考資料
// https://zenn.dev/d0ne1s/articles/ecd743f82cb59e
function gaMain() {
  const googleAnalyticsWriteSheet =
    SpreadsheetApp.openById("**書き込み用スプシid**").getSheetByName(
      "GoogleAnalytics"
    );

  const gaData = getGaData();

  writeGaSheet(googleAnalyticsWriteSheet, gaData);
}

function getGaData() {
  //グーグルアナリティクスから取得する指標データを設定する
  const metrics = "ga:users";
  //グーグルアナリティクスのディメンションでページタイトルを設定する
  const dimensions = "ga:sourceMedium";
  //表示順はページビュー順にソートする
  const sortType = "-ga:users";
  //Google Analytics APIリクエストして、グーグルアナリティクスのデータを取得する
  const result = Analytics.Data.Ga.get(
    "ga:**環境に合わせてidを記述**", // id
    "yesterday", // start day 例7daysAgo
    "yesterday", // end day
    metrics,
    {
      dimensions: dimensions,
      sort: sortType,
    }
  ).getRows();

  return result.map((data) => {
    return {
      mediaType: data[0],
      userCount: data[1],
    };
  });
}

function writeGaSheet(sheet, gaData) {
  const column = sheet.getLastColumn() + 1; //書き込み列を取得

  // headerに日付を記入
  const date = new Date();
  //昨日の日付を取得
  date.setDate(date.getDate() - 1);
  const dateFormat = Utilities.formatDate(date, "JST", "YYYY/MM/dd");
  sheet.getRange(1, column).setValues([[dateFormat]]);

  let youtubeCount = 0;
  let totalCount = 0;
  let organicCount = 0;
  for (let i = 0; i < gaData.length; i++) {
    const userCnt = gaData[i].userCount;
    const mediaType = gaData[i].mediaType;

    totalCount = totalCount + Number(userCnt);
    if (mediaType == "google / cpc") {
      sheet.getRange(3, column).setValues([[Number(userCnt)]]);
    } else if (mediaType == "yahoo / cpc") {
      sheet.getRange(4, column).setValues([[Number(userCnt)]]);
    } else if (mediaType == "yahoo / display") {
      sheet.getRange(5, column).setValues([[Number(userCnt)]]);
    } else if (mediaType == "facebook / display") {
      sheet.getRange(6, column).setValues([[Number(userCnt)]]);
    } else if (mediaType.match(/youtube/)) {
      youtubeCount = youtubeCount + Number(userCnt);
    } else {
      organicCount = organicCount + Number(userCnt);
    }
  }
  sheet.getRange(2, column).setValues([[organicCount]]);
  sheet.getRange(7, column).setValues([[youtubeCount]]);
  sheet.getRange(8, column).setValues([[totalCount]]);
}
