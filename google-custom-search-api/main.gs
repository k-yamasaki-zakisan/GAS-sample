const APIKEY = ""; //Google Custom Search APIのAPIキー
const CSEID = ""; //カスタム検索エンジン（CSE）のID

// 参考資料
// https://qiita.com/zak_y/items/42ca0f1ea14f7046108c
function googleSearchMain() {
  // 書き込み用シート取得
  const sheet =
    SpreadsheetApp.openById("**書き込み用スプシのid**").getSheetByName(
      "検索順位シート"
    );
  const column = sheet.getLastColumn() + 1;

  const date = new Date();
  const dateFormat = Utilities.formatDate(date, "JST", "YYYY/MM/dd");
  sheet.getRange(1, column).setValues([[dateFormat]]);

  // 検索ワードのソースを取得
  // const words = [["歌舞伎症候群"],["プラダ―・ウィリ症候群"], ["18トリソミー症候群"], [["13トリソミー症候群"]]]
  const words = SpreadsheetApp.openById("**ワード検索ソーススプシのid**")
    .getSheetByName("疾患解説")
    .getRange(2, 9, sheet.getLastRow() - 1)
    .getValues();

  // 検索apiの無料枠99を上限とする
  for (let v = 0; v < Math.min(words.length, 99); v++) {
    try {
      const word = words[v][0];

      const resultResponse = searchApi(word);
      const ranking = rankChecker(resultResponse);

      const row = v + 2; // シートの列位置調整
      writeSheet(ranking, row, column, sheet, word);
    } catch (e) {
      Logger.log(e);
      continue;
    }
  }
}

function searchApi(word) {
  const apiUrl = `https://www.googleapis.com/customsearch/v1?key=${APIKEY}&cx=${CSEID}&q=${word}`;
  const apiOptions = {
    method: "get",
  };
  const responseApi = UrlFetchApp.fetch(apiUrl, apiOptions);
  const responseJson = JSON.parse(responseApi.getContentText());
  return responseJson;
}

function rankChecker(response) {
  let geneticsRank = null;
  let medicalnoteRank = null;
  for (let v = 0; v < response["items"].length; v++) {
    const link = response["items"][v]["link"];
    if (link.match("genetics.qlife.jp")) {
      geneticsRank = v + 1;
    } else if (link.match("medicalnote.jp")) {
      medicalnoteRank = v + 1;
    }
  }
  return { genetics: geneticsRank, medicalnote: medicalnoteRank };
}

function writeSheet(ranking, row, column, sheet, word) {
  // 検索ワードを更新
  sheet.getRange(row, 1).setValues([[word]]);

  const geneticsRank = ranking.genetics;
  const medicalnoteRank = ranking.medicalnote;
  const content = `${geneticsRank ? geneticsRank + "位" : "11位以上"} / ${
    medicalnoteRank ? medicalnoteRank + "位" : "11位以上"
  }`;
  sheet.getRange(row, column).setValues([[content]]);
  // メディカルノートのランキングがいでプラより上位の場、青色を付ける
  if (
    (geneticsRank &&
      medicalnoteRank &&
      Number(geneticsRank) > Number(medicalnoteRank)) ||
    (!geneticsRank && medicalnoteRank)
  ) {
    sheet.getRange(row, column).setFontColor("blue");
    // いでプラが勝っている場合、赤色を付ける
  } else if (
    (geneticsRank &&
      medicalnoteRank &&
      Number(geneticsRank) < Number(medicalnoteRank)) ||
    (geneticsRank && !medicalnoteRank)
  ) {
    sheet.getRange(row, column).setFontColor("red");
  } else {
    sheet.getRange(row, column).setFontColor("black");
  }
}
