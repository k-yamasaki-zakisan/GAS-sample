function slideMain() {
  const orgSlideId = "1AbvFT0Ru3WjWe__p_XuSLslc8eZ8ApPqRzTtPQ3W1mE";
  const slideUrl = `https://docs.google.com/presentation/d/${orgSlideId}/`;
  const presentation = SlidesApp.openByUrl(slideUrl);
  const slides = presentation.getSlides();
  if (slides.length == 0) throw new Error("スライドがありません");
  const slide = slides[0];

  // 印刷用の新規スライドファイルの作成
  const sourceFile = DriveApp.getFileById(orgSlideId);
  const date = new Date();
  const dateFormat = Utilities.formatDate(date, "JST", "YYYY/MM/dd HH:mm");
  const newFile = sourceFile.makeCopy(dateFormat + " 差し込み印刷ファイル");
  const newSlideUrl = newFile.getUrl();
  const newPresentation = SlidesApp.openByUrl(newSlideUrl);

  // 該当箇所を動的に変更するためにスプシから行を取得
  const listSheet = SpreadsheetApp.openById(
    "1CYLrLtkc4fEQeCY3W4r092CBcTyQiruclazzHnL7WWE"
  ).getSheetByName("処方箋原本郵送リスト");
  const lastRow = listSheet.getLastRow();
  // TODO：iはリストを見て微調整
  for (let i = 3; i <= lastRow; i++) {
    const postCode = listSheet.getRange(i, 3).getValue();
    const address = listSheet.getRange(i, 4).getValue();
    const name = listSheet.getRange(i, 5).getValue();
    const honorificTitle = listSheet.getRange(i, 6).getValue();
    if (!(postCode && address && name && honorificTitle)) {
      SpreadsheetApp.getUi().alert(`${i}行目は項目不足でスキップします`);
      continue;
    }

    // テンプレートslideを印刷用スライドの最後に入れて文字を変更する
    const newSlide = newPresentation.appendSlide(slide);
    newSlide.replaceAllText("{郵便番号}", postCode);
    newSlide.replaceAllText("{住所}", address);
    newSlide.replaceAllText("{氏名}", name);
    newSlide.replaceAllText("{敬称}", honorificTitle);
  }

  // 作業完了をユーザーにお知らせする
  const htmlOutput = HtmlService.createHtmlOutput(
    "<p>印刷ファイルの作成が完了しました</p>"
  )
    .setWidth(250)
    .setHeight(100);
  SpreadsheetApp.getUi().showModelessDialog(htmlOutput, "結果");
}
