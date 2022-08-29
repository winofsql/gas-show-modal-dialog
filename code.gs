// ************************************
// メニューの追加
// ************************************
function onOpen(e) {

  var ui = SpreadsheetApp.getUi();
  ui.createMenu('HTMLページ')
	.addItem('✅ リンク表示　　　　', 'createLinkDialog')
	.addItem('ソース', 'createSrcDialog')
	.addToUi();
		
}

// ************************************
// HTML をダイアログとして表示
// ************************************
function createLinkDialog() {

  var html = HtmlService.createHtmlOutputFromFile('index');
  html.setWidth(1200);
  html.setHeight(1200);
  SpreadsheetApp.getUi().showModalDialog(html, "リンク");

}

// ************************************
// セルデータをソースコードとして
// ダイアログに表示
// ************************************
function createSrcDialog() {

  var html = HtmlService.createHtmlOutputFromFile('src');
  html.setWidth(1200);
  html.setHeight(1200);

  var currentCell = SpreadsheetApp.getCurrentCell();
  var targetString = currentCell.getDisplayValue();
  
  targetString = targetString.replace(/&/g,'&amp;');
  targetString = targetString.replace(/</g,'&lt;');
  targetString = targetString.replace(/>/g,'&gt;');
  targetString = targetString.replace(/    /g,'\t');

  var start = '<textarea readonly wrap="off">';
  var end = '</textarea>';

  html.append(start + targetString + end);
  SpreadsheetApp.getUi().showModalDialog(html, "ソースの表示");

}
