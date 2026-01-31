
// *選択しているセルのB列をC列に翻訳する関数
function TranslateEngToJpn(){
    const sheet = SpreadsheetApp.getActiveSpreadsheet(); // アクティブなシートを取得
    const english = sheet.getActiveRange().getValue();
    console.log(english);
    const japanese = LanguageApp.translate(english, "en", "ja");
    sheet.getActiveRange().offset(0, 1).setValue(japanese);
}
// *選択しているシートのB列をC列に翻訳する関数
function TranslateEngToJpnAll(){
    const sheet = SpreadsheetApp.getActiveSpreadsheet();    // アクティブなシートを取得
    const english = sheet.getActiveRange().getValue();
    console.log(english);
    const japanese = LanguageApp.translate(english, "en", "ja");
    sheet.getActiveRange().offset(0, 1).setValue(japanese);
}
