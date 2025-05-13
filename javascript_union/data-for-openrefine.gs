// この関数だけ実行する
function forOpenrefine() {
  Logger.log("=== 処理 開始 ===");

  prepareORDataSheet();
  insertColumnsAndPopulateValues();

  Logger.log("=== 処理 完了 ===");
}


// 以下、補助関数

function prepareORDataSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sourceSheet = ss.getSheetByName("統合データ");
  if (!sourceSheet) {
    Logger.log("統合データシートが見つかりません");
    return;
  }

  // OR用データシートの初期化または新規作成
  let orSheet = ss.getSheetByName("OR用データ");
  if (orSheet) {
    orSheet.clear();  // 中身だけ消す
  } else {
    orSheet = ss.insertSheet("OR用データ");
  }

  // 統合データのコピー
  const data = sourceSheet.getDataRange().getValues();

  // === 特定列の文字列処理 ===
    const columnsToRemoveHyphen = [1, 5, 7, 11, 20, 23, 31, 33]; // A, E, G, K, T, W, AE, AG
  const columnToRemoveUnderscore = 17; // Q

  for (let i = 1; i < data.length; i++) { // 1行目（ヘッダー）を除く
    for (const col of columnsToRemoveHyphen) {
      if (typeof data[i][col - 1] === 'string') {
        data[i][col - 1] = data[i][col - 1].replace(/-/g, '');
      }
    }
    if (typeof data[i][columnToRemoveUnderscore - 1] === 'string') {
      data[i][columnToRemoveUnderscore - 1] = data[i][columnToRemoveUnderscore - 1].replace(/_/g, '');
    }
  }


  orSheet.getRange(1, 1, data.length, data[0].length).setValues(data);

  // === D列（4列目）の書式を「整数」表示に変更 ===
  const lastRow = orSheet.getLastRow();
  const dRange = orSheet.getRange(2, 4, lastRow - 1); // 2行目以降
  dRange.setNumberFormat("0"); // 「整数」フォーマット

  // ヘッダーを上書き
  const newHeaders = [
    "Agenda", "AssemblyID", "startDate", "トピック順", "mentionedAsSubsequent", "agenda_type",
    "sc_recogito", "recogito", "screc_url", "付与チェック", "sc_gallica", "gallica",
    "ga_pageStart", "orig_pageStart", "orig_lineStart", "or_volume", "source", "sourceID",
    "ga_url", "sc_iiif", "iiif", "iiif_url", "sc_orig", "original", "入力済確認", "メモ", "開始ページ", "報告形式",
    "作成日", "初出", "初出のトピックID", "報告者", "Topic", "topic_description", "備考",
    "初出トピックID", "日付", "報告者（表示名）", "役職・職階", "役職・職階（正規化）", "分野",
    "前任者（表示名）", "報告日", "送信日", "関連トピックID"
  ];
  orSheet.getRange(1, 1, 1, newHeaders.length).setValues([newHeaders]);

  Logger.log("✅ OR用データシートが作成され、D列の表示形式を整数に設定しました。また指定された列の記号も削除済みです。");
}

function insertColumnsAndPopulateValues() {
  const sheetName = "OR用データ";
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  if (!sheet) {
    Logger.log(`シート「${sheetName}」が見つかりません`);
    return;
  }

  // 1. A列の左に2列挿入
  sheet.insertColumnsBefore(1, 2);

  // 新しいA列とB列のヘッダーを設定
  sheet.getRange("A1").setValue("ididid");
  sheet.getRange("B1").setValue("line");

  // データの行数を取得
  const lastRow = sheet.getLastRow();

  // D列（元のB列=2列目だったものが今D列=4列目）を取得
  const dValues = sheet.getRange(2, 4, lastRow - 1).getValues();  // 2行目から
  const idididValues = [];
  const lineValues = [];

  for (let i = 0; i < dValues.length; i++) {
    const lineNum = (i + 2).toString(); // 2行目から開始
    lineValues.push([lineNum]);
    idididValues.push([dValues[i][0] + lineNum]);
  }

  // A列にididid（D列の値+B列の値）、B列にline（行番号）を設定
  sheet.getRange(2, 1, idididValues.length).setValues(idididValues);
  sheet.getRange(2, 2, lineValues.length).setValues(lineValues);
}
