function preprocessing() {
  const sheetName1 = "トピックID全体管理";
  const sheetName2 = "人物管理"
  const sheetName3 = "研究報告・実験・査読";
  const topicTypeSheetName = "トピック種別";

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet1 = ss.getSheetByName(sheetName1);
  const sheet2 = ss.getSheetByName(sheetName2);
  const sheet3 = ss.getSheetByName(sheetName3);
  const topicTypeSheet = ss.getSheetByName(topicTypeSheetName);

  if (!sheet1 || !topicTypeSheet || !sheet3) {
    Logger.log("シートが見つかりません。");
    return;
  }

  // トピック種別シートのA列の範囲を取得
  const topicTypes = getColumnValues(topicTypeSheet, 'A');

  // sheetName1のD列3行目以降にプルダウンリストを設定
  setDropdownList(sheet1, 'D3:D', topicTypes);

  // sheetName3のF列3行目以降に「本人報告」と「代理人報告」のプルダウンリストを設定
  const options = ["本人報告", "代理人報告"];
  setDropdownList(sheet3, 'F3:F', options);

  // デフォルトでチェックの入っていないチェックボックスを作成
  createCheckboxes(sheet1, 'F3:F');
  createCheckboxes(sheet1, 'M3:M');
  createCheckboxes(sheet2, 'O3:O');
  createCheckboxes(sheet3, 'H3:H');
  createCheckboxes(sheet3, 'L3:L');

}

function getColumnValues(sheet, column) {
  return sheet.getRange(column + ':' + column).getValues().filter(row => row[0] !== "").flat();
}

// チェックボックスを作る関数
function createCheckboxes(sheet, range) {
  const rangeObj = sheet.getRange(range);
  rangeObj.insertCheckboxes();
  const values = rangeObj.getValues();
  for (let i = 0; i < values.length; i++) {
    values[i][0] = false;  // デフォルトでチェックなし
  }
  rangeObj.setValues(values);
}