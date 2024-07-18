function handleResearchManagement(e) {
  const range = e.range;
  const sheet = range.getSheet();
  const column = range.getColumn();
  const row = range.getRow();
  const sheetName3 = "研究報告・実験・査読";
  const sheetName4 = "初出研究報告管理"

  // アクティブなスプレッドシートを取得
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet3 = ss.getSheetByName(sheetName3);

  // 現在のシートがsheetName3でない場合は処理を終了
  if (!sheet3 || sheet.getName() !== sheetName3) {
    return;
  }

  // F列で"本人報告"が選択されたらG列をdisableCell、そうでなければenableCell
  if (column === 6) { // F列
    const value = range.getValue();
    if (value === "本人報告") {
      disableCell(sheet3, `G${row}`);
    } else {
      enableCell(sheet3, `G${row}`);
    }
  }

  // H列でチェックボックスにチェックが入ったらI列をdisableCell、そうでなければenableCell
  if (column === 8) { // H列
    const value = range.getValue();
    if (value === true) {
      disableCell(sheet3, `I${row}`);
    } else {
      enableCell(sheet3, `I${row}`);
    }
  }

  // L列とH列にチェックが入った場合の処理
  if ((column === 12 || column === 8) && sheet.getRange(`L${row}`).getValue() === true && sheet.getRange(`H${row}`).getValue() === true) {
    const targetSheet = ss.getSheetByName(sheetName4);
    if (targetSheet) {
      sendData(sheet3, row, ["A", "B", "E", "AA", "J"], sheetName4, ["B", "C", "D", "E", "F"], "A"); // AA列は空欄

      // targetSheetのA列に行番号を参考にしたIDを入力する
      const lastRow = targetSheet.getLastRow();
      const newID = generateNewID(lastRow);
      targetSheet.getRange(`A${lastRow}`).setValue(newID);
    }
  }
}

function generateNewID(row) {
  const year = 1716;
  const idNumber = row * 10;
  return `t${year}-${String(idNumber).padStart(7, '0')}`;
}
