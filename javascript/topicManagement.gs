function handleTopicManagement(e) {
  const range = e.range;
  const sheet = range.getSheet();
  const column = range.getColumn();
  const row = range.getRow();

  // B列が編集された場合のみ処理を実行
  if (column === 2) {
    copyValue(sheet, 'J2', `J${row}`); // J2の値をJ列の編集行にコピー
    copyValue(sheet, 'K2', `K${row}`); // K2の値をK列の編集行にコピー
    copyValue(sheet, 'L2', `L${row}`); // L2の値をL列の編集行にコピー
  }

  // B列とC列の両方に値が入力されている場合にA列にトピックIDを設定
  const dateValue = sheet.getRange(row, 2).getValue();
  const numericValue = sheet.getRange(row, 3).getValue();

  /*
  if (dateValue && numericValue) {
    const topicID = generateTopicID(dateValue, numericValue);
    sheet.getRange(row, 1).setValue(topicID);
  }
  */

  if (row >= 3) {
    // B列とC列の両方に値が入力されている場合にA列にトピックIDを設定
    const dateValue = sheet.getRange(row, 2).getValue();
    const numericValue = sheet.getRange(row, 3).getValue();

    if (dateValue && numericValue) {
      const topicID = generateTopicID(dateValue, numericValue);
      sheet.getRange(row, 1).setValue(topicID);
    }
  }

  // 開始頁（G列）が編集された場合にアノテーションURI（E列）の値を設定
  if (column === 7) {
    const baseurl = sheet.getRange(2, 5).getValue(); // E2の値
    const gValue = sheet.getRange(row, 7).getValue(); // G列の値
    const eValue = baseurl + gValue.toString() + "/edit";
    sheet.getRange(row, 5).setValue(eValue); // E列に値を設定
  }

  // 入力済確認(M列)にチェックが入ったら共通メタデータを他のシートに送る
  // 入力済確認(M列)にチェックが入ったら共通メタデータを他のシートに送る
  if (column === 13) {
    const status = sheet.getRange(row, 4).getValue(); // D列の値
    if (sheet.getRange(row, 13).getValue() === true) { // M列のチェック
      switch (status) {
        case "出席":
          sendData(sheet, row, ["A", "B", "G", "I"], "出席", ["A", "B", "C", "D"], "A");
          break;
        case "会員の任命":
        case "制度改定":
          sendData(sheet, row, ["A", "B", "G", "I", "B"], "会員の任命", ["A", "B", "C", "D", "E"], "A");
          break;
        case "投票・推薦":
          sendData(sheet, row, ["A", "B", "G", "I"], "投票・推薦", ["A", "B", "C", "D"], "A");
          break;
        case "研究報告":
        case "実験・提示":
        case "査読依頼":
        case "査読結果":
          sendData(sheet, row, ["A", "B", "G", "I", "D"], "研究報告・実験・査読", ["A", "B", "C", "D", "E"], "A");
          break;
      }
    }
  }

}


// トピックIDを生成ルール
function generateTopicID(dateValue, numericValue) {
  // 日付をYYYYMMDD形式に変換
  const formattedDate = Utilities.formatDate(new Date(dateValue), Session.getScriptTimeZone(), 'yyyyMMdd');

  // 数値を4桁の文字列に変換
  const formattedNumber = (numericValue * 10).toFixed(0).padStart(4, '0');

  return `${formattedDate}-${formattedNumber}`;
}
