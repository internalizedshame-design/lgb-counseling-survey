function doPost(e) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    var data = JSON.parse(e.postData.contents);

    var headers = [
      "timestamp", "assigned_scenario", "gender", "age", "orientation", "counseling_exp",
      "check_content", "manip_check",
      "out_1", "out_2", "out_3", "out_4", "out_5", "out_6", "out_7", "out_8", "out_9", "out_10", "out_11", "out_12",
      "disc_1", "disc_2", "disc_3", "disc_4", "disc_5", "disc_6", "disc_7", "disc_8",
      "rej_1", "rej_2", "rej_3", "rej_4", "rej_5", "rej_6", "rej_7", "rej_8", "rej_9", "rej_10", "rej_11", "rej_12", "rej_13", "rej_14", "rej_15"
    ];

    if (sheet.getLastRow() === 0) {
      sheet.appendRow(headers);
    }

    var row = headers.map(function(h) {
      return data[h] !== undefined ? data[h] : "";
    });
    sheet.appendRow(row);

    return ContentService
      .createTextOutput(JSON.stringify({ status: "success" }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ status: "error", message: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}
