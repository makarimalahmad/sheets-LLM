// Menu tombol otomatis di Google Sheets
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu("Sync AI")
    .addItem("Gabungkan Semua Sheet", "sendAllSheetsToBackend")
    .addToUi();
}

// Fungsi membaca semua sheet (tab 1 sampai 5 dst) dan kirim ke backend
function sendAllSheetsToBackend() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = spreadsheet.getSheets();
  var sheetsPayload = [];

  for (var i = 0; i < sheets.length; i++) {
    var sheet = sheets[i];
    var data = sheet.getDataRange().getValues();

    if (data.length >= 2) {
      var headers = data[0].map(function (h) {
        return String(h).trim();
      });
      var rows = data.slice(1);

      sheetsPayload.push({
        sheet_name: sheet.getName(),
        headers: headers,
        rows: rows,
      });
    }
  }

  if (sheetsPayload.length === 0) {
    Logger.log("Tidak ada sheet yang memiliki data.");
    return;
  }

  var payload = {
    sheets: sheetsPayload,
  };

  // Masukkan URL backend VPS Anda
  var backendUrl = "http://172.232.224.220/api/consolidate";

  var options = {
    method: "post",
    contentType: "application/json",
    payload: JSON.stringify(payload),
    muteHttpExceptions: true,
  };

  try {
    var response = UrlFetchApp.fetch(backendUrl, options);
    Logger.log("Berhasil: " + response.getContentText());
    SpreadsheetApp.getUi().alert(
      "Sukses! Data dari " +
        sheetsPayload.length +
        " sheet berhasil digabungkan oleh AI.",
    );
  } catch (error) {
    Logger.log("Error: " + error.toString());
    SpreadsheetApp.getUi().alert("Gagal kirim: " + error.toString());
  }
}
