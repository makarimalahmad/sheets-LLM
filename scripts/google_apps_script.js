// Kode ini ditempel di Google Sheets -> Menu Ekstensi -> Apps Script

function sendSheetDataToBackend() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var data = sheet.getDataRange().getValues();
  
  if (data.length < 2) {
    Logger.log("Data kurang dari 2 baris (header + isi).");
    return;
  }

  var headers = data[0];
  var rows = data.slice(1);

  var payload = {
    sheet_name: sheet.getName(),
    headers: headers,
    rows: rows
  };

  var options = {
    method: "post",
    contentType: "application/json",
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  // Masukkan URL backend FastAPI Anda di sini (misal URL ngrok / server cloud)
  var backendUrl = "http://localhost:8000/api/ingest"; 

  try {
    var response = UrlFetchApp.fetch(backendUrl, options);
    Logger.log("Response: " + response.getContentText());
  } catch (error) {
    Logger.log("Error sending data: " + error.toString());
  }
}
