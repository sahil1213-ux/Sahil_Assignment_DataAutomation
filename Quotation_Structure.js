function setQuotationHeaders() {
  const ss = SpreadsheetApp.openById(
    "11cHTSTjVQC96aTKilPms0tdPbugQRn5uU7GCfCYIh6o"
  );
  const sheet = ss.getSheetByName("Quotations");

  const headers = ["Date", "Sender", "Subject", "Product", "Quantity"];

  sheet.getRange(1, 1, 1, headers.length).setFontWeight("bold");
}

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu("📩 Email Actions")
    .addItem("🔄 Refresh List", "processQuotationEmails")
    .addToUi();
}
