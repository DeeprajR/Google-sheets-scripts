function extractHyperlinks() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var range = sheet.getDataRange();
  var values = range.getValues();
  
  for (var i = 0; i < values.length; i++) {
    var cell = sheet.getRange(i + 1, 1); // Assuming company names are in column A
    var formula = cell.getFormula();
    
    if (formula) {
      var match = formula.match(/HYPERLINK\("([^"]+)"/);
      if (match) {
        sheet.getRange(i + 1, 2).setValue(match[1]); // Place hyperlink in column B
      }
    } else {
      var hyperlink = cell.getRichTextValue() ? cell.getRichTextValue().getLinkUrl() : "";
      if (hyperlink) {
        sheet.getRange(i + 1, 2).setValue(hyperlink); // Place hyperlink in column B
      }
    }
  }
}
