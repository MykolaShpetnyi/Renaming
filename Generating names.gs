function processFilesOnSheet() {
  var sheetName = "Processing";  // Названия листа, с которым работаем 
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  
  if (sheet == null) { //остановка при ненаходе
    Logger.log("Лист '" + sheetName + "' не найден.");
    return;
  }
  
  var lastRow = sheet.getLastRow();
  var values = sheet.getRange("C2:E" + lastRow).getValues(); // Опеределения, где нужно остановиться

  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet2 = spreadsheet.getSheetByName("Sheet1"); 
  for (var i = 0; i < values.length; i++) {
    var file = values[i][0];
    var status = values[i][2];
    var cell1 = sheet2.getRange("A10");
    var cell2 = sheet2.getRange("B10");
    var value1 = cell2.getValue();
    var cell3 = sheet2.getRange("C10");
    var value3 = cell3.getValue();
    var cell4 = sheet2.getRange("D10");
    var value2 = cell4.getValue();
    var cell5 = sheet2.getRange("E10");
    if (value3 !== "") {
    if (value2 !== "") {
    if (value1 !== "") {
          if (file && status) {
          var fileType = determineFileType(file);
          var outputValue = cell1.getValue() + " " + cell2.getValue() + " " + cell3.getValue() + " " + cell4.getValue() + " " + cell5.getValue(); // задаём название для файла
                
          sheet.getRange("F" + (i + 2)).setValue(outputValue);
        }} else {
          if (file && status) {
          var fileType = determineFileType(file);
          var outputValue = cell1.getValue() + " " + fileType.fileText.toUpperCase()+ " " + cell3.getValue() + " " + cell4.getValue() + " " + cell5.getValue(); // задаём название для файла
                    
          sheet.getRange("F" + (i + 2)).setValue(outputValue);
    }
  }
    }
    else {
    if (value1 !== "") {
          if (file && status) {
          var fileType = determineFileType(file);
          var outputValue = cell1.getValue() + " " + cell2.getValue() + " " + cell3.getValue() + " " + (status === "Done" || status === "successful refund " ? "ARCHIVED " : "FINAL") + " " + cell5.getValue(); // задаём название для файла
                
          sheet.getRange("F" + (i + 2)).setValue(outputValue);
        }} else {
          if (file && status) {
          var fileType = determineFileType(file);
          var outputValue = cell1.getValue() + " " + fileType.fileText.toUpperCase()+ " " + cell3.getValue() + " " + (status === "Done" || status === "successful refund " ? "ARCHIVED " : "FINAL") + " " + cell5.getValue(); // задаём название для файла
                    
          sheet.getRange("F" + (i + 2)).setValue(outputValue);
    }
  }
}}
else {
    if (value2 !== "") {
    if (value1 !== "") {
          if (file && status) {
          var fileType = determineFileType(file);
          var outputValue = cell1.getValue() + " " + cell2.getValue() + " " + fileType.fileTitle + " " + cell4.getValue() + " " + cell5.getValue(); // задаём название для файла
                
          sheet.getRange("F" + (i + 2)).setValue(outputValue);
        }} else {
          if (file && status) {
          var fileType = determineFileType(file);
          var outputValue = cell1.getValue() + " " + fileType.fileText.toUpperCase()+ " " + fileType.fileTitle + " " + cell4.getValue() + " " + cell5.getValue(); // задаём название для файла
                    
          sheet.getRange("F" + (i + 2)).setValue(outputValue);
    }
  }
    }
    else {
    if (value1 !== "") {
          if (file && status) {
          var fileType = determineFileType(file);
          var outputValue = cell1.getValue() + " " + cell2.getValue() + " " + fileType.fileTitle + " " + (status === "Done" || status === "successful refund " ? "ARCHIVED " : "FINAL") + " " + cell5.getValue(); // задаём название для файла
                
          sheet.getRange("F" + (i + 2)).setValue(outputValue);
        }} else {
          if (file && status) {
          var fileType = determineFileType(file);
          var outputValue = cell1.getValue() + " " + fileType.fileText.toUpperCase()+ " " + fileType.fileTitle + " " + (status === "Done" || status === "successful refund " ? "ARCHIVED " : "FINAL") + " " + cell5.getValue(); // задаём название для файла
                    
          sheet.getRange("F" + (i + 2)).setValue(outputValue);
    }
  }
}
  }}
}
function determineFileType(filename) {  // Определение формата файла
  var extension = filename.split('.').pop().toLowerCase();
  var fileType = "";
  var fileText = "";
  var fileTitle = filename.substr(0, filename.lastIndexOf('.')); // Получение названия файла без формата

  if (extension === "mp4" || extension === "avi" || extension === "mov" || extension === "webm") { 
    fileType = "Видео";
    fileText = "VID";
  } else if (extension === "jpeg" ||extension === "jpg" || extension === "png" || extension === "gif") {
    fileType = "Изображение";
    fileText = "IMG";
  } else {
    fileType = "Другой";
    fileText = "WDC";
  }

  return {
    fileType: fileType,
    fileText: fileText,
    fileTitle: fileTitle
  };
}

