function repeatCopyValuesFromSheet() {
  for (var i = 0; i < 30; i++) { // 30 - число повторений, рекомендуется поднять значение если в папке файлов больше 30 
    copyValuesFromSheet();
  }
}

function copyValuesFromSheet() {
  var spreadsheet1 = SpreadsheetApp.openById("1234567"); // ID вашей данной таблицы в Google Drive
  var sheet1 = spreadsheet1.getSheetByName("Processing");  // Названия листа, с которым работаем 

  var values1 = sheet1.getDataRange().getValues();

  for (var i = 1; i < values1.length; i++) {
    var valueB = values1[i][1]; // Значение столбца B текущей строки
    var valueC = values1[i][2]; // Значение столбца C текущей строки
    
    if (valueC !== "" && valueB === "") {
      var prevValueE = values1[i-1][4]; // Значение столбца E предыдущей строки
      sheet1.getRange("E" + (i+1)).setValue(prevValueE);
    }
  }
}
