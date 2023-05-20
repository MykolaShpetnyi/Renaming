function copyValues() {
  var sheet1 = SpreadsheetApp.openById('1234567'); // ID вашей таблицы в Google Drive, куда выгрузили названия файлов
  var tab1 = sheet1.getSheetByName('Processing'); // Названия листа, с которым работаем 
  
  var spreadsheet = SpreadsheetApp.openById('1234567');
  var sheet = spreadsheet.getActiveSheet();
  
  var cellC3 = sheet.getRange("C3");
  var valueC3 = cellC3.getValue();
  
  var cellC4 = sheet.getRange("C4");
  var valueC4 = cellC4.getValue();
  
  var sheet2 = SpreadsheetApp.openById(valueC3); // ID таблицы в Google Drive, откуда берём статусы
  var tab2 = sheet2.getSheetByName(valueC4); // Названия листа, откуда берём статусы
  
  var values = tab1.getRange('B:B').getValues(); // Буква столбца, по которой выполняем поиск
  var columnG = tab2.getRange('G:G').getValues(); // Буква столбца, где выполняем поиск
  
  for (var i = 0; i < values.length; i++) {
    var value = values[i][0];
    if (value !== "") {  // Пропуск пустых полей
      for (var j = 0; j < columnG.length; j++) {
        if (columnG[j][0] === value) {
          var valueL = tab2.getRange('L' + (j+1)).getValue(); // Буква столбца со статусами, которые передаём
          tab1.getRange('E' + (i+1)).setValue(valueL); // Буква столбца, куда записываем статусы 
          break;
        }
      }
    }
  }
}
