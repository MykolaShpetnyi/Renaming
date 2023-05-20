function listFolders() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadsheet.getActiveSheet();
  var cellC2 = sheet.getRange("C2");
  var folderId = cellC2.getValue(); // ID вашей папки в Google Drive
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheetName = "Processing"; // Создание уникального имени листа
  var sheet = spreadsheet.insertSheet(sheetName);
  
  sheet.appendRow(["main folder", "folder in folder", "files"]); // Заголовки столбцов

  var folder = DriveApp.getFolderById(folderId);
  processFolder(folder, sheet, "", "");
}

function processFolder(folder, sheet, parentFolderName, parentFolderPath) {
  var folderName = folder.getName();
  var folderPath = parentFolderPath + "/" + folderName;
  
  sheet.appendRow([parentFolderName, folderName, ""]); // Записываем текущую папку в таблицу
  
  var subfolders = folder.getFolders();
  while (subfolders.hasNext()) {
    var subfolder = subfolders.next();
    processFolder(subfolder, sheet, folderName, folderPath); // Рекурсивный вызов для обработки вложенных папок
  }

  var files = folder.getFiles();
  while (files.hasNext()) {
    var file = files.next();
    sheet.appendRow([parentFolderName, "", file.getName()]); // Записываем название файла в таблицу
  }
}
