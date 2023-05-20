function renameFiles() {
  var spreadsheet = SpreadsheetApp.openById('1234567');
  var sheetName = 'Processing'; // Названия листа, с которым работаем 
  var sheet = spreadsheet.getSheetByName(sheetName);
  var data = sheet.getDataRange().getValues();
  var sheetName1 = 'Sheet1'; // Названия листа, с которым работаем 
  var sheet2 = spreadsheet.getSheetByName(sheetName1);
  var cellC1 = sheet2.getRange("C1");
  var folderName = cellC1.getValue();
 // Названия папки, с которой работаем 

  // Проверка наличия значения имени папки
  if (folderName) {
    renameFilesInFolder(data, folderName);
  }
}

function renameFilesInFolder(data, folderName) {
  var folder = DriveApp.getFoldersByName(folderName).next();
  var files = folder.getFiles();

  while (files.hasNext()) {
    var file = files.next();
    var fileName = file.getName();
    var newFileName = getNewFileName(data, fileName);

    if (newFileName) {
      file.setName(newFileName);
    }
  }

  var subfolders = folder.getFolders();

  while (subfolders.hasNext()) {
    var subfolder = subfolders.next();
    var subfolderName = subfolder.getName();
    renameFilesInFolder(data, subfolderName);
  }
}

function getNewFileName(data, fileName) {
  for (var i = 1; i < data.length; i++) {
    if (data[i][2] === fileName) { // колонка со старым именем файла
      return data[i][5]; // колонка с новым именем файла
    }
  }

  return null;
}
