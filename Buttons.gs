function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Меню')
    .addItem('Выгрузить данные папки в таблицу', 'function1')
    .addItem('Подтянуть статусы из второй таблицы', 'copyValues')
    .addItem('Распространить статусы на файлы', 'function2')
    .addItem('Сгенерировать новые названия', 'function3')
    .addItem('Переимновать файлы', 'function4')
    .addToUi();
}

function function1() {
  listFolders();
  processFolder();
}

function function2() {
  repeatCopyValuesFromSheet();
  copyValuesFromSheet();
}

function function3() {
  processFilesOnSheet();
  determineFileType(filename);
}

function function4() {
  renameFiles();
  renameFilesInFolder(data, folderName);
  getNewFileName(data, fileName);
}
