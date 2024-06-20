function onEdit(e) {
  try {
    Logger.log("onEdit trigger fired");

    if (!e || !e.range) {
      throw new Error('Event object is undefined or range is not found');
    }

    var range = e.range;
    var sheet = range.getSheet();
    var sheetName = sheet.getName();
    Logger.log("Edited sheet: " + sheetName);
    Logger.log("Edited range: " + range.getA1Notation());

    var settings = [
      { sheetName: 'DRINK', startRow: 4, columnB: 7, startColumn: 10, n: 1 },
      { sheetName: 'DRINK', startRow: 4, columnB: 5, startColumn: 11, n: 1 },
      { sheetName: 'DRINK', startRow: 4, columnB: 6, startColumn: 11, n: 1 },
      { sheetName: 'RAK6', startRow: 2, columnB: 6, startColumn: 8, n: 1 },
      { sheetName: 'RAK5', startRow: 2, columnB: 6, startColumn: 8, n: 1 },
      { sheetName: 'RAK4', startRow: 2, columnB: 6, startColumn: 8, n: 1 },
      { sheetName: 'RAK3', startRow: 2, columnB: 6, startColumn: 8, n: 1 },
      { sheetName: 'RAK2', startRow: 2, columnB: 6, startColumn: 8, n: 1 },
      { sheetName: 'RAK1', startRow: 2, columnB: 6, startColumn: 8, n: 1 },
      { sheetName: 'RETUR Mie', startRow: 2, columnB: 3, startColumn: 34, n: 1 },
      { sheetName: 'SCRAP', startRow: 2, columnB: 3, startColumn: 34, n: 1 }
    ];

    for (var j = 0; j < settings.length; j++) {
      var setting = settings[j];
      if (sheetName === setting.sheetName && range.getColumn() === setting.columnB && range.getRow() >= setting.startRow) {
        Logger.log("Match found in settings: " + JSON.stringify(setting));
        var row = range.getRow();
        var startColumn = setting.startColumn;
        var n = setting.n;

        setTimestamp(sheet, row, startColumn, n);
        setBackground(range);
        saveEditedCellInfo(range, sheetName);
        break; // Break out of the loop once a match is found
      }
    }

    logEditHistory(e, sheetName, range);

    Logger.log("Script execution completed");
  } catch (error) {
    Logger.log('Error: ' + error.message);
    Logger.log(error.stack);
  }
}

function setTimestamp(sheet, row, startColumn, n) {
  for (var col = startColumn; col < startColumn + n; col++) {
    var cell = sheet.getRange(row, col);
    var formattedDate = getFormattedTimestamp();
    cell.setValue(formattedDate);
    Logger.log("Set timestamp in cell: " + cell.getA1Notation() + " with value: " + formattedDate);
  }
}

function getFormattedTimestamp() {
  var timestamp = new Date();
  var dayNames = ["Minggu,", "Senin,", "Selasa,", "Rabu,", "Kamis,", "Jum'at,", "Sabtu,"];
  var dayName = dayNames[timestamp.getDay()];
  return dayName + ' ' + Utilities.formatDate(timestamp, Session.getScriptTimeZone(), "dd/MM/yyyy H:mm:ss");
}

function setBackground(range) {
  var timestamp = new Date();
  var dayIndex = (timestamp.getDate() - 1) % 9;
  var colors = ["#d9ead3", "#d9ead3", "#d9ead3", "#cfe2f3", "#cfe2f3", "#cfe2f3", "#e4d7f5", "#e4d7f5", "#e4d7f5"];
  var color = colors[dayIndex];
  range.setBackground(color);
  Logger.log("Set cell background color to: " + color);
}

function saveEditedCellInfo(range, sheetName) {
  var editedCells = PropertiesService.getDocumentProperties().getProperty('editedCells');
  editedCells = editedCells ? JSON.parse(editedCells) : [];
  editedCells.push({ range: range.getA1Notation(), sheet: sheetName });
  PropertiesService.getDocumentProperties().setProperty('editedCells', JSON.stringify(editedCells));
}

function logEditHistory(e, sheetName, range) {
  var timestamp = new Date();
  var row = range.getRow();
  var column = range.getColumn();
  var editedPcs, editedBox;

  if (['RAK1', 'RAK2', 'RAK3', 'RAK4', 'RAK5', 'RAK6'].includes(sheetName)) {
    if (column === 6 && row >= 3 && row <= 200) {
      editedPcs = range.getValue();
    } else {
      return; // Do not log changes outside column 6 or row range 3-100
    }
  } else if (sheetName === 'DRINK') {
    if ((column === 5 || column === 6) && row >= 5 && row <= 200) {
      var col5Value = range.getSheet().getRange(row, 5).getValue();
      var col6Value = range.getSheet().getRange(row, 6).getValue();
      if (col5Value !== "" && col6Value !== "") {
        editedPcs = range.getSheet().getRange(row, 8).getValue(); // Column 8
      } else {
        return;
      }
    } else if (column === 7 && row >= 5 && row <= 200) {
      editedBox = range.getValue();
    } else {
      return;
    }
  } else if (sheetName === 'RETUR Mie' || sheetName === 'SCRAP') {
    if (column >= 3 && column <= 33) { // Columns C to AG
      editedBox = range.getSheet().getRange(row, 34).getValue(); // Column AH
    } else {
      return;
    }
  } else if (sheetName === 'RETUR SMU' && column === 3 && row >= 2 && row <= 200) {
    editedBox = range.getValue();
  } else {
    return;
  }

  var material, mid;
  if (['DRINK', 'RAK1', 'RAK2', 'RAK3', 'RAK4', 'RAK5', 'RAK6'].includes(sheetName)) {
    material = range.getSheet().getRange(row, 3).getValue();
    mid = range.getSheet().getRange(row, 2).getValue();
  } else if (['RETUR Mie', 'SCRAP', 'RETUR SMU'].includes(sheetName)) {
    material = range.getSheet().getRange(row, 2).getValue();
    mid = range.getSheet().getRange(row, 1).getValue();
  }

  var editedBoth = editedPcs !== undefined && editedBox !== undefined;

  if (!editedBoth) {
    cleanDuplicateEntries(material, mid, sheetName);
  }

  var historySheetName = "EditHistory";
  var historySheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(historySheetName);

  if (!historySheet) {
    historySheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet(historySheetName);
    historySheet.appendRow(["Timestamp", "Sheet Name", "MID", "Material", "BOX", "PCS"]);
  }

  historySheet.appendRow([timestamp, sheetName, mid, material, editedBox, editedPcs]);

  if (editedPcs !== undefined) {
    var messagePcs = `Data di sheet '${sheetName}' telah diubah : \nMID: ${mid}\nMaterial: ${material}\nPCS: ${editedPcs} Pcs `;
    sendTelegramMessage(messagePcs);
  }
  if (editedBox !== undefined) {
    var messageBox = `Data di sheet '${sheetName}' telah diubah : \nMID: ${mid}\nMaterial: ${material}\nBOX: ${editedBox} Box `;
    sendTelegramMessage(messageBox);
  }
}

function cleanDuplicateEntries(material, mid, sheetName) {
  var historySheetName = "EditHistory";
  var historySheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(historySheetName);

  if (!historySheet) {
    Logger.log("Sheet " + historySheetName + " not found.");
    return;
  }

  var lastRow = historySheet.getLastRow();
  if (lastRow < 2) {
    Logger.log("No data in sheet " + historySheetName + ".");
    return;
  }

  var dataRange = historySheet.getRange(2, 1, lastRow - 1, 6);
  var dataValues = dataRange.getValues();

  var rowsToDelete = [];
  for (var i = dataValues.length - 1; i >= 0; i--) {
    var row = dataValues[i];
    Logger.log("Matching row " + (i + 2) + ": " + row.join(", "));
    if (row[1] === sheetName && row[2] === mid && row[3] === material) {
      Logger.log("Marking duplicate row for deletion at " + (i + 2));
      rowsToDelete.push(i + 2);
    }
  }

  for (var i = rowsToDelete.length - 1; i >= 0; i--) {
    historySheet.deleteRow(rowsToDelete[i]);
  }
}


  var payload = {
    'chat_id': chatId,
    'text': message
  };

  var options = {
    'method': 'post',
    'payload': payload
  };

  UrlFetchApp.fetch(url, options);
}






