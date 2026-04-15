function jsonResponse(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

function clean(value) {
  if (value === null || value === undefined) {
    return "";
  }
  return String(value).trim();
}

function toNumber(value, fallback) {
  const number = Number(value);
  return Number.isFinite(number) ? number : fallback;
}

function isTruthy(value) {
  if (value === true || value === false) {
    return value;
  }

  const normalized = clean(value).toLowerCase();
  return normalized === "true" || normalized === "1" || normalized === "yes" || normalized === "y";
}

function isValidEmail(email) {
  return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email);
}

function getScriptProperty(key) {
  return clean(PropertiesService.getScriptProperties().getProperty(key));
}

function getSheet(sheetName) {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = spreadsheet.getSheetByName(sheetName);

  if (!sheet) {
    throw new Error("Sheet '" + sheetName + "' not found");
  }

  return sheet;
}

function getHeaderMap(sheet) {
  const lastColumn = sheet.getLastColumn();
  if (lastColumn < 1) {
    throw new Error("Sheet '" + sheet.getName() + "' does not have headers");
  }

  const headers = sheet.getRange(1, 1, 1, lastColumn).getValues()[0];
  const headerMap = {};

  headers.forEach(function(header, index) {
    const key = clean(header);
    if (key) {
      headerMap[key] = index + 1;
    }
  });

  return headerMap;
}

function requireHeaders(sheet, names) {
  const headerMap = getHeaderMap(sheet);

  names.forEach(function(name) {
    if (!headerMap[name]) {
      throw new Error("Missing required column '" + name + "' in sheet '" + sheet.getName() + "'");
    }
  });

  return headerMap;
}

function getSheetObjects(sheet) {
  const lastRow = sheet.getLastRow();
  const lastColumn = sheet.getLastColumn();

  if (lastRow < 2 || lastColumn < 1) {
    return [];
  }

  const headers = sheet.getRange(1, 1, 1, lastColumn).getValues()[0];
  const rows = sheet.getRange(2, 1, lastRow - 1, lastColumn).getValues();

  return rows.map(function(row, rowIndex) {
    const item = { _rowNumber: rowIndex + 2 };

    headers.forEach(function(header, columnIndex) {
      const key = clean(header);
      if (key) {
        item[key] = row[columnIndex];
      }
    });

    return item;
  });
}

function appendObjectRow(sheet, valuesByHeader) {
  const headerMap = getHeaderMap(sheet);
  const lastColumn = sheet.getLastColumn();
  const row = [];

  for (let column = 1; column <= lastColumn; column += 1) {
    row.push("");
  }

  Object.keys(valuesByHeader).forEach(function(header) {
    if (headerMap[header]) {
      row[headerMap[header] - 1] = valuesByHeader[header];
    }
  });

  sheet.appendRow(row);
  return sheet.getLastRow();
}

function setCellIfHeaderExists(sheet, rowNumber, headerMap, headerName, value) {
  if (headerMap[headerName]) {
    sheet.getRange(rowNumber, headerMap[headerName]).setValue(value);
  }
}

function formatDateValue(value) {
  if (!value) {
    return "";
  }

  if (Object.prototype.toString.call(value) === "[object Date]" && !isNaN(value.getTime())) {
    return Utilities.formatDate(value, Session.getScriptTimeZone(), "yyyy-MM-dd");
  }

  return clean(value);
}

function formatDateTimeValue(value) {
  if (!value) {
    return "";
  }

  if (Object.prototype.toString.call(value) === "[object Date]" && !isNaN(value.getTime())) {
    return Utilities.formatDate(value, Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm:ss");
  }

  return clean(value);
}

function sortByDateDesc(items, fieldName) {
  return items.sort(function(left, right) {
    return String(right[fieldName] || "").localeCompare(String(left[fieldName] || ""));
  });
}
