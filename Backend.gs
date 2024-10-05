// Función que sirve la página HTML principal cuando el usuario accede
// Devuelve el contenido de 'index.html'

function doGet() {
  // Usamos createTemplateFromFile para procesar las plantillas correctamente
  return HtmlService.createTemplateFromFile('Index-7')
    .evaluate()
    .setTitle('Configuración de Correspondencia')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}


// Descripción de lo que hace esta función (ajusta según sea necesario)
// Explica qué parámetros recibe y qué devuelve

function include(filename) {
  // Función para incluir otros archivos HTML
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function verifyConnection(googleDocId, googleSheetId, outputFolderId) {
  var logs = [];
  try {
    logs.push('Iniciando verifyConnection');
    logs.push('googleDocId: ' + googleDocId);
    logs.push('googleSheetId: ' + googleSheetId);
    logs.push('outputFolderId: ' + outputFolderId);

    // Verificar el documento y obtener su nombre
    var doc = DocumentApp.openById(googleDocId);
    var docName = doc.getName();
    logs.push('Documento abierto correctamente: ' + docName);

    // Verificar la hoja de cálculo y obtener su nombre
    var sheet = SpreadsheetApp.openById(googleSheetId);
    var sheetName = sheet.getName();
    logs.push('Hoja de cálculo abierta correctamente: ' + sheetName);

    // Verificar la carpeta de salida y obtener su nombre
    var folder = DriveApp.getFolderById(outputFolderId);
    var folderName = folder.getName();
    logs.push('Carpeta de salida abierta correctamente: ' + folderName);

    return {
      success: true,
      docName: docName,
      sheetName: sheetName,
      folderName: folderName,
      logs: logs
    };
  } catch (error) {
    logs.push('Error en verifyConnection: ' + error.toString());
    return {
      success: false,
      error: error.toString(),
      logs: logs
    };
  }
}

function verifySheetName(sheetName, googleSheetId) {
  var logs = [];
  try {
    logs.push('Iniciando verifySheetName');
    logs.push('sheetName: ' + sheetName);
    logs.push('googleSheetId: ' + googleSheetId);

    var spreadsheet = SpreadsheetApp.openById(googleSheetId);
    var sheets = spreadsheet.getSheets();
    var sheetNames = sheets.map(function(sheet) {
      return sheet.getName();
    });

    logs.push('Pestañas disponibles: ' + sheetNames.join(', '));

    if (sheetNames.indexOf(sheetName) !== -1) {
      logs.push('La pestaña existe en el Google Sheet.');
      return {
        exists: true,
        message: 'La pestaña "' + sheetName + '" existe en el Google Sheet.',
        logs: logs
      };
    } else {
      logs.push('La pestaña no existe en el Google Sheet.');
      return {
        exists: false,
        message: 'La pestaña "' + sheetName + '" no existe en el Google Sheet.',
        availableSheets: sheetNames,
        logs: logs
      };
    }
  } catch (error) {
    logs.push('Error en verifySheetName: ' + error.toString());
    return {
      exists: false,
      message: 'Error al verificar el nombre de la pestaña.',
      logs: logs,
      error: error.toString()
    };
  }
}

function verifyMappings(mappings, googleDocId, googleSheetId, sheetName) {
  var logs = [];
  try {
    logs.push('Iniciando verifyMappings');
    logs.push('googleDocId: ' + googleDocId);
    logs.push('googleSheetId: ' + googleSheetId);
    logs.push('sheetName: ' + sheetName);

    var doc = DocumentApp.openById(googleDocId);
    var docText = doc.getBody().getText();
    logs.push('Documento abierto correctamente');

    var spreadsheet = SpreadsheetApp.openById(googleSheetId);
    var targetSheet = spreadsheet.getSheetByName(sheetName);
    if (!targetSheet) {
      logs.push('La pestaña especificada no existe en el Google Sheet.');
      throw 'La pestaña especificada no existe en el Google Sheet.';
    }
    logs.push('Hoja de cálculo y pestaña abiertas correctamente');

    var headers = targetSheet.getRange(1, 1, 1, targetSheet.getLastColumn()).getValues()[0];
    logs.push('Encabezados obtenidos: ' + headers.join(', '));

    var results = mappings.map(function(mapping) {
      var result = {
        docMarkerExists: false,
        sheetColumnExists: false
      };

      logs.push('Verificando mapeo: ' + JSON.stringify(mapping));

      // Si ambos están vacíos, los consideramos como válidos pero serán ignorados
      if (mapping.docMarker === '' && mapping.sheetColumn === '') {
        result.docMarkerExists = true;
        result.sheetColumnExists = true;
        return result;
      }

      // Verificar marcador en el documento
      if (mapping.docMarker && docText.indexOf(mapping.docMarker) !== -1) {
        result.docMarkerExists = true;
      }

      // Verificar columna en la hoja de cálculo
      if (mapping.sheetColumn && headers.indexOf(mapping.sheetColumn) !== -1) {
        result.sheetColumnExists = true;
      }

      logs.push('Resultado de la verificación: ' + JSON.stringify(result));
      return result;
    });

    logs.push('Verificación completada');

    // Devolvemos los resultados y los logs al cliente
    return {
      success: true,
      results: results,
      logs: logs
    };
  } catch (error) {
    logs.push('Error en verifyMappings: ' + error);
    // Devolvemos el error y los logs al cliente
    return {
      success: false,
      error: error.toString(),
      logs: logs
    };
  }
}

function generateDocuments(googleDocId, googleSheetId, sheetName, outputFolderId, columnMappings) {
  var logs = [];
  try {
    logs.push('Iniciando generateDocuments');
    logs.push('googleDocId: ' + googleDocId);
    logs.push('googleSheetId: ' + googleSheetId);
    logs.push('sheetName: ' + sheetName);
    logs.push('outputFolderId: ' + outputFolderId);
    logs.push('columnMappings: ' + JSON.stringify(columnMappings));

    // Abrir el documento de plantilla
    var templateDoc = DocumentApp.openById(googleDocId);
    logs.push('Documento de plantilla abierto correctamente');

    // Abrir la hoja de cálculo
    var spreadsheet = SpreadsheetApp.openById(googleSheetId);
    var sheet = spreadsheet.getSheetByName(sheetName);
    if (!sheet) {
      logs.push('La pestaña especificada no existe en el Google Sheet.');
      throw 'La pestaña especificada no existe en el Google Sheet.';
    }
    logs.push('Hoja de cálculo y pestaña abiertas correctamente');

    // Obtener los datos de la hoja de cálculo
    var dataRange = sheet.getDataRange();
    var data = dataRange.getValues();
    var headers = data[0];
    logs.push('Datos obtenidos de la hoja de cálculo');

    // Obtener la carpeta de salida
    var outputFolder = DriveApp.getFolderById(outputFolderId);
    logs.push('Carpeta de salida obtenida correctamente');

    // Recorremos cada fila de datos (excepto la primera que es de encabezados)
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      var newDoc = templateDoc.copy('Documento Generado - Fila ' + i);
      var body = newDoc.getBody();

      // Reemplazar los marcadores con los datos correspondientes
      columnMappings.forEach(function(mapping) {
        var docMarker = mapping.docMarker;
        var sheetColumn = mapping.sheetColumn;
        // Obtener el índice de la columna
        var columnIndex = headers.indexOf(sheetColumn);
        if (columnIndex !== -1) {
          var value = row[columnIndex];
          body.replaceText(docMarker, value);
        }
      });

      // Mover el documento a la carpeta de salida
      var newDocFile = DriveApp.getFileById(newDoc.getId());
      outputFolder.addFile(newDocFile);
      DriveApp.getRootFolder().removeFile(newDocFile); // Remover de Mi Unidad
      logs.push('Documento generado y movido a la carpeta de salida: Documento Generado - Fila ' + i);
    }

    logs.push('Generación de documentos completada');

    return {
      success: true,
      logs: logs
    };
  } catch (error) {
    logs.push('Error en generateDocuments: ' + error);
    return {
      success: false,
      error: error.toString(),
      logs: logs
    };
  }
}

function generateDocuments(googleDocId, googleSheetId, outputFolderId, columnMappings) {
  try {
    var template = DocumentApp.openById(googleDocId);
    var sheet = SpreadsheetApp.openById(googleSheetId);
    var outputFolder = DriveApp.getFolderById(outputFolderId);
    var data = sheet.getDataRange().getValues();
    var headers = data[0];

    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      var newDoc = template.makeCopy(outputFolder);
      var body = newDoc.getBody();

      columnMappings.forEach(function(mapping) {
        var columnIndex = headers.indexOf(mapping.sheetColumn);
        if (columnIndex !== -1) {
          body.replaceText(mapping.docMarker, row[columnIndex] || '');
        }
      });

      newDoc.setName('Documento generado - ' + i);
      newDoc.saveAndClose();
    }

    return true;
  } catch (error) {
    Logger.log('Error en la generación de documentos: ' + error.toString());
    return false;
  }
}

function verifyMappings(mappings, googleDocId, googleSheetId, sheetName) {
  try {
    var logs = [];
    logs.push('Iniciando verifyMappings');
    logs.push('googleDocId: ' + googleDocId);
    logs.push('googleSheetId: ' + googleSheetId);
    logs.push('sheetName: ' + sheetName);

    var doc = DocumentApp.openById(googleDocId);
    var docText = doc.getBody().getText();
    logs.push('Documento abierto correctamente');

    var spreadsheet = SpreadsheetApp.openById(googleSheetId);
    var targetSheet = spreadsheet.getSheetByName(sheetName);
    if (!targetSheet) {
      throw 'La pestaña especificada no existe en el Google Sheet.';
    }
    logs.push('Hoja de cálculo y pestaña abiertas correctamente');

    var headers = targetSheet.getRange(1, 1, 1, targetSheet.getLastColumn()).getValues()[0];
    logs.push('Encabezados obtenidos: ' + headers.join(', '));

    var results = mappings.map(function(mapping) {
      var result = {
        docMarkerExists: false,
        sheetColumnExists: false
      };

      logs.push('Verificando mapeo: ' + JSON.stringify(mapping));

      // Si ambos están vacíos, los consideramos como válidos pero serán ignorados
      if (mapping.docMarker === '' && mapping.sheetColumn === '') {
        result.docMarkerExists = true;
        result.sheetColumnExists = true;
        return result;
      }

      // Verificar marcador en el documento
      if (mapping.docMarker && docText.indexOf(mapping.docMarker) !== -1) {
        result.docMarkerExists = true;
      }

      // Verificar columna en la hoja de cálculo
      if (mapping.sheetColumn && headers.indexOf(mapping.sheetColumn) !== -1) {
        result.sheetColumnExists = true;
      }

      logs.push('Resultado de la verificación: ' + JSON.stringify(result));
      return result;
    });

    logs.push('Verificación completada');

    // Devolvemos los resultados y los logs al cliente
    return {
      success: true,
      results: results,
      logs: logs
    };
  } catch (error) {
    logs.push('Error en verifyMappings: ' + error);
    // Devolvemos el error y los logs al cliente
    return {
      success: false,
      error: error.toString(),
      logs: logs
    };
  }
}

// Backend-4.gs

function verifyMappingsServer(mappings, googleDocId, googleSheetId, sheetName) {
  try {
    var doc = DocumentApp.openById(googleDocId);
    var docText = doc.getBody().getText();

    var spreadsheet = SpreadsheetApp.openById(googleSheetId);
    var sheet = spreadsheet.getSheetByName(sheetName);
    if (!sheet) {
      return {
        success: false,
        error: 'La pestaña especificada no existe en el Google Sheet.'
      };
    }

    var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];

    var results = mappings.map(function(mapping) {
      return {
        docMarkerExists: docText.indexOf(mapping.docMarker) !== -1,
        sheetColumnExists: headers.indexOf(mapping.sheetColumn) !== -1
      };
    });

    return {
      success: true,
      results: results
    };
  } catch (error) {
    return {
      success: false,
      error: error.toString()
    };
  }
}

function verifySheetName(sheetName, googleSheetId) {
  var logs = [];
  try {
    logs.push("Iniciando verificación para la pestaña: " + sheetName);
    logs.push("ID de la hoja de cálculo: " + googleSheetId);

    var spreadsheet = SpreadsheetApp.openById(googleSheetId);
    logs.push("Hoja de cálculo abierta correctamente");

    var allSheets = spreadsheet.getSheets();
    var sheetNames = allSheets.map(function(s) { return s.getName(); });
    logs.push("Nombres de todas las pestañas: " + JSON.stringify(sheetNames));

    var targetSheet = spreadsheet.getSheetByName(sheetName);
    if (targetSheet !== null) {
      logs.push("Pestaña encontrada");
      return {
        exists: true,
        message: "Pestaña '" + sheetName + "' encontrada",
        sheetInfo: {
          name: targetSheet.getName(),
          index: targetSheet.getIndex(),
          id: targetSheet.getSheetId()
        },
        logs: logs
      };
    } else {
      logs.push("Pestaña no encontrada");
      return {
        exists: false,
        message: "Pestaña '" + sheetName + "' no encontrada",
        availableSheets: sheetNames,
        logs: logs
      };
    }
  } catch (error) {
    logs.push("Error: " + error.toString());
    logs.push("Stack: " + error.stack);
    return {
      exists: false,
      message: "Error en la verificación de la pestaña",
      error: error.toString(),
      logs: logs
    };
  }
}

function getSpreadsheetInfo(googleSheetId) {
  try {
    var spreadsheet = SpreadsheetApp.openById(googleSheetId);
    var sheets = spreadsheet.getSheets();
    var activeSheet = spreadsheet.getActiveSheet();

    var sheetInfo = sheets.map(function(sheet) {
      return {
        name: sheet.getName(),
        index: sheet.getIndex(),
        id: sheet.getSheetId(),
        isActive: (sheet.getSheetId() === activeSheet.getSheetId())
      };
    });

    return {
      success: true,
      name: spreadsheet.getName(),
      url: spreadsheet.getUrl(),
      sheets: sheetInfo,
      activeSheet: activeSheet.getName()
    };
  } catch (error) {
    return {
      success: false,
      error: error.toString(),
      stack: error.stack
    };
  }
}

function getRange(googleSheetId, sheetName, columnName, fromRow, toRow) {
  try {
    var spreadsheet = SpreadsheetApp.openById(googleSheetId);
    var sheet = spreadsheet.getSheetByName(sheetName);
    if (!sheet) {
      return {
        success: false,
        error: 'La pestaña especificada no existe en el Google Sheet.'
      };
    }

    var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    var columnIndex = headers.indexOf(columnName) + 1;
    if (columnIndex === 0) {
      return {
        success: false,
        error: 'La columna especificada no existe en la hoja.'
      };
    }

    var columnLetter = String.fromCharCode(64 + columnIndex);
    var range = columnLetter + fromRow + ':' + columnLetter + toRow;
    var data = sheet.getRange(range).getValues();
    var flatData = data.map(function(row) { return row[0]; });

    return {
      success: true,
      range: range,
      data: flatData
    };
  } catch (error) {
    return {
      success: false,
      error: error.toString()
    };
  }
}

function generateDocuments(googleDocId, googleSheetId, sheetName, outputFolderId, columnMappings, rangeFrom, rangeTo) {
  try {
    var template = DocumentApp.openById(googleDocId);
    var sheet = SpreadsheetApp.openById(googleSheetId).getSheetByName(sheetName);
    var outputFolder = DriveApp.getFolderById(outputFolderId);

    var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    var data = sheet.getRange(rangeFrom, 1, rangeTo - rangeFrom + 1, sheet.getLastColumn()).getValues();

    for (var i = 0; i < data.length; i++) {
      var newDoc = DriveApp.getFileById(template.getId()).makeCopy(outputFolder);
      var doc = DocumentApp.openById(newDoc.getId());
      var body = doc.getBody();

      columnMappings.forEach(function(mapping) {
        var columnIndex = headers.indexOf(mapping.sheetColumn);
        if (columnIndex !== -1) {
          body.replaceText(mapping.docMarker, data[i][columnIndex] || '');
        }
      });

      doc.saveAndClose();
    }

    return {
      success: true,
      message: 'Documentos generados con éxito.'
    };
  } catch (error) {
    return {
      success: false,
      error: error.toString()
    };
  }
}
