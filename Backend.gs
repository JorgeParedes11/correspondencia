function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
      .setTitle('Configuración de Correspondencia')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function conectarArchivos(docTemplateId, sheetId, folderId) {
  // Verify if the provided IDs are valid and accessible
  try {
    DocumentApp.openById(docTemplateId);
    SpreadsheetApp.openById(sheetId);
    DriveApp.getFolderById(folderId);
    
    PropertiesService.getScriptProperties().setProperties({
      'docTemplateId': docTemplateId,
      'sheetId': sheetId,
      'folderId': folderId
    });
    
    return 'Archivos conectados exitosamente';
  } catch (error) {
    throw new Error('Error al conectar archivos: ' + error.message);
  }
}

function verificarPestana(sheetName) {
  const sheetId = PropertiesService.getScriptProperties().getProperty('sheetId');
  if (!sheetId) {
    throw new Error('ID de la hoja de cálculo no configurado');
  }
  
  const spreadsheet = SpreadsheetApp.openById(sheetId);
  const sheet = spreadsheet.getSheetByName(sheetName);
  
  if (sheet) {
    return 'Pestaña verificada exitosamente';
  } else {
    throw new Error('La pestaña especificada no existe');
  }
}

function verificarMarcadores(mapeos) {
  const docTemplateId = PropertiesService.getScriptProperties().getProperty('docTemplateId');
  const sheetId = PropertiesService.getScriptProperties().getProperty('sheetId');
  
  if (!docTemplateId || !sheetId) {
    throw new Error('IDs de documento o hoja de cálculo no configurados');
  }
  
  const docTemplate = DocumentApp.openById(docTemplateId);
  const sheet = SpreadsheetApp.openById(sheetId).getActiveSheet();
  
  const docContent = docTemplate.getBody().getText();
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  
  let errores = [];
  
  mapeos.forEach(mapeo => {
    if (!docContent.includes(mapeo.marcador)) {
      errores.push(`Marcador "${mapeo.marcador}" no encontrado en el documento`);
    }
    if (!headers.includes(mapeo.columna)) {
      errores.push(`Columna "${mapeo.columna}" no encontrada en la hoja de cálculo`);
    }
  });
  
  if (errores.length > 0) {
    throw new Error(errores.join('\n'));
  }
  
  return 'Todos los marcadores y columnas verificados exitosamente';
}

function reporteMovimientos() {
  // Implementar lógica para generar reporte de movimientos
  return 'Reporte de movimientos generado';
}

function generarDocumentos() {
  // Implementar lógica para generar documentos
  return 'Documentos generados exitosamente';
}
