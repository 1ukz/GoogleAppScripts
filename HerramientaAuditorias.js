// CREATOR: LUCAS SAETA (@1ukz) lucassaeta5@gmail.com

/**
 * Esta función se ejecuta al abrir la hoja de cálculo. 
 * Crea un menú personalizado llamado 'Herramienta Auditorías' en la interfaz de usuario de Google Sheets.
 * Al seleccionar la opción correspondiente en el Item en el menú, se invoca la función que tenga al lado.
 */
function onOpen() {
  const UI = SpreadsheetApp.getUi(); 

  UI.createMenu('Herramienta Auditorías')
    .addItem('Crear carpetas', 'createFoldersAndSpreadsheets')
    .addToUi();
}


/** FUNCIÓN AUXILIAR
 * Crea una hoja de log en la hoja de cálculo activa.
 * Si la hoja de registro ya existe, se limpia su contenido.
 * Se registra la fecha y hora de la creación de la hoja en la celda A1.
 *
 * @param {string} nombreHojaLogs - El nombre para la hoja de registro que se va a crear o limpiar.
 */
function createLogSheet(nombreHojaLogs){
  const ID = SpreadsheetApp.getActiveSpreadsheet().getId();
  const MAIN_SHEET = SpreadsheetApp.openById(ID);
  var newLogSheet = MAIN_SHEET.getSheetByName(nombreHojaLogs);

    if (!newLogSheet) {
      newLogSheet = MAIN_SHEET.insertSheet(nombreHojaLogs);
    } else {
      newLogSheet.clear();
    }
  var date = new Date();
  var cell = newLogSheet.getRange(1,1); 
  cell.setValue('Logs generados para la ejecución de: "' + nombreHojaLogs + '" a las ' + date.getHours() + ':' + date.getMinutes());
  cell.setFontSize(14); // Ajusta el tamanio de la fuente 
  cell.setFontWeight('bold'); // Aplica negrita al texto
}

/** FUNCIÓN AUXILIAR
 * Registra un mensaje en la hoja de log especificada.
 * Agrega el mensaje en la siguiente fila vacía de la hoja de registro.
 *
 * @param {string} nombreHoja - El nombre de la hoja donde se registrará el mensaje.
 * @param {string} message - El mensaje que se va a registrar.
 */
function logToSheet(nombreHoja, message){
  const ID = SpreadsheetApp.getActiveSpreadsheet().getId();
  const MAIN_SHEET = SpreadsheetApp.openById(ID);

  var logSheet = MAIN_SHEET.getSheetByName(nombreHoja);  
  const lastRow = logSheet.getLastRow();
  logSheet.getRange(lastRow + 1, 1).setValue(message);  

}


//Esta función crea carpetas padres e hijas desde una spreadsheet
//crea una hoja de cálculo en la padre 
//Le pone el nombre del padre + el texto que queramos
//Escribe en el spreadsheet inicial el número de id de la hoja que crea para utilizarlo en otro momento.

function createFoldersAndSpreadsheets() {

  const ID = SpreadsheetApp.getActiveSpreadsheet().getId();
  const SS = SpreadsheetApp.openById(ID);
  const UI = SpreadsheetApp.getUi(); 
  var date = new Date();

  var sheet = SS.getSheetByName('Crear Carpetas');
  var extension = sheet.getRange('C9').getValue().trim();
  var nombreHojaLogs = '(' + date.getDate() + '/' + (date.getMonth() + 1) + '/' + date.getFullYear() + ') - Logs CREAR CARPETAS'; 
  createLogSheet(nombreHojaLogs);

  UI.alert('Comenzando ejecución de la función CREAR CARPETAS. Por favor, espere pacientemente hasta el siguiente mensaje que indique que el proceso ha terminado.\n\nSe ha generado una Hoja con los Logs de lo realizado para esta ejecución, disponible en: "' + nombreHojaLogs + '"')

  // Get the data range assuming the parent folder names are in column A and the child folder names are in column B
  var range = sheet.getDataRange();
  var values = range.getValues();

  // Specify the ID or the name of the existing folder to hold all parent folders
  var topLevelFolderId = sheet.getRange('C7').getValue().trim(); 

  // Get the existing top-level folder
  var topLevelFolder = DriveApp.getFolderById(topLevelFolderId);
  if (!topLevelFolder) {
    // If the folder is not found, you can handle it here or log an error
    UI.alert("La carpeta con el ID proporcionado en la celda 'C7' no existe. Por favor, revise que el ID es correcto.");
    logToSheet(nombreHojaLogs, "La carpeta con el ID proporcionado en la celda 'C7' no existe. Por favor, revise que el ID es correcto.")
    return;
  }

  // Get the child folders from the first row of column B (13 porque empieza en la fila 13)
  var childFolders = values[13][1].toString().split(",");

  // Loop through the rows (13 para omitir todos los values que haya antes de la fila 13)
  for (var i = 0; i < values.length - 13; i++) {
    var parentFolderName = values[i+13][0]; // Get parent folder name from column A

    if (parentFolderName) { // Check if parent name is provided
      var parentFolder = topLevelFolder.createFolder(parentFolderName); // Create parent folder
      logToSheet(nombreHojaLogs, "--------------------------------------------------------------------------------------");
      logToSheet(nombreHojaLogs, "La carpeta: '" + parentFolderName + "' ha sido creada.");
      // Loop through child folder names and create subfolders
      var msg = "Las subcarpetas:" 
      for (var j = 0; j < childFolders.length; j++) {
        var childFolderName = childFolders[j].trim();
        msg += "  '" + childFolderName + "'";
        var childFolder = parentFolder.createFolder(childFolderName); // Create child folder

        if (childFolderName === "ROLL FORWARD") {
            var rollForwardSpreadsheetName = parentFolderName + "_ROLL FORWARD" + extension;
            var rollForwardSpreadsheet = SpreadsheetApp.create(rollForwardSpreadsheetName);
            var rollForwardSpreadsheetFile = DriveApp.getFileById(rollForwardSpreadsheet.getId());
            rollForwardSpreadsheetFile.moveTo(childFolder);
            // Write the ID of the created spreadsheet to column d
            sheet.getRange(i + 1, 4).setValue(rollForwardSpreadsheetFile.getId()); //Change range for place in other columns
          }
        }
      
      // Create spreadsheet in parent folder with specified name
      var spreadsheetName = parentFolderName + extension; 
      //var sheetName = "Infraestructura";
      var spreadsheet = SpreadsheetApp.create(spreadsheetName, 100, 50); // Change number of columns and rows
      //var sheetRename = spreadsheet.renameActiveSheet(sheetName);
      //var sheetCreate = spreadsheet.insertSheet();
      var spreadsheetFile = DriveApp.getFileById(spreadsheet.getId());
      spreadsheetFile.moveTo(parentFolder)

      // Write the ID of the created spreadsheet to column C
      sheet.getRange(i + 1, 3).setValue(spreadsheetFile.getId()); //Change range for place in other columns
    }
    msg += " han sido creadas para la carpeta " + parentFolderName;
    logToSheet(nombreHojaLogs, msg);
    msg2 = "El sheet: '" + spreadsheetName;
    if (rollForwardSpreadsheetName) {
        msg2 += " y el sheet: '" + rollForwardSpreadsheetName + "'";
    }
    msg2 += "' se ha creado para la carpeta " + parentFolderName;
    logToSheet(nombreHojaLogs, msg2);
    logToSheet(nombreHojaLogs, "--------------------------------------------------------------------------------------");
  }
  UI.alert("Ejecución de CREAR CARPETAS terminada.")
}

//copia una sheet en diferentes spreatsheets ya creadas cuyos id están en una lista en el spreadsheet sobre el que estamos realizando el código

// function copySheetToMultipleSpreadsheets() {
//   var sourceSpreadsheetId ='1abc0xhSt9XhuhdHSknVGGskl4PLQAp1NobIq7Qo2JxE'; //Cambiar por el id de la página que queremos que sea copiada
//   var sourceSheetName = 'Alcance'; // Cambiar por el nombre de la página que queremos que copie.

//   var sourceSpreadsheet = SpreadsheetApp.openById(sourceSpreadsheetId);
//   var sourceSheet = sourceSpreadsheet.getSheetByName(sourceSheetName);
//   //Toma los valores de los IDs de la hoja actual columna [2] que es la C.
//   var ss = SpreadsheetApp.getActiveSpreadsheet();
//   var sheet = ss.getActiveSheet();
//   var range = sheet.getDataRange();
//   var values = range.getValues();

//   for (var i = 0; i < values.length; i++) {
//     var targetSpreadsheetId = values[i][2];
//     var targetSpreadsheet = SpreadsheetApp.openById(targetSpreadsheetId);
//     var newSheet = sourceSheet.copyTo(targetSpreadsheet);
//     newSheet.setName('Infraestructura'); // Cambiar por el nombre que queramos
//  }
// }

function deleteSheet1FromMultipleSpreadsheets() {
  var sourceSpreadsheetId = '1EKbDjzi3xnONmfL2QHDUYu_JqDi8VcIUUxryU3gWneo'; // ID de la spreatsheet en la que está la lista de id
  var targetColumn = 'C'; // Column containing spreadsheet IDs

  var sourceSpreadsheet = SpreadsheetApp.openById(sourceSpreadsheetId);
  var sourceSheet = sourceSpreadsheet.getSheetByName('Plantilla'); // Nombre de la página en la que están los id
  var lastRow = sourceSheet.getLastRow();
  var spreadsheetIds = sourceSheet.getRange(targetColumn + '36:' + targetColumn + lastRow).getValues(); //Cambiar número por la fila que queremos que empiece
  
  for (var i = 0; i < spreadsheetIds.length; i++) {
    var targetSpreadsheetId = spreadsheetIds[i][0];
    console.log(targetSpreadsheetId);

    var targetSpreadsheet = SpreadsheetApp.openById(targetSpreadsheetId);
    var targetSheet1 = targetSpreadsheet.getSheets()[0]; // Get the first sheet
    targetSpreadsheet.deleteSheet(targetSheet1);
  }
}

function renameSpreadsheets() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet();
  
  var lastRow = sheet.getLastRow();
  
  // Loop through each row
  for (var i = 1; i <= lastRow; i++) {
    var parentName = sheet.getRange(i, 1).getValue().toString().trim();
    var spreadsheetId = sheet.getRange(i, 3).getValue().toString().trim();
    
    if (parentName && spreadsheetId) {
      try {
        // Construct the new name for the spreadsheet
        var newSpreadsheetName = parentName + "_EY EEFF - FASE II 2024";
        
        // Get the spreadsheet by ID
        var spreadsheet = SpreadsheetApp.openById(spreadsheetId);
        
        // Rename the spreadsheet
        spreadsheet.rename(newSpreadsheetName);
        
        Logger.log('Spreadsheet renamed to: ' + newSpreadsheetName);
      } catch (e) {
        Logger.log('Error renaming spreadsheet with ID: ' + spreadsheetId + ' - ' + e.message);
      }
    }
  }
}

function copySheetToMultipleSpreadsheets() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet();
  var data = sheet.getDataRange().getValues();
  var lastRow = sheet.getLastRow();

  // Pedir el ID de la hoja secundaria
  var ui = SpreadsheetApp.getUi();
  var response = ui.prompt('Introduce el ID de la hoja secundaria:');
  var secondarySpreadsheetId = response.getResponseText().trim();
  
  if (!secondarySpreadsheetId) {
    ui.alert('No se ha introducido un ID válido.');
    return;
  }
  
  try {
    var secondarySpreadsheet = SpreadsheetApp.openById(secondarySpreadsheetId);
    var secondarySheet = secondarySpreadsheet.getSheetByName('Infraestructura');
    if (!secondarySheet) {
      ui.alert('No se ha encontrado la hoja "Infraestructura" en la hoja secundaria.');
      return;
    }
    
    // Loop through each row in the current sheet
    for (var i = 1; i <= lastRow; i++) {
      var targetSpreadsheetId = data[i-1][2].toString().trim(); // ID de la columna C
      
      if (targetSpreadsheetId) {
        try {
          var targetSpreadsheet = SpreadsheetApp.openById(targetSpreadsheetId);
          var targetSheet = targetSpreadsheet.getSheetByName('Infraestructura');
          if (targetSheet) {
            targetSpreadsheet.deleteSheet(targetSheet);
          }
          
          // Copiar la hoja "2024 Infraestructura" a la hoja objetivo
          secondarySheet.copyTo(targetSpreadsheet).setName('Infraestructura');
          
        } catch (e) {
          Logger.log('Error processing ID: ' + targetSpreadsheetId + ' - ' + e.message);
        }
      }
    }
  } catch (e) {
    ui.alert('Error al abrir la hoja secundaria: ' + e.message);
  }
}

function filterRowsInTargetSheets() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet();
  var data = sheet.getDataRange().getValues();
  var lastRow = sheet.getLastRow();

  // Loop through each row in the current sheet
  for (var i = 1; i <= lastRow; i++) {
    var keyword = data[i - 1][0].toString().trim(); // Palabra clave de la columna A
    var targetSpreadsheetId = data[i - 1][2].toString().trim(); // ID de la columna C

    if (keyword && targetSpreadsheetId) {
      try {
        var targetSpreadsheet = SpreadsheetApp.openById(targetSpreadsheetId);
        var targetSheet = targetSpreadsheet.getSheetByName('Infraestructura');
        if (targetSheet) {
          var range = targetSheet.getDataRange();
          var values = range.getValues();

          var rowsToDelete = [];
          var rowsToKeep = [];
          var header = values[0]; // Suponemos que la primera fila es el encabezado

          // Mantener el encabezado
          rowsToKeep.push(header);

          // Filtrar las filas que contienen la palabra clave en la columna C
          for (var j = 1; j < values.length; j++) {
            if (values[j][2].toString().includes(keyword)) {
              rowsToKeep.push(values[j]);
            } else {
              rowsToDelete.push(j + 1); // Guardar el índice de la fila (1-based) para eliminar
            }
          }

          // Eliminar filas desde el final para evitar problemas de desplazamiento
          for (var k = rowsToDelete.length - 1; k >= 0; k--) {
            targetSheet.deleteRow(rowsToDelete[k]);
          }

          // Opcional: volver a escribir las filas que se deben mantener
          // targetSheet.clear(); 
          // targetSheet.getRange(1, 1, rowsToKeep.length, rowsToKeep[0].length).setValues(rowsToKeep);

        } else {
          Logger.log('No se encontró la hoja "Infraestructura" en el spreadsheet con ID: ' + targetSpreadsheetId);
        }

      } catch (e) {
        Logger.log('Error procesando ID: ' + targetSpreadsheetId + ' - ' + e.message);
      }
    }
  }
}

function deleteSheetFromSpreadsheets() {
  // Obtén la hoja activa
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  
  // Obtén todos los valores de la columna C (index 3)
  var spreadsheetIds = sheet.getRange("C2:C" + sheet.getLastRow()).getValues();
  
  // Itera a través de cada ID de la hoja de cálculo
  for (var i = 0; i < spreadsheetIds.length; i++) {
    var id = spreadsheetIds[i][0];
    if (id) {  // Si hay un ID válido
      try {
        var spreadsheet = SpreadsheetApp.openById(id);
        var sheetToDelete = spreadsheet.getSheetByName('Hoja 1');
        if (sheetToDelete) {
          spreadsheet.deleteSheet(sheetToDelete);
          Logger.log('Hoja 1 eliminada de la hoja de cálculo con ID: ' + id);
        } else {
          Logger.log('Hoja 1 no encontrada en la hoja de cálculo con ID: ' + id);
        }
      } catch (e) {
        Logger.log('Error al abrir o modificar la hoja de cálculo con ID: ' + id + ' - ' + e.message);
      }
    }
  }
}

function copyContentWithStyles() {
  // Obtener la hoja activa y los datos de la columna C
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var activeSheet = ss.getActiveSheet();
  var ids = activeSheet.getRange("C:C").getValues().flat().filter(String); // Obtener todos los valores no vacíos de la columna C

  // Crear una nueva pestaña
  var newSheet = ss.insertSheet("Consolidado");

  // Definir la fila inicial en la nueva pestaña
  var newSheetRow = 1;

  // Iterar a través de los IDs
  ids.forEach(function(id) {
    try {
      // Abrir la hoja de cálculo por ID
      var externalSs = SpreadsheetApp.openById(id);
      var sheet = externalSs.getSheetByName("Infraestructura");

      if (sheet) {
        // Obtener el rango con datos (menos la primera fila)
        var lastRow = sheet.getLastRow();
        if (lastRow > 1) {
          var dataRange = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn());
          var data = dataRange.getValues();
          var formats = dataRange.getNumberFormats();
          var backgrounds = dataRange.getBackgrounds();
          var fontColors = dataRange.getFontColors();
          var fontFamilies = dataRange.getFontFamilies();
          var fontSizes = dataRange.getFontSizes();
          var fontWeights = dataRange.getFontWeights();
          var fontLines = dataRange.getFontLines();
          var textStyles = dataRange.getTextStyles();
          var horizontalAlignments = dataRange.getHorizontalAlignments();
          var verticalAlignments = dataRange.getVerticalAlignments();
          
          // Obtener la última fila con datos en la nueva pestaña
          var lastNewSheetRow = newSheet.getLastRow();
          
          // Calcular la fila inicial en la nueva pestaña para pegar los datos
          if (lastNewSheetRow > 0) {
            newSheetRow = lastNewSheetRow + 1;
          }
          
          // Pegarlo en la nueva pestaña
          var targetRange = newSheet.getRange(newSheetRow, 1, data.length, data[0].length);
          targetRange.setValues(data);
          targetRange.setNumberFormats(formats);
          targetRange.setBackgrounds(backgrounds);
          targetRange.setFontColors(fontColors);
          targetRange.setFontFamilies(fontFamilies);
          targetRange.setFontSizes(fontSizes);
          targetRange.setFontWeights(fontWeights);
          targetRange.setFontLines(fontLines);
          targetRange.setTextStyles(textStyles);
          targetRange.setHorizontalAlignments(horizontalAlignments);
          targetRange.setVerticalAlignments(verticalAlignments);
        }
      } else {
        Logger.log("Hoja 'Infraestructura' no encontrada en el archivo con ID: " + id);
      }
    } catch (e) {
      Logger.log("Error al abrir la hoja de cálculo con ID: " + id + ". Error: " + e.message);
    }
  });

  Logger.log("Proceso completado");
}

function copyContentWithStylesBordes() {
  // Pedir el ID de la hoja de cálculo de destino
  var ui = SpreadsheetApp.getUi();
  var response = ui.prompt('Por favor, ingresa el ID de la hoja de cálculo de destino');
  if (response.getSelectedButton() != ui.Button.OK) {
    ui.alert('Proceso cancelado.');
    return;
  }
  var destSpreadsheetId = response.getResponseText();
  
  // Obtener la hoja activa y los datos de la columna C
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var activeSheet = ss.getActiveSheet();
  var ids = activeSheet.getRange("C:C").getValues().flat().filter(String); // Obtener todos los valores no vacíos de la columna C

  if (ids.length === 0) {
    ui.alert('No se encontraron IDs en la columna C.');
    return;
  }

  // Abrir la hoja de cálculo de destino y crear una nueva pestaña
  var destSpreadsheet = SpreadsheetApp.openById(destSpreadsheetId);
  var newSheet = destSpreadsheet.insertSheet("Consolidado");
  var firstCopy = true;
  var newSheetRow = 2; // Empezamos desde la fila 2 porque la fila 1 será la cabecera

  // Iterar a través de los IDs
  ids.forEach(function(id, index) {
    try {
      // Abrir la hoja de cálculo por ID
      var externalSs = SpreadsheetApp.openById(id);
      var sheet = externalSs.getSheetByName("Infraestructura");

      if (sheet) {
        // Determinar el rango a copiar
        var startRow = firstCopy ? 1 : 2;
        var numRows = sheet.getLastRow() - startRow + 1;
        var numCols = sheet.getLastColumn();
        
        if (numRows > 0) {
          var dataRange = sheet.getRange(startRow, 1, numRows, numCols);
          var data = dataRange.getValues();
          var formats = dataRange.getNumberFormats();
          var backgrounds = dataRange.getBackgrounds();
          var fontColors = dataRange.getFontColors();
          var fontFamilies = dataRange.getFontFamilies();
          var fontSizes = dataRange.getFontSizes();
          var fontWeights = dataRange.getFontWeights();
          var fontLines = dataRange.getFontLines();
          var textStyles = dataRange.getTextStyles();
          var horizontalAlignments = dataRange.getHorizontalAlignments();
          var verticalAlignments = dataRange.getVerticalAlignments();

          // Calcular la fila inicial en la nueva pestaña para pegar los datos
          var lastNewSheetRow = newSheet.getLastRow();
          if (lastNewSheetRow > 0 && newSheetRow == 2) {
            newSheetRow = lastNewSheetRow + 1;
          }

          // Pegarlo en la nueva pestaña
          var targetRange = newSheet.getRange(newSheetRow, 1, data.length, data[0].length);
          targetRange.setValues(data);
          targetRange.setNumberFormats(formats);
          targetRange.setBackgrounds(backgrounds);
          targetRange.setFontColors(fontColors);
          targetRange.setFontFamilies(fontFamilies);
          targetRange.setFontSizes(fontSizes);
          targetRange.setFontWeights(fontWeights);
          targetRange.setFontLines(fontLines);
          targetRange.setTextStyles(textStyles);
          targetRange.setHorizontalAlignments(horizontalAlignments);
          targetRange.setVerticalAlignments(verticalAlignments);

          // Actualizar la fila inicial para la siguiente iteración
          newSheetRow += data.length;

          if (firstCopy) {
            // Copiar la primera fila del primer ID a la fila 1 de la nueva pestaña
            var headerRange = sheet.getRange(1, 1, 1, numCols);
            var headerData = headerRange.getValues();
            var headerFormats = headerRange.getNumberFormats();
            var headerBackgrounds = headerRange.getBackgrounds();
            var headerFontColors = headerRange.getFontColors();
            var headerFontFamilies = headerRange.getFontFamilies();
            var headerFontSizes = headerRange.getFontSizes();
            var headerFontWeights = headerRange.getFontWeights();
            var headerFontLines = headerRange.getFontLines();
            var headerTextStyles = headerRange.getTextStyles();
            var headerHorizontalAlignments = headerRange.getHorizontalAlignments();
            var headerVerticalAlignments = headerRange.getVerticalAlignments();

            var targetHeaderRange = newSheet.getRange(1, 1, 1, numCols);
            targetHeaderRange.setValues(headerData);
            targetHeaderRange.setNumberFormats(headerFormats);
            targetHeaderRange.setBackgrounds(headerBackgrounds);
            targetHeaderRange.setFontColors(headerFontColors);
            targetHeaderRange.setFontFamilies(headerFontFamilies);
            targetHeaderRange.setFontSizes(headerFontSizes);
            targetHeaderRange.setFontWeights(headerFontWeights);
            targetHeaderRange.setFontLines(headerFontLines);
            targetHeaderRange.setTextStyles(headerTextStyles);
            targetHeaderRange.setHorizontalAlignments(headerHorizontalAlignments);
            targetHeaderRange.setVerticalAlignments(headerVerticalAlignments);
          }

          // Marcar que ya se ha realizado la primera copia
          firstCopy = false;
        }
      } else {
        Logger.log("Hoja 'Infraestructura' no encontrada en el archivo con ID: " + id);
      }
    } catch (e) {
      Logger.log("Error al abrir la hoja de cálculo con ID: " + id + ". Error: " + e.message);
    }
  });

  // Aplicar bordes a todas las celdas con contenido en la nueva pestaña
  var lastRow = newSheet.getLastRow();
  var lastColumn = newSheet.getLastColumn();
  if (lastRow > 0 && lastColumn > 0) {
    var rangeWithContent = newSheet.getRange(1, 1, lastRow, lastColumn);
    rangeWithContent.setBorder(true, true, true, true, true, true);
  }

  Logger.log("Proceso completado");
}

function copyContentWithStylesAndRename() {
  // Pedir el ID de la hoja de cálculo de destino
  var ui = SpreadsheetApp.getUi();
  var response = ui.prompt('Por favor, ingresa el ID de la hoja de cálculo de destino');
  if (response.getSelectedButton() != ui.Button.OK) {
    ui.alert('Proceso cancelado.');
    return;
  }
  var destSpreadsheetId = response.getResponseText();
  
  // Obtener la hoja activa y los datos de la columna C
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var activeSheet = ss.getActiveSheet();
  var ids = activeSheet.getRange("C:C").getValues().flat().filter(String); // Obtener todos los valores no vacíos de la columna C

  if (ids.length === 0) {
    ui.alert('No se encontraron IDs en la columna C.');
    return;
  }

  // Abrir la hoja de cálculo de destino y crear una nueva pestaña
  var destSpreadsheet = SpreadsheetApp.openById(destSpreadsheetId);
  var newSheet = destSpreadsheet.insertSheet("Consolidado");
  var firstCopy = true;
  var newSheetRow = 2; // Empezamos desde la fila 2 porque la fila 1 será la cabecera

  // Iterar a través de los IDs
  ids.forEach(function(id) {
    try {
      // Abrir la hoja de cálculo por ID
      var externalSs = SpreadsheetApp.openById(id);
      var sheet = externalSs.getSheetByName("Infraestructura");

      if (sheet) {
        // Determinar el rango a copiar
        var startRow = firstCopy ? 1 : 2;
        var numRows = sheet.getLastRow() - startRow + 1;
        var numCols = sheet.getLastColumn();
        
        if (numRows > 0) {
          var dataRange = sheet.getRange(startRow, 1, numRows, numCols);
          var data = dataRange.getValues();
          var formats = dataRange.getNumberFormats();
          var backgrounds = dataRange.getBackgrounds();
          var fontColors = dataRange.getFontColors();
          var fontFamilies = dataRange.getFontFamilies();
          var fontSizes = dataRange.getFontSizes();
          var fontWeights = dataRange.getFontWeights();
          var fontLines = dataRange.getFontLines();
          var textStyles = dataRange.getTextStyles();
          var horizontalAlignments = dataRange.getHorizontalAlignments();
          var verticalAlignments = dataRange.getVerticalAlignments();

          // Calcular la fila inicial en la nueva pestaña para pegar los datos
          var lastNewSheetRow = newSheet.getLastRow();
          if (lastNewSheetRow > 0 && newSheetRow == 2) {
            newSheetRow = lastNewSheetRow + 1;
          }

          // Pegarlo en la nueva pestaña
          var targetRange = newSheet.getRange(newSheetRow, 1, data.length, data[0].length);
          targetRange.setValues(data);
          targetRange.setNumberFormats(formats);
          targetRange.setBackgrounds(backgrounds);
          targetRange.setFontColors(fontColors);
          targetRange.setFontFamilies(fontFamilies);
          targetRange.setFontSizes(fontSizes);
          targetRange.setFontWeights(fontWeights);
          targetRange.setFontLines(fontLines);
          targetRange.setTextStyles(textStyles);
          targetRange.setHorizontalAlignments(horizontalAlignments);
          targetRange.setVerticalAlignments(verticalAlignments);

          // Actualizar la fila inicial para la siguiente iteración
          newSheetRow += data.length;

          if (firstCopy) {
            // Copiar la primera fila del primer ID a la fila 1 de la nueva pestaña
            var headerRange = sheet.getRange(1, 1, 1, numCols);
            var headerData = headerRange.getValues();
            var headerFormats = headerRange.getNumberFormats();
            var headerBackgrounds = headerRange.getBackgrounds();
            var headerFontColors = headerRange.getFontColors();
            var headerFontFamilies = headerRange.getFontFamilies();
            var headerFontSizes = headerRange.getFontSizes();
            var headerFontWeights = headerRange.getFontWeights();
            var headerFontLines = headerRange.getFontLines();
            var headerTextStyles = headerRange.getTextStyles();
            var headerHorizontalAlignments = headerRange.getHorizontalAlignments();
            var headerVerticalAlignments = headerRange.getVerticalAlignments();

            var targetHeaderRange = newSheet.getRange(1, 1, 1, numCols);
            targetHeaderRange.setValues(headerData);
            targetHeaderRange.setNumberFormats(headerFormats);
            targetHeaderRange.setBackgrounds(headerBackgrounds);
            targetHeaderRange.setFontColors(headerFontColors);
            targetHeaderRange.setFontFamilies(headerFontFamilies);
            targetHeaderRange.setFontSizes(headerFontSizes);
            targetHeaderRange.setFontWeights(headerFontWeights);
            targetHeaderRange.setFontLines(headerFontLines);
            targetHeaderRange.setTextStyles(headerTextStyles);
            targetHeaderRange.setHorizontalAlignments(headerHorizontalAlignments);
            targetHeaderRange.setVerticalAlignments(headerVerticalAlignments);

            // Marcar que ya se ha realizado la primera copia
            firstCopy = false;
          }
        }
      } else {
        Logger.log("Hoja 'Infraestructura' no encontrada en el archivo con ID: " + id);
      }
    } catch (e) {
      Logger.log("Error al abrir la hoja de cálculo con ID: " + id + ". Error: " + e.message);
    }
  });

  // Aplicar bordes a todas las celdas con contenido en la nueva pestaña
  var lastRow = newSheet.getLastRow();
  var lastColumn = newSheet.getLastColumn();
  if (lastRow > 0 && lastColumn > 0) {
    var rangeWithContent = newSheet.getRange(1, 1, lastRow, lastColumn);
    rangeWithContent.setBorder(true, true, true, true, true, true);
  }

  // Eliminar "Hoja 1" si existe
  var sheetToDelete = destSpreadsheet.getSheetByName("Hoja 1");
  if (sheetToDelete) {
    destSpreadsheet.deleteSheet(sheetToDelete);
  }

  // Renombrar la nueva pestaña como "Infraestructura"
  newSheet.setName("Infraestructura");

  Logger.log("Proceso completado");
}

function copyContentWithStylesAndRenameEncabezado() {
  // Pedir el ID de la hoja de cálculo de destino
  var ui = SpreadsheetApp.getUi();
  var response = ui.prompt('Por favor, ingresa el ID de la hoja de cálculo de destino');
  if (response.getSelectedButton() != ui.Button.OK) {
    ui.alert('Proceso cancelado.');
    return;
  }
  var destSpreadsheetId = response.getResponseText();
  
  // Obtener la hoja activa y los datos de la columna C
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var activeSheet = ss.getActiveSheet();
  var ids = activeSheet.getRange("C:C").getValues().flat().filter(String); // Obtener todos los valores no vacíos de la columna C

  if (ids.length === 0) {
    ui.alert('No se encontraron IDs en la columna C.');
    return;
  }

  // Abrir la hoja de cálculo de destino y crear una nueva pestaña
  var destSpreadsheet = SpreadsheetApp.openById(destSpreadsheetId);
  var newSheet = destSpreadsheet.insertSheet("Consolidado");
  var firstCopy = true;
  var newSheetRow = 2; // Empezamos desde la fila 2 porque la fila 1 será la cabecera

  // Iterar a través de los IDs
  ids.forEach(function(id) {
    try {
      // Abrir la hoja de cálculo por ID
      var externalSs = SpreadsheetApp.openById(id);
      var sheet = externalSs.getSheetByName("Infraestructura");

      if (sheet) {
        // Determinar el rango a copiar
        var startRow = firstCopy ? 1 : 2;
        var numRows = sheet.getLastRow() - startRow + 1;
        var numCols = sheet.getLastColumn();
        
        if (numRows > 0) {
          var dataRange = sheet.getRange(startRow, 1, numRows, numCols);
          var data = dataRange.getValues();
          var formats = dataRange.getNumberFormats();
          var backgrounds = dataRange.getBackgrounds();
          var fontColors = dataRange.getFontColors();
          var fontFamilies = dataRange.getFontFamilies();
          var fontSizes = dataRange.getFontSizes();
          var fontWeights = dataRange.getFontWeights();
          var fontLines = dataRange.getFontLines();
          var textStyles = dataRange.getTextStyles();
          var horizontalAlignments = dataRange.getHorizontalAlignments();
          var verticalAlignments = dataRange.getVerticalAlignments();

          // Calcular la fila inicial en la nueva pestaña para pegar los datos
          var lastNewSheetRow = newSheet.getLastRow();
          if (lastNewSheetRow > 0 && newSheetRow == 2) {
            newSheetRow = lastNewSheetRow + 1;
          }

          // Pegarlo en la nueva pestaña
          var targetRange = newSheet.getRange(newSheetRow, 1, data.length, data[0].length);
          targetRange.setValues(data);
          targetRange.setNumberFormats(formats);
          targetRange.setBackgrounds(backgrounds);
          targetRange.setFontColors(fontColors);
          targetRange.setFontFamilies(fontFamilies);
          targetRange.setFontSizes(fontSizes);
          targetRange.setFontWeights(fontWeights);
          targetRange.setFontLines(fontLines);
          targetRange.setTextStyles(textStyles);
          targetRange.setHorizontalAlignments(horizontalAlignments);
          targetRange.setVerticalAlignments(verticalAlignments);

          // Actualizar la fila inicial para la siguiente iteración
          newSheetRow += data.length;

          if (firstCopy) {
            // Copiar la primera fila del primer ID a la fila 1 de la nueva pestaña
            var headerRange = sheet.getRange(1, 1, 1, numCols);
            var headerData = headerRange.getValues();
            var headerFormats = headerRange.getNumberFormats();
            var headerBackgrounds = headerRange.getBackgrounds();
            var headerFontColors = headerRange.getFontColors();
            var headerFontFamilies = headerRange.getFontFamilies();
            var headerFontSizes = headerRange.getFontSizes();
            var headerFontWeights = headerRange.getFontWeights();
            var headerFontLines = headerRange.getFontLines();
            var headerTextStyles = headerRange.getTextStyles();
            var headerHorizontalAlignments = headerRange.getHorizontalAlignments();
            var headerVerticalAlignments = headerRange.getVerticalAlignments();

            var targetHeaderRange = newSheet.getRange(1, 1, 1, numCols);
            targetHeaderRange.setValues(headerData);
            targetHeaderRange.setNumberFormats(headerFormats);
            targetHeaderRange.setBackgrounds(headerBackgrounds);
            targetHeaderRange.setFontColors(headerFontColors);
            targetHeaderRange.setFontFamilies(headerFontFamilies);
            targetHeaderRange.setFontSizes(headerFontSizes);
            targetHeaderRange.setFontWeights(headerFontWeights);
            targetHeaderRange.setFontLines(headerFontLines);
            targetHeaderRange.setTextStyles(headerTextStyles);
            targetHeaderRange.setHorizontalAlignments(headerHorizontalAlignments);
            targetHeaderRange.setVerticalAlignments(headerVerticalAlignments);

            // Marcar que ya se ha realizado la primera copia
            firstCopy = false;
          }
        }
      } else {
        Logger.log("Hoja 'Infraestructura' no encontrada en el archivo con ID: " + id);
      }
    } catch (e) {
      Logger.log("Error al abrir la hoja de cálculo con ID: " + id + ". Error: " + e.message);
    }
  });

  // Eliminar la fila 1 de la nueva pestaña
  newSheet.deleteRow(1);

  var newSheet = destSpreadsheet.getActiveSheet();
  newSheet.getRange(1, 1, newSheet.getMaxRows(), newSheet.getMaxColumns()).activate();
  destSpreadsheet.getActiveRangeList().setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);
  destSpreadsheet.getActiveSheet().setRowHeights(1, 999, 104);
  destSpreadsheet.getRange('1:1').activate();
  destSpreadsheet.getActiveSheet().setColumnWidth(6, 302);
  destSpreadsheet.getActiveSheet().setColumnWidth(21, 135);

  // Aplicar bordes a todas las celdas con contenido en la nueva pestaña
  var lastRow = newSheet.getLastRow();
  var lastColumn = newSheet.getLastColumn();
  if (lastRow > 0 && lastColumn > 0) {
    var rangeWithContent = newSheet.getRange(1, 1, lastRow, lastColumn);
    rangeWithContent.setBorder(true, true, true, true, true, true);
  }

  // Eliminar "Hoja 1" si existe
  var sheetToDelete = destSpreadsheet.getSheetByName("Hoja 1");
  if (sheetToDelete) {
    destSpreadsheet.deleteSheet(sheetToDelete);
  }

  // Renombrar la nueva pestaña como "Infraestructura"
  newSheet.setName("Infraestructura");



  Logger.log("Proceso completado");
}

function createAndRenameSheet() {
  // Pedir el ID de la hoja de cálculo de destino
  var ui = SpreadsheetApp.getUi();
  var response = ui.prompt('Por favor, ingresa el ID de la hoja de cálculo de destino');
  if (response.getSelectedButton() != ui.Button.OK) {
    ui.alert('Proceso cancelado.');
    return;
  }
  var destSpreadsheetId = response.getResponseText();
  
  // Abrir la hoja de cálculo de destino
  var destSpreadsheet = SpreadsheetApp.openById(destSpreadsheetId);
  
  // Eliminar la hoja "Consolidado" si existe
  var existingSheet = destSpreadsheet.getSheetByName("Consolidado");
  if (existingSheet) {
    destSpreadsheet.deleteSheet(existingSheet);
  }
  
  // Crear una nueva hoja llamada "Consolidado"
  var newSheet = destSpreadsheet.insertSheet("Consolidado");
  
  // Cambiar el nombre de la hoja a "Infraestructura compilada 2024"
  newSheet.setName("Infraestructura compilada 2024");
  
  // Renombrar la hoja "Infraestructura compilada 2024" a "Infraestructura compilada Fase II 2024"
  var renamedSheet = destSpreadsheet.getSheetByName("Infraestructura compilada 2024");
  if (renamedSheet) {
    renamedSheet.setName("Infraestructura compilada Fase II 2024");
  }
  
  Logger.log("Proceso completado");
}



