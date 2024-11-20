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
