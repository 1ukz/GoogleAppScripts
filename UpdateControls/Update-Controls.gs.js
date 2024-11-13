// CREATOR: LUCAS SAETA (@1ukz) lucassaeta5@gmail.com

/**
 * Esta función se ejecuta al abrir la hoja de cálculo. 
 * Crea un menú personalizado llamado 'Update-Controls' en la interfaz de usuario de Google Sheets.
 * Al seleccionar 'Ejecutar programa' en el menú, se invoca la función 'main'.
 */
function onOpen() {
  const UI = SpreadsheetApp.getUi(); 

  UI.createMenu('Update-Controls')
    .addItem('Ejecutar programa', 'main')
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

/** FUNCIÓN AUXILIAR
 * Verifica si una hoja con el nombre especificado existe en la hoja de cálculo activa.
 *
 * @param {string} sheetName - El nombre de la hoja que se va a verificar.
 * @return {boolean} - Retorna true si la hoja existe, de lo contrario false.
 */
function checkSheetExists(sheetName) {
  var sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
  for (var i = 0; i < sheets.length; i++) {
    if (sheets[i].getName() === sheetName) {
      return true; // La hoja existe
    }
  }
  return false; // La hoja no existe
}

/** FUNCIÓN AUXILIAR
 * Solicita al usuario un nombre de hoja a través de un cuadro de diálogo.
 * Si el usuario cierra el diálogo, se retorna null. 
 * Si el nombre de la hoja no existe, se solicita nuevamente.
 *
 * @param {string} mensaje - El mensaje que se mostrará en el cuadro de diálogo.
 * @return {string|null} - Retorna el nombre de la hoja si existe, o null si se cierra el diálogo.
 */
function solicitarNombreHoja(mensaje) {
  const UI = SpreadsheetApp.getUi()
  var response;
  do {
    response = UI.prompt(mensaje);
    if (response.getSelectedButton() === UI.Button.CLOSE) {
      return null; // Si el usuario cierra el diálogo, retorna null
    }
    var nombreHoja = response.getResponseText().trim();
    if (!checkSheetExists(nombreHoja)) {
      UI.alert('La hoja "' + response.getResponseText() + '" no existe. Por favor, inténtelo de nuevo.');
    }
  } while (!checkSheetExists(nombreHoja));
  return nombreHoja; // Retorna el nombre de la hoja si existe
}

/** FUNCIÓN AUXILIAR
 * Muestra un cuadro de diálogo (pop-up) al usuario para preguntar una respuesta de sí o no.
 * El usuario debe ingresar 'y' para sí o 'n' para no. 
 * Si se ingresa una respuesta no válida, se solicita nuevamente.
 *
 * @param {string} mensaje - El mensaje que se mostrará en el pop-up.
 * @return {boolean} - Retorna true si la respuesta es 'y', o false si la respuesta es 'n'.
 */
function preguntaMenu(mensaje) {

  const UI = SpreadsheetApp.getUi(); 
  var bool;
  var bucle = false;
  do {
    var resp = UI.prompt(mensaje);
    if (resp.getResponseText().toLowerCase().trim() === 'y') {
      bool = true;
      bucle = true;
    } else if (resp.getResponseText().toLowerCase().trim() === 'n') {
      bool = false;
      bucle = true;
    } else {
      UI.alert('Respuesta no válida. Por favor, introduzca "y"  o "n".');
    }
  } while (!bucle);
  return bool;

}

/** FUNCIÓN AUXILIAR
 * Verifica si la hoja del control cumple con el formato estándar requerido.
 * Si no cumple, actualiza el formato y registra una advertencia en la hoja de logs.
 *
 * @param {Object} hoja - La hoja de cálculo que se va a verificar.
 * @param {string} nombreHojaLogs - El nombre de la hoja de logs donde se registrará la advertencia.
 * @param {string} message - Mensaje descriptivo para el registro.
 */
function checkFormat(hoja, nombreHojaLogs, message){
  //Si no cumple todas las especificaciones del formato estándar: 
  if(!(hoja.getRange('A1:F1').getValue().includes('DOCUMENTACIÓN DEL CONTROL') && hoja.getRange('A2').getValue().includes('Tipo de Control') && hoja.getRange('A6').getValue().includes('Descripción') && hoja.getRange('A8').getValue().includes('Evidencia') && hoja.getRange('A11:F11').getValue().includes('DESCRIPCIÓN DE LA PRUEBA A EJECUTAR') && hoja.getRange('A12').getValue().includes('Prueba a realizar') && hoja.getRange('E14').getValue().includes('Tamaño Muestra'))){
    
    //Actualiza la hoja del control
    updateFormat(hoja); 
    logToSheet(nombreHojaLogs, 'ADVERTENCIA: Se ha actualizado el formato de la ' + message + ': "' + hoja.getName() + '" ya que no seguía el formato estándar.')
  }
}

/** FUNCIÓN AUXILIAR
 * Actualiza el formato de la hoja según ciertas condiciones.
 * Elimina filas o inserta nuevas filas según el contenido de las celdas.
 *
 * @param {Object} hoja - La hoja de cálculo que se va a actualizar.
 */
function updateFormat(hoja){
  const ID = SpreadsheetApp.getActiveSpreadsheet().getId();
  const MAIN_SHEET = SpreadsheetApp.openById(ID);
  const UI = SpreadsheetApp.getUi();
  
  //Comparación fila extra arriba
  if(hoja.getRange('A3').getValue() === 'Tipo de Control' || hoja.getRange('A3').getValue() === 'Clase'){

    if(hoja.getRange('A2:F2').getValue().includes('DOCUMENTACIÓN DEL CONTROL')){
      hoja.deleteRow('1');
    }
  }

  //Comparación fila anterior a descripción prubea a ejecutar
  if(hoja.getRange('A10:F10').getValue().includes('DESCRIPCIÓN DE LA PRUEBA A EJECUTAR') && hoja.getRange('A9').getValue().includes('Evidencia. Actualizaciones')){

    hoja.insertRowBefore(10);
    hoja.getRange('A10:F10').setBackground(null);  // Restablece el fondo de la nueva fila en blanco
    hoja.getRange('A10:F10').mergeAcross();  // Fusiona las celdas en la nueva fila
  }

}

/** FUNCIÓN AUXILIAR
 * Verifica el contenido de celdas específicas en la hoja de destino (control documentado).
 * Registra advertencias en la hoja de logs si las celdas no cumplen con lo esperado.
 *
 * @param {Object} hojaDestino - La hoja de cálculo que se va a verificar.
 * @param {string} nombreHojaLogs - El nombre de la hoja de logs donde se registrarán las advertencias.
 */
function verificarCeldas(hojaDestino, nombreHojaLogs) {

  var celdasAComprobar = [
    'B13:F13', 'B9:F9', 'B7:F7', 'A5', 'B5', 'C5', 'D5', 'E5', 'F5'
  ];

  var celdasNoVacias = [
    'B6:F6', 'B8:F8', 'B12:F12', 'E14', 'F14'
  ];

  // Verificar celdas que deben estar vacías
  for (var i = 0; i < celdasAComprobar.length; i++) {
    var rango = hojaDestino.getRange(celdasAComprobar[i]);
    var valores = rango.getValues();
    var hayValores = false;

    // Iterar por las celdas para verificar si alguna contiene datos
    for (var k = 0; k < valores.length; k++) {
      for (var l = 0; l < valores[k].length; l++) {
        if (valores[k][l].toString().trim() !== '') { // Si la celda no está vacía
          hayValores = true;
          break; // Salir del bucle interno
        }
      }
      if (hayValores) {
        break; // Salir del bucle externo
      }
    }

    // Si hay valores en las celdas que deberían estar vacías
    if (hayValores && hojaDestino.getRange('A12').getValue() === 'Prueba a realizar') {
      logToSheet(nombreHojaLogs,'ADVERTENCIA: El campo "' + celdasAComprobar[i] + '" NO está vacío cuando debería estarlo.');
    }
  }

  // Verificar celdas que NO deben estar vacías
  for (var j = 0; j < celdasNoVacias.length; j++) {
    var rangoNoVacio = hojaDestino.getRange(celdasNoVacias[j]);
    var valoresNoVacios = rangoNoVacio.getValues();
    var hayValoresNoVacios = false;

    // Iterar por las celdas para verificar si están vacías
    for (var m = 0; m < valoresNoVacios.length; m++) {
      for (var n = 0; n < valoresNoVacios[m].length; n++) {
        if (valoresNoVacios[m][n].toString().trim() !== '') { // Si la celda no está vacía
          hayValoresNoVacios = true;
          break; // Salir del bucle interno
        }
      }
      if (hayValoresNoVacios) {
        break; // Salir del bucle externo
      }
    }

    // Si NO hay valores en las celdas que deberían tener datos
    if (!hayValoresNoVacios && hojaDestino.getRange('A12').getValue() === 'Prueba a realizar') {
      logToSheet(nombreHojaLogs,'ADVERTENCIA: El campo "' + celdasNoVacias[j] + '" está vacío, pero debería tener datos.');
    }
  }

  // Verificar si la hoja tiene más de 14 filas, lo que indicaría que F15 existe
  var totalFilas = hojaDestino.getMaxRows();
  if (totalFilas > 14 && hojaDestino.getRange('E14').getValue() === 'Tamaño Muestra.') {
    logToSheet(nombreHojaLogs, 'ADVERTENCIA: Existe una o unas filas adicionales que no deberían de estar presentes.');
  }
}

/** FUNCIÓN AUXILIAR
 * Almacena los IDs de los controles en una hoja de cálculo a partir de los archivos en carpetas en una carpeta de Google Drive.
 * Crea o limpia una hoja especificada y registra los IDs de los archivos de Google Sheets y Excel.
 * Contiene también una función auxiliar dentro de la función para quitar lo que sobre al nombre del
 control (cleanControlName).
 * 
 * @param {string} folderId - El ID de la carpeta en Google Drive que contiene las carpetas con    
 archivos.
 * @param {string} sheetName - El nombre de la hoja donde se almacenarán los IDs.
 * @param {string} nombreHojaLogs - El nombre de la hoja de logs donde se registrarán las acciones realizadas.
 */
function almacenarIDs(folderId, sheetName, nombreHojaLogs) {
  const ID = SpreadsheetApp.getActiveSpreadsheet().getId();
  const MAIN_SHEET = SpreadsheetApp.openById(ID);
  const UI = SpreadsheetApp.getUi(); 

  var newSheet = MAIN_SHEET.getSheetByName(sheetName);
  UI.alert('Capturando los IDs de los Controles en: "' + folderId + '" en la Hoja: "' + sheetName + '". \n\nPOR FAVOR, ESPERE PACIENTEMENTE HASTA EL PRÓXIMO POP-UP QUE INDICANDO QUE EL PROCESO HA TERMINADO. \n\nSe ha generado una Hoja con los Logs de lo realizado para esta ejecución, disponible en: "' + nombreHojaLogs + '"');
  if (!newSheet) {
    newSheet = MAIN_SHEET.insertSheet(sheetName);
  } else {
    newSheet.clear();
  }

  var folder = DriveApp.getFolderById(folderId);
  var folderName = folder.getName();

  newSheet.getRange(1, 1).setValue(folderName);
  newSheet.getRange(2, 1).setValue("ID de control");
  newSheet.getRange(2, 2).setValue("Hoja de control");
  newSheet.getRange(2, 3).setValue("Nombre de Sheet");

  var row = 3;

  function cleanControlName(name) {
    var patterns = ["_PASA", "_FALLA", "_INCONCLUSO"];
    for (var i = 0; i < patterns.length; i++) {
      var index = name.indexOf(patterns[i]);
      if (index !== -1) {
        return name.substring(0, index);
      }
    }
    return name;
  }

  var subfolders = folder.getFolders();

  while (subfolders.hasNext()) {
    var subfolder = subfolders.next();
    var subfolderId = subfolder.getId();
    var subfolderName = cleanControlName(subfolder.getName());

    // Obtener todos los archivos de Google Sheets y Excel
    var files = subfolder.getFiles();
    while (files.hasNext()) {
      var file = files.next();
      var mimeType = file.getMimeType();

      if (mimeType === MimeType.GOOGLE_SHEETS || mimeType === MimeType.MICROSOFT_EXCEL) {
        var fileId;
        var spreadsheet;

        if (mimeType === MimeType.MICROSOFT_EXCEL) {

          logToSheet(nombreHojaLogs, 'ERROR. El control: "' + subfolderName + '" tiene formato EXCEL y no se puede extraer su ID.');
        } else {
          fileId = file.getId();
          spreadsheet = SpreadsheetApp.openById(fileId);
        

        var firstSheetName = spreadsheet.getSheets()[0].getName();

        newSheet.getRange(row, 1).setValue(fileId);
        newSheet.getRange(row, 2).setValue(firstSheetName);
        newSheet.getRange(row, 3).setValue(subfolderName);
        logToSheet(nombreHojaLogs, 'Se ha copiado el ID del control: "' + subfolderName + '" correctamente.');
        row++;
        }
      }
    }
  }
}

/** FUNCIÓN AUXILIAR
 * Almacena los IDs de los documentos en una hoja de cálculo a partir de los archivos en una carpeta de Google Drive.
 * Crea o limpia una hoja especificada y registra los IDs de los archivos de Google Sheets y Excel.
 *
 * @param {string} folderId - El ID de la carpeta en Google Drive que contiene los archivos.
 * @param {string} sheetName - El nombre de la hoja donde se almacenarán los IDs.
 * @param {string} nombreHojaLogs - El nombre de la hoja de logs donde se registrarán las acciones realizadas.
 */
function almacenarIDsDocus(folderId, sheetName, nombreHojaLogs) {
  const ID = SpreadsheetApp.getActiveSpreadsheet().getId();
  const MAIN_SHEET = SpreadsheetApp.openById(ID);
  const UI = SpreadsheetApp.getUi(); 

  UI.alert('Capturando los IDs de los Controles en: "' + folderId + '" en la Hoja: "' + sheetName + '". \n\nPOR FAVOR, ESPERE PACIENTEMENTE HASTA EL PRÓXIMO POP-UP INDICANDO QUE EL PROCESO HA TERMINADO. \n\nSe ha generado una Hoja con los Logs de lo realizado para esta ejecución, disponible en: "' + nombreHojaLogs + '"');
  var newSheet = MAIN_SHEET.getSheetByName(sheetName);
  
  if (!newSheet) {
    newSheet = MAIN_SHEET.insertSheet(sheetName);
  } else {
    newSheet.clear();
  }

  var folder = DriveApp.getFolderById(folderId);
  var folderName = folder.getName();

  newSheet.getRange(1, 1).setValue(folderName);
  newSheet.getRange(2, 1).setValue("ID de control");
  newSheet.getRange(2, 2).setValue("Hoja de control");
  newSheet.getRange(2, 3).setValue("Nombre de Sheet");

  var row = 3;

  // Obtener todos los archivos de Google Sheets y Excel
  var files = folder.getFiles();
  while (files.hasNext()) {
    var file = files.next();
    var mimeType = file.getMimeType();
    
    if (mimeType === MimeType.GOOGLE_SHEETS || mimeType === MimeType.MICROSOFT_EXCEL) {
      var fileId;
      var spreadsheet;
      var fileName;

      if (mimeType === MimeType.MICROSOFT_EXCEL) {
          logToSheet(nombreHojaLogs, 'ERROR. El control: "' + file + '" tiene formato EXCEL y no se puede extraer su ID.');
      } else {
        fileId = file.getId();
        fileName = file.getName();
        spreadsheet = SpreadsheetApp.openById(fileId);
      

      var firstSheetName = spreadsheet.getSheets()[0].getName();

      newSheet.getRange(row, 1).setValue(fileId);
      newSheet.getRange(row, 2).setValue(firstSheetName);
      newSheet.getRange(row, 3).setValue(fileName);
      logToSheet(nombreHojaLogs, 'Se ha copiado el ID del control: "' + fileName + '" correctamente.');
      row++;
      }
    }
  }
}

/** FUNCIÓN AUXILIAR
 * Compara los datos de dos hojas de cálculo (la de los IDs) y registra las coincidencias en una nueva hoja.
 * Si no se encuentra una coincidencia, pregunta al usuario si desea crear un nuevo control.
 *
 * @param {string} nombreHojaLogs - El nombre de la hoja de logs donde se registrarán las acciones realizadas.
 * @param {string} sheet1Name - El nombre de la primera hoja a comparar (siempre debe ser la de controles testeados).
 * @param {string} sheet2Name - El nombre de la segunda hoja a comparar (siempre debe ser la de los controles documentados).
 * @param {string} sheet2Location - La ubicación de la carpeta (de los controles documentados) donde se pueden crear nuevos controles.
 * @param {string} newSheetName - El nombre de la nueva hoja donde se registrarán los resultados de la comparación (la hoja del mapeo).
 */  
function compararSheets(nombreHojaLogs, sheet1Name, sheet2Name, sheet2Location, newSheetName) { 

  const ID = SpreadsheetApp.getActiveSpreadsheet().getId();
  const MAIN_SHEET = SpreadsheetApp.openById(ID);
  const UI = SpreadsheetApp.getUi(); 

  UI.alert('Comparando los datos de los Controles de la Hoja: "' + sheet1Name + '" y la Hoja: "' + sheet2Name + '"');
  
  var sheet1 = MAIN_SHEET.getSheetByName(sheet1Name); // Hoja 1
  var sheet2 = MAIN_SHEET.getSheetByName(sheet2Name); // Hoja 2
  var newSheet = MAIN_SHEET.getSheetByName(newSheetName); // Nueva hoja para resultados

  if (!newSheet) {
    newSheet = MAIN_SHEET.insertSheet(newSheetName); // Crear nueva hoja si no existe
  } else {
    newSheet.clear(); // Limpiar la hoja si ya existe
  }

  var data1 = sheet1.getDataRange().getValues(); // Obtener datos de sheet1
  var data2 = sheet2.getDataRange().getValues(); // Obtener datos de sheet2

  var row = 1;
  var sheet2Map = {};

  // Construir el mapa de sheet2 con los nombres de las hojas (Columna C de sheet2)
  for (var j = 1; j < data2.length; j++) {
    var sheetName2 = data2[j][2]; // Eliminar espacios y convertir a minúsculas
    sheet2Map[sheetName2] = {
      id: data2[j][0],  // Columna A (ID de control)
      name: data2[j][1] // Columna B (Hoja de control)
    };
  }

  // Comparar los nombres de las hojas (Columna C de sheet1 con Columna C de sheet2)
  for (var i = 1; i < data1.length; i++) {
    var id1 = data1[i][0];    // Columna A de sheet1
    var name1 = data1[i][1];  // Columna B de sheet1
    var sheetName1 = data1[i][2];  // Eliminar espacios y convertir a minúsculas
    var matchFound = false;

    // Verificar coincidencia exacta o si sheetName1 es parte de sheetName2
    for (var sheetName2 in sheet2Map) {
      if (sheetName1.trim().toLowerCase() === sheetName2.trim().toLowerCase() || sheetName2.trim().toLowerCase().includes(sheetName1.trim().toLowerCase())) {
        var id2 = sheet2Map[sheetName2].id;
        var name2 = sheet2Map[sheetName2].name;

        // Guardar valores en la nueva hoja
        newSheet.getRange(row, 1).setValue(id1);      // ID de control de sheet1
        newSheet.getRange(row, 2).setValue(name1);    // Hoja de control de sheet1
        newSheet.getRange(row, 3).setValue(id2);      // ID de control de sheet2
        newSheet.getRange(row, 4).setValue(name2);    // Hoja de control de sheet2
        newSheet.getRange(row, 5).setValue(sheetName2); // Nombre de sheet que coincide
        row++;
        matchFound = true;
        break; // Salimos del loop una vez encontramos una coincidencia
      }
    }

    // Si no se encuentra coincidencia, preguntar al usuario si quiere crear el control
    if (!matchFound) {   

      var crearControlMenu = true;
      do{     
        var response = preguntaMenu('El Control: "' + sheetName1 + '" NO se ha encontrado en los Controles dentro de "' + sheet2Name + '"\n¿Desea crear el Control "' + sheetName1 + '" en la Carpeta de los Controles documentados? (y/n)');
        
        if (response) {
          try {
            // Copiar el control al folder destino
            var responseName = UI.prompt('¿Que nombre desea ponerle al control:  ' + sheetName1 + ' en la carpeta destino?');
            
            var sourceFile = DriveApp.getFileById(id1);
            var folderDestino = DriveApp.getFolderById(sheet2Location);
            var copiedFile = sourceFile.makeCopy(responseName.getResponseText().trim(), folderDestino);
  
            // Obtener nuevo ID y añadirlo al sheet
            var newId = copiedFile.getId();
            var newName = copiedFile.getName();
            var spreadsheet = SpreadsheetApp.openById(newId);
            var firstSheetName = spreadsheet.getSheets()[0].getName(); // Obtener nombre de la primera hoja
  
            // Guardar el nuevo control en la nueva hoja
            newSheet.getRange(row, 1).setValue(id1);
            newSheet.getRange(row, 2).setValue(name1);
            newSheet.getRange(row, 3).setValue(newId);
            newSheet.getRange(row, 4).setValue(firstSheetName);
            newSheet.getRange(row, 5).setValue(newName); // Guardar el nombre de la hoja original (en este caso, la que se ha copiado y se ha escogido como nombre en la hoja destino)
            row++;           
            logToSheet(nombreHojaLogs, 'Control creado: "' + newName + '" con ID: "' + newId + '"');
            crearControlMenu = false;
          } catch (error) {
            UI.alert('ERROR: Error general al copiar el control testeado: "' + error.message + '"');
          }
        }else{
          logToSheet(nombreHojaLogs, 'ADVERTENCIA: Control no creado: "' + sheet1Name + '"');
          crearControlMenu = false;
        }
      }while(crearControlMenu);
    }
  }
}

/** FUNCIÓN AUXILIAR
 * Copia celdas específicas desde una hoja de control (control testeado) a la otra hoja del control documentado.
 * Verifica que los formatos sean correctos y registra los cambios en una hoja de logs.
 *
 * @param {string} nombreHojaLogs - El nombre de la hoja de logs donde se registrarán las acciones realizadas.
 * @param {string} idFile - El nombre de la hoja que contiene los IDs (la hoja del mapeo de controles).
 * @param {string} nombreHojaPrincipal - El nombre de la hoja principal donde se registrarán los resultados.
 */  
function copiarCeldasDesdeControl(nombreHojaLogs, idFile, nombreHojaPrincipal) {
  const UI = SpreadsheetApp.getUi(); 
  try {
    //Abre la spreadsheet desde donde se esta ejecutando el script y pilla la hoja que contiene las comparaciones
    var hojaConIds = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(idFile); 

    //Hoja principal (Agenda) con todos los controles
    var hojaAgenda = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(nombreHojaPrincipal);

    //Pilla la referencia del rango que contiene datos en las hojas y almacena los datos del rango especificado en un array 2d
    var datos = hojaConIds.getDataRange().getValues();
    var datosNombres = hojaAgenda.getDataRange().getValues();

    for (var i = 1; i < datos.length; i++) {
      //De la hoja de los datos donde estan las URLs, por cada linea que haya, hace todo a continuacion
      var idOrigen = datos[i][0];
      var hojaOrigenNombre = datos[i][1];
      var idDestino = datos[i][2];
      var hojaDestinoNombre = datos[i][3];
      var controlActual = datos[i][4];
      
      logToSheet(nombreHojaLogs, '--------------------------------------------------------------------')
      logToSheet(nombreHojaLogs, 'CONTROL actual: ' + controlActual);
      logToSheet(nombreHojaLogs, 'Hoja Origen: ' + hojaOrigenNombre);
      logToSheet(nombreHojaLogs, 'Hoja Destino: ' + hojaDestinoNombre);

      try {
        //Abre la hoja origen (el control de este anio) dado el ID (url) correspondiente
        var archivoOrigen = SpreadsheetApp.openById(idOrigen);
        var hojaOrigen = archivoOrigen.getSheetByName(hojaOrigenNombre);
        if (!hojaOrigen) {
          UI.alert('ERROR: No se encontró la hoja de origen con el nombre especificado: "' + hojaOrigenNombre + '"');
        }
      } catch (e) {
        logToSheet(nombreHojaLogs, 'ERROR al abrir el archivo de origen con ID: "' + idOrigen + '" de la hoja origen: "' + hojaOrigenNombre);
        continue;
      }

      try {
        //Abre la hoja destino (control documentado formal) de la misma manera al origen
        var archivoDestino = SpreadsheetApp.openById(idDestino);
        var hojaDestino = archivoDestino.getSheetByName(hojaDestinoNombre);
        if (!hojaDestino) {
          UI.alert('No se encontró la hoja de destino con el nombre especificado: ' + hojaDestinoNombre);
        }
      } catch (e) {
        logToSheet(nombreHojaLogs, 'ERROR al abrir el archivo de destino con ID: "' + idDestino + '" de la hoja destino: "' + hojaDestinoNombre);
        continue;
      }

      //Verificar que los ficheros siguen el formato estandar y sino retocarlos
      checkFormat(hojaOrigen, nombreHojaLogs, 'hoja origen'); 
      checkFormat(hojaDestino, nombreHojaLogs, 'hoja destino');
    

      var celdas = [
        {nombreOrigen: 'A4', origen: 'A5', nombreDestino: 'A2', destino: 'A3'},
        {nombreOrigen: 'B4', origen: 'B5', nombreDestino: 'B2', destino: 'B3'},
        {nombreOrigen: 'C4', origen: 'C5', nombreDestino: 'C2', destino: 'C3'},
        {nombreOrigen: 'D4', origen: 'D5', nombreDestino: 'D2', destino: 'D3'},
        {nombreOrigen: 'E4', origen: 'E5', nombreDestino: 'E2', destino: 'E3'},
        {nombreOrigen: 'F4', origen: 'F5', nombreDestino: 'F2', destino: 'F3'},
        {nombreOrigen: 'A7', origen: 'B7:F7', nombreDestino: 'A6', destino: 'B6:F6'},
        {nombreOrigen: 'A9', origen: 'B9:F9', nombreDestino: 'A8', destino: 'B8:F8'},
        {nombreOrigen: 'A13', origen: 'B13:F13', nombreDestino: 'A12', destino: 'B12:F12'}
      ];

      var textoCopiado = [];  // Para almacenar los campos copiados y pegarlos al excel con los logs

      for (var j = 0; j < celdas.length; j++) {
        var rangoOrigen = celdas[j].origen;
        var rangoDestino = celdas[j].destino;
        var valores = hojaOrigen.getRange(rangoOrigen).getValues();
        //Verifica si existen valores en el campo origen en los campos de actualizaciones
        var hayValores = valores.some(fila => fila.some(valor => valor.trim() !== ''));
          
        //Si existe algo (es decir, ha habido un cambio y se ha rellenado la celda correspondiente) 
        if (hayValores) {
          hojaDestino.getRange(rangoDestino).setValues(valores);
          textoCopiado.push(hojaOrigen.getRange(celdas[j].nombreOrigen).getValue().toString());
          logToSheet(nombreHojaLogs, 'Se ha copiado el campo: "' + hojaOrigen.getRange(celdas[j].nombreOrigen).getValue().toString() + '" de la hoja origen al campo: "' + hojaDestino.getRange(celdas[j].nombreDestino).getValue().toString() + '" de la hoja destino');
        }
      }

      //para comparar muestra tamanio
      if(hojaOrigen.getRange('F14').getValue() !== hojaDestino.getRange('F14').getValue()){
        hojaDestino.getRange('F14').setValue(hojaOrigen.getRange('F14').getValue());
        textoCopiado.push(hojaOrigen.getRange(hojaOrigen.getRange('E14')).getValue().toString());
        logToSheet(nombreHojaLogs, 'Se ha copiado el campo: "' + hojaOrigen.getRange('E14').getValue().toString() + '" de la hoja origen al campo: "' + hojaDestino.getRange('E14').getValue().toString() + '" de la hoja destino');
      }
      
      
      verificarCeldas(hojaDestino, nombreHojaLogs);

      //Pega en el log los valores de los campos que se han ido cambiando en cada iteracion de cada control 
      for (var j = 2; j < datosNombres.length; j++) {
        if(datosNombres[j][0].trim().toLowerCase() === controlActual.trim().toLowerCase()) {
          hojaAgenda.getRange(j+1, 2).setValue('OK');          
          // Columna donde queremos poner el resultado (la siguiente a la columna "OK")
          hojaAgenda.getRange(j+1, 3).setValue(textoCopiado.join(', '));
          //Columna donde ponemos frecuencias y numero de muestras
          hojaAgenda.getRange(j+1, 4).setValue(hojaDestino.getRange('D3').getValue().toString() + ', ' + hojaDestino.getRange('F14').getValue().toString());
          hojaAgenda.getRange(j+1, 5).setValue(hojaDestino.getRange('F14').getValue().toString());
          logToSheet(nombreHojaLogs, 'Celda de la hoja principal actualizada con OK y campos copiados para "' + hojaDestinoNombre + '"');
        }
      }
    }
  } catch (e) {
    UI.alert('ERROR: Error general durante el proceso de actualizar los controles: ' + e.message);
  }
}

/** FUNCIÓN PRINCIPAL
 * Gestiona el menú de opciones para el usuario.
 * Permite al usuario seleccionar diferentes acciones relacionadas con el manejo de controles documentados y testeados.
 * Las opciones incluyen copiar IDs de controles, comparar controles y actualizar controles documentados.
 */
function main(){
  const ID = SpreadsheetApp.getActiveSpreadsheet().getId();
  const MAIN_SHEET = SpreadsheetApp.openById(ID);
  const UI = SpreadsheetApp.getUi(); 

  var menu = true;
  var date = new Date();

  do {
    var option = UI.prompt('Introduce el número de la opción que quieres ejecutar:\n (1) -> Copiar IDs de Carpeta Controles Documentados. \nRecuerda: Los ficheros de Controles deben de estar todos en una sola carpeta.\n (2) -> Copiar IDs de Controles Testeados. \nRecuerda: Los ficheros de Controles deben de encontrarse en carpetas separadas, las cuales se encuentran en una sola carpeta.\n (3) -> Comparar Controles.\n (4) -> Actualizar Controles Documentados.\n (5) -> Salir.');
    switch (option.getResponseText().trim()) {
      case '1':
        var response = UI.prompt("Introduzca el ID de la *CARPETA* que contiene los Controles Documentados:"); 
        var response2 = UI.prompt("Introduzca el NOMBRE para la *HOJA* que se va a crear conteniendo los IDs de los Controles: ");
        var nombreHojaLogs = '(' + date.getDate() + '/' + (date.getMonth() + 1) + '/' + date.getFullYear() + ') - Logs COPIAR IDs Documentados para "' + response2.getResponseText() + '"';      
        createLogSheet(nombreHojaLogs);
        almacenarIDsDocus(response.getResponseText().trim(), response2.getResponseText().trim(), nombreHojaLogs);
        UI.alert('Se ha terminado la ejecución de copiar los IDs de los Controles.');
        break;
      case '2':
        var response = UI.prompt("Introduzca el ID de la *CARPETA* que contiene carpetas con los Controles Testeados: ");
        var response2 = UI.prompt("Introduzca el NOMBRE para la *HOJA* que se va a crear conteniendo los IDs de los controles: ");
        var nombreHojaLogs = '(' + date.getDate() + '/' + (date.getMonth() + 1) + '/' + date.getFullYear() + ') - Logs COPIAR IDs Controles para "' + response2.getResponseText() + '"';      
        createLogSheet(nombreHojaLogs);
        almacenarIDs(response.getResponseText().trim(), response2.getResponseText().trim(), nombreHojaLogs);
        UI.alert('Se ha terminado la ejecución de copiar los IDs de los Controles.');
        break;
      case '3':
        var nombreControlesTesteados = solicitarNombreHoja("Introduzca el NOMBRE de la *HOJA* que contiene los IDs de los Controles Testeados que quieres copiar y actualizar en los Controles Documentados: ");
        if (nombreControlesTesteados === null) break; // Si el usuario cierra el diálogo, salir
        
        var nombreControlesDocumentados = solicitarNombreHoja("Introduzca el NOMBRE de la *HOJA* que contiene los IDs de los Controles Documentados que quieres actualizar: ");
        if (nombreControlesDocumentados === null) break; // Si el usuario cierra el diálogo, salir
        
        var response3 = UI.prompt("Introduzca el ID de la *CARPETA* que contiene los Controles Documentados: ");
        var response4 = UI.prompt('Introduzca el NOMBRE para la *HOJA* que se va a crear para representar la comparación de los Controles en ' + nombreControlesTesteados + ' y ' + nombreControlesDocumentados);
        var nombreHojaLogs = '(' + date.getDate() + '/' + (date.getMonth() + 1) + '/' + date.getFullYear() + ') - Logs COMPARAR IDs Controles para "' + response4.getResponseText() + '"';      
        createLogSheet(nombreHojaLogs);
        compararSheets(nombreHojaLogs, nombreControlesTesteados, nombreControlesDocumentados, response3.getResponseText().trim(), response4.getResponseText().trim()); 
        UI.alert('Ejecución terminada. \nYa se ha creado la hoja de comparación necesaria para actualizar los Controles Documentados con los Controles de "' + nombreControlesTesteados + '".');
        break;
      case '4': 
        var nombreHojaPrincipal = solicitarNombreHoja("Introduzca el NOMBRE de la *HOJA* principal que contiene todos los controles para dejar un registro de las actualizaciones: ");
        if (nombreHojaPrincipal === null) break; // Si el usuario cierra el diálogo, salir

        var nombreHojaComparacion = solicitarNombreHoja("Introduzca el NOMBRE de la *HOJA* que contiene los IDs de la comparación de Controles Testados y Controles Documentados: ");
        if (nombreHojaComparacion === null) break; // Si el usuario cierra el diálogo, salir

        var nombreHojaLogs = '(' + date.getDate() + '/' + (date.getMonth() + 1) + '/' + date.getFullYear() + ') - Logs ACTUALIZACIONES Controles para "' + nombreHojaPrincipal + '"';      
        UI.alert('Se procede a ejecutar las actualizaciones.\n\nESPERE PACIENTEMENTE HASTA EL POP UP INFORMANDO DE QUE EL PROCESO HA TERMINADO.\n\nPuede hacer un seguimiento y una revisión en la hoja de Logs que puede encontrar en: "' + nombreHojaLogs + '"');
        createLogSheet(nombreHojaLogs);
        copiarCeldasDesdeControl(nombreHojaLogs, nombreHojaComparacion, nombreHojaPrincipal);
        UI.alert('Ejecución terminada. \n\nLas actualizaciones de los controles que se encuentran en: "' + nombreHojaComparacion + '" se pueden repasar en la hoja: "' + nombreHojaPrincipal + '". ');
        break;
      case '5':
        menu = false;
        break;
      default:
        UI.alert('Opción no válida: "' + option.getResponseText() + '". Por favor, introduce un número de opción válido.');
    }
    menu = preguntaMenu('¿Quieres realizar otra acción? (y/n)');
  } while (menu);

  UI.alert('EL PROGRAMA HA TERMINADO. Gracias por utilizar UPDATE-CONTROLS! :)');
  return;
}
