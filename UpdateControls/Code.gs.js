// CREATOR: LUCAS SAETA (@1ukz) - lucassaeta9@gmail.com

//TODOOOOO 
// en la funcion de limpiar control creado, hay que hacer verificaciones para no borrar vainas que no debiera por no estar en las posiciones enviadas por args. 
// copiarControles no funciona bien, mirar logs y hoja principal PRUEBA AGENDA
// en la docu, al principio el punto 3, explicar las key hojas y apartados importantes: rollo poner - Hoja principal: esta sera l ahoja etc, - Hoja Docus, - Hoja SOX, - Hoja comparaciones.


// GLOBAL VARIABLES
// const ID = SpreadsheetApp.getActiveSpreadsheet().getId();
// const MAIN_SHEET = SpreadsheetApp.openById(ID);
// const UI = SpreadsheetApp.getUi(); 

function onOpen() {
  const UI = SpreadsheetApp.getUi(); 

  UI.createMenu('Update-Controls')
    .addItem('Ejecutar programa', 'main')
    .addToUi();
}

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
  cell.setFontSize(14); // Ajusta el tamaño de la fuente según tus necesidades
  cell.setFontWeight('bold'); // Aplica negrita al texto
}

function logToSheet(nombreHoja, message){
  const ID = SpreadsheetApp.getActiveSpreadsheet().getId();
  const MAIN_SHEET = SpreadsheetApp.openById(ID);

  var logSheet = MAIN_SHEET.getSheetByName(nombreHoja);  
  const lastRow = logSheet.getLastRow();
  logSheet.getRange(lastRow + 1, 1).setValue(message);  

}

function checkSheetExists(sheetName) {
  var sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
  for (var i = 0; i < sheets.length; i++) {
    if (sheets[i].getName() === sheetName) {
      return true; // La hoja existe
    }
  }
  return false; // La hoja no existe
}

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

function preguntaMenu() {

  const UI = SpreadsheetApp.getUi(); 
  var bool;
  var bucle = false;
  do {
    var resp = UI.prompt('¿Quieres realizar otra acción? (y/n)');
    if (resp.getResponseText().toLowerCase() === 'y') {
      bool = true;
      bucle = true;
    } else if (resp.getResponseText().toLowerCase() === 'n') {
      bool = false;
      bucle = true;
    } else {
      UI.alert('Respuesta no válida. Por favor, introduzca "y" para realizar otra acción, o "n" para salir del programa.');
    }
  } while (!bucle);
  return bool;

}

function limpiarControlCreado(sheet, celdasAComprobar) {
  celdasAComprobar.forEach(function(rango) {
    sheet.getRange(rango).clearContent();
  });
}

function almacenarIDs(folderId, sheetName, nombreHojaLogs) {
    const ID = SpreadsheetApp.getActiveSpreadsheet().getId();
    const MAIN_SHEET = SpreadsheetApp.openById(ID);
    const UI = SpreadsheetApp.getUi(); 

    var newSheet = MAIN_SHEET.getSheetByName(sheetName);
    UI.alert('Capturando los IDs de los Controles en: "' + folderId + '" en la Hoja: "' + sheetName + '". \nPOR FAVOR, ESPERE PACIENTEMENTE HASTA EL PROXIMO POP UP QUE INDICANDO QUE EL PROCESO HA TERMINADO. \nSe ha generado una Hoja con los Logs de lo realizado para esta ejecución, disponible en: "' + nombreHojaLogs + '"');
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

            logToSheet(nombreHojaLogs, 'Error. El control: "' + subfolderName + '" tiene formato EXCEL y no se puede extraer su ID.');
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
  
  function almacenarIDsDocus(folderId, sheetName, nombreHojaLogs) {
    const ID = SpreadsheetApp.getActiveSpreadsheet().getId();
    const MAIN_SHEET = SpreadsheetApp.openById(ID);
    const UI = SpreadsheetApp.getUi(); 

    UI.alert('Capturando los IDs de los Controles en: "' + folderId + '" en la Hoja: "' + sheetName + '". \n\nPOR FAVOR, ESPERE PACIENTEMENTE HASTA EL PROXIMO POP UP INDICANDO QUE EL PROCESO HA TERMINADO. \n\nSe ha generado una Hoja con los Logs de lo realizado para esta ejecución, disponible en: "' + nombreHojaLogs + '"');
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
            logToSheet(nombreHojaLogs, 'El control: "' + file + '" tiene formato EXCEL.');
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
        var response = UI.prompt('El Control: "' + sheetName1 + '" NO se ha encontrado en los Controles dentro de "' + sheet2Name + '"\n¿Desea crear el Control "' + sheetName1 + '" en la Carpeta de los Controles documentados? (y/n)');
        
        if (response.getResponseText().toLowerCase() === 'y') {
          try {
            // Copiar el control al folder destino
            var responseName = UI.prompt('¿Que nombre desea ponerle al control:  ' + sheetName1 + ' en la carpeta destino?').getResponseText();
            
            var sourceFile = DriveApp.getFileById(id1);
            var folderDestino = DriveApp.getFolderById(sheet2Location);
            var copiedFile = sourceFile.makeCopy(responseName, folderDestino);
  
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
            limpiarControlCreado(spreadsheet.getSheets()[0], [
            'B13:F13', 'B9:F9', 'B7:F7', 'A5', 'B5', 'C5', 'D5', 'E5', 'F5', 'F15'
            ]);            
            logToSheet(nombreHojaLogs, 'Control creado: "' + newName + '" con ID: "' + newId + '"');
            crearControlMenu = false;
          } catch (error) {
            UI.alert('Error al copiar el control: "' + error.message + '"');
          }
        }else if (response.getResponseText().toLowerCase() === 'n'){
          logToSheet(nombreHojaLogs, 'Control no creado: "' + sheet1Name + '"');
          crearControlMenu = false;
        }else{
          UI.alert('Respuesta no válida. Por favor, introduzca "y" para crear una copia del Control en los Documentados, o "n" para no crearlo.');
        }
      }while(crearControlMenu);
    }
  }
}
  
function verificarCeldas(hojaDestino, nombreHojaLogs) {

  var celdasAComprobar = [
    'B13:F13', 'B9:F9', 'B7:F7', 'A5', 'B5', 'C5', 'D5', 'E5', 'F5'
  ];

  var celdasNoVacias = [
    'B6:F6', 'B8:F8', 'B12:F12'
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
    logToSheet(nombreHojaLogs, 'ADVERTENCIA: Existe una fila adicional (15) que no debería estar presente.');
  }
}
  
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
          UI.alert('No se encontró la Hoja de origen con el nombre especificado: "' + hojaOrigen + '"');
        }
      } catch (e) {
        logToSheet(nombreHojaLogs, 'Error al abrir el archivo de origen con ID: "' + idOrigen + '" de la hoja origen: "' + hojaOrigenNombre);
        continue;
      }

      try {
        //Abre la hoja destino (control documentado formal) de la misma manera al origen
        var archivoDestino = SpreadsheetApp.openById(idDestino);
        var hojaDestino = archivoDestino.getSheetByName(hojaDestinoNombre);
        if (!hojaDestino) {
          UI.alert('No se encontró la Hoja de destino con el nombre especificado: ' + hojaDestino);
        }
      } catch (e) {
        logToSheet(nombreHojaLogs, 'Error al abrir el archivo de destino con ID: "' + idDestino + '" de la hoja destino: "' + hojaDestinoNombre);
        continue;
      }

      //Segun cuantas filas tenga el documento (si tiene una arriba extra, por ejemplo) la celda de los datos a copiar se encontraran en un lugar u otro
      var numRango = 0;
      if(hojaOrigen.getRange('A2').getValue() === 'Tipo de Control' || hojaOrigen.getRange('A2').getValue() === 'Clase'){      
        numRango = 5;
      } else if (hojaOrigen.getRange('A3').getValue() === 'Tipo de Control' || hojaOrigen.getRange('A3').getValue() === 'Clase') {
        numRango = 6;
      } else {
        logToSheet(nombreHojaLogs, 'ERROR. Fichero con rango de celdas incorrecto o nombres diferentes. REVISAR: ' + hojaDestinoNombre);
        continue; 
      }

      var celdas = [
        {nombreOrigen: 'A' + (numRango-1), origen: 'A' + numRango, nombreDestino: 'A' + (numRango-3), destino: 'A' + (numRango-2)},
        {nombreOrigen: 'B' + (numRango-1), origen: 'B' + numRango, nombreDestino: 'B' + (numRango-3), destino: 'B' + (numRango-2)},
        {nombreOrigen: 'C' + (numRango-1), origen: 'C' + numRango, nombreDestino: 'C' + (numRango-3), destino: 'C' + (numRango-2)},
        {nombreOrigen: 'D' + (numRango-1), origen: 'D' + numRango, nombreDestino: 'D' + (numRango-3), destino: 'D' + (numRango-2)},
        {nombreOrigen: 'E' + (numRango-1), origen: 'E' + numRango, nombreDestino: 'E' + (numRango-3), destino: 'E' + (numRango-2)},
        {nombreOrigen: 'A' + (numRango+2), origen: 'B' + (numRango+2) + ':' + 'F' + (numRango+2), nombreDestino: 'A' + (numRango+1), destino: 'B' + (numRango+1) + ':' + 'F' + (numRango+1)},
        {nombreOrigen: 'A' + (numRango+4), origen: 'B' + (numRango+4) + ':' + 'F' + (numRango+4), nombreDestino: 'A' + (numRango+3), destino: 'B' + (numRango+3) + ':' + 'F' + (numRango+3)}
      ];

      var textoCopiado = [];  // Para almacenar los campos copiados y pegarlos al excel con los logs
      var textoFrecuentasYMuestras = []; //Para el numero de frecuencias y muestras y pegarlos al excel con los logs
      var textoSoloMuestras = [];

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
      
      //Logica para el campo de Pruebas, ya que suele ser diferente y no es tan uniforme como el resto de campos ya validados (suele a veces haber filas extras entre la Evidencia y las Pruebas)
      var valorA12 = hojaOrigen.getRange('A12').getValue();
      var numRangoMuestras = 0;

      //Si en el campo A12 esta el texto a continuacion
      if (valorA12 === 'Prueba a realizar') {
      // Se almacenan los de la fila siguiente en el rango correspondiente
      var valoresB13F13 = hojaOrigen.getRange('B13:F13').getValues();
      var hayValoresB13F13 = valoresB13F13.some(fila => fila.some(valor => valor.trim() !== '' && valor !== null));

      if (hayValoresB13F13 && hojaDestino.getRange('A12').getValue() === 'Prueba a realizar') {
          hojaDestino.getRange('B12:F12').setValues(valoresB13F13);
          textoCopiado.push('Prueba. Actualizaciones');
          logToSheet(nombreHojaLogs, 'Se ha copiado el campo: "Prueba. Actualizaciones" de la hoja origen al campo "' + hojaDestino.getRange('A12').getValue().toString() + '" de la hoja destino');
      } else if (hayValoresB13F13 && hojaDestino.getRange('A11').getValue() === 'Prueba a realizar') {
          hojaDestino.insertRowBefore(10);  // Esto agrega una fila en blanco en la fila 10
          hojaDestino.getRange('A10:F10').setBackground(null);  // Opcional: restablece el fondo de la nueva fila en blanco
          hojaDestino.getRange('A10:F10').mergeAcross();  // Opcional: fusionar las celdas en la nueva fila
          hojaDestino.getRange('B12:F12').setValues(valoresB12F12);
          logToSheet(nombreHojaLogs, 'El Control "' + controlActual + '"  no tenía el formato correcto. Se ha añadido la fila 10 para que siga el formato correcto.');
          logToSheet(nombreHojaLogs, 'Se ha copiado el campo: "Prueba. Actualizaciones" de la hoja origen al campo "' + hojaDestino.getRange('A12').getValue().toString() + '" de la hoja destino');
      }

    } else if (valorA12 === 'Prueba. Actualizaciones') { //TODO añadir verificación en lado destino para que pegue correcto

        numRangoMuestras++; // Se incrementa el contador para almacenar el campo correspondiente al número de muestras
        var valoresB12F12 = hojaOrigen.getRange('B12:F12').getValues();
        var hayValoresB12F12 = valoresB12F12.some(fila => fila.some(valor => valor.trim() !== '' && valor !== null));

        if (hayValoresB12F12 && hojaDestino.getRange('A12').getValue() === 'Prueba a realizar') {
            hojaDestino.getRange('B12:F12').setValues(valoresB12F12);
            textoCopiado.push('Prueba. Actualizaciones');
            logToSheet(nombreHojaLogs, 'Se ha copiado el campo: "Prueba. Actualizaciones" de la hoja origen al campo "' + hojaDestino.getRange('A12').getValue().toString() + '" de la hoja destino');
        } else if (hayValoresB12F12 && hojaDestino.getRange('A11').getValue() === 'Prueba a realizar') {
          hojaDestino.insertRowBefore(10);  // Esto agrega una fila en blanco en la fila 10
          hojaDestino.getRange('A10:F10').setBackground(null);  // Opcional: restablece el fondo de la nueva fila en blanco
          hojaDestino.getRange('A10:F10').mergeAcross();  // Opcional: fusionar las celdas en la nueva fila
            hojaDestino.getRange('B12:F12').setValues(valoresB12F12);
            logToSheet(nombreHojaLogs, 'El Control "' + controlActual + '"  no tenía el formato correcto. Se ha aniadido la fila 10 para que siga el formato correcto.');
            logToSheet(nombreHojaLogs, 'Se ha copiado el campo: "Prueba. Actualizaciones" de la hoja origen al campo "' + hojaDestino.getRange('A12').getValue().toString() + '" de la hoja destino');
        }
    } else{

        logToSheet(nombreHojaLogs, 'REVISAR Control Testeado"' + controlActual + '" ya que tiene la "PRUEBA A REALIZAR" en la fila incorrecta.');
    }

      //Array para el campo de frecuencias y muestras normales para ver si esta acorde, ademas de aniadir estos campos al log tambien 
      var frecuenciasYMuestras = [
        {nombreOrigen: 'D' + (numRango-3), origen: 'D' + (numRango-2), nombreDestino: 'D' + (numRango-3), destino: 'D' + (numRango-2)}, 
        {nombreOrigen: 'E' + (numRango+9-numRangoMuestras), origen: 'F' + (numRango+9-numRangoMuestras), nombreDestino: 'E' + (numRango+9), destino: 'F' + (numRango+9)}
      ]

      //Por cada linea del array 
      for (var k = 0; k < frecuenciasYMuestras.length; k++){
        var rangoOrigenMF = frecuenciasYMuestras[k].origen;
        var rangoDestinoMF = frecuenciasYMuestras[k].destino;
        var nombreRangoOrigenMF = frecuenciasYMuestras[k].nombreOrigen;
        var nombreRangoDestinoMF = frecuenciasYMuestras[k].nombreDestino;

        var muestraOrigen = hojaOrigen.getRange(rangoOrigenMF).getValue().toString();
        var muestraDestino = hojaDestino.getRange(rangoDestinoMF).getValue().toString();
        
        // //Valida si el campo frecuencia, o en la siguiente iteracion el campo numero de muestras es lo mismo
        // if(muestraOrigen !== muestraDestino && nombreRangoOrigenMF === nombreRangoDestinoMF && k === 0){
        //   hojaDestino.getRange(rangoDestinoMF).setValue(hojaOrigen.getRange(rangoOrigenMF).getValue());
        //   textoCopiado.push(hojaOrigen.getRange(frecuenciasYMuestras[k].nombreOrigen).getValue().toString());
        //   logToSheet(nombreHojaLogs, 'Se ha copiado el campo MUESTRA o FRECUENCIA: "' + hojaOrigen.getRange(frecuenciasYMuestras[k].nombreOrigen).getValue().toString() + '" de la hoja origen al campo "' + hojaDestino.getRange(frecuenciasYMuestras[k].nombreDestino).getValue().toString() + '" de la hoja destino.');
        // }else if(muestraOrigen !== muestraDestino && nombreRangoOrigenMF !== nombreRangoDestinoMF){
        //   logToSheet(nombreHojaLogs, 'Revisar porque las muestras o la frecuencia no coinciden, pero el formato no es uniforme entre el control: ' + hojaOrigenNombre + ' y: ' + hojaDestinoNombre);
        // }

        //Almacena los valores finales frecuencias y muestras del control documentado oficial 
        textoFrecuentasYMuestras.push(hojaDestino.getRange(rangoDestinoMF).getValue().toString());

      }

      verificarCeldas(hojaDestino, nombreHojaLogs);
      //Para la columna de solo muestras
      var textoSoloMuestras = hojaDestino.getRange('F' + (numRango+9)).getValue().toString();
      
      //Pega en el log los valores de los campos que se han ido cambiando en cada iteracion de cada control 
      for (var j = 2; j < datosNombres.length; j++) {
        if(datosNombres[j][0].trim().toLowerCase() === controlActual.trim().toLowerCase()) {
          hojaAgenda.getRange(j+1, 2).setValue('OK');          
          // Columna donde queremos poner el resultado (la siguiente a la columna "OK")
          hojaAgenda.getRange(j+1, 3).setValue(textoCopiado.join(', '));
          //Columna donde ponemos frecuencias y numero de muestras
          hojaAgenda.getRange(j+1, 4).setValue(textoFrecuentasYMuestras.join(', '));
          hojaAgenda.getRange(j+1, 5).setValue(textoSoloMuestras);
          logToSheet(nombreHojaLogs, 'Celda de la hoja principal actualizada con OK y campos copiados para "' + hojaDestinoNombre + '"');
        }
      }
    }
  } catch (e) {
    UI.alert('Error general: ' + e.message);
  }
}

function main(){
  const ID = SpreadsheetApp.getActiveSpreadsheet().getId();
  const MAIN_SHEET = SpreadsheetApp.openById(ID);
  const UI = SpreadsheetApp.getUi(); 

  var menu = true;
  var date = new Date();

  do {
    var option = UI.prompt('Introduce el número de la opción que quieres ejecutar:\n (1) -> Copiar IDs de Carpeta Controles Documentados.\n (2) -> Copiar IDs de Carpeta con carpetas y controles.\n (3) -> Comparar controles.\n (4) -> Actualizar Controles Documentados.\n (5) -> Salir.');
    switch (option.getResponseText()) {
      case '1':
        var response = UI.prompt("Introduzca el ID de la *CARPETA* que contiene los Controles Documentados:"); 
        var response2 = UI.prompt("Introduzca el NOMBRE para la *HOJA* que se va a crear conteniendo los IDs de los Controles: ");
        var nombreHojaLogs = '(' + date.getDate() + '/' + (date.getMonth() + 1) + '/' + date.getFullYear() + ') - Logs COPIAR IDs Documentados para "' + response2.getResponseText() + '"';      
        createLogSheet(nombreHojaLogs);
        almacenarIDsDocus(response.getResponseText(), response2.getResponseText(), nombreHojaLogs);
        UI.alert('Se ha terminado la ejecución de copiar los IDs de los Controles correctamente.');
        break;
      case '2':
        var response = UI.prompt("Introduzca el ID de la *CARPETA* que contiene carpetas con los Controles Testeados: ");
        var response2 = UI.prompt("Introduzca el NOMBRE para la *HOJA* que se va a crear conteniendo los IDs de los controles: ");
        var nombreHojaLogs = '(' + date.getDate() + '/' + (date.getMonth() + 1) + '/' + date.getFullYear() + ') - Logs COPIAR IDs Controles para "' + response2.getResponseText() + '"';      
        createLogSheet(nombreHojaLogs);
        almacenarIDs(response.getResponseText(), response2.getResponseText(), nombreHojaLogs);
        UI.alert('Se ha terminado la ejecución de copiar los IDs de los Controles correctamente.');
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
        compararSheets(nombreHojaLogs, nombreControlesTesteados, nombreControlesDocumentados, response3.getResponseText(), response4.getResponseText()); 
        UI.alert('Ejecución terminada correctamente. \nYa se ha creado la hoja de comparación necesaria para actualizar los Controles Documentados con los Controles de "' + nombreControlesTesteados + '".');
        break;
      case '4': 
        var nombreHojaPrincipal = solicitarNombreHoja("Introduzca el NOMBRE de la *HOJA* principal que contiene todos los controles para dejar un registro de las actualizaciones: ");
        if (nombreHojaPrincipal === null) break; // Si el usuario cierra el diálogo, salir

        var nombreHojaComparacion = solicitarNombreHoja("Introduzca el NOMBRE de la *HOJA* que contiene los IDs de la comparación de Controles Testados y Controles Documentados: ");
        if (nombreHojaComparacion === null) break; // Si el usuario cierra el diálogo, salir

        var nombreHojaLogs = '(' + date.getDate() + '/' + (date.getMonth() + 1) + '/' + date.getFullYear() + ') - Logs ACTUALIZACIONES Controles para "' + nombreHojaPrincipal + '"';      
        UI.alert('Se procede a ejecutar las actualizaciones.\nESPERE PACIENTEMENTE HASTA EL POP UP INFORMANDO DE QUE EL PROCESO HA TERMINADO.\n\nPuede hacer un seguimiento y una revisión en la hoja de Logs que puede encontrar en: "' + nombreHojaLogs + '"');
        createLogSheet(nombreHojaLogs);
        copiarCeldasDesdeControl(nombreHojaLogs, nombreHojaComparacion, nombreHojaPrincipal);
        UI.alert('Ejecución terminada correctamente. \nLas actualizaciones de los controles que se encuentran en: "' + nombreHojaComparacion + '" se pueden repasar en la hoja: "' + nombreHojaPrincipal + '". ');
        break;
      case '5':
        menu = false;
        break;
      default:
        UI.alert('Opción no válida: "' + option.getResponseText() + '". Por favor, introduce un número de opción válido.');
    }
    menu = preguntaMenu();
  } while (menu);

  UI.alert('EL PROGRAMA HA TERMINADO. Gracias por utilizar UPDATE-CONTROLS! :)');
  return;
}
