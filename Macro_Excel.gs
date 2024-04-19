function onOpen() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var menuItems = [
    {name: 'Mover Filas', functionName: 'moverFilas'}
  ];
  spreadsheet.addMenu('Custom Menu', menuItems);

  // Llamar a moverFilas() al abrir la hoja
  moverFilas();
}

function moverFilas() {
  var hojaOrigen = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var hojaDestino = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("verificar");
  
  var data = hojaOrigen.getDataRange().getValues();
  var newData = [];
  var filasAEliminar = []; // Almacena los índices de las filas a eliminar en la hoja de origen
  
  for (var i = 0; i < data.length; i++) {
    if (data[i].includes("Remitidos soportes a LEXER") || data[i].includes("RADICADA SOLICITUD DE MUTACIÓN CATASTRAL")) {
      newData.push(data[i]);
      filasAEliminar.push(i + 1); // Guarda el índice de la fila a eliminar (+1 para ajustar el índice)
    }
  }
  
  // Copiar los datos a la hoja de destino
  if (newData.length > 0) {
    hojaDestino.getRange(hojaDestino.getLastRow() + 1, 1, newData.length, newData[0].length).setValues(newData);
  }
  
  // Eliminar las filas de la hoja de origen
  for (var j = filasAEliminar.length - 1; j >= 0; j--) {
    hojaOrigen.deleteRow(filasAEliminar[j]);
  }
  // Llamar a organizarPorA() después de 1 minuto
  Utilities.sleep(60000); // 1 minuto = 60000 milisegundos
  copiarFilas1();
}

function copiarFilas1() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var hojaActiva = ss.getSheetByName("verificar"); // Obtener la hoja llamada "verificar"
  if (hojaActiva) { // Verificar si la hoja existe
    var hojaDestino = ss.getSheetByName("2023");
    var rangoDatos = hojaActiva.getDataRange();
    var filas = rangoDatos.getValues();
    var fechaInicio = new Date('1/1/2023');
    var fechaFin = new Date('12/31/2023');
    var colorVerde = '#00FF00';
    
    for (var i = 0; i < filas.length; i++) {
      var fecha = filas[i][9]; // Columna J, índice 9 (empezando desde 0)
      if (fecha instanceof Date && fecha >= fechaInicio && fecha <= fechaFin) {
        hojaDestino.appendRow(filas[i]);
        var ultimaFila = hojaDestino.getLastRow();
        var rangoFila = hojaDestino.getRange(ultimaFila, 1, 1, hojaDestino.getLastColumn());
        rangoFila.setBackground(colorVerde);
      }
    }
    
  } else {
    Logger.log("La hoja 'verificar' no fue encontrada.");
  }
  Utilities.sleep(30000); // 1 minuto = 60000 milisegundos
  pasar();
}

function pasar() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var hojaActiva = ss.getSheetByName("verificar"); // Obtener la hoja llamada "verificar"
  if (hojaActiva) { // Verificar si la hoja existe
    var hojaDestino = ss.getSheetByName("2024");
    var rangoDatos = hojaActiva.getDataRange();
    var filas = rangoDatos.getValues();
    var fechaInicio = new Date('1/1/2024');
    var fechaFin = new Date('12/31/2024');
    var colorVerde = '#00FF00';
    
    for (var i = 0; i < filas.length; i++) {
      var fecha = filas[i][9]; // Columna J, índice 9 (empezando desde 0)
      if (fecha instanceof Date && fecha >= fechaInicio && fecha <= fechaFin) {
        hojaDestino.appendRow(filas[i]);
        var ultimaFila = hojaDestino.getLastRow();
        var rangoFila = hojaDestino.getRange(ultimaFila, 1, 1, hojaDestino.getLastColumn());
        rangoFila.setBackground(colorVerde);
      }
    }
    rangoDatos.clear();
  } else {
    Logger.log("La hoja 'verificar' no fue encontrada.");
  }
}