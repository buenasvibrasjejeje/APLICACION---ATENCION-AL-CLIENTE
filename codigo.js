function mostrarFormulario() {
  var html = HtmlService.createHtmlOutputFromFile('formulario')
      .setWidth(600)
      .setHeight(600);
  SpreadsheetApp.getUi().showModalDialog(html, 'Formulario');
}
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
function guardarDatos(colaborador, fecha, inicio, fin, producto, institucion, nombre, atencion, detalle,solucion, estado, personaDerivada) {
  // Verificar si algún campo está vacío
  if (!colaborador || !fecha || !inicio || !producto || !institucion || !nombre || !atencion || !detalle || !estado) {
    SpreadsheetApp.getUi().alert('Debe completar todos los campos antes de guardar.');
    return; // Detiene la ejecución de la función si falta algún campo
  }
  // Guarda los datos ingresados por el usuario en la hoja de cálculo
  var hoja = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  hoja.appendRow([colaborador, fecha, inicio, fin, producto, institucion, nombre, atencion, detalle, solucion , estado, personaDerivada]);
  // Genera un código único
  var uniqueCode = generateUniqueCode();
  // Inserta el código único en la fila siguiente a la última fila
  hoja.getRange(hoja.getLastRow(), 13).setValue(uniqueCode);
  // Notificar al usuario que los datos se guardaron correctamente
  SpreadsheetApp.getUi().alert('Datos guardados correctamente. Código único generado: ' + uniqueCode);
}////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
function generateUniqueCode() {
  var currentDate = new Date();
  var year = currentDate.getFullYear().toString();
  var month = ('0' + (currentDate.getMonth() + 1)).slice(-2);
  var day = ('0' + currentDate.getDate()).slice(-2);
  var lastCode = PropertiesService.getScriptProperties().getProperty('lastCode');
  var sequentialNumber = 1;
  if (lastCode) {
    var lastSequentialNumber = parseInt(lastCode.slice(-4));
    sequentialNumber = lastSequentialNumber + 1;
  }
  var uniqueCode = year + month + day + ('000' + sequentialNumber).slice(-4);
  PropertiesService.getScriptProperties().setProperty('lastCode', uniqueCode);
  return uniqueCode;
}
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
function buscarDatosPorCodigoUnico() {
  // Solicitar al usuario que ingrese el código único
  var codigoUnico = SpreadsheetApp.getUi().prompt("Ingrese el código único a buscar:").getResponseText();
  var hoja = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var datos = hoja.getDataRange().getValues();
  var resultados = [];
  // Itera sobre los datos para encontrar las filas que coinciden con el código único
  for (var i = 1; i < datos.length; i++) { // Comenzamos desde la fila 2 para omitir la fila de encabezados
    if (datos[i][12] == codigoUnico) { // Suponiendo que el código único está en la columna K (11)
      resultados.push(datos[i]);
    }
  }
  // Construir la tabla de resultados
  var tablaHTML = "<table border='1'><tr><th>Colaborador</th><th>Fecha Atención</th><th>Inicio</th><th>Fin</th><th>Producto/Servicio</th><th>Institución/Empresa/Otros</th><th>Nombre</th><th>Tipo de Atención</th><th>Detalle - Especificar el Problema</th><th>Solución</th><th>Estado</th><th>Persona Derivada</th><th>Código</th></tr>";

  for (var j = 0; j < resultados.length; j++) {
    tablaHTML += "<tr>";
    tablaHTML += "<td>" + resultados[j][0] + "</td>"; // Fecha Atención
    tablaHTML += "<td>" + resultados[j][1] + "</td>"; // Fecha Atención
    tablaHTML += "<td>" + resultados[j][2] + "</td>"; // Inicio
    tablaHTML += "<td>" + resultados[j][3] + "</td>"; // Fin
    tablaHTML += "<td>" + resultados[j][4] + "</td>"; // Producto/Servicio
    tablaHTML += "<td>" + resultados[j][5] + "</td>"; // Institución/Empresa/Otros
    tablaHTML += "<td>" + resultados[j][6] + "</td>"; // Nombre
    tablaHTML += "<td>" + resultados[j][7] + "</td>"; // Tipo de Atención
    tablaHTML += "<td>" + resultados[j][8] + "</td>"; // Detalle
    tablaHTML += "<td>" + resultados[j][9] + "</td>"; // Estado
    tablaHTML += "<td>" + resultados[j][10] + "</td>"; // Código Único
    tablaHTML += "<td>" + resultados[j][11] + "</td>"; // Fecha Atención
    tablaHTML += "<td>" + resultados[j][12] + "</td>"; // Fecha Atención
    tablaHTML += "</tr>";
  }
  
  tablaHTML += "</table>";
  
  // Muestra la tabla de resultados en la ventana de diálogo
  var htmlOutput = HtmlService.createHtmlOutput(tablaHTML)
      .setWidth(800)
      .setHeight(400);
  
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, "Resultados de búsqueda");
}
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
function enviarCorreo(colaborador, email) {
  var asunto = "Asunto del correo";
  var cuerpo = "Hola " + colaborador + ",\n\nFuiste asignado a resolver esta atencion al cliente y fuiste derivado por :.";
  // Envía el correo electrónico
  MailApp.sendEmail(email, asunto, cuerpo);
}
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
function onSelectionChange(e) {
  var range = e.range; // Obtiene el rango seleccionado
  var sheet = range.getSheet(); // Obtiene la hoja en la que se realizó la selección
  
  // Verifica si la selección se realizó en la columna D
  if (range.getColumn() == 4) {
    // Establece la hora actual en la celda seleccionada
    sheet.getRange(range.getRow(), range.getColumn()).setValue(new Date().toLocaleTimeString());
  }
}