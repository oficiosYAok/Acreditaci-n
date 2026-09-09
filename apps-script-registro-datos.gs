function existeDni(sheet, dni) {
  var ultimaFila = sheet.getLastRow();
  if (ultimaFila < 2) return false;
  var columnaDni = sheet.getRange(2, 4, ultimaFila - 1, 1).getValues(); // Columna D = DNI
  var dniBuscado = String(dni).trim();
  for (var i = 0; i < columnaDni.length; i++) {
    if (String(columnaDni[i][0]).trim() === dniBuscado) return true;
  }
  return false;
}
 
function obtenerHojaDelMes(ss, crearSiNoExiste) {
  var ahora = new Date();
  var meses = ["Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"];
  var nombreMesActivo = meses[ahora.getMonth()] + "_" + ahora.getFullYear();
 
  var sheet = ss.getSheetByName(nombreMesActivo);
  if (!sheet && crearSiNoExiste) {
    sheet = ss.insertSheet(nombreMesActivo, 0);
 
    sheet.appendRow([
      "Fecha",              // A
      "Hora",                // B
      "Apellido y Nombre",   // C
      "DNI",                 // D
      "Celular",              // E
      "Escuela / Colegio"    // F
    ]);
 
    var rangoEncabezado = sheet.getRange("A1:F1");
    rangoEncabezado.setFontWeight("bold")
                   .setBackground("#5990fc")
                   .setFontColor("white")
                   .setHorizontalAlignment("center")
                   .setVerticalAlignment("middle")
                   .setWrap(true);
 
    sheet.setRowHeight(1, 40);
 
    var anchos = [100, 100, 220, 130, 130, 260];
    for (var i = 0; i < anchos.length; i++) {
      sheet.setColumnWidth(i + 1, anchos[i]);
    }
  }
  return sheet;
}
 
// Recibe el registro del formulario y lo guarda, bloqueando DNI repetidos.
function doPost(e) {
  // REEMPLAZÁ CON EL ID DE TU HOJA DE CÁLCULO
  var idHoja = "PEGAR_ID_DE_TU_GOOGLE_SHEET_ACA";
  var ss = SpreadsheetApp.openById(idHoja);
 
  try {
    var data = JSON.parse(e.postData.contents);
 
    var ahora = new Date();
    var soloFecha = ahora.toLocaleDateString('es-AR');
    var soloHora = ahora.toLocaleTimeString('es-AR', { hour: '2-digit', minute: '2-digit' });
 
    var sheet = obtenerHojaDelMes(ss, true);
 
    // BLOQUEO DE DNI DUPLICADO (protección real, del lado del servidor)
    if (existeDni(sheet, data.dni)) {
      return ContentService.createTextOutput("Duplicado: este DNI ya fue registrado.")
        .setMimeType(ContentService.MimeType.TEXT);
    }
 
    sheet.appendRow([
      soloFecha,       // A
      soloHora,        // B
      data.nombre,     // C
      data.dni,        // D
      data.cel,        // E
      data.escuela     // F
    ]);
 
    return ContentService.createTextOutput("Éxito").setMimeType(ContentService.MimeType.TEXT);
  } catch (err) {
    return ContentService.createTextOutput("Error: " + err.message)
      .setMimeType(ContentService.MimeType.TEXT);
  }
}
 
// Usado por el formulario (vía JSONP) para avisar ANTES de enviar si el DNI ya existe.
function doGet(e) {
  var idHoja = "1Dv-NTX7iPdpD8oMi7jgSEaqULnZXJbx_4QxSNE8HwVw";
  var ss = SpreadsheetApp.openById(idHoja);
 
  var accion = e.parameter.accion;
  var callback = e.parameter.callback;
 
  if (accion === 'verificarDni') {
    var dni = e.parameter.dni || '';
    var sheet = obtenerHojaDelMes(ss, false);
    var existe = sheet ? existeDni(sheet, dni) : false;
    var resultado = JSON.stringify({ existe: existe });
 
    if (callback) {
      return ContentService.createTextOutput(callback + '(' + resultado + ')')
        .setMimeType(ContentService.MimeType.JAVASCRIPT);
    }
    return ContentService.createTextOutput(resultado).setMimeType(ContentService.MimeType.JSON);
  }
 
  return ContentService.createTextOutput("OK");
}
 
