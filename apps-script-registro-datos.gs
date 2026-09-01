function doPost(e) {
  // REEMPLAZÁ CON EL ID DE TU HOJA DE CÁLCULO
  var idHoja = "1Dv-NTX7iPdpD8oMi7jgSEaqULnZXJbx_4QxSNE8HwVw";
  var ss = SpreadsheetApp.openById(idHoja);

  try {
    var data = JSON.parse(e.postData.contents);

    // 1. FECHA Y HORA ACTUAL (Formato Argentina)
    var ahora = new Date();
    var soloFecha = ahora.toLocaleDateString('es-AR');
    var soloHora = ahora.toLocaleTimeString('es-AR', { hour: '2-digit', minute: '2-digit' });

    var meses = ["Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"];
    var nombreMesActivo = meses[ahora.getMonth()] + "_" + ahora.getFullYear();

    // 2. VERIFICACIÓN Y CREACIÓN DE HOJA MENSUAL AUTOMÁTICA
    var sheet = ss.getSheetByName(nombreMesActivo);
    if (!sheet) {
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

      var anchos = [100, 100, 220, 130, 130, 220];
      for (var i = 0; i < anchos.length; i++) {
        sheet.setColumnWidth(i + 1, anchos[i]);
      }
    }

    // 3. INSERTAR LOS DATOS EN LA PLANILLA
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
    return ContentService.createTextOutput("Error: " + err.message);
  }
}
