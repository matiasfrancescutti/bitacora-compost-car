/**
 * ═══════════════════════════════════════════════════════════════════
 *  CAR · Bitácora de Compost — Conector con Google Sheets
 * ═══════════════════════════════════════════════════════════════════
 *
 *  Este código se pega en Google Apps Script (script.google.com)
 *  asociado a una Google Sheet. Una vez publicado como Web App,
 *  la URL resultante se carga en la bitácora HTML, dentro de
 *  "Configuración avanzada → URL del Web App".
 *
 *  PASOS DE INSTALACIÓN:
 *  1. Crear una hoja de cálculo nueva en Google Sheets.
 *  2. Menú: Extensiones → Apps Script.
 *  3. Borrar el código por defecto y pegar todo este archivo.
 *  4. Guardar (icono del disquete).
 *  5. Implementar → Nueva implementación → Tipo: Aplicación web.
 *  6. Ejecutar como: yo. Acceso: cualquier persona.
 *  7. Copiar la URL que termina en /exec y pegarla en la bitácora.
 */

// Recibe los datos enviados por la bitácora HTML y los registra en la hoja.
function doPost(e) {
  try {
    const datos = JSON.parse(e.postData.contents);
    const accion = datos.accion;
    const lote = datos.lote;
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let hoja = ss.getSheetByName("Cargas");

    // Crea la hoja con encabezados la primera vez que se usa.
    if (!hoja) {
      hoja = ss.insertSheet("Cargas");
      hoja.appendRow([
        "ID lote", "Fecha armado", "Rumen", "Hojas secas", "Hojas húmedas",
        "Ramas secas", "Ramas húmedas", "Tierra", "Otros",
        "C/N teórica", "C total (kg)", "N total (kg)",
        "Observaciones", "Calificación", "Comentario calificación", "Fecha calificación"
      ]);
      hoja.getRange("A1:P1").setFontWeight("bold").setBackground("#2D5016").setFontColor("white");
    }

    // Si la acción es una nueva carga, agrega una fila al final de la hoja.
    if (accion === "nueva_carga") {
      hoja.appendRow([
        lote.id,
        lote.fecha,
        formatearCarga(lote.cargas.rumen),
        formatearCarga(lote.cargas.hojas_secas),
        formatearCarga(lote.cargas.hojas_humedas),
        formatearCarga(lote.cargas.ramas_secas),
        formatearCarga(lote.cargas.ramas_humedas),
        formatearCarga(lote.cargas.tierra),
        formatearCarga(lote.cargas.otros),
        lote.cn_teorica.toFixed(2),
        lote.c_total_kg.toFixed(2),
        lote.n_total_kg.toFixed(3),
        lote.notas || "",
        "", "", ""
      ]);
    }

    // Si la acción es una calificación, busca el lote por ID y actualiza las últimas columnas.
    if (accion === "calificar") {
      const filas = hoja.getDataRange().getValues();
      for (let i = 1; i < filas.length; i++) {
        if (filas[i][0] === lote.id) {
          hoja.getRange(i + 1, 14).setValue(lote.calificacion);
          hoja.getRange(i + 1, 15).setValue(lote.comentario_calificacion || "");
          hoja.getRange(i + 1, 16).setValue(lote.fecha_calificacion || "");
          break;
        }
      }
    }

    return ContentService.createTextOutput(JSON.stringify({ status: "ok" }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ status: "error", msg: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// Convierte el objeto {cantidad, unidad} a texto legible para la planilla.
function formatearCarga(carga) {
  if (!carga || parseFloat(carga.cantidad) === 0) return "";
  return carga.cantidad + " " + carga.unidad;
}

// Endpoint GET de prueba para verificar que el Web App está activo.
function doGet() {
  return ContentService.createTextOutput("Bitácora CAR · endpoint activo");
}
