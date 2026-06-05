/**
 * ============================================================
 *  GOOGLE APPS SCRIPT – Bitácora Trans Ráfagas del Golfo
 * ============================================================
 *
 *  PASOS PARA INSTALAR:
 *
 *  1. Abre Google Sheets en tu cuenta → crea una hoja nueva.
 *  2. En el menú superior: Extensiones → Apps Script
 *  3. Borra el código que trae y pega TODO este archivo.
 *  4. Guarda (Ctrl+S). Ponle nombre: "BitacoraRafagas"
 *  5. Click en "Implementar" → "Nueva implementación"
 *  6. Tipo: "Aplicación web"
 *     - Ejecutar como: "Yo"
 *     - Quién tiene acceso: "Cualquier usuario" (o "Cualquiera, incluso anónimos")
 *  7. Haz click en "Implementar" → Copia la URL que aparece.
 *  8. Abre bitacora_rafagas.html en un editor de texto.
 *  9. Busca la línea: const SHEETS_URL = "PEGA_AQUI_TU_URL_..."
 *     y reemplaza ese texto con la URL copiada.
 * 10. Guarda el HTML y súbelo a donde lo necesites (hosting, etc.)
 *
 * ============================================================
 */

const SHEET_NAME = "Bitácora";   // Nombre de la pestaña de tu hoja
const DRIVE_FOLDER = "Fotos_Bitacora";  // Carpeta en Drive para guardar fotos

// ────────────────────────────────────────────────────────────
//  doPost – recibe los datos del formulario
// ────────────────────────────────────────────────────────────
function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const ss   = SpreadsheetApp.getActiveSpreadsheet();
    let sheet  = ss.getSheetByName(SHEET_NAME);

    // Crear hoja si no existe y poner encabezados
    if (!sheet) {
      sheet = ss.insertSheet(SHEET_NAME);
      agregarEncabezados(sheet);
    } else if (sheet.getLastRow() === 0) {
      agregarEncabezados(sheet);
    }

    // Guardar fotos en Drive y obtener URLs
    const fotoUrls = guardarFotos(data.folio, data.fotos || []);
    delete data.fotos;  // no meter base64 en la hoja

    // Construir fila
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const fila    = headers.map(h => {
      if (h === "FotoURLs") return fotoUrls.join(", ");
      return data[h] !== undefined ? data[h] : "";
    });

    sheet.appendRow(fila);

    return ContentService
      .createTextOutput(JSON.stringify({ ok: true, folio: data.folio }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ ok: false, error: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ────────────────────────────────────────────────────────────
//  Encabezados de la hoja
// ────────────────────────────────────────────────────────────
function agregarEncabezados(sheet) {
  const headers = [
    "folio","tipo","fecha","hora","guardia",
    "placaTractor","placa1Remolque","placa2Remolque",
    "conductor","origen","destino",
    "kilometraje","diesel","aceite",
    "estadoGeneral","documentacion","observaciones",
    "firmGuardia","firmConductor",
    // Llantas 1-36
    ...Array.from({length:36}, (_, i) => `llanta_${i+1}_estado`),
    ...Array.from({length:36}, (_, i) => `obs_llanta_${i+1}`),
    "FotoURLs",
    "timestamp"
  ];

  sheet.appendRow(headers);

  // Estilo de encabezados
  const rango = sheet.getRange(1, 1, 1, headers.length);
  rango.setBackground("#1a1a1a");
  rango.setFontColor("#ffffff");
  rango.setFontWeight("bold");
  sheet.setFrozenRows(1);
}

// ────────────────────────────────────────────────────────────
//  Guardar fotos en Google Drive
// ────────────────────────────────────────────────────────────
function guardarFotos(folio, fotos) {
  if (!fotos || fotos.length === 0) return [];

  let carpeta;
  const carpetas = DriveApp.getFoldersByName(DRIVE_FOLDER);
  carpeta = carpetas.hasNext() ? carpetas.next() : DriveApp.createFolder(DRIVE_FOLDER);

  return fotos.map((foto, i) => {
    try {
      // foto.data es "data:image/jpeg;base64,..."
      const partes    = foto.data.split(",");
      const mimeType  = partes[0].match(/:(.*?);/)[1];
      const bytes     = Utilities.base64Decode(partes[1]);
      const blob      = Utilities.newBlob(bytes, mimeType, `${folio}_foto${i+1}`);
      const archivo   = carpeta.createFile(blob);
      archivo.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      return archivo.getUrl();
    } catch(e) {
      return "Error al guardar foto";
    }
  });
}

// ────────────────────────────────────────────────────────────
//  doGet – prueba rápida
// ────────────────────────────────────────────────────────────
function doGet(e) {
  return ContentService
    .createTextOutput("✅ Script de Bitácora Ráfagas activo")
    .setMimeType(ContentService.MimeType.TEXT);
}
