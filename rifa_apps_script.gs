// ============================================================
// RIFA MUNDIALISTA 2026 — CeluCenter
// Backend Google Apps Script
//
// Google Sheet: https://docs.google.com/spreadsheets/d/16ISKttW9JEHOjUgKfVWso3jDI4uJ7Hc_KTLxpITk6C0/
//
// SETUP:
// 1. Abrir el Sheet de arriba > Extensiones > Apps Script
// 2. Pegar este código
// 3. Ejecutar setupHojaRifa() una vez
// 4. Deploy > New deployment > Web app
//    - Execute as: Me
//    - Who has access: Anyone
// 5. Copiar URL del deploy y pegarla en rifa_mundial_celucenter.html
// ============================================================

// =============================================
// CONFIG
// =============================================
const HOJA_RIFA = 'Rifa';
const MACHOTE_SHEET_ID = '19-QtiBnvjUWJ22KgqAXAQdhcN7_PkGEzkUPDZjJRch4';
const MACHOTE_TAB = 'General';
const MACHOTE_COL_FOLIO = 2; // Columna B

// =============================================
// SETUP — Ejecutar una vez para crear headers
// =============================================
function setupHojaRifa() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let hoja = ss.getSheetByName(HOJA_RIFA);

  if (!hoja) {
    hoja = ss.insertSheet(HOJA_RIFA);
  }

  // Headers
  const headers = ['Timestamp', 'Nombre', 'Telefono', 'Email', 'Sucursal', 'Ticket', 'Boleto'];
  hoja.getRange(1, 1, 1, headers.length).setValues([headers]);

  // Formato headers
  const headerRange = hoja.getRange(1, 1, 1, headers.length);
  headerRange.setFontWeight('bold');
  headerRange.setBackground('#00a651');
  headerRange.setFontColor('#ffffff');
  headerRange.setHorizontalAlignment('center');

  // Ancho de columnas
  hoja.setColumnWidth(1, 180); // Timestamp
  hoja.setColumnWidth(2, 200); // Nombre
  hoja.setColumnWidth(3, 130); // Telefono
  hoja.setColumnWidth(4, 220); // Email
  hoja.setColumnWidth(5, 180); // Sucursal
  hoja.setColumnWidth(6, 150); // Ticket
  hoja.setColumnWidth(7, 120); // Boleto

  // Congelar header
  hoja.setFrozenRows(1);

  SpreadsheetApp.getUi().alert('Hoja "Rifa" configurada correctamente.');
}

// =============================================
// WEB APP ENDPOINTS
// =============================================
function doPost(e) {
  var params;
  try {
    params = JSON.parse(e.postData.contents);
  } catch(err) {
    params = e.parameter;
  }

  const action = params.action;

  if (action === 'registrar_rifa') {
    return registrarRifa(params);
  }

  if (action === 'consultar_boletos') {
    return consultarBoletos(params);
  }

  return jsonResponse({ success: false, message: 'Accion no valida' });
}

function doGet(e) {
  return jsonResponse({ success: false, message: 'Usa POST para interactuar con la rifa' });
}

// =============================================
// REGISTRAR PARTICIPANTE
// =============================================
function registrarRifa(params) {
  const nombre = (params.nombre || '').trim();
  const telefono = (params.telefono || '').trim();
  const email = (params.email || '').trim();
  const sucursal = (params.sucursal || '').trim();
  const ticket = (params.ticket || '').trim();

  // Validaciones
  if (nombre.length < 3) {
    return jsonResponse({ success: false, message: 'El nombre debe tener al menos 3 caracteres' });
  }
  if (!/^\d{10}$/.test(telefono)) {
    return jsonResponse({ success: false, message: 'El telefono debe tener exactamente 10 digitos' });
  }
  if (!email.includes('@')) {
    return jsonResponse({ success: false, message: 'Ingresa un email valido' });
  }
  if (!sucursal) {
    return jsonResponse({ success: false, message: 'Selecciona una sucursal' });
  }
  if (ticket.length < 4) {
    return jsonResponse({ success: false, message: 'El ticket debe tener al menos 4 caracteres' });
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hoja = ss.getSheetByName(HOJA_RIFA);

  if (!hoja) {
    return jsonResponse({ success: false, message: 'Error interno: hoja no encontrada. Ejecuta setupHojaRifa() primero.' });
  }

  const datos = hoja.getDataRange().getValues();

  // Validar que el ticket exista en el machote de ventas (folios reales)
  if (!validarFolioMachote(ticket)) {
    return jsonResponse({
      success: false,
      message: 'Este numero de ticket no es valido. Verifica que sea el folio correcto de tu ticket de compra.'
    });
  }

  // Verificar que el ticket no haya sido usado
  for (let i = 1; i < datos.length; i++) {
    if (String(datos[i][5]).toLowerCase() === ticket.toLowerCase()) {
      return jsonResponse({
        success: false,
        message: 'Este ticket de compra ya fue registrado por otro participante.'
      });
    }
  }

  // Generar numero de boleto secuencial
  const siguienteNumero = datos.length; // fila 1 = headers, fila 2 = #1, etc.
  const boleto = 'RIFA-' + String(siguienteNumero).padStart(4, '0');

  // Timestamp
  const timestamp = Utilities.formatDate(new Date(), 'America/Mexico_City', 'yyyy-MM-dd HH:mm:ss');

  // Guardar en Sheet
  hoja.appendRow([timestamp, nombre, telefono, email, sucursal, ticket, boleto]);

  return jsonResponse({
    success: true,
    boleto: boleto,
    nombre: nombre,
    existing: false,
    message: 'Registro exitoso. Tu numero de boleto es ' + boleto
  });
}

// =============================================
// CONSULTAR BOLETOS POR TELEFONO
// =============================================
function consultarBoletos(params) {
  const telefono = (params.telefono || '').trim();

  if (!/^\d{10}$/.test(telefono)) {
    return jsonResponse({ success: false, message: 'Telefono invalido' });
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hoja = ss.getSheetByName(HOJA_RIFA);

  if (!hoja) {
    return jsonResponse({ success: false, message: 'Error interno' });
  }

  const datos = hoja.getDataRange().getValues();
  var boletos = [];
  var nombre = '';

  for (let i = 1; i < datos.length; i++) {
    if (String(datos[i][2]) === telefono) {
      nombre = datos[i][1];
      boletos.push({
        boleto: datos[i][6],
        sucursal: datos[i][4],
        ticket: datos[i][5]
      });
    }
  }

  if (boletos.length === 0) {
    return jsonResponse({ success: false, message: 'No se encontraron boletos con ese telefono' });
  }

  return jsonResponse({
    success: true,
    nombre: nombre,
    boletos: boletos
  });
}

// =============================================
// VALIDAR FOLIO CONTRA MACHOTE DE VENTAS
// =============================================
function validarFolioMachote(ticket) {
  try {
    const machote = SpreadsheetApp.openById(MACHOTE_SHEET_ID);
    const hoja = machote.getSheetByName(MACHOTE_TAB);
    if (!hoja) return false;

    const lastRow = hoja.getLastRow();
    if (lastRow < 2) return false;

    const folios = hoja.getRange(2, MACHOTE_COL_FOLIO, lastRow - 1, 1).getValues();
    const ticketLower = ticket.toLowerCase().trim();

    for (let i = 0; i < folios.length; i++) {
      if (String(folios[i][0]).toLowerCase().trim() === ticketLower) {
        return true;
      }
    }
    return false;
  } catch(err) {
    // Si no puede acceder al machote, dejar pasar (para no bloquear)
    Logger.log('Error validando folio: ' + err);
    return true;
  }
}

// =============================================
// HELPER: JSON Response
// =============================================
function jsonResponse(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}
