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

  // Buscar si el telefono ya esta registrado (login automatico)
  for (let i = 1; i < datos.length; i++) {
    if (String(datos[i][2]) === telefono) {
      return jsonResponse({
        success: true,
        boleto: datos[i][6],
        nombre: datos[i][1],
        existing: true,
        message: 'Ya estas registrado. Aqui esta tu numero de boleto.'
      });
    }
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
// HELPER: JSON Response
// =============================================
function jsonResponse(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}
