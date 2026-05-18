// ============================================================
// RIFA MUNDIALISTA 2026 — CeluCenter Contla
// Backend Google Apps Script
//
// Google Sheet: https://docs.google.com/spreadsheets/d/16ISKttW9JEHOjUgKfVWso3jDI4uJ7Hc_KTLxpITk6C0/
// Pestana: "Rifa Contla"
//
// SETUP:
// 1. Abrir el Sheet de arriba > Extensiones > Apps Script
// 2. Crear un nuevo archivo .gs y pegar este codigo
// 3. Ejecutar setupHojaRifaContla() una vez
// 4. Deploy > New deployment > Web app
//    - Execute as: Me
//    - Who has access: Anyone
// 5. Copiar URL del deploy y pegarla en contla/index.html
// ============================================================

// =============================================
// CONFIG
// =============================================
const HOJA_RIFA_CONTLA = 'Rifa Contla';

// =============================================
// SETUP — Ejecutar una vez para crear headers
// =============================================
function setupHojaRifaContla() {
  const ss = SpreadsheetApp.openById('16ISKttW9JEHOjUgKfVWso3jDI4uJ7Hc_KTLxpITk6C0');
  let hoja = ss.getSheetByName(HOJA_RIFA_CONTLA);

  if (!hoja) {
    hoja = ss.insertSheet(HOJA_RIFA_CONTLA);
  }

  // Headers
  const headers = ['Timestamp', 'Nombre', 'Telefono', 'Direccion', 'Boletos', 'Boleto1', 'Boleto2'];
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
  hoja.setColumnWidth(4, 300); // Direccion
  hoja.setColumnWidth(5, 80);  // Boletos (cantidad)
  hoja.setColumnWidth(6, 140); // Boleto1
  hoja.setColumnWidth(7, 140); // Boleto2

  // Congelar header
  hoja.setFrozenRows(1);

  SpreadsheetApp.getUi().alert('Hoja "Rifa Contla" configurada correctamente.');
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

  if (action === 'registrar_contla') {
    return registrarContla(params);
  }

  if (action === 'consultar_boletos_contla') {
    return consultarBoletosContla(params);
  }

  return jsonResponseContla({ success: false, message: 'Accion no valida' });
}

function doGet(e) {
  return jsonResponseContla({ success: false, message: 'Usa POST para interactuar con la rifa' });
}

// =============================================
// REGISTRAR PARTICIPANTE — CONTLA
// =============================================
function registrarContla(params) {
  const nombre = (params.nombre || '').trim();
  const telefono = (params.telefono || '').trim();
  const direccion = (params.direccion || '').trim();

  // Validaciones
  if (nombre.length < 3) {
    return jsonResponseContla({ success: false, message: 'El nombre debe tener al menos 3 caracteres' });
  }
  if (!/^\d{10}$/.test(telefono)) {
    return jsonResponseContla({ success: false, message: 'El telefono debe tener exactamente 10 digitos' });
  }

  const ss = SpreadsheetApp.openById('16ISKttW9JEHOjUgKfVWso3jDI4uJ7Hc_KTLxpITk6C0');
  const hoja = ss.getSheetByName(HOJA_RIFA_CONTLA);

  if (!hoja) {
    return jsonResponseContla({ success: false, message: 'Error interno: hoja no encontrada. Ejecuta setupHojaRifaContla() primero.' });
  }

  const datos = hoja.getDataRange().getValues();

  // Verificar duplicado por telefono
  for (let i = 1; i < datos.length; i++) {
    if (String(datos[i][2]) === telefono) {
      return jsonResponseContla({
        success: false,
        message: 'Este telefono ya fue registrado. Solo puedes participar una vez. Consulta tus boletos en la seccion de abajo.'
      });
    }
  }

  // Contar boletos existentes para generar secuencial
  let maxBoleto = 0;
  for (let i = 1; i < datos.length; i++) {
    // Revisar Boleto1
    const b1 = String(datos[i][5]);
    const m1 = b1.match(/^CONTLA-(\d+)$/);
    if (m1) {
      const n = parseInt(m1[1]);
      if (n > maxBoleto) maxBoleto = n;
    }
    // Revisar Boleto2
    const b2 = String(datos[i][6]);
    const m2 = b2.match(/^CONTLA-(\d+)$/);
    if (m2) {
      const n = parseInt(m2[1]);
      if (n > maxBoleto) maxBoleto = n;
    }
  }

  // Determinar cantidad de boletos
  const tieneDireccion = direccion.length > 0;
  const cantBoletos = tieneDireccion ? 2 : 1;

  // Generar boletos
  const boleto1 = 'CONTLA-' + String(maxBoleto + 1).padStart(4, '0');
  const boleto2 = tieneDireccion ? 'CONTLA-' + String(maxBoleto + 2).padStart(4, '0') : '';

  // Timestamp
  const timestamp = Utilities.formatDate(new Date(), 'America/Mexico_City', 'yyyy-MM-dd HH:mm:ss');

  // Guardar en Sheet
  hoja.appendRow([timestamp, nombre, telefono, direccion, cantBoletos, boleto1, boleto2]);

  return jsonResponseContla({
    success: true,
    nombre: nombre,
    boleto1: boleto1,
    boleto2: boleto2 || null,
    cantBoletos: cantBoletos,
    message: 'Registro exitoso. Tienes ' + cantBoletos + ' boleto(s).'
  });
}

// =============================================
// CONSULTAR BOLETOS POR TELEFONO — CONTLA
// =============================================
function consultarBoletosContla(params) {
  const telefono = (params.telefono || '').trim();

  if (!/^\d{10}$/.test(telefono)) {
    return jsonResponseContla({ success: false, message: 'Telefono invalido' });
  }

  const ss = SpreadsheetApp.openById('16ISKttW9JEHOjUgKfVWso3jDI4uJ7Hc_KTLxpITk6C0');
  const hoja = ss.getSheetByName(HOJA_RIFA_CONTLA);

  if (!hoja) {
    return jsonResponseContla({ success: false, message: 'Error interno' });
  }

  const datos = hoja.getDataRange().getValues();
  var boletos = [];
  var nombre = '';

  for (let i = 1; i < datos.length; i++) {
    if (String(datos[i][2]) === telefono) {
      nombre = datos[i][1];
      if (datos[i][5]) boletos.push(String(datos[i][5]));
      if (datos[i][6]) boletos.push(String(datos[i][6]));
    }
  }

  if (boletos.length === 0) {
    return jsonResponseContla({ success: false, message: 'No se encontraron boletos con ese telefono' });
  }

  return jsonResponseContla({
    success: true,
    nombre: nombre,
    boletos: boletos
  });
}

// =============================================
// HELPER: JSON Response
// =============================================
function jsonResponseContla(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}
