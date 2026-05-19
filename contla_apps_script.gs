// ============================================================
// RIFA MUNDIALISTA 2026 — CeluCenter Contla
// Funciones para la rifa de Contla (se integran al doPost principal)
//
// Usa la misma pestana "Rifa" y numeracion RIFA-XXXX consecutiva
// Sucursal = "Contla", Direccion va en campo Ticket
// ============================================================

// =============================================
// REGISTRAR PARTICIPANTE — CONTLA
// =============================================
function registrarContla(params) {
  const nombre = (params.nombre || '').trim();
  const telefono = (params.telefono || '').trim();
  const direccion = (params.direccion || '').trim();

  if (nombre.length < 3) {
    return jsonResponse({ success: false, message: 'El nombre debe tener al menos 3 caracteres' });
  }
  if (!/^\d{10}$/.test(telefono)) {
    return jsonResponse({ success: false, message: 'El telefono debe tener exactamente 10 digitos' });
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hoja = ss.getSheetByName(HOJA_RIFA);

  if (!hoja) {
    return jsonResponse({ success: false, message: 'Error interno: hoja no encontrada.' });
  }

  const datos = hoja.getDataRange().getValues();

  // Verificar duplicado por telefono + sucursal Contla
  for (let i = 1; i < datos.length; i++) {
    if (String(datos[i][2]) === telefono && String(datos[i][4]) === 'Contla') {
      return jsonResponse({
        success: false,
        message: 'Este telefono ya fue registrado. Solo puedes participar una vez. Consulta tus boletos en la seccion de abajo.'
      });
    }
  }

  // Obtener siguiente numero de boleto (secuencial con todos los demas)
  const siguienteNumero = datos.length; // fila 1 = headers, fila 2 = #1
  const tieneDireccion = direccion.length > 0;
  const cantBoletos = tieneDireccion ? 2 : 1;

  const timestamp = Utilities.formatDate(new Date(), 'America/Mexico_City', 'yyyy-MM-dd HH:mm:ss');

  // Boleto 1 — siempre
  const boleto1 = 'RIFA-' + String(siguienteNumero).padStart(4, '0');
  hoja.appendRow([timestamp, nombre, telefono, '', 'Contla', direccion || 'QR Contla', boleto1]);

  // Boleto 2 — solo si dio direccion
  var boleto2 = null;
  if (tieneDireccion) {
    const siguienteNumero2 = siguienteNumero + 1;
    boleto2 = 'RIFA-' + String(siguienteNumero2).padStart(4, '0');
    hoja.appendRow([timestamp, nombre, telefono, '', 'Contla', 'Bonus direccion', boleto2]);
  }

  return jsonResponse({
    success: true,
    nombre: nombre,
    boleto1: boleto1,
    boleto2: boleto2,
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
    if (String(datos[i][2]) === telefono && String(datos[i][4]) === 'Contla') {
      nombre = datos[i][1];
      boletos.push(String(datos[i][6]));
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
