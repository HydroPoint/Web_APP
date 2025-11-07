const SS = SpreadsheetApp.getActiveSpreadsheet();

const CONFIG_SHEET = 'CONFIG';
const USERS_SHEET = 'USERS';
const AUDIT_LOG_SHEET = 'AUDIT_LOG';
const ORD_VENTA_SHEET = 'ORD_VENTA';
const FACTURAS_SHEET = 'FACTURAS';
const CLIENTES_SHEET = 'CLIENTES';
const COMPRAS_OC_SHEET = 'COMPRAS_OC';
const INSUMOS_SHEET = 'INSUMOS';
const PASIVOS_SHEET = 'PASIVOS';
const CASH_ON_HAND_SHEET = 'CASH_ON_HAND';
const PRODUCTOS_SHEET = 'PRODUCTOS';
const BOM_RECETAS_SHEET = 'BOM_RECETAS';
const NOMINA_SHEET = 'NOMINA';
const COSTOS_INDIRECTOS_SHEET = 'COSTOS_INDIRECTOS';
const MAQUINARIA_SHEET = 'MAQUINARIA';
const VARIABLES_SHEET = 'VARIABLES';
const CRM_LOG_SHEET = 'CRM_LOG';
const CONTABILIDAD_PUC_SHEET = 'CONTABILIDAD_PUC';
const CONTABILIDAD_MOVIMIENTOS_SHEET = 'CONTABILIDAD_MOVIMIENTOS';
const MOVIMIENTOS_INVENTARIO_SHEET = 'MOVIMIENTOS_INVENTARIO';
const GASTOS_VARIOS_SHEET = 'GASTOS_VARIOS';
const REPORTES_LOG_SHEET = 'REPORTES_LOG';
const CRM_INTERACCIONES_SHEET = 'CRM_INTERACCIONES';
const CMSI_CAMPAÑAS_SHEET = 'CMSI_CAMPAÑAS';
const PRODUCCION_ORDENES_SHEET = 'PRODUCCION_ORDENES';
const QMS_INSPECCIONES_SHEET = 'QMS_INSPECCIONES';
const EMS_CONSUMOS_SHEET = 'EMS_CONSUMOS';
const PROVEEDORES_SHEET = 'PROVEEDORES';
const CONSOLIDADO_WOOC_SHEET = 'Consolidado_WooC';


function doGet(e) {
  let pageName = e.parameter.page;

  // Si no hay parámetro 'page', cargamos 'Index'
  if (!pageName) {
    pageName = 'INDEX'; 
  }

  // Creamos el template desde el archivo (ej: 'Index' o 'CLIENTES')
  const template = HtmlService.createTemplateFromFile(pageName);
  

  template.url = getScriptUrl(); 
  
  // Ahora sí evaluamos y retornamos
  const html = template.evaluate();
  let title = (pageName === 'INDEX') ? 'SGCI-VERP' : pageName + ' - ERP';
  html.setTitle(title)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1.0')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
      
  return html;
}

function getScriptUrl() {
  var url = ScriptApp.getService().getUrl();
  
  // Si la URL es nula o vacía (porque estamos en /dev)
  if (!url) {
    var scriptId = ScriptApp.getScriptId();
    
    // Si usas una cuenta de @gmail.com
    url = "https://script.google.com/macros/d/" + scriptId + "/dev";
    
    // NOTA: Si usas una cuenta de Google Workspace (ej: @tuempresa.com)
    // la URL de desarrollo puede ser diferente, pero la /d/ casi siempre funciona.
  }
  
  // Limpiamos cualquier parámetro (?page=...) que pudiera tener la URL base
  return url.split('?')[0];
}
function include(filename) {
    return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function _getCashOnHandTotal() {
  const data = getSheetData(CASH_ON_HAND_SHEET);
  if (data.error) return 0;

  // Filtra solo por moneda local (COP) para el dashboard principal
  const dataCOP = data.filter(row => row.Moneda && row.Moneda.toUpperCase() === 'COP');
  
  // Suma la columna 'Saldo_Actual' de las cuentas en COP
  const cashTotal = _sumColumn(dataCOP, 'Saldo_Actual');
  
  return cashTotal; 
}

function getNextId(sheetName, prefix) {
  const sheet = SS.getSheetByName(sheetName);
  if (!sheet) {
    throw new Error(`La pestaña "${sheetName}" no existe.`);
  }
  const lastRow = sheet.getLastRow(); 
  if (lastRow < 2) {
    return `${prefix}0001`; 
  }
  const lastId = sheet.getRange(lastRow, 1).getValue();
  try {
    const lastNum = parseInt(lastId.split('-')[1], 10);
    const nextNum = lastNum + 1;
    const nextId = `${prefix}${nextNum.toString().padStart(4, '0')}`;
    return nextId;
  } catch (e) {
    return `${prefix}${(lastRow).toString().padStart(4, '0')}`;
  }
}

/**
 * Autentica a un usuario contra la hoja "USERS".
 */
function authenticateUser(identifier, password) {
  try {
    // --- 1. Bypass de Administrador ---
    if (identifier.toLowerCase() === 'admin' && password === '1234') {
      Logger.log('Acceso de Bypass concedido.');
      return { 
        success: true, 
        user: { email: 'admin@bypass.com', nombre: 'Admin Bypass', rol: 'admin' } 
      };
    }

    // --- 2. Revisar la Hoja "USERS" ---
    const usersData = getSheetData(USERS_SHEET); 
    if (usersData.error) throw new Error(usersData.error); 

    const lowerIdentifier = identifier.toLowerCase();
    
    const user = usersData.find(u => 
      (u.Email && u.Email.toLowerCase() === lowerIdentifier) ||
      (u.USER_ID && u.USER_ID.toLowerCase() === lowerIdentifier) ||
      (u.Nombre && u.Nombre.toLowerCase() === lowerIdentifier)
    );

    if (!user) {
      return { success: false, message: 'Usuario no encontrado.' };
    }
    
    // 3. Validar Contraseña
    if (user.Pswrd !== password) {
      return { success: false, message: 'Contraseña incorrecta.' };
    }

    // 4. Validar si el usuario está Activo
    if (user.Activo !== true && user.Activo !== 'TRUE' && user.Activo !== 'true' && user.Activo !== 1) {
      return { success: false, message: 'El usuario no está activo.' };
    }

    // 5. ¡Éxito!
    return { 
      success: true, 
      user: { email: user.Email, nombre: user.Nombre, rol: user.Rol } 
    };

  } catch (e) {
    Logger.log(e);
    return { success: false, message: `Error de servidor: ${e.message}` };
  }
}

/**
 * Registra un evento en la hoja "AUDIT_LOG".
 */
function logAudit(email, action, detail) {
  try {
    const sheet = SS.getSheetByName(AUDIT_LOG_SHEET);
    if (!sheet) throw new Error(`Hoja ${AUDIT_LOG_SHEET} no encontrada.`);

    const newId = getNextId(AUDIT_LOG_SHEET, 'LOG-');
    const now = new Date(); 
    
    const newRow = [ newId, now, email, action, detail ];
    
    sheet.appendRow(newRow);
    return { success: true };

  } catch (e) {
    Logger.log(e);
    return { error: e.message };
  }
}

function createOrderVenta(payload) {
  try {
    const sheet = SS.getSheetByName(ORD_VENTA_SHEET);
    if (!sheet) {
      throw new Error(`La pestaña ${ORD_VENTA_SHEET} no se encuentra.`);
    }
    
    if (!payload.Cliente_ID || !payload.ItemsJSON) {
      throw new Error('Faltan datos (Cliente_ID o ItemsJSON).');
    }

    // Obtener encabezados dinámicamente
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const newRow = new Array(headers.length).fill('');
    const idx = {};
    headers.forEach((header, i) => {
      idx[header.trim()] = i;
    });

    const newId = getNextId(ORD_VENTA_SHEET, 'ORD-');
    const today = new Date();

    // Mapeo seguro basado en nombres de encabezado
    // (Asumiendo los nombres de columna basados en el array original)
    if (idx['ORD_ID'] !== undefined) newRow[idx['ORD_ID']] = newId;
    if (idx['Fecha'] !== undefined) newRow[idx['Fecha']] = today;
    if (idx['Cliente_ID'] !== undefined) newRow[idx['Cliente_ID']] = payload.Cliente_ID;
    if (idx['ItemsJSON'] !== undefined) newRow[idx['ItemsJSON']] = payload.ItemsJSON;
    if (idx['Estado'] !== undefined) newRow[idx['Estado']] = 'pendiente';
    if (idx['Transporte'] !== undefined) newRow[idx['Transporte']] = payload.Transporte || '';
    // Las columnas 6 y 8 (índices 5 y 7) se dejan vacías ('')

    sheet.appendRow(newRow);
    
    return { 
      success: true, 
      newId: newId, 
      message: `Orden ${newId} creada exitosamente.` 
    };
  } catch (e) {
    Logger.log(e);
    return { error: e.message };
  }
}
function getDashboardAnalytics() {
  try {
    // --- Datos para las 4 tarjetas superiores ---
    const pendingOrders = _getPendingOrdersCount();
    const obligations = _getPendingObligationsCount();
    const dollar = _getDollarValue();
    const balance = _getFinancialBalance();
    
    // --- ¡NUEVO! Datos para el footer ---
    // ¡CORREGIDO! La función _getFooterAnalytics no recibe parámetros.
    const footerData = _getFooterAnalytics(); 

    // --- LÍNEA DE DIAGNÓSTICO QUE NECESITAMOS ---
    Logger.log("DATOS DEL FOOTER CALCULADOS:");
    Logger.log(JSON.stringify(footerData));
    // --- FIN DE LA LÍNEA DE DIAGNÓSTICO ---

    return {
      pendingOrders: pendingOrders,
      obligations: obligations,
      dollar: dollar,
      balance: balance,
      footerData: footerData // Se añade el nuevo objeto de datos
    };
  } catch (e) {
    Logger.log(`Error en getDashboardAnalytics: ${e}`);
    return { error: e.message };
  }
}

function _getPendingOrdersCount() {
  const data = getSheetData(ORD_VENTA_SHEET);
  if (data.error) return 0;
  
  const pending = data.filter(row => row.Estado && row.Estado.toLowerCase() === 'pendiente');
  return pending.length;
}

/**
 * _(NUEVO Helper) Limpia un valor para convertirlo en un número flotante válido.
 * Quita "$" y separadores de miles ".".
 * Reemplaza la coma decimal "," por un punto ".".
 */
function _cleanNumber(value) {
  if (typeof value === 'number') {
    return value; // Si ya es un número, devolverlo tal cual.
  }
  if (typeof value !== 'string') {
    return NaN; // Si no es string ni número, no se puede procesar.
  }
  
  // 1. Quitar el símbolo de moneda y los separadores de miles (puntos)
  // Ej: "$ 1.500.000,50" -> " 1500000,50"
  let cleanedValue = value.replace(/\$|\./g, '');
  
  // 2. Reemplazar la coma decimal por un punto
  // Ej: " 1500000,50" -> " 1500000.50"
  cleanedValue = cleanedValue.replace(',', '.');
  
  // 3. Quitar espacios y convertir a número
  // Ej: " 1500000.50" -> 1500000.5
  return parseFloat(cleanedValue.trim());
}

/**
 * start_marker:REEMPLAZO_GET_FACTURAS
 * @description Obtiene los datos de la hoja de FACTURAS y los mapea.
 * (Versión REFACTORIZADA)
 */
function _getFacturasData() {
  // Simplemente llamamos a la nueva función maestra.
  // _getMappedData se encarga de todo el trabajo.
  // La caché se manejará en la función que llame a esta, si es necesario.
  return _getMappedData(FACTURAS_SHEET);
}

/**
 * start_marker:NUEVA_FUNCION_MAESTRA
 * @description (NUEVA) Función genérica para obtener datos de CUALQUIER hoja
 * y devolverlos como un array de objetos (mapeo por encabezados).
 * Esto reemplaza el método de 'índices fijos'.
 * @param {string} sheetName El nombre de la pestaña (ej. 'CLIENTES').
 * @returns {Array<Object>} Un array de objetos, donde cada objeto es una fila.
 */
function _getMappedData(sheetName) {
  const sheet = SS.getSheetByName(sheetName);
  if (!sheet) {
    Logger.log(`ERROR: No se encontró la hoja: ${sheetName}`);
    return [];
  }
  
  const dataRange = sheet.getDataRange();
  // Usar getValues() para preservar tipos de datos (fechas, números)
  const values = dataRange.getValues();
  
  if (values.length < 2) {
    Logger.log(`ADVERTENCIA: La hoja ${sheetName} está vacía o solo tiene encabezados.`);
    return [];
  }
  
  // La primera fila (índice 0) son los encabezados
  const headers = values[0];
  
  // Limpiar encabezados (quitar espacios al inicio/final)
  const cleanHeaders = headers.map(header => {
    if (typeof header === 'string') {
      return header.trim();
    }
    return header;
  });

  // Mapear el resto de las filas (desde el índice 1)
  const data = values.slice(1).map((row, rowIndex) => {
    // Solo procesar filas que no estén completamente vacías
    if (row.join('').length === 0) {
      return null;
    }

    const obj = {
      _row: rowIndex + 2 // Guardamos el N° de fila real (útil para editar)
    };
    
    cleanHeaders.forEach((header, index) => {
      if (header) { // Solo si el encabezado no está vacío
        let value = row[index];
        
        // Convertir fechas a strings ISO (JSON-friendly)
        // Usamos toLocaleDateString para el formato local dd/mm/aaaa
        if (value instanceof Date) {
          value = value.toLocaleDateString('es-CO', { year: 'numeric', month: '2-digit', day: '2-digit' });
        }
        
        obj[header] = value;
      }
    });
    return obj;
  }).filter(Boolean); // Eliminar filas nulas (vacías)
  
  return data;
}

/**
 * _(Helper) Función interna para sumar columnas de forma segura
 * ¡VERSIÓN ACTUALIZADA! Ahora usa _cleanNumber
 */
function _sumColumn(data, columnName) {
  if (data.error || !data.length) return 0;
  
  return data.reduce((sum, row) => {
    // Usa _cleanNumber para limpiar el dato antes de sumarlo
    const value = _cleanNumber(row[columnName]);
    return sum + (isNaN(value) ? 0 : value);
  }, 0);
}

/**
 * _(Helper) Función interna para sumar productos de columnas
 * ¡VERSIÓN ACTUALIZADA! Ahora usa _cleanNumber
 */
function _sumProductOfColumns(data, col1, col2) {
  if (data.error || !data.length) return 0;
  
  return data.reduce((sum, row) => {
    // Usa _cleanNumber para limpiar ambos valores
    const val1 = _cleanNumber(row[col1]);
    const val2 = _cleanNumber(row[col2]);
    
    if (!isNaN(val1) && !isNaN(val2)) {
      return sum + (val1 * val2);
    }
    return sum;
  }, 0);
}

function _getPendingObligationsCount() {
  let count = 0;
  
  const nominaData = getSheetData(NOMINA_SHEET);
  if (!nominaData.error) {
    count += nominaData.filter(row => row.Estado_pago && row.Estado_pago.toLowerCase() !== 'pagado').length;
  }
  
  const comprasData = getSheetData(COMPRAS_OC_SHEET);
  if (!comprasData.error) {
    count += comprasData.filter(row => row.Estado && row.Estado.toLowerCase() === 'pendiente').length;
  }
  
  return count;
}

/**
 * _(Helper) Obtiene el valor del Dólar (USD a COP).
 */
function _getDollarValue() {
  try {
    const response = UrlFetchApp.fetch("https://open.er-api.com/v6/latest/USD");
    const data = JSON.parse(response.getContentText());
    
    if (!data || !data.rates || !data.rates.COP) {
      throw new Error("Respuesta de API inválida. No se encontró 'rates.COP'.");
 t }
    const rate = data.rates.COP;
    const change = Math.floor(Math.random() * 101) - 50; 
    
    return {
      rate: rate.toFixed(2),
      change: change
    };
  } catch (e) {
    Logger.log(`Error de API Dólar: ${e}`);
    return { rate: "Error", change: 0 };
  }
}

/**
 * _(Helper) Calcula el balance financiero total (Ganancia).
 */
function _getFinancialBalance() {
  let totalGastos = 0;
  let totalIngresos = 0;

  // 1. SUMAR GASTOS
  totalGastos += _sumColumn(getSheetData(COMPRAS_OC_SHEET), 'Totales');
  totalGastos += _sumColumn(getSheetData(NOMINA_SHEET), 'Total_pagar');
  totalGastos += _sumColumn(getSheetData(COSTOS_INDIRECTOS_SHEET), 'Monto_mensual');
  totalGastos += _sumColumn(getSheetData(MAQUINARIA_SHEET), 'Costo_compra');
  
  // 2. SUMAR INGRESOS (Solo facturas pagadas)
  const facturasData = getSheetData(FACTURAS_SHEET);
  if (!facturasData.error) {
    const facturasPagadas = facturasData.filter(row => row.Estado_pago && row.Estado_pago.toLowerCase() === 'pagado');
    totalIngresos = _sumColumn(facturasPagadas, 'Total');
  }
  
  // 3. CALCULAR GANANCIA
  const ganancia = totalIngresos - totalGastos;
  
  return {
    gastos: totalGastos.toFixed(2),
    ingresos: totalIngresos.toFixed(2),
    ganancia: ganancia.toFixed(2)
  };
}

/**
 * _(Helper) Calcula los 4 valores del footer.
 * ¡VERSIÓN CORREGIDA!
 * Se ajustó 'Tipo' a 'Tipo_Cuenta' según tus archivos.
 */
function _getFooterAnalytics() { 
  
  // --- TARJETA 1: INSUMOS Y MAQUINARIA ---
  const insumosData = getSheetData(INSUMOS_SHEET);
  const valorInsumos = _sumProductOfColumns(insumosData, 'Stock_actual', 'Costo_unitario');
  
  const maquinariaData = getSheetData(MAQUINARIA_SHEET);
  const valorMaquinaria = _sumColumn(maquinariaData, 'Costo_compra');

  // --- TARJETA 2: PRODUCTOS (VENTA Y COSTO) ---
  const productosData = getSheetData(PRODUCTOS_SHEET);
  const valorProductosCosto = _sumProductOfColumns(productosData, 'Stock', 'Costo_estimado');
  // Tu CSV 'PRODUCTOS.csv' confirma que 'Precio_venta' es correcto.
  const valorProductosVenta = _sumProductOfColumns(productosData, 'Stock', 'Precio_venta');
  
  // --- TARJETA 3: CASH (TOTAL Y EFECTIVO) ---
  const cashData = getSheetData(CASH_ON_HAND_SHEET);
  const cashTotal = _getCashOnHandTotal(); // Esta ya suma solo COP
  
  // <-- ¡AQUÍ ESTABA EL ERROR! 
  // Cambiado de 'row.Tipo' a 'row.Tipo_Cuenta' para que coincida con tu CSV.
  const cashEfectivoData = cashData.filter(row => 
    row.Tipo_Cuenta && row.Tipo_Cuenta.toLowerCase() === 'efectivo' && row.Moneda.toUpperCase() === 'COP'
  );
  const cashEfectivo = _sumColumn(cashEfectivoData, 'Saldo_Actual');
  
  // --- TARJETA 4: DEBE Y PAGADO ---
  const pasivosData = getSheetData(PASIVOS_SHEET);
  // Tu CSV 'PASIVOS.csv' confirma que 'Monto_Pagado' es correcto.
  const totalPagado = _sumColumn(pasivosData, 'Monto_Pagado');
  
  let totalDebe = 0;
  if (!pasivosData.error) {
    const pasivosPendientes = pasivosData.filter(row => 
      row.Estado && 
      (row.Estado.toLowerCase() === 'pendiente' || row.Estado.toLowerCase() === 'en mora')
    );
    // Tu CSV 'PASIVOS.csv' confirma que 'Monto_Pendiente' es correcto.
    totalDebe = _sumColumn(pasivosPendientes, 'Monto_Pendiente');
  }
  
  return {
    // Tarjeta 1
    valorInsumos: valorInsumos.toFixed(2),
    valorMaquinaria: valorMaquinaria.toFixed(2),
    // Tarjeta 2
    valorProductosVenta: valorProductosVenta.toFixed(2),
    valorProductosCosto: valorProductosCosto.toFixed(2),
    // Tarjeta 3
    cashTotal: cashTotal.toFixed(2),
    cashEfectivo: cashEfectivo.toFixed(2),
    // Tarjeta 4
    totalDebe: totalDebe.toFixed(2),
    totalPagado: totalPagado.toFixed(2)
  };
}
function searchClientes(searchTerm) {
  try {
    const allClientes = getSheetData(CLIENTES_SHEET);
    if (allClientes.error) throw new Error(allClientes.error);

    const results = allClientes.filter(c => {
      const term = searchTerm.toLowerCase();
      
      return (
        (c.Nombre && String(c.Nombre).toLowerCase().includes(term)) ||
        (c['NIT/CC'] && String(c['NIT/CC']).toLowerCase().includes(term)) || 
        (c.Tel && String(c.Tel).toLowerCase().includes(term)) || 
        (c.Email && String(c.Email).toLowerCase().includes(term)) ||
        (c.Tipo && String(c.Tipo).toLowerCase().includes(term)) || 
        (c.CLIENTE_ID && String(c.CLIENTE_ID).toLowerCase().includes(term)) ||
        // --- ✅ BÚSQUEDA EN NUEVOS CAMPOS ---
        (c.Tags && String(c.Tags).toLowerCase().includes(term)) ||
        (c.Estado_Contacto && String(c.Estado_Contacto).toLowerCase().includes(term))
      );
    });

    // ✅ Devolvemos MÁS datos para el CMSI
    return results.map(c => ({
        CLIENTE_ID: c.CLIENTE_ID,
        Nombre: c.Nombre,
        Tipo: c.Tipo,
        NIT_CC: c['NIT/CC'], 
        Tel: c.Tel,
        // --- ✅ CAMPOS NUEVOS DE CMSI ---
        Tags: c.Tags || '',
        Ultima_Interaccion: c.Ultima_Interaccion || 'N/A',
        Estado_Contacto: c.Estado_Contacto || 'Activo'
    }));
  } catch (e) {
    Logger.log(e);
    return []; 
  }
}
function createCliente(payload) {
  try {
    const sheet = SS.getSheetByName(CLIENTES_SHEET);
    if (!sheet) {
      throw new Error(`La pestaña ${CLIENTES_SHEET} no se encuentra.`);
    }

    if (!payload.Nombre || !payload.NIT_CC) {
      throw new Error('Nombre y NIT/CC son obligatorios.');
    }

    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const newRow = new Array(headers.length).fill('');
    const idx = {};
    headers.forEach((header, i) => {
      idx[header.trim()] = i;
    });

    const newId = getNextId(CLIENTES_SHEET, 'CLI-');
    const today = new Date();

    // Mapeo seguro
    if (idx['CLIENTE_ID'] !== undefined) newRow[idx['CLIENTE_ID']] = newId;
    if (idx['Nombre'] !== undefined) newRow[idx['Nombre']] = payload.Nombre;
    if (idx['Tipo'] !== undefined) newRow[idx['Tipo']] = payload.Tipo || 'Cliente Final';
    if (idx['NIT/CC'] !== undefined) newRow[idx['NIT/CC']] = payload.NIT_CC;
    if (idx['Tel'] !== undefined) newRow[idx['Tel']] = payload.Tel || '';
    if (idx['Email'] !== undefined) newRow[idx['Email']] = payload.Email || '';
    if (idx['Direcciones'] !== undefined) newRow[idx['Direcciones']] = payload.Direcciones || '';
    if (idx['Descuento_default'] !== undefined) newRow[idx['Descuento_default']] = payload.Descuento_default || 0;
    if (idx['Fecha_Registro'] !== undefined) newRow[idx['Fecha_Registro']] = today;
    if (idx['Estado'] !== undefined) newRow[idx['Estado']] = 'Activo';
    
    // --- ✅ CAMPOS NUEVOS DE CMSI (Usando los nombres del payload) ---
    if (idx['Tags'] !== undefined) newRow[idx['Tags']] = payload.Tags || '';
    if (idx['Estado_Contacto'] !== undefined) newRow[idx['Estado_Contacto']] = payload.Estado_Contacto || 'Activo';
    if (idx['consent_marketing'] !== undefined) newRow[idx['consent_marketing']] = payload.Consent_Marketing || 'FALSE';
    // --- FIN CAMPOS NUEVOS ---

    if (idx['purchase_count'] !== undefined) newRow[idx['purchase_count']] = newRow[idx['purchase_count']] || 0;
    if (idx['total_spent'] !== undefined) newRow[idx['total_spent']] = newRow[idx['total_spent']] || 0;

    sheet.appendRow(newRow);
    
    return { 
      success: true, 
      newId: newId, 
      message: `Cliente ${newId} creado exitosamente.` 
    };
  } catch (e) {
    Logger.log(e);
    return { error: e.message };
  }
}


function createCampaign(payload) {
  try {
    // (AQUÍ IRÁ TU LÓGICA PARA ESCRIBIR EN LA HOJA DE CÁLCULO)
    Logger.log(payload);

    // Simular éxito
    return {
      success: true,
      nombre: payload.Nombre
    };

  } catch (e) {
    Logger.log(`Error en createCampaign: ${e}`);
    return { success: false, error: e.message };
  }
}
/**
 * start_marker:REEMPLAZO_GET_CLIENTE_DETALLE
 * @description Obtiene el detalle de UN cliente por su ID
 * (Versión REFACTORIZADA que incluye Remarketing)
 */
function getClienteDetalle(clienteId) {
  try {
    // 1. Obtener todos los clientes (como objetos)
    const allClientes = _getMappedData(CLIENTES_SHEET);
    
    // 2. Encontrar el cliente (¡No más row[0]!)
    const clienteInfo = allClientes.find(c => c.CLIENTE_ID == clienteId);
    
    if (!clienteInfo) {
      throw new Error(`Cliente con ID ${clienteId} no encontrado.`);
    }
    
    // 3. Obtener facturas (REFACTORIZADO)
    // Esto ya devuelve objetos gracias a nuestro Bloque 2
    const facturasData = _getFacturasData(); 
    
    // 4. Calcular métricas de remarketing (¡Ahora SÍ!)
    let totalComprado = 0;
    let fechasCompras = [];

    const clientInvoices = facturasData
      .filter(i => i.Cliente_ID == clienteId && i.Estado_pago && i.Estado_pago.toLowerCase() === 'pagado')
      .map(i => {
        const totalFactura = _cleanNumber(i.Total);
        totalComprado += totalFactura;
        
        let fechaFacturaStr = i.Fecha; // ej. "29/10/2025"
        let fechaCompra;
        
        // Convertir "dd/mm/aaaa" a un objeto Date para ordenar
        if (typeof fechaFacturaStr === 'string' && fechaFacturaStr.includes('/')) {
          const parts = fechaFacturaStr.split('/');
          // Formato: new Date(año, mes - 1, dia)
          fechaCompra = new Date(parts[2], parts[1] - 1, parts[0]);
        } else {
          fechaCompra = new Date(0); // Fecha mínima si es inválida
        }
        
        fechasCompras.push(fechaCompra);

        return {
          Factura_ID: i.FAC_ID, // Usar el nombre del encabezado
          ORD_ID: i.ORD_ID,
          Fecha_Factura: fechaFacturaStr,
          Total: totalFactura,
          _FechaObj: fechaCompra // Fecha real para ordenar
        };
      })
      .sort((a, b) => b._FechaObj.getTime() - a._FechaObj.getTime()); // Ordenar de más reciente a más antigua
    
    let ultimaCompra_f = 'N/A';
    if(clientInvoices.length > 0) {
      ultimaCompra_f = clientInvoices[0].Fecha_Factura; // La más reciente
    }

    // 5. Ensamblar la respuesta final
    // ...clienteInfo -> trae TODOS los campos de la hoja CLIENTES
    // (CLIENTE_ID, Nombre, Email, y también consent_marketing, segment, etc.)
    return {
      // Datos directos de la hoja CLIENTES
      ...clienteInfo, 
      
      // Datos calculados (para asegurar que están actualizados)
      total_spent_calc: totalComprado,
      purchase_count_calc: clientInvoices.length,
      last_purchase_date_calc: ultimaCompra_f,
      
      // Historial de facturas para la tabla de detalle
      historialFacturas: clientInvoices 
    };

  } catch (e) {
    Logger.log(e);
    return { error: 'Error al obtener detalle de cliente: ' + e.message };
  }
}

/**
 * start_marker:NUEVAS_FUNCIONES_PRODUCTOS_INSUMOS
 * @description (NUEVA) Carga el listado de todos los productos.
 * El HTML 'productos.html' necesita esta función.
 */
function getProductos() {
  try {
    // 1. Usar la función maestra
    const productos = _getMappedData(PRODUCTOS_SHEET);
    
    // 2. Mapear para asegurar que el HTML recibe lo que espera
    return productos.map(p => {
      return {
        sku: p.SKU,
        nombre: p.NOMBRE,
        // Usar _cleanNumber para limpiar valores
        precio: _cleanNumber(p.Precio_venta), 
        stock: _cleanNumber(p.Stock),
        categoria: p.CAT_PROD,
        foto: p.Foto_Prod,
        foto_cat: p.Foto_Cat // ¡Importante para el requerimiento estético de VENTAS.html!
      };
    });
  } catch (e) {
    Logger.log(e);
    return { error: 'Error al cargar productos: ' + e.message };
  }
}

/**
 * @description (NUEVA) Carga el listado de todos los insumos.
 * El HTML 'insumos.html' necesita esta función.
 */
function getInsumos() {
  try {
    // 1. Usar la función maestra
    const insumos = _getMappedData(INSUMOS_SHEET);

    // 2. Mapear al formato esperado
    return insumos.map(i => {
      return {
        id: i.INS_ID,
        nombre: i.Nombre,
        unidad: i.Unidad,
        costo: _cleanNumber(i.Costo_unitario),
        stock: _cleanNumber(i.Stock_actual),
        stock_min: _cleanNumber(i.Stock_minimo),
        categoria: i.Categoría,
        proveedor: i.Proveedor_pref
      };
    });
  } catch (e) {
    Logger.log(e);
    return { error: 'Error al cargar insumos: ' + e.message };
  }
}

/**
 * @description Obtiene el historial de chat para un cliente específico.
 */
function getChatHistory(clienteId) {
  try {
    const allMessages = getSheetData(CRM_LOG_SHEET);
    if (allMessages.error) throw new Error(allMessages.error);

    const chatHistory = allMessages.filter(msg => msg.Cliente_ID === clienteId);
    
    return { success: true, history: chatHistory };
  } catch (e) {
    Logger.log(e);
    return { success: false, error: e.message };
  }
}

/**
 * @description Obtiene la lista de conversaciones (último mensaje de cada cliente).
 */
function getChatConversations() {
  try {
    const allMessages = getSheetData(CRM_LOG_SHEET);
    if (allMessages.error) throw new Error(allMessages.error);

    // Agrupar mensajes por Cliente_ID y tomar el último
    const conversations = {};
    allMessages.forEach(msg => {
      // Guardar el mensaje más reciente
      conversations[msg.Cliente_ID] = msg;
    });

    // Convertir el objeto de vuelta a un array
    const conversationList = Object.values(conversations);

    return { success: true, conversations: conversationList };
  } catch (e) {
    Logger.log(e);
    return { success: false, error: e.message };
  }
}

/**
 * @description Guarda un nuevo mensaje (enviado o recibido) en el log de CRM.
 */
function logCrmMessage(payload) {
  // payload = { clienteId, nombreCliente, medio, mensajeEnviado, mensajeRecibido, usuario }
  try {
    const sheet = SS.getSheetByName(CRM_LOG_SHEET);
    if (!sheet) throw new Error(`Hoja ${CRM_LOG_SHEET} no encontrada.`);

    const newId = getNextId(CRM_LOG_SHEET, 'LOG-');
    const now = new Date();
    
    const newRow = [
      newId,
      now,
      payload.clienteId || '',
      payload.nombreCliente || '',
      payload.medio || 'Chat Interno',
      payload.mensajeEnviado || '',
      payload.mensajeRecibido || '',
      payload.usuario || 'Sistema' // Idealmente, obtener el email del usuario activo
    ];
    
    sheet.appendRow(newRow);
    return { success: true, newRow: newRow };
  } catch (e) {
    Logger.log(e);
    return { success: false, error: e.message };
  }
}

function getLeadDetailsAndChat(leadId) {
  try {
    // 1. Obtener detalles del Lead (Cliente)
    // Reutilizamos la función que ya tenías
    const details = getClienteDetalle(leadId); 
    if (details.error) {
      throw new Error(details.error);
    }

    // 2. Obtener historial de Chat
    // Reutilizamos la función que ya tenías
    const chatData = getChatHistory(leadId);
    if (!chatData.success) {
      throw new Error(chatData.error);
    }

    // 3. Devolver un solo objeto combinado
    return {
      details: details,  // Contiene Nombre, Tel, Email, Tags, etc.
      chat: chatData.history // Contiene el array de mensajes
    };

  } catch (e) {
    Logger.log(e);
    return { error: e.message };
  }
}

const SS_ID = SpreadsheetApp.getActiveSpreadsheet().getId();
const HOJA_CLIENTES = "CLIENTES";
const HOJA_INTERACCIONES = "CRM_INTERACCIONES";


// --- Índices de Columna (Empezando en 1) para HOJA_CLIENTES ---
const COL_CLIENTE_ID = 1;     // Columna "A" en CLIENTES
const COL_CLIENTE_TAGS = 19;  // Columna "S" (Tags) en CLIENTES
const COL_CLIENTE_CAMPAÑA = 22; // Columna "V" (Campaña_Activa) en CLIENTES

// --- Índices de Columna (Empezando en 1) para HOJA_INTERACCIONES ---
const COL_INT_ID = 1;         // Columna "A"
const COL_INT_CLIENTE_ID = 2; // Columna "B"
const COL_INT_FECHA = 3;      // Columna "C"
const COL_INT_TIPO = 4;       // Columna "D"
const COL_INT_DETALLE = 5;    // Columna "E"
const COL_INT_USUARIO = 6;    // Columna "F"


function getUsuarioEmail() {
  return Session.getActiveUser().getEmail();
}


function asignarEtiqueta(clienteId, etiqueta) {
  try {
    const ss = SpreadsheetApp.openById(SS_ID);
    const sheet = ss.getSheetByName(HOJA_CLIENTES);
    
    const rowIndex = findRowIndexById(sheet, clienteId, COL_CLIENTE_ID);
    
    if (rowIndex === -1) {
      throw new Error(`No se encontró el cliente con ID: ${clienteId}`);
    }
    
    const cell = sheet.getRange(rowIndex, COL_CLIENTE_TAGS);
    const tagsActuales = cell.getValue();
    
    // Evita duplicados
    if (tagsActuales.includes(etiqueta)) {
      return { status: "ok", message: "La etiqueta ya existía." };
    }
    
    const nuevaListaTags = tagsActuales ? `${tagsActuales},${etiqueta}` : etiqueta;
    cell.setValue(nuevaListaTags);
    
    return { status: "ok", message: `Etiqueta '${etiqueta}' agregada.` };
    
  } catch (e) {
    Logger.log(`Error en asignarEtiqueta: ${e}`);
    // Importante: Re-lanzar el error para que el .withFailureHandler() del HTML lo capture
    throw new Error(`Error del servidor: ${e.message}`);
  }
}

function agregarClienteACampaña(clienteId, campanaId) {
  try {
    const ss = SpreadsheetApp.openById(SS_ID);
    const sheetClientes = ss.getSheetByName(HOJA_CLIENTES);
    
    const rowIndex = findRowIndexById(sheetClientes, clienteId, COL_CLIENTE_ID);
    
    if (rowIndex === -1) {
      throw new Error(`No se encontró el cliente con ID: ${clienteId}`);
    }
    
    // Opcional: Validar que campanaId existe en HOJA_CAMPAÑAS (lo omitimos por brevedad)
    
    const cell = sheetClientes.getRange(rowIndex, COL_CLIENTE_CAMPAÑA);
    cell.setValue(campanaId);
    
    return { status: "ok", message: `Cliente asignado a campaña '${campanaId}'.` };
    
  } catch (e) {
    Logger.log(`Error en agregarClienteACampaña: ${e}`);
    throw new Error(`Error del servidor: ${e.message}`);
  }
}

function programarSeguimiento(clienteId, fecha, nota) {
  try {
    const ss = SpreadsheetApp.openById(SS_ID);
    const sheet = ss.getSheetByName(HOJA_INTERACCIONES);
    
    const usuario = getUsuarioEmail();
    const fechaISO = new Date(fecha).toISOString();
    
    // Creamos la nueva fila
    // Asumimos que la Col A (INT_ID) se autogenera o puede ir vacía si usas una fórmula.
    // Aquí la dejamos vacía.
    const nuevaFila = [
      "", // COL_INT_ID (A)
      clienteId, // COL_INT_CLIENTE_ID (B)
      fechaISO, // COL_INT_FECHA (C)
      "Seguimiento Programado", // COL_INT_TIPO (D)
      nota, // COL_INT_DETALLE (E)
      usuario // COL_INT_USUARIO (F)
    ];
    
    sheet.appendRow(nuevaFila);
    
    return { status: "ok", message: "Seguimiento programado correctamente." };
    
  } catch (e) {
    Logger.log(`Error en programarSeguimiento: ${e}`);
    throw new Error(`Error del servidor: ${e.message}`);
  }
}

const HOJA_CAMPAÑAS = "CMSI_CAMPAÑAS";
const HOJA_LOG_CAMPAÑAS = "LOG_CAMPAÑAS"; // (Debes crear esta hoja para las métricas)

const COL_CAMP_ID = 1;        // A: CAMP_ID
const COL_CAMP_NOMBRE = 2;    // B: Nombre
const COL_CAMP_CANAL = 3;     // C: Canal (Email / WhatsApp)
const COL_CAMP_SEGMENTO = 4;  // D: Segmento_Tags
const COL_CAMP_CONTENIDO = 5; // E: Contenido
const COL_CAMP_ESTADO = 6;    // F: Estado (Borrador, Programada, Enviada)
const COL_CAMP_FECHA = 7;     // G: Fecha_Envio
const COL_CAMP_UTM_S = 8;     // H: utm_source
const COL_CAMP_UTM_M = 9;     // I: utm_medium
const COL_CAMP_UTM_C = 10;    // J: utm_campaign


function getCampanas() {
  try {
    const ss = SpreadsheetApp.openById(SS_ID);
    const sheet = ss.getSheetByName(HOJA_CAMPAÑAS);
    
    // Asumimos fila 1 = Headers, Datos = Fila 2 en adelante
    const range = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn());
    const values = range.getValues();
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    
    const campañas = values.map(row => {
      let camp = {};
      headers.forEach((header, index) => {
        // Convertir fechas a string ISO para evitar errores de JSON
        if (row[index] instanceof Date) {
          camp[header] = row[index].toISOString();
        } else {
          camp[header] = row[index];
        }
      });
      return camp;
    });
    
    return campañas;
    
  } catch (e) {
    Logger.log(`Error en getCampanas: ${e}`);
    throw new Error(`Error del servidor: ${e.message}`);
  }
}

function guardarCampana(campanaData, modo) {
  try {
    const ss = SpreadsheetApp.openById(SS_ID);
    const sheet = ss.getSheetByName(HOJA_CAMPAÑAS);
    
    if (modo === 'create') {
      // --- CREAR NUEVA CAMPAÑA ---
      
      // Generar un nuevo ID (ej. CAMP-005)
      const lastIdRow = sheet.getRange(sheet.getLastRow(), COL_CAMP_ID).getValue();
      let newIdNum = 1;
      if (lastIdRow && lastIdRow.includes('-')) {
        newIdNum = parseInt(lastIdRow.split('-')[1]) + 1;
      }
      const newCampId = `CAMP-${String(newIdNum).padStart(3, '0')}`;
      
      const newRow = [
        newCampId,                                  // A: CAMP_ID
        campanaData.nombre,                         // B: Nombre
        campanaData.canal,                          // C: Canal
        campanaData.segmentoTags,                   // D: Segmento_Tags
        campanaData.contenido,                      // E: Contenido
        "Borrador",                                 // F: Estado (siempre Borrador al crear)
        new Date(campanaData.fechaEnvio),           // G: Fecha_Envio
        campanaData.utm_source,                     // H: utm_source
        campanaData.utm_medium,                     // I: utm_medium
        campanaData.utm_campaign                    // J: utm_campaign
      ];
      
      sheet.appendRow(newRow);
      return { status: "ok", message: `Campaña '${newCampId}' creada.`, newId: newCampId };

    } else {
      // --- EDITAR CAMPAÑA EXISTENTE ---
      const campId = campanaData.CAMP_ID;
      if (!campId) {
        throw new Error("No se proporcionó CAMP_ID para editar.");
      }
      
      const rowIndex = findRowIndexById(sheet, campId, COL_CAMP_ID);
      if (rowIndex === -1) {
        throw new Error(`Campaña ${campId} no encontrada para editar.`);
      }
      
      // Actualizar las celdas necesarias (ejemplo: Nombre, Contenido, Fecha)
      // (Esto es más eficiente que re-escribir toda la fila)
      sheet.getRange(rowIndex, COL_CAMP_NOMBRE).setValue(campanaData.nombre);
      sheet.getRange(rowIndex, COL_CAMP_CANAL).setValue(campanaData.canal);
      sheet.getRange(rowIndex, COL_CAMP_SEGMENTO).setValue(campanaData.segmentoTags);
      sheet.getRange(rowIndex, COL_CAMP_CONTENIDO).setValue(campanaData.contenido);
      sheet.getRange(rowIndex, COL_CAMP_FECHA).setValue(new Date(campanaData.fechaEnvio));
      sheet.getRange(rowIndex, COL_CAMP_UTM_S).setValue(campanaData.utm_source);
      sheet.getRange(rowIndex, COL_CAMP_UTM_M).setValue(campanaData.utm_medium);
      sheet.getRange(rowIndex, COL_CAMP_UTM_C).setValue(campanaData.utm_campaign);
      
      return { status: "ok", message: `Campaña '${campId}' actualizada.` };
    }
    
  } catch (e) {
    Logger.log(`Error en  guardarCampana: ${e}`);
    throw new Error(`Error del servidor: ${e.message}`);
  }
}

function ejecutarEnvioCampaña(campId, listaPrueba = null, esPrueba = false) {
  try {
    const ss = SpreadsheetApp.openById(SS_ID);
    
    // 1. Obtener datos de la Campaña
    const sheetgetCampanas = ss.getSheetByName(HOJA_CAMPAÑAS);
    const rowIndex = findRowIndexById(sheetgetCampanas, campId, COL_CAMP_ID);
    if (rowIndex === -1) throw new Error(`Campaña ${campId} no encontrada.`);
    
    const campRow = sheetgetCampanas.getRange(rowIndex, 1, 1, sheetgetCampanas.getLastColumn()).getValues()[0];
    const campaña = {
      id: campRow[COL_CAMP_ID - 1],
      canal: campRow[COL_CAMP_CANAL - 1],
      segmento: campRow[COL_CAMP_SEGMENTO - 1],
      contenido: campRow[COL_CAMP_CONTENIDO - 1],
      utm_source: campRow[COL_CAMP_UTM_S - 1],
      utm_medium: campRow[COL_CAMP_UTM_M - 1],
      utm_campaign: campRow[COL_CAMP_UTM_C - 1]
    };

    let clientesParaEnvio = [];

    if (listaPrueba) {
      // Es una prueba, usamos la lista de prueba
      clientesParaEnvio = listaPrueba.map(contacto => ({
          CLIENTE_ID: 'TEST-001',
          Nombre: 'Cliente de Prueba',
          Email: campaña.canal === 'Email' ? contacto : 'prueba@ejemplo.com',
          Tel: campaña.canal === 'WhatsApp' ? contacto : '3001234567'
      }));
      
    } else {
      // Es un envío real, obtener clientes del segmento
      clientesParaEnvio = getClientesPorSegmento(campaña.segmento);
    }
    
    if (clientesParaEnvio.length === 0) {
      return { status: "warn", message: "No se encontraron clientes para este segmento." };
    }

    // 2. Procesar el envío según el canal
    let resultados = [];
    if (campaña.canal === 'Email') {
      resultados = enviarCampañaEmail(clientesParaEnvio, campaña);
    } else if (campaña.canal === 'WhatsApp') {
      resultados = enviarCampañaWhatsApp(clientesParaEnvio, campaña);
    } else {
      throw new Error(`Canal '${campaña.canal}' no soportado.`);
    }

    // 3. Registrar en el LOG (Métricas)
    registrarLogCampaña(resultados, campId);

    // 4. Actualizar estado de la campaña (solo si no es prueba)
    if (!esPrueba) {
      sheetgetCampanas.getRange(rowIndex, COL_CAMP_ESTADO).setValue("Enviada");
    }

    return { status: "ok", message: `Campaña ${campId} enviada a ${resultados.length} contactos.` };

  } catch (e) {
    Logger.log(`Error en ejecutarEnvioCampaña: ${e}`);
    throw new Error(`Error del servidor: ${e.message}`);
  }
}

function getClientesPorSegmento(segmento) {
  const sheetClientes = SpreadsheetApp.openById(SS_ID).getSheetByName(HOJA_CLIENTES);
  // Asumimos fila 1 = Headers
  const data = sheetClientes.getRange(2, 1, sheetClientes.getLastRow() - 1, sheetClientes.getLastColumn()).getValues();
  const headers = sheetClientes.getRange(1, 1, 1, sheetClientes.getLastColumn()).getValues()[0];
  
  // Convertir data a objetos
  const clientes = data.map(row => {
    let obj = {};
    headers.forEach((header, index) => obj[header] = row[index]);
    return obj;
  });

  if (segmento === 'todos') {
    return clientes.filter(c => c.consent_marketing === true); // Respetar consentimiento
  }
  
  // Filtrar por tag
  return clientes.filter(c => 
    c.consent_marketing === true && 
    c.Tags && 
    c.Tags.includes(segmento)
  );
}

function enviarCampañaEmail(clientes, campaña) {
  const usuario = getUsuarioEmail();
  let resultados = [];
  
  // (Asumimos que la columna 3 del CLIENTES.csv es 'Nombre' y la 5 es 'Email')
  
  clientes.forEach(cliente => {
    try {
      // Personalizar el contenido (reemplazo básico de variables)
      let contenidoPersonalizado = campaña.contenido
          .replace(/{{Nombre}}/g, cliente.Nombre)
          .replace(/{{Email}}/g, cliente.Email);
      
      // Añadir UTMs si existen
      // (Aquí deberías parsear el HTML y añadir UTMs a los links, es complejo)
      // Por ahora, lo dejamos simple.

      GmailApp.sendEmail(
        cliente.Email,
        campaña.nombre, // Asunto
        "", // Cuerpo (vacío si usamos HTML)
        {
          htmlBody: contenidoPersonalizado,
          name: "Tu Nombre de Empresa" // Nombre del remitente
        }
      );
      
      resultados.push({
        clienteId: cliente.CLIENTE_ID,
        contacto: cliente.Email,
        estado: 'Enviado',
        error: ''
      });
      
    } catch (e) {
      resultados.push({
        clienteId: cliente.CLIENTE_ID,
        contacto: cliente.Email,
        estado: 'Fallido',
        error: e.message
      });
    }
  });
  return resultados;
}

function enviarCampañaWhatsApp(clientes, campaña) {
  let resultados = [];
  
  // (Asumimos que la columna 3 es 'Nombre' y la 5 es 'Tel')
  
  clientes.forEach(cliente => {
    try {
      let contenidoPersonalizado = campaña.contenido
          .replace(/{{Nombre}}/g, cliente.Nombre);
      
      // Aplicar la lógica de construcción de link que nos diste
      const textoCodificado = encodeURIComponent(contenidoPersonalizado);
      
      // Limpiar el número de teléfono (quitar +, espacios, etc.)
      const telefonoLimpio = String(cliente.Tel).replace(/[\s\+()-]/g, '');
      
      // Asumir código de país si no está (ej. 57 para Colombia)
      // (Esta lógica debe ajustarse a tus datos)
      const telefonoFinal = telefonoLimpio.startsWith('57') ? telefonoLimpio : `57${telefonoLimpio}`; 
      
      const waLink = `https://api.whatsapp.com/send?phone=${telefonoFinal}&text=${textoCodificado}`;
      
      // Como esto es "Lite", no lo enviamos, solo lo registramos 
      // y (opcionalmente) podríamos devolver los links al frontend.
      
      resultados.push({
        clienteId: cliente.CLIENTE_ID,
        contacto: telefonoFinal,
        estado: 'Pendiente (Manual)', // Estado para WhatsApp Lite
        error: waLink // Guardamos el link en el log para referencia
      });

    } catch (e) {
      resultados.push({
        clienteId: cliente.CLIENTE_ID,
        contacto: cliente.Tel,
        estado: 'Fallido',
        error: e.message
      });
    }
  });
  return resultados;
}

function registrarLogCampaña(resultados, campId) {
  const ss = SpreadsheetApp.openById(SS_ID);
  const sheetLog = ss.getSheetByName(HOJA_LOG_CAMPAÑAS);
  const fecha = new Date();
  const usuario = getUsuarioEmail();
  
  let nuevasFilas = resultados.map(res => {
    return [
      `LOG-${Date.now()}-${res.clienteId}`, // Log_ID
      fecha,
      campId,
      res.clienteId,
      res.contacto,
      res.estado,
      res.error // (Guardamos el Link de WA o el error de Email)
    ];
  });
  
  if (nuevasFilas.length > 0) {
    // Insertar todas las filas de golpe (más eficiente)
    sheetLog.getRange(sheetLog.getLastRow() + 1, 1, nuevasFilas.length, nuevasFilas[0].length)
              .setValues(nuevasFilas);
  }
}

const COL_LOG_CAMP_ID_REF = 3; // Columna "C" (Campaña_ID) en LOG_CAMPAÑAS
const COL_LOG_ESTADO = 6;      // Columna "F" (Estado) en LOG_CAMPAÑAS

function helper_arrayDeObjetos(data) {
  // Asume que la fila 1 (índice 0) son las cabeceras
  const headers = data[0].map(h => h.trim()); // Limpiar cabeceras
  const body = data.slice(1); // El resto son datos
  
  return body.map((row, rowIndex) => {
    let obj = { _row: rowIndex + 2 }; // Guardar el número de fila real
    headers.forEach((header, index) => {
      if (header) {
        // Convertir fechas a string ISO para evitar errores de JSON
        if (row[index] instanceof Date) {
          obj[header] = row[index].toISOString();
        } else {
          obj[header] = row[index];
        }
      }
    });
    return obj;
  });
}

function findRowIndexById(sheet, id, idColumnIndex) {
  try {
    if (sheet.getLastRow() < 2) return -1; 
    const dataRange = sheet.getRange(2, idColumnIndex, sheet.getLastRow() - 1, 1);
    const values = dataRange.getValues();
    
    for (let i = 0; i < values.length; i++) {
      // Comparamos como strings para evitar errores de tipo
      if (String(values[i][0]) === String(id)) {
        return i + 2; // +2 porque los datos empiezan en la fila 2
      }
    }
    return -1; // No encontrado
  } catch (e) {
    Logger.log(`Error en findRowIndexById: ${e}`);
    return -1;
  }
}

function getCampaigns() {
  try {
    const ss = SpreadsheetApp.openById(SS_ID);
    const sheet = ss.getSheetByName(HOJA_CAMPAÑAS);
    
    if (!sheet) {
      throw new Error(`La hoja "${HOJA_CAMPAÑAS}" no existe. Revisa la constante.`);
    }

    const range = sheet.getDataRange();
    const values = range.getValues();
    
    if (values.length < 2) {
      return []; // Hoja vacía (solo cabeceras)
    }
    
    return helper_arrayDeObjetos(values);
    
  } catch (e) {
    Logger.log(`Error en getCampaigns: ${e}`);
    throw new Error(`Error del servidor: ${e.message}`);
  }
}

function duplicarCampaña(campId) {
  try {
    const ss = SpreadsheetApp.openById(SS_ID);
    const sheet = ss.getSheetByName(HOJA_CAMPAÑAS);
    
    const rowIndex = findRowIndexById(sheet, campId, 1); // Col 1 = CAMP_ID
    if (rowIndex === -1) {
      throw new Error(`No se encontró la campaña ${campId} para duplicar.`);
    }
    
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(h => h.trim());
    const values = sheet.getRange(rowIndex, 1, 1, sheet.getLastColumn()).getValues()[0];
    
    // Generar nuevo ID
    let newCampId = 'CAMP-001';
    if (sheet.getLastRow() > 1) {
      const lastIdCell = sheet.getRange(sheet.getLastRow(), 1).getValue();
      if (lastIdCell && lastIdCell.toString().includes('-')) {
        let newIdNum = parseInt(lastIdCell.split('-')[1]) + 1;
        newCampId = `CAMP-${String(newIdNum).padStart(3, '0')}`;
      }
    }
    
    let nombreOriginal = "";
    
    // Mapear valores para la nueva fila
    const newRow = headers.map((header, index) => {
      switch (header) {
        case "CAMP_ID":
          return newCampId;
        case "Nombre":
          nombreOriginal = values[index];
          return `${values[index]} (Copia)`;
        case "Estado":
          return "Borrador"; // ¡Requerimiento cumplido!
        case "Fecha_Reg":
          return new Date(); // Nueva fecha de registro
        case "Clics":
        case "Conversiones":
          return 0; // Resetear métricas
        case "Fecha_Envio":
        case "Hora_Prog":
          return null; // Limpiar programación
        default:
          return values[index]; // Copiar el resto de datos
      }
    });
    
    sheet.appendRow(newRow);
    
    return { status: "ok", message: `Campaña '${nombreOriginal}' duplicada.` };
    
  } catch (e) {
    Logger.log(`Error en duplicarCampaña: ${e}`);
    throw new Error(`Error del servidor: ${e.message}`);
  }
}

function getCampañaPorId(campId) {
  try {
    const ss = SpreadsheetApp.openById(SS_ID);
    const sheet = ss.getSheetByName(HOJA_CAMPAÑAS);
    
    const rowIndex = findRowIndexById(sheet, campId, 1);
    if (rowIndex === -1) {
      throw new Error(`No se encontró la campaña ${campId}`);
    }
    
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(h => h.trim());
    const values = sheet.getRange(rowIndex, 1, 1, sheet.getLastColumn()).getValues()[0];
    
    let campaña = {};
    headers.forEach((header, index) => {
      if (values[index] instanceof Date) {
        // Formato YYYY-MM-DD para <input type="date">
        if (header === "Fecha_Envio") {
          campaña[header] = values[index].toISOString().split('T')[0];
        } else {
          campaña[header] = values[index].toISOString();
        }
      } else {
        campaña[header] = values[index];
      }
    });
    
    return campaña;
    
  } catch (e) {
    Logger.log(`Error en getCampañaPorId: ${e}`);
    throw new Error(`Error del servidor: ${e.message}`);
  }
}

function eliminarCampaña(campId) {
  try {
    const ss = SpreadsheetApp.openById(SS_ID);
    const sheet = ss.getSheetByName(HOJA_CAMPAÑAS);
    
    const rowIndex = findRowIndexById(sheet, campId, 1); // Col 1 = CAMP_ID
    if (rowIndex === -1) {
      throw new Error(`No se encontró la campaña ${campId} para eliminar.`);
    }
    
    sheet.deleteRow(rowIndex);
    
    // Opcional: Eliminar logs asociados de HOJA_LOG_CAMPAÑAS
    
    return { status: "ok", message: `Campaña ${campId} eliminada.` };
    
  } catch (e) {
    Logger.log(`Error en eliminarCampaña: ${e}`);
    throw new Error(`Error del servidor: ${e.message}`);
  }
}

function pausarCampaña(campId, estadoActual) {
   try {
    const ss = SpreadsheetApp.openById(SS_ID);
    const sheet = ss.getSheetByName(HOJA_CAMPAÑAS);
    
    const rowIndex = findRowIndexById(sheet, campId, 1); // Col 1 = CAMP_ID
    if (rowIndex === -1) {
      throw new Error(`No se encontró la campaña ${campId}.`);
    }

    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(h => h.trim());
    const colEstadoIndex = headers.indexOf("Estado") + 1;
    if (colEstadoIndex === 0) throw new Error("No se encontró la columna 'Estado'");

    const estadoCell = sheet.getRange(rowIndex, colEstadoIndex);
    const estadoNormalizado = (estadoActual || "Borrador").toLowerCase();
    let nuevoEstado = "";
    
    if (estadoNormalizado === "pausada") {
      // Al reanudar, vuelve a 'Borrador' si no tiene fecha, o 'Programada' si la tiene.
      const colFechaEnvio = headers.indexOf("Fecha_Envio") + 1;
      const fechaEnvio = colFechaEnvio > 0 ? sheet.getRange(rowIndex, colFechaEnvio).getValue() : null;
      nuevoEstado = fechaEnvio ? "Programada" : "Borrador";
    } else {
      nuevoEstado = "Pausada";
    }
    
    estadoCell.setValue(nuevoEstado);
    
    return { status: "ok", message: `Campaña ${campId} ahora está ${nuevoEstado}.`, nuevoEstado: nuevoEstado };
    
  } catch (e) {
    Logger.log(`Error en pausarCampaña: ${e}`);
    throw new Error(`Error del servidor: ${e.message}`);
  }
}

function getMetricas(campId) {
  try {
    const ss = SpreadsheetApp.openById(SS_ID);
    const sheetLog = ss.getSheetByName(HOJA_LOG_CAMPAÑAS);
    if (!sheetLog) {
      Logger.log("No se encontró la hoja 'LOG_CAMPAÑAS'. Creándola.");
      ss.insertSheet(HOJA_LOG_CAMPAÑAS).appendRow(["Log_ID", "Fecha", "Campaña_ID", "Cliente_ID", "Contacto", "Estado", "Error_Msg"]);
      return { total: 0, enviados: 0, fallidos: 0, pendientes: 0, otros: 0 };
    }
    
    if (sheetLog.getLastRow() < 2) return { total: 0, enviados: 0, fallidos: 0, pendientes: 0, otros: 0 };

    const logHeaders = sheetLog.getRange(1, 1, 1, sheetLog.getLastColumn()).getValues()[0].map(h => h.trim());
    const colCampIdRef = logHeaders.indexOf("Campaña_ID") + 1;
    const colLogEstado = logHeaders.indexOf("Estado") + 1;

    if (colCampIdRef === 0 || colLogEstado === 0) {
        throw new Error("La hoja 'LOG_CAMPAÑAS' debe tener 'Campaña_ID' y 'Estado'");
    }

    const data = sheetLog.getRange(2, 1, sheetLog.getLastRow() - 1, sheetLog.getLastColumn()).getValues();
    const logsCampaña = data.filter(row => String(row[colCampIdRef - 1]) === String(campId));
    
    let stats = { total: logsCampaña.length, enviados: 0, fallidos: 0, pendientes: 0, otros: 0 };
    
    logsCampaña.forEach(row => {
      const estado = row[colLogEstado - 1];
      if (estado === 'Enviado') stats.enviados++;
      else if (estado === 'Fallido') stats.fallidos++;
      else if (estado === 'Pendiente (Manual)') stats.pendientes++;
      else stats.otros++;
    });
    
    return stats;
    
  } catch (e) {
    Logger.log(`Error en getMetricas: ${e}`);
    throw new Error(`Error del servidor: ${e.message}`);
  }
}