/**
 * @file main.gs
 * Punto de entrada principal del proyecto Pasar Modular.
 * Esta versión está preparada para ejecutarse como una aplicación web
 * (idéntica al comportamiento original del proyecto ggeip.gpon.modular).
 * 
 * Cuando se despliegue como WebApp, Google Apps Script llamará
 * automáticamente a la función doGet(), que devuelve la interfaz HTML.
 */

// ============================================================
// 🗂️ CONFIGURACIÓN BASE
// ============================================================
const SPREADSHEET_ID = "1Dzy49UHef2vi2GIxtzNkW6xneWTmqxhpqQG2-OkH37o";
const SHEET_NAME = "Applications";
const LISTASFIJAS_SHEET = "ListasFijas";

// ============================================================
// 🌐 FUNCIÓN PRINCIPAL DE ACCESO WEB
// ============================================================
function doGet(e) {
  if (e && e.parameter.mod) {
    return HtmlService.createHtmlOutputFromFile(`ui/${e.parameter.mod}`);
  }

  return HtmlService.createHtmlOutputFromFile('ui/index')
    .setTitle("📋 Pasar Modular — Formulario Principal")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ============================================================
// ⚙️ FUNCIONES DE UTILIDAD
// ============================================================
function include(filename) {
  return HtmlService.createHtmlOutputFromFile("ui/" + filename).getContent();
}

function getModuloHTML(nombre) {
  return HtmlService.createHtmlOutputFromFile('ui/' + nombre).getContent();
}

// ============================================================
// 📋 FUNCIONES DE DATOS (Sheets Integration)
// ============================================================
/**
 * 📋 Devuelve las listas fijas desde la hoja "ListasFijas".
 * Cada columna es una lista y la fila 1 contiene su nombre.
 */
function getListasFijas() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(LISTASFIJAS_SHEET);
  if (!sheet) throw new Error("❌ No se encontró la hoja 'ListasFijas'");

  const data = sheet.getDataRange().getValues();
  const headers = data[0];

  const map = {
    // las que ya funcionaban (las dejamos)
    numero: [],
    situacion_actual: [],
    odn: [],
    // 🔧 las que reportaste que no cargan (mapeo EXACTO)
    cod_area: [],
    modelo_equipo: [],
    estatus_local: [],
    situacion_abordaje: [],
    oferta_entregada: [],
    oferta_aceptada: []
  };

  for (let c = 0; c < headers.length; c++) {
    const header = (headers[c] || "").toString().trim(); // ← SIN lower/upper flexible: EXACTO
    if (!header) continue;

    const values = [];
    for (let r = 1; r < data.length; r++) {
      const v = (data[r][c] || "").toString().trim();
      if (v) values.push(v);
    }

    switch (header) {
      case "Lista_Numero":               map.numero = values; break;
      case "Lista_Situacion_Actual":     map.situacion_actual = values; break;
      case "Lista_ODN":                  map.odn = values; break;

      // ⬇️ LAS QUE PEDISTE (exactamente como están nombradas en tu hoja)
      case "Lista_Cod":                  map.cod_area = values; break;
      case "Lista_Equipo":               map.modelo_equipo = values; break;
      case "Lista_Estatus":              map.estatus_local = values; break;
      case "Lista_Situacion":            map.situacion_abordaje = values; break;
      case "Lista_Resp":
        map.oferta_entregada = values;
        map.oferta_aceptada = values;
        break;
    }
  }

  console.log("✅ Listas fijas (mapeo exacto):", JSON.stringify(map, null, 2));
  return map;
} // ← 🔒 ESTA LLAVE CIERRA CORRECTAMENTE LA FUNCIÓN

/**
 * 🔹 Devuelve todas las regiones desde la hoja "Lista_Region".
 */
function getRegiones() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName("Lista_Region");
  if (!sheet) throw new Error("❌ No se encontró la hoja 'Lista_Region'");
  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues();
  const regiones = data.flat().filter(String);
  console.log("🌍 Regiones:", regiones);
  return regiones;
}

/**
 * 🔹 Devuelve los gerentes asociados a una región desde la hoja "Lista_Gerente".
 * Estructura esperada: Región | Gerente
 */
function getGerentesPorRegion(region) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName("Lista_Gerente");
  if (!sheet) throw new Error("❌ No se encontró la hoja 'Lista_Gerente'");
  const data = sheet.getDataRange().getValues();
  const gerentes = data
    .filter(r => r[0] === region)
    .map(r => r[1])
    .filter(String);
  console.log(`👤 Gerentes de ${region}:`, gerentes);
  return gerentes;
}

/**
 * 🔹 Devuelve los estados asociados a una región desde la hoja "Lista_Estado".
 * Estructura esperada: Región | Estado
 */
function getEstadosPorRegion(region) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName("Lista_Estado");
  if (!sheet) throw new Error("❌ No se encontró la hoja 'Lista_Estado'");
  const data = sheet.getDataRange().getValues();
  const estados = data
    .filter(r => r[0] === region)
    .map(r => r[1])
    .filter(String);
  console.log(`🗺️ Estados de ${region}:`, estados);
  return estados;
}

/**
 * 🔹 Devuelve los municipios asociados a un estado desde la hoja "Lista_Municipio".
 * Estructura esperada: Estado | Municipio
 */
function getMunicipiosPorEstado(estado) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName("Lista_Municipio");
  if (!sheet) throw new Error("❌ No se encontró la hoja 'Lista_Municipio'");
  const data = sheet.getDataRange().getValues();
  const municipios = data
    .filter(r => r[0] === estado)
    .map(r => r[1])
    .filter(String);
  console.log(`🏙️ Municipios de ${estado}:`, municipios);
  return municipios;
}

// ============================================================
// 🧾 GUARDAR FORMULARIO (Versión normalizada para encabezados)
// ============================================================
function submitApplication(datos) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000); // Previene concurrencia (espera hasta 30s)

  try {
    console.log("🟢 DATOS RECIBIDOS:", JSON.stringify(datos, null, 2));

    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(SHEET_NAME);
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];

    console.log("📋 ENCABEZADOS:", headers);

    // 🔧 Normalizamos encabezados del Sheet para que coincidan con las claves técnicas del formulario
    const normalizar = h => h
      .toLowerCase()
      .replace(/\s+/g, '_')                     // reemplaza espacios por _
      .normalize('NFD').replace(/[\u0300-\u036f]/g, ''); // quita acentos/ñ

    // ============================================================
// 🆔 GENERAR ID UNICO gpon-[número]-[DDMMAAAA]-[HHMMSS]
// ============================================================
const idCol = headers.findIndex(h => normalizar(h) === 'id') + 1;  // Columna 'ID' en Sheets
let nextNum = 1;

if (idCol > 0 && sheet.getLastRow() > 1) {
  const ids = sheet.getRange(2, idCol, sheet.getLastRow() - 1, 1).getValues().flat();
  const numeros = ids
    .map(v => {
      if (v && typeof v === "string" && v.startsWith("gpon-")) {
        const n = parseInt(v.split("-")[1]);
        return isNaN(n) ? 0 : n;
      }
      return 0;
    })
    .filter(n => n > 0);

  if (numeros.length > 0) {
    nextNum = Math.max(...numeros) + 1; // Toma el mayor número existente, sin importar el orden
  }
}

const fecha = Utilities.formatDate(new Date(), "America/Caracas", "ddMMyyyy");
const hora  = Utilities.formatDate(new Date(), "America/Caracas", "HHmmss");
const nuevoID = `gpon-${nextNum}-${fecha}-${hora}`;
datos.id_registro = nuevoID;

    // ============================================================
    // 🧱 CREAR FILA ORDENADA SEGÚN ENCABEZADOS
    // ============================================================
    // Compatibilidad: la columna en el Sheet es "ID" pero el campo técnico es "id_registro"
    const row = headers.map(h => {
    const key = normalizar(h);
    if (key === 'id') return datos['id_registro'] || '';  // ← Mapea id_registro → ID
    return datos[key] || '';
    });

    console.log("📦 FILA A INSERTAR:", row);

    // ============================================================
    // 📍 INSERTAR SIEMPRE EN LA FILA 2
    // ============================================================
    sheet.insertRows(2, 1);
    sheet.getRange(2, 1, 1, row.length).setValues([row]);
    console.log(`✅ FILA INSERTADA EN LA FILA 2 CON ID: ${nuevoID}`);

    return `✅ Registro guardado correctamente en con ID ${nuevoID}.`;

  } catch (err) {
    console.error("❌ ERROR AL GUARDAR:", err);
    throw err;
  } finally {
    lock.releaseLock(); // libera el bloqueo
  }
}

// ============================================================
// 🔍 MÓDULO: MODIFICAR REGISTRO
// ============================================================

/**
 * Devuelve todos los registros que coinciden con el RIF buscado.
 * Se usa en el módulo "Modificar Registro".
 */
function getRegistrosPorRif(rifQuery) {
  rifQuery = (rifQuery || '').toString().trim().toLowerCase();
  if (!rifQuery) return [];

  console.log("🔎 Buscando registros por RIF:", rifQuery);

  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) throw new Error(`❌ No se encontró la hoja "${SHEET_NAME}"`);

  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow < 2) return [];

  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];

  // Normalizador: minúsculas + sin acentos
  const norm = s => (s || '').toString().trim().toLowerCase()
    .normalize('NFD').replace(/[\u0300-\u036f]/g, '');

  // Índice de encabezados
  const idx = {};
  headers.forEach((h, i) => {
    const key = norm(h);
    if (key) idx[key] = i;
  });

// 🔎 Log de todos los encabezados normalizados  -  LUEGO ELIMINAR FILA 280 Y 281
console.log("📑 Índices de encabezados:", JSON.stringify(idx));

  // Aux: obtener por uno de varios nombres/alias
  const pick = (row, ...names) => {
    for (let n of names) {
      const i = idx[norm(n)];
      if (i !== undefined) return row[i];
    }
    return '';
  };

  const data = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
  const resultados = [];

  data.forEach(row => {
    const rif = (pick(row, 'rif') || '').toString().toLowerCase();
    if (!rif.includes(rifQuery)) return;

// 🔎 Log puntual de Fecha Abordaje   -   LUEGO ELIMINAR LINEA 296 A 298
        const fechaRaw = pick(row, 'fecha_abordaje', 'Fecha Abordaje', 'fecha abordaje');
    console.log("🗓 Fecha Abordaje cruda:", fechaRaw, typeof fechaRaw);

      let fechaISO = '';
      if (fechaRaw instanceof Date) {
        const y = fechaRaw.getFullYear();
        const m = String(fechaRaw.getMonth() + 1).padStart(2, '0');
        const d = String(fechaRaw.getDate()).padStart(2, '0');
        fechaISO = `${y}-${m}-${d}`;
      } else if (typeof fechaRaw === 'string' && /^\d{2}\/\d{2}\/\d{4}$/.test(fechaRaw)) {
        const [d, m, y] = fechaRaw.split('/');
        fechaISO = `${y}-${m.padStart(2,'0')}-${d.padStart(2,'0')}`;
      }

    const reg = {
      id_registro: pick(row, 'id_registro', 'id'),
      fecha_registro: pick(row, 'fecha_registro'),
      tipo_data: pick(row, 'tipo_data'),
      numero: pick(row, 'numero'),
      nombre_contacto: pick(row, 'nombre contacto'),
      cedula: pick(row, 'cedula'),
      correo: pick(row, 'correo'),
      cod_area: pick(row, 'cod area', 'codigo_area'),
      telefono_completo: pick(row, 'telefono_completo', 'telefono'),
      cliente: pick(row, 'cliente', 'razon_social', 'cliente / razón social', 'cliente / razon social'),
      rif: pick(row, 'rif'),
      direccion_rif: pick(row, 'direccion rif'),
      calle: pick(row, 'calle'),
      edificio: pick(row, 'edificio'),
      piso: pick(row, 'piso'),
      local: pick(row, 'local'),

      // 🔧 Tolerantes a acentos/variantes:
      region: pick(row, 'Región', 'Region', 'region', 'region gpon'),
      gerente_region: pick(row, 'gerente_region', 'gerente region'),
      estado: pick(row, 'estado'),
      municipio: pick(row, 'municipio'),
      ciudad: pick(row, 'ciudad'),

      contrato: pick(row, 'contrato'),
      modelo_equipo: pick(row, 'modelo_equipo', 'modelo equipo'),
      serial_asignado: pick(row, 'serial asignado', 'serial'),
      fecha_abordaje: fechaISO,
      estatus_local: pick(row, 'estatus local', 'estatus'),
      oferta_entregada: pick(row, 'oferta entregada'),
      oferta_aceptada: pick(row, 'oferta aceptada'),
      situacion_abordaje: pick(row, 'situacion_abordaje', 'situación abordaje'),

      // 🟦 Campos que pediste:
      situacion_actual: pick(row, 'Situacion Actual', 'Situación Actual', 'situacion_actual'),
      numero_servicio_700: pick(row, 'Número Servicio 700', 'Numero Servicio 700', 'N° Servicio 700', 'Nº Servicio 700', 'numero_servicio_700', 'nro_servicio_700'),

      equipo: pick(row, 'equipo'),
      odn: pick(row, 'odn'),
      observaciones: pick(row, 'observaciones')
    };

    resultados.push(reg);
  });

  console.log(`✅ ${resultados.length} registro(s) encontrado(s) para el RIF ${rifQuery}`);
  return resultados;
}

/**
 * Devuelve un registro completo según su ID (columna "ID" en Sheets).
 * Se usa cuando el usuario pulsa "Modificar" en el overlay.
 */
function getRegistroPorId(idRegistro) {
  idRegistro = (idRegistro || '').toString().trim().toLowerCase();
  if (!idRegistro) return null;

  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) throw new Error(`Hoja no encontrada: ${SHEET_NAME}`);

  const lastRow = sh.getLastRow(), lastCol = sh.getLastColumn();
  if (lastRow < 2) return null;

  const norm = s => (s || '').toString().trim().toLowerCase()
    .normalize('NFD').replace(/[\u0300-\u036f]/g, '');

  const headers = sh.getRange(1,1,1,lastCol).getValues()[0];
  const headersNorm = headers.map(h => norm(h));

  // Si tu columna de ID se llama distinto, ajusta aquí:
  let idCol = headersNorm.indexOf('id');
  if (idCol === -1) idCol = headersNorm.indexOf('id_registro');
  if (idCol === -1) throw new Error('No existe la columna "ID" ni "id_registro" en encabezados');

  const rows = sh.getRange(2,1,lastRow-1,lastCol).getValues();

  for (let r = 0; r < rows.length; r++) {
    const row = rows[r];
    const id = (row[idCol] || '').toString().trim().toLowerCase();
    if (id === idRegistro) {
      const obj = {};
      headers.forEach((h, i) => { obj[norm(h)] = row[i]; });
      return obj;
    }
  }

  return null;
}

/**
 * Igual que getRegistroPorId, pero además retorna el índice de fila (2-based).
 * Respuesta: { data: { ...campos... }, row_index: number }
 */
function getRegistroPorIdYFila(idRegistro) {
  idRegistro = (idRegistro || '').toString().trim().toLowerCase();
  if (!idRegistro) return null;

  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) throw new Error(`Hoja no encontrada: ${SHEET_NAME}`);

  const lastRow = sh.getLastRow(), lastCol = sh.getLastColumn();
  if (lastRow < 2) return null;

  const norm = s => (s || '').toString().trim().toLowerCase()
    .normalize('NFD').replace(/[\u0300-\u036f]/g, '');

  const headers = sh.getRange(1,1,1,lastCol).getValues()[0];
  const headersNorm = headers.map(h => norm(h));

  let idCol = headersNorm.indexOf('id');
  if (idCol === -1) idCol = headersNorm.indexOf('id_registro');
  if (idCol === -1) throw new Error('No existe la columna "ID" ni "id_registro" en encabezados');

  const rows = sh.getRange(2,1,lastRow-1,lastCol).getValues();

  for (let r = 0; r < rows.length; r++) {
    const row = rows[r];
    const id = (row[idCol] || '').toString().trim().toLowerCase();
    if (id === idRegistro) {
      const obj = {};
      headers.forEach((h, i) => { obj[norm(h)] = row[i]; });

      // Normaliza: si el encabezado es "id", mapea a "id_registro" para el form
      if (obj['id'] && !obj['id_registro']) obj['id_registro'] = obj['id'];

      return { data: obj, row_index: r + 2 }; // +2 porque la primera fila de datos es la 2
    }
  }
  return null;
}

/**
 * 🔄 Actualiza un registro existente en la hoja según su ID.
 * @param {string} idRegistro - ID único del registro (columna "ID" o "id_registro").
 * @param {Object} nuevosDatos - Objeto con los valores del formulario.
 */
function updateRegistro(idRegistro, nuevosDatos) {
  if (!idRegistro) {
    throw new Error("❌ No se recibió un ID válido desde el formulario.");
  }

  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sh = ss.getSheetByName(SHEET_NAME);
  const lastRow = sh.getLastRow();
  const lastCol = sh.getLastColumn();
  if (lastRow < 2) return "❌ No hay registros para actualizar.";

  const headers = sh.getRange(1, 1, 1, lastCol).getValues()[0];

  // Normalizador idéntico al de submitApplication (local a updateRegistro)
  const normalizar = h => h
    .toLowerCase()
    .replace(/\s+/g, '_')
    .normalize('NFD').replace(/[\u0300-\u036f]/g, '');

  const headersNorm = headers.map(h => normalizar(h));

  // Buscar columna ID
  let idCol = headersNorm.indexOf('id');
  if (idCol === -1) idCol = headersNorm.indexOf('id_registro');
  if (idCol === -1) throw new Error("❌ No existe columna ID en la hoja.");

  console.log("📨 updateRegistro recibido con ID:", idRegistro);
  console.log("📨 nuevosDatos:", JSON.stringify(nuevosDatos, null, 2));

  const hoy = new Date();
  const fechaISO = Utilities.formatDate(hoy, "America/Caracas", "dd/MM/yyyy, HH:mm:ss");
  const usuario = Session.getActiveUser().getEmail() || "desconocido";

  const rows = sh.getRange(2, 1, lastRow - 1, lastCol).getValues();
  for (let r = 0; r < rows.length; r++) {
    const id = (rows[r][idCol] || '').toString().trim().toLowerCase();
    if (id === idRegistro.trim().toLowerCase()) {

      // Construcción de fila usando el mismo normalizador
      const row = headers.map(h => {
        const key = normalizar(h);

        if (key === 'id' || key === 'id_registro') return idRegistro;
        if (key === 'usuario_modificacion') return usuario;
        if (key === 'fecha_modificacion') return fechaISO;

        // Caso especial teléfono si tu hoja tiene "telefono" como columna única
        if (key === 'telefono') {
          const telConcat = `${(nuevosDatos.cod_area || '').toString().trim()}${(nuevosDatos.telefono_completo || '').toString().trim()}`;
          return telConcat || '';
        }

        return nuevosDatos[key] ?? '';
      });

      sh.getRange(r + 2, 1, 1, row.length).setValues([row]);
      console.log(`✅ Registro ${idRegistro} actualizado en fila ${r+2}`);
      return `✅ Registro ${idRegistro} actualizado correctamente.`;
    }
  }

  return `❌ No se encontró el registro con ID ${idRegistro}.`;
}
