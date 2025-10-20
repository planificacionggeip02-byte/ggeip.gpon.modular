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

    const row = headers.map(h => {
      const key = normalizar(h);
      return datos[key] || '';
    });

    console.log("📦 FILA A INSERTAR:", row);

    sheet.appendRow(row);
    console.log("✅ FILA AGREGADA CORRECTAMENTE.");

    return `✅ Registro guardado correctamente en "${SHEET_NAME}".`;

  } catch (err) {
    console.error("❌ ERROR AL GUARDAR:", err);
    throw err;
  }
}
