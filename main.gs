/**
 * @file main.gs
 * Punto de entrada principal del proyecto Pasar Modular.
 * Esta versión está preparada para ejecutarse como una aplicación web
 * (exactamente igual que tu versión original del repositorio GitHub).
 * 
 * Cuando se despliegue como WebApp, Google Apps Script llamará
 * automáticamente a la función doGet(), que devuelve el formulario HTML.
 */

// ============================================================
// 🌐 FUNCIÓN PRINCIPAL DE ACCESO WEB
// ============================================================

/**
 * doGet() se ejecuta automáticamente cuando alguien accede a la URL
 * del despliegue web (Implementar > Implementar como aplicación web).
 * 
 * Su objetivo es devolver la interfaz HTML (index.html) del formulario.
 * Este comportamiento es idéntico al de tu código anterior.
 * 
 * @param {Object} e - Parámetros de la solicitud (no obligatorios).
 * @returns {HtmlOutput} Interfaz HTML del formulario principal.
 */
function doGet(e) {
  return HtmlService.createHtmlOutputFromFile("ui/index")
    .setTitle("📋 Pasar Modular — Formulario Principal")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL); 
    // Permite abrir el formulario en iframe si lo embebes en otra página.
}

// ============================================================
// ⚙️ FUNCIONES DE UTILIDAD (opcional, pero disponibles)
// ============================================================

/**
 * Carga un archivo HTML desde la carpeta ui/
 * y lo devuelve como contenido dinámico dentro del index.
 * 
 * @param {string} filename - Nombre del archivo HTML a incluir.
 * @returns {string} Contenido renderizado listo para insertarse.
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile("ui/" + filename).getContent();
}

function doGet(e) {
  // Si se solicita un módulo específico (?mod=nuevoRegistro)
  if (e && e.parameter.mod) {
    return HtmlService.createHtmlOutputFromFile(`ui/${e.parameter.mod}`);
  }

  // Carga principal
  return HtmlService.createHtmlOutputFromFile('ui/index');
}

// ✅ 🔧 CERRADO correctamente y función independiente
function getModuloHTML(nombre) {
  // Devuelve el contenido crudo del archivo ui/<nombre>.html (sin wrappers del sandbox)
  return HtmlService.createHtmlOutputFromFile('ui/' + nombre).getContent();
}
