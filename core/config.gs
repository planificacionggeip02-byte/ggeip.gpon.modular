/**
 * @file core/config.gs
 * 📁 CONFIGURACIÓN BASE DEL PROYECTO PASAR MODULAR
 * Este archivo concentra las variables globales que identifican
 * el archivo principal de datos y las hojas que utiliza el sistema.
 */

// ============================================================
// 🗂️ CONFIGURACIÓN BASE
// ============================================================

// ID del archivo principal de Google Sheets donde se almacenan los datos.
// ⚠️ IMPORTANTE: Este ID no debe cambiarse a menos que muevas el sistema
// a otro documento.
const SPREADSHEET_ID = "1Dzy49UHef2vi2GIxtzNkW6xneWTmqxhpqQG2-OkH37o";

// Nombre de la hoja principal donde se registran las aplicaciones.
const SHEET_NAME = "Applications";

// Nombre de la hoja que contiene las listas fijas o catálogos de apoyo.
const LISTASFIJAS_SHEET = "ListasFijas";

// ============================================================
// ⚙️ OBJETOS Y FUNCIONES DE CONFIGURACIÓN
// ============================================================

/**
 * Devuelve una referencia al archivo principal del sistema.
 * @returns {SpreadsheetApp.Spreadsheet} Archivo de cálculo activo según SPREADSHEET_ID.
 */
function getMainSpreadsheet() {
  return SpreadsheetApp.openById(SPREADSHEET_ID);
}

/**
 * Devuelve una hoja específica dentro del archivo principal.
 * @param {string} name - Nombre de la hoja que se desea obtener.
 * @returns {GoogleAppsScript.Spreadsheet.Sheet} Hoja solicitada.
 */
function getSheetByName(name) {
  const ss = getMainSpreadsheet();
  const sh = ss.getSheetByName(name);
  if (!sh) throw new Error("❌ No se encontró la hoja: " + name);
  return sh;
}
