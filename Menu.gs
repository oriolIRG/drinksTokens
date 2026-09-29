// ═══════════════════════════════════════════════════════════════════
// MENU.GS — menú personalizado en la Sheet, para no tener que entrar
// al editor de Apps Script cada vez que hay que refrescar costes o
// correr la consolidación. Solo llama a funciones que ya existen en
// TspoonAPI.gs / Consolidate.gs / CategoryMap.gs — no duplica lógica.
// ═══════════════════════════════════════════════════════════════════

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('528 Ibiza')
    .addItem('Refrescar costes tSpoon (mes actual)', 'refreshCostMappingFromTspoonMenu_')
    .addItem('Refrescar costes tSpoon (elegir mes)…', 'refreshCostMappingFromTspoonPrompt_')
    .addSeparator()
    .addItem('Correr consolidación', 'runConsolidationMenu_')
    .addSeparator()
    .addItem('Actualizar Category_Mapping (categorías Square)', 'generateCategoryMappingSkeletonMenu_')
    .addToUi();
}

// ── Refrescar costes: mes actual, un click ──────────────────────────
function refreshCostMappingFromTspoonMenu_() {
  try {
    refreshCostMappingFromTspoon(); // sin argumentos = año/mes actual
  } catch (e) {
    SpreadsheetApp.getUi().alert('Fallo al refrescar costes: ' + e.message);
  }
}

// ── Refrescar costes: pide año y mes (para refrescar un mes pasado) ──
function refreshCostMappingFromTspoonPrompt_() {
  const ui = SpreadsheetApp.getUi();

  const respYear = ui.prompt('Refrescar costes tSpoon', 'Año (ej. 2026):', ui.ButtonSet.OK_CANCEL);
  if (respYear.getSelectedButton() !== ui.Button.OK) return;

  const respMonth = ui.prompt('Refrescar costes tSpoon', 'Mes (1-12):', ui.ButtonSet.OK_CANCEL);
  if (respMonth.getSelectedButton() !== ui.Button.OK) return;

  const year = Number(respYear.getResponseText());
  const month = Number(respMonth.getResponseText());
  if (!year || !month || month < 1 || month > 12) {
    ui.alert('Año o mes inválido — no se hizo ningún cambio.');
    return;
  }

  try {
    refreshCostMappingFromTspoon(year, month);
  } catch (e) {
    ui.alert('Fallo al refrescar costes: ' + e.message);
  }
}

// ── Correr consolidación, con confirmación previa (sobrescribe Consolidated) ──
function runConsolidationMenu_() {
  const ui = SpreadsheetApp.getUi();
  const resp = ui.alert(
    'Correr consolidación',
    'Esto puede tardar unos minutos y va a sobrescribir la pestaña Consolidated. ¿Continuar?',
    ui.ButtonSet.YES_NO
  );
  if (resp !== ui.Button.YES) return;

  try {
    runConsolidation();
    ui.alert('Consolidación terminada. Revisá Ejecuciones > Registros si querés el detalle de match_status/cascada.');
  } catch (e) {
    ui.alert('Fallo en la consolidación: ' + e.message + '\n\nMirá Ejecuciones > Registros para el detalle completo.');
  }
}

// ── Actualizar Category_Mapping (trae categorías nuevas de Square, respeta Bucket ya corregido a mano) ──
function generateCategoryMappingSkeletonMenu_() {
  try {
    generateCategoryMappingSkeleton();
  } catch (e) {
    SpreadsheetApp.getUi().alert('Fallo al actualizar Category_Mapping: ' + e.message);
  }
}