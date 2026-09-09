// ═══════════════════════════════════════════════════════════════════
// CATALOGSYNC.GS
// Trae el catálogo de Square y lo deja listo para mapear coste.
// ═══════════════════════════════════════════════════════════════════

// ── 1. Refresca Catalog_Cache completo desde Square (se puede pisar) ─
function refreshCatalogCache() {
  const catalog = fetchCatalog();
  const itemNameById = {};
  (catalog.items || []).forEach(item => {
    itemNameById[item.id] = item.item_data && item.item_data.name || '';
  });

  const rows = (catalog.variations || []).map(v => {
    const vd = v.item_variation_data || {};
    const itemName = itemNameById[vd.item_id] || '';
    const variationName = vd.name || '';
    // Muchos POS de barra nombran el item ya con el detalle completo
    // (ej. "ESTRELLA DAMM REGULAR"), y la variación es solo "Regular".
    // Guardamos ambas formas para no perder el match contra line_item.name.
    const fullName = variationName && variationName.toLowerCase() !== 'regular'
      ? `${itemName} ${variationName}`.trim()
      : itemName;

    return {
      catalog_object_id: v.id,
      item_id: vd.item_id || '',
      item_name: itemName,
      variation_name: variationName,
      full_name: fullName,
      price: vd.price_money ? (vd.price_money.amount / 100) : ''
    };
  });

  const sheet = SpreadsheetApp.getActive().getSheetByName('Catalog_Cache')
    || SpreadsheetApp.getActive().insertSheet('Catalog_Cache');
  sheet.clearContents();

  const headers = ['catalog_object_id', 'item_id', 'item_name', 'variation_name', 'full_name', 'price'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  if (rows.length) {
    const values = rows.map(r => headers.map(h => r[h]));
    sheet.getRange(2, 1, values.length, headers.length).setValues(values);
  }
  logMessage_('INFO', `Catalog_Cache actualizado: ${rows.length} variaciones.`);
}

// ── 2. Añade a Cost_Mapping solo lo nuevo, sin pisar lo ya rellenado ──
function syncCostMappingSkeleton() {
  const catalogRows = sheetToObjects_('Catalog_Cache');
  const costSheet = SpreadsheetApp.getActive().getSheetByName('Cost_Mapping')
    || SpreadsheetApp.getActive().insertSheet('Cost_Mapping');

  let existing = [];
  if (costSheet.getLastRow() > 0) {
    existing = sheetToObjects_('Cost_Mapping');
  } else {
    costSheet.getRange(1, 1, 1, 4).setValues([['catalog_object_id', 'item_name', 'cost_price', 'notes']]);
  }
  const existingIds = new Set(existing.map(r => r.catalog_object_id));

  const newRows = catalogRows
    .filter(c => !existingIds.has(c.catalog_object_id))
    .map(c => [c.catalog_object_id, c.full_name || c.item_name, '', 'nuevo — pendiente de coste']);

  if (newRows.length) {
    costSheet.getRange(costSheet.getLastRow() + 1, 1, newRows.length, 4).setValues(newRows);
  }
  logMessage_('INFO', `Cost_Mapping: ${newRows.length} ítems nuevos añadidos (sin tocar los existentes).`);
}