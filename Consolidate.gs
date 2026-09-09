// ═══════════════════════════════════════════════════════════════════
// CONSOLIDATE.GS
// Cruza Rdmp_raw + Purch_raw (precio) + Square (orders) + Cost_Mapping (tSpoon) → Consolidated
// ═══════════════════════════════════════════════════════════════════

const VAT_RATE = 0.10; // precio de venta incluye 10% IVA; coste (tSpoon) es neto

function runConsolidation() {
  const redemptions = sheetToObjects_('Rdmp_raw');
  const costMap = buildCostMap_();
  const priceMap = buildPriceMap_();
  const packageMap = buildPackageMap_();

  const uniqueOrderIds = [...new Set(redemptions.map(r => r.square_order_id).filter(Boolean))];
  logMessage_('INFO', `Consultando ${uniqueOrderIds.length} órdenes en Square...`);
  const orders = fetchOrdersByIds(uniqueOrderIds);
  const orderIndex = buildOrderIndex_(orders);

  const consolidated = redemptions.map(r => enrichRedemption_(r, orderIndex, costMap, priceMap, packageMap));

  writeConsolidated_(consolidated);
  logMatchSummary_(consolidated);
}

// ── Lectura genérica de una pestaña como objetos ─────────────────────
function sheetToObjects_(sheetName) {
  const sheet = SpreadsheetApp.getActive().getSheetByName(sheetName);
  if (!sheet) throw new Error(`No existe la pestaña "${sheetName}".`);
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  return data.slice(1).map(row => {
    const obj = {};
    headers.forEach((h, i) => obj[h] = row[i]);
    return obj;
  });
}

// ── Cost_Mapping: catalog_object_id → coste (tSpoon) ─────────────────
function buildCostMap_() {
  const rows = sheetToObjects_('Cost_Mapping');
  const map = {};
  rows.forEach(r => {
    if (r.catalog_object_id) map[r.catalog_object_id] = Number(r.cost_price) || null;
  });
  return map;
}

// ── Purch_raw: ticket_code (token individual) → calculated_price_paid ──
// Filas hijas (ticket_code real "FU-XXXXXX", total_redemptions === 1):
// calculated_price_paid ya viene con el precio unitario correcto
// (total pagado / qty del pack), no hace falta recalcularlo a mano.
function buildPriceMap_() {
  const rows = sheetToObjects_('Purch_raw');
  const map = {};
  rows.forEach(r => {
    const totalRedemptions = Number(r.total_redemptions);
    const ticketCode = String(r.ticket_code || '');
    const isTokenLine = totalRedemptions === 1 && ticketCode !== 'PENDING' && ticketCode !== '';
    if (isTokenLine) {
      const price = Number(r.calculated_price_paid);
      if (!isNaN(price)) map[ticketCode] = price;
    }
  });
  return map;
}

// ── Purch_raw: ticket_code (token) → paquete real (tipo + tamaño) ────
// IMPORTANTE, dos saltos necesarios:
//  1. La línea HIJA de cada token (ticket_code real, tipo "FU-XXXXXX")
//     solo dice "PACKAGE DRINKS // ONLINE ONLY" (o BEER/WATER) — tiene
//     el TIPO pero nunca el TAMAÑO (3 o 5).
//  2. El TAMAÑO real solo vive en la línea PADRE de esa misma compra
//     (ticket_code = "PENDING", mismo square_order_id), con nombres
//     como "5 Drink Pack" o, si compró varios de una vez, "2x 5 Drink
//     Pack" — el prefijo "Nx " es multiplicador de compra, no una
//     categoría nueva: cada token dentro sigue siendo un "5 Drink Pack".
// No existen tokens sueltos fuera de estos 6 paquetes (3/5 × Drink/Beer/Water).
function buildPackageMap_() {
  const rows = sheetToObjects_('Purch_raw');

  // Índice: square_order_id → { package_type, package_size, package_name } del padre
  const packageByOrder = {};
  rows.forEach(r => {
    if (Number(r.total_redemptions) > 1) {
      const parsed = parsePackageName_(r.ticket_type_name);
      if (parsed) packageByOrder[r.square_order_id] = parsed;
    }
  });

  // Índice: ticket_code (hijo) → paquete, heredado de su compra padre
  const map = {};
  rows.forEach(r => {
    const totalRedemptions = Number(r.total_redemptions);
    const ticketCode = String(r.ticket_code || '');
    const isTokenLine = totalRedemptions === 1 && ticketCode !== 'PENDING' && ticketCode !== '';
    if (isTokenLine) {
      map[ticketCode] = packageByOrder[r.square_order_id] || null;
    }
  });
  return map;
}

// "2x 5 Drink Pack" -> { package_type:'DRINK', package_size:5, package_name:'5 Drink Pack' }
function parsePackageName_(rawName) {
  const cleaned = String(rawName || '').replace(/^\d+x\s+/i, '').trim();
  const match = cleaned.match(/^(\d+)\s+(Drink|Beer|Water)\s+Pack$/i);
  if (!match) return null; // no debería pasar dado que solo existen estos 6 paquetes
  return {
    package_size: Number(match[1]),
    package_type: match[2].toUpperCase(),
    package_name: cleaned
  };
}

// ── order_id → line items normalizadas ───────────────────────────────
function buildOrderIndex_(orders) {
  const index = {};
  orders.forEach(order => {
    index[order.id] = (order.line_items || []).map(li => {
      const fullName = [li.name, li.variation_name]
        .filter(Boolean)
        .join(' ')
        .trim()
        .toLowerCase();
      return {
        catalog_object_id: li.catalog_object_id || null,
        name: fullName,
        quantity: Number(li.quantity) || 1
      };
    });
  });
  return index;
}

// ── Enriquecer una redención con catalog_object_id, precio, paquete y coste ──
function enrichRedemption_(r, orderIndex, costMap, priceMap, packageMap) {
  const result = Object.assign({}, r, {
    catalog_object_id: null,
    match_status: 'no_order_found',      // resolución de producto vía Square
    price_status: 'no_purch_match',      // resolución de precio vía Purch_raw
    package_status: 'no_package_match',  // resolución de paquete (tipo+tamaño) vía Purch_raw
    calculated_price_gross: null,        // precio pagado, IVA incluido (tal cual Purch_raw)
    calculated_price_net: null,          // mismo precio, neto de IVA — es el que se usa para margen
    package_type: null,                  // DRINK / BEER / WATER
    package_size: null,                  // 3 / 5
    package_name: null,                  // "5 Drink Pack" — combinado legible
    cost_price_resolved: null,           // ya viene neto de tSpoon
    cost_source: 'none',
    margin_resolved: null
  });

  const ticketCode = String(r.ticket_code || '');

  // -- Precio: join directo por ticket_code contra Purch_raw --
  if (priceMap[ticketCode] !== undefined) {
    const gross = priceMap[ticketCode];
    result.calculated_price_gross = gross;
    result.calculated_price_net = Math.round((gross / (1 + VAT_RATE)) * 100) / 100;
    result.price_status = 'matched';
  }

  // -- Paquete: ticket_code (hijo) -> square_order_id -> línea padre --
  const pkg = packageMap[ticketCode];
  if (pkg) {
    result.package_type = pkg.package_type;
    result.package_size = pkg.package_size;
    result.package_name = pkg.package_name;
    result.package_status = 'matched';
  }

  // -- Producto: Square order → line item → catalog_object_id --
  const lineItems = orderIndex[r.square_order_id];
  if (lineItems) {
    const targetName = (r.item_name || '').trim().toLowerCase();
    const matches = lineItems.filter(li => li.name === targetName);

    if (matches.length === 0) {
      result.match_status = 'no_line_item_match';
    } else {
      const distinctCatalogIds = [...new Set(matches.map(m => m.catalog_object_id))];
      result.catalog_object_id = matches[0].catalog_object_id;
      result.match_status = distinctCatalogIds.length === 1
        ? 'matched' // 1+ líneas, todas el mismo producto real — resuelto
        : 'ambiguous_conflicting_ids'; // mismo nombre, productos DISTINTOS — esto sí hay que mirarlo
    }
  }

  // -- Coste: catalog_object_id → Cost_Mapping (tSpoon) --
  if (result.catalog_object_id && costMap[result.catalog_object_id] != null) {
    result.cost_price_resolved = costMap[result.catalog_object_id];
    result.cost_source = 'cost_mapping';
  }

  // -- Margen real: precio NETO de IVA menos coste (ya neto) --
  if (result.calculated_price_net != null && result.cost_price_resolved != null) {
    result.margin_resolved = Math.round((result.calculated_price_net - result.cost_price_resolved) * 100) / 100;
  }

  return result;
}

// ── Escribir Consolidated ────────────────────────────────────────────
function writeConsolidated_(rows) {
  const sheet = SpreadsheetApp.getActive().getSheetByName('Consolidated')
    || SpreadsheetApp.getActive().insertSheet('Consolidated');
  sheet.clearContents();
  if (rows.length === 0) return;

  const headers = [
    'redemption_id', 'ticket_id', 'ticket_code', 'square_order_id',
    'catalog_object_id', 'item_name', 'location_name', 'operational_date',
    'package_type', 'package_size', 'package_name',
    'calculated_price_gross', 'calculated_price_net', 'cost_price_resolved', 'margin_resolved',
    'cost_source', 'price_status', 'package_status', 'match_status'
  ];
  const values = rows.map(r => headers.map(h => r[h] !== undefined && r[h] !== null ? r[h] : ''));
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(2, 1, values.length, headers.length).setValues(values);
}

// ── Resumen de calidad del match, sin abrir la hoja ──────────────────
function logMatchSummary_(rows) {
  const matchCounts = {};
  const priceCounts = {};
  const packageCounts = {};
  rows.forEach(r => {
    matchCounts[r.match_status] = (matchCounts[r.match_status] || 0) + 1;
    priceCounts[r.price_status] = (priceCounts[r.price_status] || 0) + 1;
    packageCounts[r.package_status] = (packageCounts[r.package_status] || 0) + 1;
  });
  Logger.log('── Resumen de match_status (producto, vía Square) ──');
  Object.keys(matchCounts).forEach(k => Logger.log('%s: %s', k, matchCounts[k]));
  Logger.log('── Resumen de price_status (precio, vía Purch_raw) ──');
  Object.keys(priceCounts).forEach(k => Logger.log('%s: %s', k, priceCounts[k]));
  Logger.log('── Resumen de package_status (paquete, vía Purch_raw) ──');
  Object.keys(packageCounts).forEach(k => Logger.log('%s: %s', k, packageCounts[k]));
}