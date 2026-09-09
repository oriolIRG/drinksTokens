function debugTspoonUnCliente() {
  const items = tspoonFetchCostsForCustomer_('157004632467353328240857992780583296909'); // SQUARE FOOD, del ejemplo
  Logger.log('Total items: ' + items.length);
  Logger.log(JSON.stringify(items.slice(0, 5), null, 2));
  SpreadsheetApp.getUi().alert(`${items.length} productos encontrados para ese cliente. Mira Ejecuciones > Registros para el detalle.`);
}

function debugTspoonVariosClientes() {
  const idsAProbar = [
    '157004632467353328240857992780583296909', // SQUARE FOOD
    '42497413789494174231678491378524725459',  // SQUARE GARDEN
    '8893488071478067228075406580608412101'    // SQUARE PATRON
  ];

  const inicio = new Date().getTime();
  let totalItems = 0;
  const resumen = [];

  idsAProbar.forEach(id => {
    const t0 = new Date().getTime();
    const items = tspoonFetchCostsForCustomer_(id);
    const segundos = ((new Date().getTime() - t0) / 1000).toFixed(1);
    totalItems += items.length;
    resumen.push(`${id}: ${items.length} items en ${segundos}s`);
    Logger.log(`${id}: ${items.length} items en ${segundos}s`);
  });

  const totalSegundos = ((new Date().getTime() - inicio) / 1000).toFixed(1);
  const mensaje = resumen.join('\n') + `\n\nTotal: ${totalItems} items en ${totalSegundos}s para 3 clientes.` +
    `\nProyección para 10: ~${(totalSegundos / 3 * 10).toFixed(0)}s`;

  Logger.log(mensaje);
  SpreadsheetApp.getUi().alert(mensaje);
}

function debugLocationsFilter() {
  const locations = sheetToObjects_('Locations');
  Logger.log('Total filas leídas de Locations: ' + locations.length);
  Logger.log('Primera fila completa: ' + JSON.stringify(locations[0]));

  const customerIds = [...new Set(
    locations.filter(l => l.match_status === 'matched' && l.tspoon_idCustomer)
             .map(l => l.tspoon_idCustomer)
  )];
  Logger.log('customerIds encontrados: ' + customerIds.length);
  Logger.log(JSON.stringify(customerIds));

  SpreadsheetApp.getUi().alert(
    `Locations: ${locations.length} filas.\ncustomerIds válidos: ${customerIds.length}\n\nMira Ejecuciones > Registros para el detalle.`
  );
}

function debugUnaOrdenSquare() {
  const redemptions = sheetToObjects_('Rdmp_raw');
  const someOrderIds = [...new Set(redemptions.map(r => r.square_order_id).filter(Boolean))].slice(0, 2);

  Logger.log('Consultando order_ids: ' + JSON.stringify(someOrderIds));
  const orders = fetchOrdersByIds(someOrderIds);
  Logger.log('Órdenes recibidas: ' + orders.length);

  orders.forEach(order => {
    Logger.log('--- Orden ' + order.id + ' ---');
    Logger.log('line_items: ' + JSON.stringify(order.line_items, null, 2));
  });

  // Y su redención correspondiente, para comparar item_name tal cual está en Rdmp
  const susRedenciones = redemptions.filter(r => someOrderIds.includes(r.square_order_id));
  Logger.log('Redenciones de Rdmp para esas órdenes: ' + JSON.stringify(susRedenciones.map(r => ({square_order_id: r.square_order_id, item_name: r.item_name})), null, 2));
}

// ═══════════════════════════════════════════════════════════════════
// DIAGNOSTIC — Catalog API: descuentos, pricing rules y product sets
// Correr manualmente, mirar el log, pegarme el output. No integrado
// todavía a Consolidate.gs — es solo para ver la forma real del dato
// antes de construir lógica sobre una suposición.
// ═══════════════════════════════════════════════════════════════════

function diagnoseDiscountCategories() {
  const token = PropertiesService.getScriptProperties().getProperty('SQUARE_API_KEY');
  if (!token) throw new Error('Falta SQUARE_API_KEY en Script Properties');

  const objectTypes = ['DISCOUNT', 'PRICING_RULE', 'PRODUCT_SET', 'CATEGORY'];
  const results = {};

  objectTypes.forEach(type => {
    results[type] = searchCatalogByType_(type, token);
  });

  Logger.log('── DISCOUNT objects ──');
  results.DISCOUNT.forEach(o => {
    Logger.log('id=%s name=%s', o.id, o.discount_data && o.discount_data.name);
  });

  Logger.log('── PRICING_RULE objects (vinculan discount ↔ product set) ──');
  results.PRICING_RULE.forEach(o => {
    const d = o.pricing_rule_data || {};
    Logger.log('id=%s name=%s discount_id=%s match_products_id=%s',
      o.id, d.name, d.discount_id, d.match_products_id);
  });

  Logger.log('── PRODUCT_SET objects (qué productos/categorías incluye cada regla) ──');
  results.PRODUCT_SET.forEach(o => {
    const d = o.product_set_data || {};
    Logger.log('id=%s all_products=%s product_ids_any=%s',
      o.id, d.all_products, JSON.stringify(d.product_ids_any || []));
  });

  Logger.log('── CATEGORY objects (nombre real, para cruzar los IDs que ya pegaste a mano) ──');
  results.CATEGORY.forEach(o => {
    Logger.log('id=%s name=%s', o.id, o.category_data && o.category_data.name);
  });

  return results; // por si querés inspeccionarlo desde el editor con el debugger
}

function searchCatalogByType_(objectType, token) {
  const url = 'https://connect.squareup.com/v2/catalog/search';
  let cursor = null;
  const all = [];

  do {
    const payload = { object_types: [objectType], include_related_objects: true };
    if (cursor) payload.cursor = cursor;

    const response = UrlFetchApp.fetch(url, {
      method: 'post',
      contentType: 'application/json',
      headers: { Authorization: 'Bearer ' + token },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    });

    const body = JSON.parse(response.getContentText());
    if (response.getResponseCode() !== 200) {
      Logger.log('Error consultando %s: %s', objectType, response.getContentText());
      return all;
    }
    (body.objects || []).forEach(o => all.push(o));
    cursor = body.cursor || null;
  } while (cursor);

  return all;
}

// ═══════════════════════════════════════════════════════════════════
// DIAGNOSTIC — qué descuento se aplicó realmente en una orden de barra
// Correr manualmente con 2-3 square_order_id reales de Rdmp_raw,
// pegar el log completo antes de construir la lógica de cascada.
// ═══════════════════════════════════════════════════════════════════

function diagnoseOrderDiscount(squareOrderId) {
  const key = PropertiesService.getScriptProperties().getProperty('SQUARE_API_KEY');
  if (!key) throw new Error('Falta square_api_key en Script Properties');

  const url = 'https://connect.squareup.com/v2/orders/' + squareOrderId;
  const response = UrlFetchApp.fetch(url, {
    method: 'get',
    headers: { Authorization: 'Bearer ' + key },
    muteHttpExceptions: true
  });

  if (response.getResponseCode() !== 200) {
    Logger.log('Error: %s', response.getContentText());
    return;
  }

  const order = JSON.parse(response.getContentText()).order;
  const discountsById = {};
  (order.discounts || []).forEach(d => discountsById[d.uid] = d.name);

  Logger.log('── Order %s ──', squareOrderId);
  Logger.log('order.discounts: %s', JSON.stringify(order.discounts || [], null, 2));

  (order.line_items || []).forEach(li => {
    const appliedNames = (li.applied_discounts || []).map(ad => discountsById[ad.discount_uid] || ad.discount_uid);
    Logger.log('Line item: %s (%s) | catalog_object_id=%s | applied_discounts=%s',
      li.name, li.variation_name, li.catalog_object_id, JSON.stringify(appliedNames));
  });
}

// Corré esto una vez por cada square_order_id que quieras chequear, ej:
// diagnoseOrderDiscount('L6p84OQ3z73vqDUpEBbHheKOZSOZY');

function diagnoseUnaOrdenEnConcreto (){
  diagnoseOrderDiscount ('PSgQ3mcucKo6wp1x8z99v9ynCnbZY');
}