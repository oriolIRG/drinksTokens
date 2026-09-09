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