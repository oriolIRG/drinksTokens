// ═══════════════════════════════════════════════════════════════════
// TSPOONAPI.GS
// Solo lo necesario para este proyecto: costComponent real por PLU.
// No trae inventarios, compras ni mermas — eso vive en el otro proyecto.
// ═══════════════════════════════════════════════════════════════════

const TSPOON_BASE_URL = 'https://app.tspoonlab.com/recipes/api';

function tspoonListCustomers_() {
  const headers = { rememberme: getTspoonToken_() };
  const customers = [];
  let start = 0;
  const rows = 200;
  let page;
  do {
    const url = `${TSPOON_BASE_URL}/listCustomersPaged?start=${start}&rows=${rows}`;
    const resp = UrlFetchApp.fetch(url, { method: 'get', headers, muteHttpExceptions: true });
    if (resp.getResponseCode() !== 200) {
      throw new Error('Error listando clientes tSpoon: ' + resp.getContentText());
    }
    page = JSON.parse(resp.getContentText());
    customers.push(...page);
    start += rows;
  } while (page.length === rows);
  return customers;
}

// ── Trae TODOS los productos a la venta de un cliente, recorriendo
// sus grupos uno a uno (el endpoint no los devuelve todos de golpe) ──
// ── Filtra solo lo relevante: activo y con ventas reales.
// Esto recorta el volumen drásticamente sin perder nada que necesitemos.
function tspoonFetchCostsForCustomer_(idCustomer) {
  const headers = { rememberme: getTspoonToken_() };

  const respCustomer = UrlFetchApp.fetch(`${TSPOON_BASE_URL}/customer/${idCustomer}`, {
    method: 'get', headers, muteHttpExceptions: true
  });
  if (respCustomer.getResponseCode() !== 200) {
    throw new Error(`Error obteniendo cliente ${idCustomer}: ` + respCustomer.getContentText());
  }
  const customer = JSON.parse(respCustomer.getContentText());
  const groupIds = (customer.listGroups || []).map(g => g.id);
  const idGroupsToFetch = [...groupIds, null];

  const items = [];
  idGroupsToFetch.forEach(idGroup => {
    let start = 0;
    const rows = 200;
    let page;
    do {
      let url = `${TSPOON_BASE_URL}/customer/${idCustomer}/components/paged?start=${start}&rows=${rows}`;
      if (idGroup !== null) url += `&idGroup=${idGroup}`;
      const resp = UrlFetchApp.fetch(url, { method: 'get', headers, muteHttpExceptions: true });
      if (resp.getResponseCode() !== 200) {
        throw new Error(`Error en cliente ${idCustomer}, grupo ${idGroup}: ` + resp.getContentText());
      }
      page = JSON.parse(resp.getContentText());
      // deleted/locked = ya no existen en el POS; sin plu no sirven para el cruce
      items.push(...page.filter(c => c.plu && !c.deleted && !c.locked));
      start += rows;
    } while (page.length === rows);
  });

  return items;
}

// ── Rellena Cost_Mapping directamente desde tSpoon (plu = CODI) ──────
// customerIds: array de idCustomer a consultar (los venues de bebida reales).
function refreshCostMappingFromTspoon() {
  Logger.log('=== INICIO refreshCostMappingFromTspoon ===');

  try {
    const locations = sheetToObjects_('Locations');
    const customerIds = [...new Set(
      locations.filter(l => l.match_status === 'matched' && l.tspoon_idCustomer)
               .map(l => l.tspoon_idCustomer)
    )];
    Logger.log(`Locations leídas: ${locations.length}. customerIds: ${customerIds.length}`);

    if (customerIds.length === 0) {
      Logger.log('ABORTA: no hay customerIds.');
      SpreadsheetApp.getUi().alert('No hay ningún venue matcheado en Locations.');
      return;
    }

    const acumulado = {};

    for (let i = 0; i < customerIds.length; i++) {
      const idCustomer = customerIds[i];
      Logger.log(`[${i + 1}/${customerIds.length}] Empezando venue ${idCustomer}...`);

      const items = tspoonFetchCostsForCustomer_(idCustomer);
      Logger.log(`[${i + 1}/${customerIds.length}] Venue ${idCustomer}: ${items.length} items recibidos.`);

      items.forEach(c => {
        if (!c.plu) return;
        const prev = acumulado[c.plu];
        if (prev && prev.source !== idCustomer && Math.abs(prev.cost_price - (c.costComponent || 0)) > 0.01) {
          prev.conflict = true;
        }
        if (prev && prev.source === idCustomer && prev.lastModified && c.lastModifiedShort && prev.lastModified >= c.lastModifiedShort) return;

        acumulado[c.plu] = {
          cost_price: c.costComponent || 0,
          item_name: c.component || '',
          source: idCustomer,
          lastModified: c.lastModifiedShort,
          conflict: prev ? prev.conflict : false
        };
      });

      Logger.log(`[${i + 1}/${customerIds.length}] Acumulado hasta ahora: ${Object.keys(acumulado).length} PLU únicos.`);
    }

    Logger.log(`Bucle terminado. Total PLU acumulados: ${Object.keys(acumulado).length}. Escribiendo en la hoja...`);

    const sheet = SpreadsheetApp.getActive().getSheetByName('Cost_Mapping')
      || SpreadsheetApp.getActive().insertSheet('Cost_Mapping');
    sheet.clearContents();
    Logger.log('clearContents() ejecutado.');

    const headers = ['catalog_object_id', 'item_name', 'cost_price', 'notes'];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

    const rows = Object.keys(acumulado).map(plu => {
      const item = acumulado[plu];
      return [plu, item.item_name, item.cost_price, item.conflict ? 'CONFLICTO entre venues — revisar' : 'auto — tSpoon API'];
    });
    Logger.log(`Filas a escribir: ${rows.length}`);

    if (rows.length) sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
    Logger.log('setValues() de filas ejecutado. Escritura completa.');

    const conflictCount = rows.filter(r => r[3].startsWith('CONFLICTO')).length;
    SpreadsheetApp.getUi().alert(`Cost_Mapping: ${rows.length} PLU cargados.${conflictCount ? ' ⚠️ ' + conflictCount + ' con conflicto.' : ''}`);
    Logger.log('=== FIN OK ===');

  } catch (e) {
    Logger.log('❌ EXCEPCIÓN: ' + e.message);
    Logger.log('Stack: ' + e.stack);
    SpreadsheetApp.getUi().alert('❌ Falló a mitad: ' + e.message + '\n\nMira Ejecuciones > Registros para el detalle completo.');
    throw e; // para que también quede marcada como "fallida" en el panel de Ejecuciones
  }
}
// ── Función pública para poder ejecutarla desde el editor ────────
function verClientesTspoon() {
  const customers = tspoonListCustomers_();
  const texto = customers
    .map(c => `${c.id}  |  ${c.descr}  |  codi: ${c.codi || '(sin codi)'}`)
    .join('\n');

  Logger.log(texto); // por si son muchos y el alert se queda corto
  SpreadsheetApp.getUi().alert(
    `${customers.length} clientes encontrados:\n\n` + texto.substring(0, 1500) +
    (texto.length > 1500 ? '\n\n(...ver el resto en Ejecuciones > Registros)' : '')
  );
}