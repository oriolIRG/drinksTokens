// ═══════════════════════════════════════════════════════════════════
// TSPOONAPI.GS
// Solo lo necesario para este proyecto: coste real por PLU (catalog_object_id).
// No trae inventarios, compras ni mermas - eso vive en el otro proyecto.
//
// HISTORIAL DE DISEÑO (por que esta version, no una mas simple):
// - v1: sumaba costComponent de components/paged por plu. Nunca incluia
//   el mixer (esa API solo devuelve el spirit), y ademas el loop de
//   idGroup podia repetir la misma fila, inflando el coste.
// - v2 (esta): components/paged se usa SOLO para el mapeo plu -> idComponent
//   (que producto de Square corresponde a que receta de tSpoon). El COSTE
//   real (ya reconciliado, spirit+mixer, tal como se vendio) sale de
//   reportex/sales, que lo calcula tSpoon mismo a nivel de Order Center
//   completo - no hay nada que reconstruir ni sumar de nuestro lado.
// ═══════════════════════════════════════════════════════════════════

const TSPOON_BASE_URL = 'https://app.tspoonlab.com/recipes/api';

// idOrderCenter de "528 IBIZA" (todos los venues juntos).
const TSPOON_ORDER_CENTER_ID = '151126213482189357415371932181933566892';

function tspoonHeaders_() {
  return {
    rememberme: getTspoonToken_(),
    order: TSPOON_ORDER_CENTER_ID
  };
}

function tspoonListCustomers_() {
  const headers = tspoonHeaders_();
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

// Trae TODOS los productos a la venta de un cliente, recorriendo sus
// grupos uno a uno (el endpoint no los devuelve todos de golpe).
// Usado solo para el mapeo plu -> idComponent, no para el coste en si.
function tspoonFetchCostsForCustomer_(idCustomer) {
  const headers = tspoonHeaders_();

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
      items.push(...page.filter(c => c.plu && !c.deleted && !c.locked));
      start += rows;
    } while (page.length === rows);
  });

  return items;
}

// ── Coste reconciliado real (spirit + mixer), fuente definitiva ─────

function tspoonFetchSalesReport_(year, month) {
  const url = `${TSPOON_BASE_URL}/reportex/sales/${year}/${month}/D/all/${TSPOON_ORDER_CENTER_ID}/none/none/all/all/null/all`;
  const resp = UrlFetchApp.fetch(url, { method: 'get', headers: tspoonHeaders_(), muteHttpExceptions: true });
  if (resp.getResponseCode() !== 200) {
    throw new Error(`Error en reporte de ventas tSpoon (${year}/${month}): ` + resp.getContentText());
  }
  return JSON.parse(resp.getContentText());
}

// { idComponent: costPerUnit } - coste reconciliado real por receta, para
// TODO el periodo/mes pedido, combinando todos los venues de una vez.
function buildReconciledCostByIdComponent_(year, month) {
  const report = tspoonFetchSalesReport_(year, month);
  const map = {};
  (report.listTopComponentsSalesQuantity || []).forEach(item => {
    const q = (item.listQuantities || [])[0];
    if (q && q.costPerUnit != null) {
      map[item.id] = q.costPerUnit;
    }
  });
  return map;
}

// { plu: idComponent } - viene de components/paged, deduplicando por "id"
// (el loop de idGroup puede repetir la misma fila varias veces).
function buildPluToIdComponentMap_(customerIds) {
  const map = {};
  const conflicts = {};

  customerIds.forEach(idCustomer => {
    const items = tspoonFetchCostsForCustomer_(idCustomer);
    const seenIds = {};
    items.forEach(c => {
      if (!c.plu || !c.idComponent || !c.id) return;
      if (seenIds[c.id]) return;
      seenIds[c.id] = true;

      if (map[c.plu] && map[c.plu] !== c.idComponent) {
        conflicts[c.plu] = true;
      }
      map[c.plu] = c.idComponent;
    });
  });

  return { map: map, conflicts: conflicts };
}

// ── Rellena Cost_Mapping (funcion principal, correr manualmente) ────
// year/month opcionales, default al mes actual. Upsert: preserva filas
// existentes de plu sin ventas en el periodo pedido, para no perder
// cobertura de productos que no se vendieron ese mes exacto.
function refreshCostMappingFromTspoon(year, month) {
  year = year || new Date().getFullYear();
  month = month || (new Date().getMonth() + 1);
  Logger.log('=== INICIO refreshCostMappingFromTspoon %s/%s ===', year, month);

  try {
    const locations = sheetToObjects_('Locations');
    const customerIds = [...new Set(
      locations.filter(l => l.match_status === 'matched' && l.tspoon_idCustomer)
               .map(l => l.tspoon_idCustomer)
    )];
    Logger.log('customerIds: %s', customerIds.length);

    if (customerIds.length === 0) {
      SpreadsheetApp.getUi().alert('No hay ningun venue matcheado en Locations.');
      return;
    }

    Logger.log('Trayendo coste reconciliado real (spirit+mixer) desde reportex/sales...');
    const costByIdComponent = buildReconciledCostByIdComponent_(year, month);
    Logger.log('Componentes con coste reconciliado este periodo: %s', Object.keys(costByIdComponent).length);

    Logger.log('Trayendo mapeo plu -> idComponent desde components/paged...');
    const result = buildPluToIdComponentMap_(customerIds);
    const pluToIdComponent = result.map;
    const conflicts = result.conflicts;
    Logger.log('PLU mapeados: %s. Conflictos plu->idComponent entre venues: %s', Object.keys(pluToIdComponent).length, Object.keys(conflicts).length);

    const sheet = SpreadsheetApp.getActive().getSheetByName('Cost_Mapping')
      || SpreadsheetApp.getActive().insertSheet('Cost_Mapping');

    const existing = {};
    const existingData = sheet.getDataRange().getValues();
    if (existingData.length > 1) {
      const eh = existingData[0];
      const pluIdx = eh.indexOf('catalog_object_id');
      const compIdx = eh.indexOf('idComponent');
      const costIdx = eh.indexOf('cost_price');
      const notesIdx = eh.indexOf('notes');
      if (pluIdx !== -1) {
        existingData.slice(1).forEach(row => {
          if (!row[pluIdx]) return;
          existing[row[pluIdx]] = {
            idComponent: compIdx !== -1 ? row[compIdx] : null,
            cost_price: costIdx !== -1 ? row[costIdx] : null,
            notes: notesIdx !== -1 ? row[notesIdx] : null
          };
        });
      }
    }

    let updatedCount = 0, noCostThisPeriod = 0;
    Object.keys(pluToIdComponent).forEach(plu => {
      const idComponent = pluToIdComponent[plu];
      const cost = costByIdComponent[idComponent];
      if (cost == null) { noCostThisPeriod++; return; }
      existing[plu] = {
        idComponent: idComponent,
        cost_price: Math.round(cost * 100) / 100,
        notes: conflicts[plu]
          ? 'CONFLICTO plu->idComponent entre venues - revisar'
          : `auto - tSpoon reportex/sales ${year}/${month} (reconciliado)`
      };
      updatedCount++;
    });

    Logger.log('PLU actualizados con coste fresco de este periodo: %s. Sin ventas este periodo (se preserva valor anterior si existia): %s', updatedCount, noCostThisPeriod);

    sheet.clearContents();
    const headers = ['catalog_object_id', 'idComponent', 'cost_price', 'notes'];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

    const rows = Object.keys(existing).map(plu => {
      const item = existing[plu];
      return [plu, item.idComponent || '', item.cost_price || '', item.notes || ''];
    });
    if (rows.length) sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);

    const conflictCount = rows.filter(r => String(r[3]).startsWith('CONFLICTO')).length;
    SpreadsheetApp.getUi().alert(
      `Cost_Mapping: ${rows.length} PLU totales.\n` +
      `Actualizados con coste reconciliado de ${year}/${month}: ${updatedCount}.\n` +
      `Sin ventas este periodo (valor preservado): ${noCostThisPeriod}.` +
      `${conflictCount ? '\nADVERTENCIA: ' + conflictCount + ' con conflicto.' : ''}`
    );
    Logger.log('=== FIN OK ===');

  } catch (e) {
    Logger.log('EXCEPCION: ' + e.message);
    Logger.log('Stack: ' + e.stack);
    SpreadsheetApp.getUi().alert('Fallo: ' + e.message + '\n\nMira Ejecuciones > Registros para el detalle completo.');
    throw e;
  }
}

function verClientesTspoon() {
  const customers = tspoonListCustomers_();
  const texto = customers
    .map(c => `${c.id}  |  ${c.descr}  |  codi: ${c.codi || '(sin codi)'}`)
    .join('\n');

  Logger.log(texto);
  SpreadsheetApp.getUi().alert(
    `${customers.length} clientes encontrados:\n\n` + texto.substring(0, 1500) +
    (texto.length > 1500 ? '\n\n(...ver el resto en Ejecuciones > Registros)' : '')
  );
}