// ═══════════════════════════════════════════════════════════════════
// SQUAREAPI.GS
// Llamadas a Square para este proyecto: batch-retrieve de orders
// por order_id (no búsqueda por fecha/location) + catálogo.
// ═══════════════════════════════════════════════════════════════════

const SQ_VERSION = '2024-01-18'; // ajusta a la versión que uses en tus otros scripts

function _sqHeaders() {
  return {
    'Authorization' : 'Bearer ' + getSquareToken(),
    'Square-Version': SQ_VERSION,
    'Content-Type'  : 'application/json'
  };
}

function _sqFetchWithRetry(url, options, maxRetries) {
  maxRetries = maxRetries || 3;
  let lastError;
  for (let attempt = 0; attempt < maxRetries; attempt++) {
    try {
      const r = UrlFetchApp.fetch(url, options);
      if (r.getResponseCode() === 429) {
        const wait = Math.pow(2, attempt) * 2000;
        Utilities.sleep(wait);
        continue;
      }
      return r;
    } catch (e) {
      lastError = e;
      if (e.message && e.message.includes('Bandwidth quota')) {
        Utilities.sleep(Math.pow(2, attempt) * 5000);
      } else {
        throw e;
      }
    }
  }
  throw lastError;
}

function _sqPost(path, body) {
  const r = _sqFetchWithRetry('https://connect.squareup.com/v2' + path, {
    method: 'post', headers: _sqHeaders(),
    payload: JSON.stringify(body), muteHttpExceptions: true
  });
  return JSON.parse(r.getContentText());
}

// ── Orders por ID exacto (no por fecha) ──────────────────────────────
// Square permite hasta 500 order_ids por llamada.
function fetchOrdersByIds(orderIds) {
  const orders = [];
  const chunks = [];
  for (let i = 0; i < orderIds.length; i += 500) {
    chunks.push(orderIds.slice(i, i + 500));
  }
  chunks.forEach(chunk => {
    const d = _sqPost('/orders/batch-retrieve', { order_ids: chunk });
    if (d.errors) logMessage_('ERROR', 'fetchOrdersByIds: ' + JSON.stringify(d.errors));
    if (d.orders) orders.push(...d.orders);
  });
  return orders;
}

// ── Catálogo (igual que tu versión probada) ──────────────────────────
function fetchCatalog() {
  function searchType(type) {
    const objs = [];
    let cursor;
    do {
      const body = { object_types: [type], include_deleted_objects: true, limit: 1000 };
      if (cursor) body.cursor = cursor;
      const d = _sqPost('/catalog/search', body);
      if (d.objects) objs.push(...d.objects);
      cursor = d.cursor;
    } while (cursor);
    return objs;
  }
  return {
    variations: searchType('ITEM_VARIATION'),
    items: searchType('ITEM')
  };
}

function logMessage_(level, msg) {
  Logger.log('[%s] %s', level, msg);
}