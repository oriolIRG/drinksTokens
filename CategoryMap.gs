// ═══════════════════════════════════════════════════════════════════
// CATALOG CATEGORY MAP — catalog_object_id (item o variación) → category_id
// Usa el mismo patrón de búsqueda que DiagnoseDiscounts.gs (searchCatalogByType_).
// Se integra a Consolidate.gs para resolver item_category por redención.
// ═══════════════════════════════════════════════════════════════════

// Devuelve { catalog_object_id: category_id }, cubriendo tanto items como
// variaciones (una redención matchea contra el ID de la VARIACIÓN via Square
// Orders, así que hay que resolver también variación → categoría del item padre).
function buildItemCategoryMap_() {
  const key = PropertiesService.getScriptProperties().getProperty('SQUARE_API_KEY');
  if (!key) throw new Error('Falta SQUARE_API_KEY en Script Properties');

  const items = searchCatalogByType_('ITEM', key);
  const map = {};

  items.forEach(item => {
    const d = item.item_data;
    if (!d) return;

    // Prioridad: reporting_category > categories[0] > category_id (legacy)
    let catId = null;
    if (d.reporting_category && d.reporting_category.id) catId = d.reporting_category.id;
    else if (d.categories && d.categories.length) catId = d.categories[0].id;
    else if (d.category_id) catId = d.category_id;
    if (!catId) return;

    map[item.id] = catId; // el item "padre" en sí
    (d.variations || []).forEach(v => { map[v.id] = catId; }); // cada variación hereda la categoría
  });

  return map;
}

// { category_id: category_name } — para join contra tu hoja de Bucket
function buildCategoryNameMap_() {
  const key = PropertiesService.getScriptProperties().getProperty('SQUARE_API_KEY');
  if (!key) throw new Error('Falta SQUARE_API_KEY en Script Properties');

  const categories = searchCatalogByType_('CATEGORY', key);
  const map = {};
  categories.forEach(c => {
    if (c.category_data && c.category_data.name) map[c.id] = c.category_data.name;
  });
  return map;
}

// Lee tu hoja manual (Nombre Categoria | Id Categoria | Bucket) → { category_id: bucket }
function buildCategoryBucketMap_() {
  const rows = sheetToObjects_('Category_Mapping'); // ajustar nombre de pestaña si es distinto
  const map = {};
  rows.forEach(r => {
    const catId = r['Id Categoria'];
    const bucket = r['Bucket'];
    if (catId && bucket) map[catId] = String(bucket).trim().toUpperCase();
  });
  return map;
}

// ═══════════════════════════════════════════════════════════════════
// GENERADOR DE ESQUELETO — arma/actualiza Category_Mapping desde Square,
// con un Bucket sugerido por palabra clave. Correr una vez para crear la
// hoja, y de nuevo cada tanto para traer categorías nuevas sin pisar las
// que ya corregiste a mano (upsert: respeta el Bucket existente).
// ═══════════════════════════════════════════════════════════════════

const CATEGORY_MAPPING_SHEET = 'Category_Mapping';

function generateCategoryMappingSkeleton() {
  const nameMap = buildCategoryNameMap_(); // { category_id: name }, ya usa SQUARE_API_KEY

  const sheet = SpreadsheetApp.getActive().getSheetByName(CATEGORY_MAPPING_SHEET)
    || SpreadsheetApp.getActive().insertSheet(CATEGORY_MAPPING_SHEET);

  // Preservar buckets ya corregidos a mano si la hoja ya existía
  const existing = {};
  const existingData = sheet.getDataRange().getValues();
  if (existingData.length > 1) {
    const headers = existingData[0];
    const idIdx = headers.indexOf('Id Categoria');
    const bucketIdx = headers.indexOf('Bucket');
    if (idIdx !== -1 && bucketIdx !== -1) {
      existingData.slice(1).forEach(row => {
        if (row[idIdx] && row[bucketIdx]) existing[row[idIdx]] = row[bucketIdx];
      });
    }
  }

  const rows = Object.keys(nameMap).map(id => {
    const name = nameMap[id];
    const bucket = existing[id] || guessBucket_(name); // respeta lo ya corregido, sugiere lo nuevo
    return [id, name, bucket];
  });

  // Orden alfabético por nombre para que sea más fácil de revisar a mano
  rows.sort((a, b) => String(a[1]).localeCompare(String(b[1])));

  sheet.clearContents();
  sheet.getRange(1, 1, 1, 3).setValues([['Id Categoria', 'Nombre Categoria', 'Bucket']]);
  if (rows.length) sheet.getRange(2, 1, rows.length, 3).setValues(rows);

  Logger.log('Category_Mapping generado/actualizado: %s categorías (%s ya tenían Bucket confirmado).',
    rows.length, Object.keys(existing).length);
  Logger.log('Revisá la pestaña "%s" antes de usarla en Consolidate.gs — el Bucket es una SUGERENCIA por palabra clave, no una verdad de negocio.', CATEGORY_MAPPING_SHEET);
}

// Heurística por palabra clave — deliberadamente conservadora: ante la duda,
// devuelve 'REVIEW' en vez de adivinar mal, para que no se filtre un bucket
// incorrecto sin que lo veas.
function guessBucket_(name) {
  const n = String(name || '').toUpperCase();

  if (n.indexOf('BEER') !== -1) return 'BEER';
  if (n.indexOf('SOFT DRINK') !== -1 || n.indexOf('WATER') !== -1) return 'WATER_SOFT';

  const spiritKeywords = ['GIN', 'VODKA', 'RON', 'TEQUILA', 'LICOR', 'WHISKY', 'MEZCAL',
    'BRANDY', 'COÑAC', 'COCKTAIL', 'APERITIVO', 'WINE', 'CAVA', 'CHAMPAGNE'];
  if (spiritKeywords.some(k => n.indexOf(k) !== -1)) return 'SPIRIT_COCKTAIL';

  const otherKeywords = ['FOOD', 'BURGER', 'PIZZA', 'SIDE', 'MERCHANDISE', 'DOOR',
    'SERVICE', 'PREPAYMENT', 'BACKSTAGE', 'STAFF DRINK', 'HELADO', 'COFFEE', 'TICKET'];
  if (otherKeywords.some(k => n.indexOf(k) !== -1)) return 'OTHER';

  return 'REVIEW'; // no matcheó ninguna keyword — mirala vos antes de confiar en ella
}