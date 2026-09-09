/**
 * Config.gs
 * Único punto de acceso a las Script Properties.
 * Nadie más en el proyecto debe llamar a PropertiesService directamente.
 */

const CONFIG_KEYS = {
  SQUARE_API_KEY: 'SQUARE_API_KEY',
  SQUARE_ENV: 'SQUARE_ENV',                 // "sandbox" | "production"
  SQUARE_LOCATION_IDS: 'SQUARE_LOCATION_IDS' // coma-separado si hay varias
};

const SQUARE_API_BASE = {
  sandbox: 'https://connect.squareupsandbox.com',
  production: 'https://connect.squareup.com'
};

/**
 * Lee una propiedad obligatoria y lanza error claro si falta.
 * Evita que un script a medio ejecutar falle con un undefined silencioso.
 */
function getRequiredProperty_(key) {
  const value = PropertiesService.getScriptProperties().getProperty(key);
  if (!value) {
    throw new Error(
      `Falta la Script Property "${key}". Ve a Configuración del proyecto > ` +
      `Propiedades del script y añádela antes de ejecutar.`
    );
  }
  return value;
}

function getSquareToken() {
  return getRequiredProperty_(CONFIG_KEYS.SQUARE_API_KEY);
}

function getSquareEnv() {
  const env = getRequiredProperty_(CONFIG_KEYS.SQUARE_ENV).toLowerCase();
  if (env !== 'sandbox' && env !== 'production') {
    throw new Error(`SQUARE_ENV debe ser "sandbox" o "production", no "${env}".`);
  }
  return env;
}

function getSquareApiBaseUrl() {
  return SQUARE_API_BASE[getSquareEnv()];
}

function getSquareLocationIds() {
  const sheet = SpreadsheetApp.getActive().getSheetByName('Locations');
  if (!sheet) {
    throw new Error('No existe la pestaña "Locations". Créala con columnas: location_id, location_name, venue_type.');
  }
  const data = sheet.getDataRange().getValues();
  const header = data[0];
  const idCol = header.indexOf('location_id');
  if (idCol === -1) {
    throw new Error('La pestaña "Locations" necesita una columna "location_id".');
  }
  return data.slice(1)
    .map(row => row[idCol])
    .filter(Boolean);
}

/**
 * Ejecútala manualmente (Ejecutar > validateConfig) tras configurar las
 * propiedades. Confirma que todo está bien SIN imprimir el valor de la key
 * en los logs — solo dice si existe y tiene pinta correcta.
 */
function validateConfig() {
  const apiKey = getSquareApiKey();
  const env = getSquareEnv();
  const locationIds = getSquareLocationIds();

  Logger.log('SQUARE_ENV: %s', env);
  Logger.log('SQUARE_API_KEY: %s (longitud %s, %s)',
    apiKey.substring(0, 4) + '••••••••',
    apiKey.length,
    apiKey.length > 20 ? 'longitud plausible' : 'revisa, parece corta'
  );
  Logger.log('SQUARE_LOCATION_IDS: %s ubicación(es) → %s', locationIds.length, locationIds.join(', '));
  Logger.log('Base URL: %s', getSquareApiBaseUrl());
  Logger.log('✓ Configuración OK');
}

// ── Añadir a CONFIG_KEYS en Config.gs ────────────────────────
// CONFIG_KEYS.TSPOON_USER = 'TSPOON_USER'
// CONFIG_KEYS.TSPOON_PASSWORD = 'TSPOON_PASSWORD'

function getTspoonToken_() {
  const cache = CacheService.getScriptCache();
  const cached = cache.get('tspoon_token');
  if (cached) return cached;

  const user = getRequiredProperty_('TSPOON_USER');
  const pass = getRequiredProperty_('TSPOON_PASSWORD');
  const resp = UrlFetchApp.fetch('https://app.tspoonlab.com/recipes/api/login', {
    method: 'post',
    payload: 'username=' + user + '&password=' + pass,
    muteHttpExceptions: true
  });
  if (resp.getResponseCode() !== 200) {
    throw new Error('Login tSpoon falló: ' + resp.getContentText());
  }
  const token = resp.getContentText().trim();
  cache.put('tspoon_token', token, 1200); // 20 min
  return token;
}