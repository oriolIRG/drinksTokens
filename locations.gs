// ═══════════════════════════════════════════════════════════════════
// LOCATIONS.GS
// Cruza los venues de Rdmp (square_location_id) con los "clientes"
// de tSpoon (idCustomer), por nombre. Deja constancia de lo que no
// matchea automático para revisar a mano — nunca lo asume en silencio.
// ═══════════════════════════════════════════════════════════════════

function normalizeVenueName_(name) {
  return String(name || '')
    .toUpperCase()
    .replace(/^SQUARE\s+/, '')
    .replace(/\s+/g, ' ')
    .trim();
}

function buildLocationCrosswalk() {
  const rdmpLocations = {};
  sheetToObjects_('Rdmp_raw').forEach(r => {
    if (r.square_location_id) {
      rdmpLocations[r.square_location_id] = r.location_name || '';
    }
  });

  const tspoonCustomers = tspoonListCustomers_();
  const tspoonByName = {};
  tspoonCustomers.forEach(c => {
    tspoonByName[normalizeVenueName_(c.descr)] = c;
  });

  const rows = [];
  Object.keys(rdmpLocations).forEach(locationId => {
    const locationName = rdmpLocations[locationId];
    const match = tspoonByName[normalizeVenueName_(locationName)];
    rows.push([
      locationId,
      locationName,
      match ? match.id : '',
      match ? match.descr : '',
      match ? 'matched' : 'SIN MATCH — revisar a mano'
    ]);
  });

  const sheet = SpreadsheetApp.getActive().getSheetByName('Locations')
    || SpreadsheetApp.getActive().insertSheet('Locations');
  sheet.clearContents();
  const headers = ['square_location_id', 'location_name', 'tspoon_idCustomer', 'tspoon_descr', 'match_status'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  if (rows.length) sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);

  const sinMatch = rows.filter(r => r[4] !== 'matched').length;
  SpreadsheetApp.getUi().alert(
    `Locations: ${rows.length} venues de Rdmp, ${rows.length - sinMatch} matcheados con tSpoon.` +
    (sinMatch ? `\n⚠️ ${sinMatch} sin match — revisa la pestaña y rellena tspoon_idCustomer a mano donde falte.` : '')
  );
}
