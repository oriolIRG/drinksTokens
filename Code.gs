/**
 * 528 Ibiza — Pre-Sale Packages Analysis Webapp
 * Reads Consolidated (redemption-level) for financial/product views,
 * and Purch_raw + Rdmp_raw directly (via helpers shared with Consolidate.gs,
 * same Apps Script project) for breakage/behavior views, since a
 * never-redeemed token by definition never appears in Consolidated.
 */

const SHEET_NAME = 'Consolidated';
const VAT_RATE_WEBAPP = 0.10; // must match VAT_RATE in Consolidate.gs

// ---------------------------------------------------------------------------
// Request-scoped cache — SpreadsheetApp.getDataRange().getValues() on
// Purch_raw/Rdmp_raw/Consolidated is the slow part of every tab load. Each
// google.script.run call is a fresh execution (this cache doesn't survive
// between separate client calls), but within ONE call — e.g. getPackagesTabData()
// fanning out to 4 helper functions that each used to re-read the same
// sheets — this turns N reads into 1. Combined with batching the Packages
// tab into a single server call (see getPackagesTabData), this is the fix
// for "charts take forever to load".
// ---------------------------------------------------------------------------
const _sheetCache_ = {};
function getCachedSheetRows_(sheetName) {
  if (!_sheetCache_[sheetName]) _sheetCache_[sheetName] = sheetToObjects_(sheetName);
  return _sheetCache_[sheetName];
}

let _consolidatedCache_ = null;

// Cutoff for "has this event actually happened yet" — an unredeemed token
// for a FUTURE event isn't breakage, it's just an event that hasn't
// happened yet. Using start-of-today in the script's timezone: a token for
// today's event only counts once today has fully passed.
function getPastEventCutoff_() {
  const tz = Session.getScriptTimeZone();
  const now = new Date();
  const todayStr = Utilities.formatDate(now, tz, 'yyyy-MM-dd');
  return new Date(todayStr + 'T00:00:00').getTime();
}

function isPastEvent_(eventDateValue, cutoffMs) {
  const eventMs = parseDMY_(eventDateValue);
  return !isNaN(eventMs) && eventMs < cutoffMs;
}

// ---------------------------------------------------------------------------
// Entry point
// ---------------------------------------------------------------------------

function doGet(e) {
  const template = HtmlService.createTemplateFromFile('Index');
  template.userEmail = Session.getActiveUser().getEmail();
  return template.evaluate()
    .setTitle('528 Ibiza — Pre-Sale Packages Analysis')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function includeHtml_(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// ---------------------------------------------------------------------------
// Consolidated — redemption-level read
// ---------------------------------------------------------------------------

function getConsolidatedData_() {
  if (_consolidatedCache_) return _consolidatedCache_;

  const sheet = SpreadsheetApp.getActive().getSheetByName(SHEET_NAME);
  if (!sheet) throw new Error('Sheet "' + SHEET_NAME + '" not found');

  const values = sheet.getDataRange().getValues();
  const headers = values.shift();
  const idx = {};
  headers.forEach((h, i) => idx[h] = i);

  const required = ['redemption_id', 'ticket_id', 'ticket_code', 'item_name', 'catalog_object_id',
    'location_name', 'operational_date', 'package_type', 'package_size', 'package_name',
    'item_category_bucket', 'is_cascade', 'cascade_direction',
    'calculated_price_net', 'cost_price_resolved', 'margin_resolved', 'cost_source',
    'price_status', 'package_status', 'match_status'];
  required.forEach(col => {
    if (!(col in idx)) throw new Error('Missing column "' + col + '" in ' + SHEET_NAME);
  });

  const tz = Session.getScriptTimeZone();

  const mapped = values.map(r => ({
    redemption_id: r[idx['redemption_id']],
    ticket_id: r[idx['ticket_id']],
    ticket_code: r[idx['ticket_code']],
    item_name: r[idx['item_name']],
    catalog_object_id: r[idx['catalog_object_id']],
    location_name: r[idx['location_name']],
    operational_date: Utilities.formatDate(new Date(r[idx['operational_date']]), tz, 'yyyy-MM-dd'),
    package_type: r[idx['package_type']],
    package_size: r[idx['package_size']],
    package_name: r[idx['package_name']] || 'Unknown package',
    item_category_bucket: r[idx['item_category_bucket']],
    is_cascade: r[idx['is_cascade']],
    cascade_direction: r[idx['cascade_direction']],
    // net_revenue = calculated_price_net (already ex-VAT). Named net_revenue in
    // the webapp to match the financial vocabulary used in the P&L review.
    net_revenue: Number(r[idx['calculated_price_net']]) || 0,
    cos: Number(r[idx['cost_price_resolved']]) || 0,          // cost of sale
    contribution: Number(r[idx['margin_resolved']]) || 0,     // net_revenue - cos
    cost_source: r[idx['cost_source']],
    price_status: r[idx['price_status']],
    package_status: r[idx['package_status']],
    match_status: r[idx['match_status']]
  }));
  _consolidatedCache_ = mapped;
  return mapped;
}

// ---------------------------------------------------------------------------
// Overview
// ---------------------------------------------------------------------------

function getOverviewData() {
  const rows = getConsolidatedData_();
  const total = rows.length;
  const matched = rows.filter(r => r.match_status === 'matched').length;
  const priceMatched = rows.filter(r => r.price_status === 'matched').length;
  const netRevenue = sum_(rows, 'net_revenue');
  const cos = sum_(rows, 'cos');
  const contribution = sum_(rows, 'contribution');

  // Breakage across all packages, past events only (see getBreakageData).
  const breakage = getBreakageData();
  const unitsSoldPast = breakage.reduce((s, p) => s + p.units_sold_past_events, 0);
  const unredeemed = breakage.reduce((s, p) => s + p.unredeemed, 0);
  const breakageNetValue = breakage.reduce((s, p) => s + p.unredeemed_net, 0);

  return {
    total_redemptions: total,
    net_revenue: round2_(netRevenue),
    cos: round2_(cos),
    cos_pct: netRevenue ? round2_((cos / netRevenue) * 100) : 0,
    contribution: round2_(contribution),
    margin_pct: netRevenue ? round2_((contribution / netRevenue) * 100) : 0,
    match_pct: total ? round2_((matched / total) * 100) : 0,
    price_match_pct: total ? round2_((priceMatched / total) * 100) : 0,
    breakage_pct: unitsSoldPast ? round2_((unredeemed / unitsSoldPast) * 100) : 0,
    breakage_units_sold: unitsSoldPast,
    breakage_unredeemed: unredeemed,
    breakage_net_value: round2_(breakageNetValue)
  };
}

// ---------------------------------------------------------------------------
// Locations (== points of service inside the 528 Ibiza venue — NOT venues)
// ---------------------------------------------------------------------------

function getLocationData() {
  const rows = getConsolidatedData_();
  const byLocation = {};

  rows.forEach(r => {
    if (!byLocation[r.location_name]) {
      byLocation[r.location_name] = { location_name: r.location_name, count: 0, cos: 0, contribution: 0, net_revenue: 0 };
    }
    const l = byLocation[r.location_name];
    l.count += 1;
    l.cos += r.cos;
    l.contribution += r.contribution;
    l.net_revenue += r.net_revenue;
  });

  return Object.values(byLocation).map(l => ({
    location_name: l.location_name,
    count: l.count,
    cos: round2_(l.cos),
    contribution: round2_(l.contribution),
    net_revenue: round2_(l.net_revenue),
    avg_cos_per_redemption: l.count ? round2_(l.cos / l.count) : 0
  })).sort((a, b) => b.cos - a.cos);
}

// ---------------------------------------------------------------------------
// Products (Pareto)
// ---------------------------------------------------------------------------

function getProductData() {
  const rows = getConsolidatedData_();
  const byProduct = {};

  rows.forEach(r => {
    const key = r.catalog_object_id;
    if (!byProduct[key]) {
      byProduct[key] = {
        catalog_object_id: key, item_name: r.item_name, count: 0,
        cos: 0, net_revenue: 0, contribution: 0, costs: []
      };
    }
    const p = byProduct[key];
    p.count += 1;
    p.cos += r.cos;
    p.net_revenue += r.net_revenue;
    p.contribution += r.contribution;
    p.costs.push(r.cos);
  });

  let products = Object.values(byProduct).map(p => ({
    catalog_object_id: p.catalog_object_id,
    item_name: p.item_name,
    count: p.count,
    net_revenue: round2_(p.net_revenue),
    cos: round2_(p.cos),
    contribution: round2_(p.contribution),
    margin_pct: p.net_revenue ? round2_((p.contribution / p.net_revenue) * 100) : 0,
    avg_unit_cost: round2_(p.cos / p.count),
    min_unit_cost: round2_(Math.min.apply(null, p.costs)),
    max_unit_cost: round2_(Math.max.apply(null, p.costs))
  }));

  products.sort((a, b) => b.cos - a.cos);

  const grandTotal = products.reduce((s, p) => s + p.cos, 0);
  let running = 0;
  products = products.map(p => {
    running += p.cos;
    return Object.assign({}, p, {
      cumulative_pct: grandTotal ? round2_((running / grandTotal) * 100) : 0
    });
  });

  return products;
}

// ---------------------------------------------------------------------------
// Packages — margin table (mirrors the PDF's "Margin by package")
// Grain: redemption (Consolidated). Only redeemed units, same as the PDF's
// "Redeemed units" column — breakage is a separate concern, see below.
// ---------------------------------------------------------------------------

function getPackageMarginData() {
  const rows = getConsolidatedData_();
  const byPackage = {};

  rows.forEach(r => {
    const key = r.package_name;
    if (!byPackage[key]) {
      byPackage[key] = {
        package_name: key, package_type: r.package_type, package_size: r.package_size,
        redeemed: 0, net_revenue: 0, cos: 0
      };
    }
    const p = byPackage[key];
    p.redeemed += 1;
    p.net_revenue += r.net_revenue;
    p.cos += r.cos;
  });

  const packages = Object.values(byPackage).map(p => {
    const contribution = round2_(p.net_revenue - p.cos);
    return {
      package_name: p.package_name,
      package_type: p.package_type,
      package_size: p.package_size,
      redeemed: p.redeemed,
      net_revenue: round2_(p.net_revenue),
      cos: round2_(p.cos),
      cos_pct: p.net_revenue ? round2_((p.cos / p.net_revenue) * 100) : 0,
      contribution: contribution,
      margin_pct: p.net_revenue ? round2_((contribution / p.net_revenue) * 100) : 0
    };
  });

  packages.sort((a, b) => b.net_revenue - a.net_revenue);
  return packages;

  // NOTE — not included, needs data we don't have yet:
  // "Effective discount vs bar" (PDF) needs the full menu/bar price per
  // product, which isn't in Consolidated or Purch_raw. Flag if this becomes
  // a priority; it requires a new price source (Square catalog list price).
}

// ---------------------------------------------------------------------------
// Packages — cascade usage (package_type entitlement vs item_category_bucket
// actually redeemed, e.g. a DRINK pack redeemed as beer or water)
// Grain: package_name (6 rows), not package_type (3 rows) — the old grouping
// hid the 3-vs-5 size split. Also reports avg contribution for cascaded vs
// non-cascaded redemptions within the same package, since "X% cascaded" on
// its own doesn't say whether that's actually costing margin or not.
// ---------------------------------------------------------------------------

function getCascadeData() {
  const rows = getConsolidatedData_();
  const byPackage = {};

  rows.forEach(r => {
    if (!r.package_type || !r.package_name) return;
    if (!byPackage[r.package_name]) {
      byPackage[r.package_name] = {
        package_name: r.package_name, package_type: r.package_type, package_size: r.package_size,
        total: 0, cascaded: 0, unknown: 0,
        contribution_cascaded: 0, contribution_other: 0
      };
    }
    const b = byPackage[r.package_name];
    b.total += 1;
    if (r.is_cascade === 'yes') {
      b.cascaded += 1;
      b.contribution_cascaded += r.contribution;
    } else {
      b.contribution_other += r.contribution;
      if (r.is_cascade === 'unknown') b.unknown += 1;
    }
  });

  return Object.values(byPackage).map(b => {
    const nonCascaded = b.total - b.cascaded;
    return {
      package_name: b.package_name,
      package_type: b.package_type,
      package_size: b.package_size,
      total: b.total,
      cascaded: b.cascaded,
      unknown: b.unknown,
      cascade_pct: b.total ? round2_((b.cascaded / b.total) * 100) : 0,
      unknown_pct: b.total ? round2_((b.unknown / b.total) * 100) : 0,
      avg_contribution_cascaded: b.cascaded ? round2_(b.contribution_cascaded / b.cascaded) : null,
      avg_contribution_other: nonCascaded ? round2_(b.contribution_other / nonCascaded) : null
    };
  }).sort((a, b) => b.total - a.total);
}

// ---------------------------------------------------------------------------
// Packages — single entry point for the whole tab. Used to be 4 separate
// google.script.run calls (overview, cascade, window breakage, location
// mix) from the client, each independently re-reading Consolidated /
// Purch_raw / Rdmp_raw. Bundling them into one call, combined with the
// request-scoped cache above, means each sheet is read once per tab load
// instead of ~6 times — this is the fix for "the charts take forever".
// ---------------------------------------------------------------------------

function getPackagesTabData() {
  return {
    overview: getPackageOverviewData(),
    cascade: getCascadeData(),
    windowBreakage: getBreakageByPurchaseWindowData(),
    windowBreakageDaily: getBreakageByExactDaysData(),
    locationMix: getLocationPackageData()
  };
}

// ---------------------------------------------------------------------------
// Packages — overview scorecard: one row per package (6 rows), every
// dimension side by side (redemption, breakage, cascade, purchase window).
// Deliberately a merge of the four independent server calls above by
// package_name — no new source of truth, just a cross-cut view so the
// person doesn't have to hold four tables in their head to compare a
// package's margin against its own breakage or cascade rate.
// ---------------------------------------------------------------------------

function getPackageOverviewData() {
  const margin = getPackageMarginData();
  const breakage = getBreakageData();
  const cascade = getCascadeData();
  const purchaseWindow = getPurchaseWindowData();

  const byName = {};
  function ensureRow_(name, type, size) {
    if (!byName[name]) byName[name] = { package_name: name, package_type: type || null, package_size: size || null };
    const row = byName[name];
    if (type && !row.package_type) row.package_type = type;
    if (size && !row.package_size) row.package_size = size;
    return row;
  }

  margin.forEach(m => {
    const row = ensureRow_(m.package_name, m.package_type, m.package_size);
    row.redeemed = m.redeemed;
    row.net_revenue = m.net_revenue;
    row.margin_pct = m.margin_pct;
    row.contribution = m.contribution;
  });
  breakage.forEach(b => {
    const row = ensureRow_(b.package_name, b.package_type, b.package_size);
    row.units_sold = b.units_sold;
    row.pct_unredeemed = b.pct_unredeemed;
  });
  cascade.forEach(c => {
    const row = ensureRow_(c.package_name, c.package_type, c.package_size);
    row.cascade_pct = c.cascade_pct;
  });
  purchaseWindow.forEach(w => {
    const row = ensureRow_(w.package_name);
    row.avg_days_before_event = w.avg_days_before_event;
  });

  return Object.values(byName)
    .filter(r => r.package_name && r.package_name !== 'Unknown package')
    .sort((a, b) => (b.units_sold || 0) - (a.units_sold || 0));
}

// ---------------------------------------------------------------------------
// Packages — breakage (Purch_raw + Rdmp_raw directly, NOT Consolidated:
// an unredeemed token never generates a Consolidated row by definition)
//
// IMPORTANT: an unredeemed token only counts as breakage once its event has
// actually happened — a token for an event next month isn't "breakage",
// it just hasn't had its chance to be redeemed yet. `units_sold` below is
// total ever sold (all events, informational); `pct_unredeemed` and
// `unredeemed_net` are computed only over `units_sold_past_events` to avoid
// inflating breakage with future events that simply haven't happened yet.
// ---------------------------------------------------------------------------

function getBreakageData() {
  const purchRows = getCachedSheetRows_('Purch_raw');
  const rdmpRows = getCachedSheetRows_('Rdmp_raw');
  const redeemedCodes = new Set(rdmpRows.map(r => r.ticket_code));
  const cutoff = getPastEventCutoff_();

  const packageByOrder = buildPackageByOrder_(purchRows);
  const byPackage = {};
  let skippedNotTokenLine = 0, skippedNoPackageMatch = 0, skippedFutureEvent = 0;

  purchRows.forEach(r => {
    if (!isTokenLine_(r)) { skippedNotTokenLine++; return; }
    const pkg = packageByOrder[r.square_order_id];
    if (!pkg) { skippedNoPackageMatch++; return; }

    const key = pkg.package_name;
    if (!byPackage[key]) {
      byPackage[key] = {
        package_name: key, package_type: pkg.package_type, package_size: pkg.package_size,
        units_sold: 0, gross: 0, units_sold_past_events: 0, unredeemed: 0, unredeemed_gross: 0
      };
    }
    const b = byPackage[key];
    b.units_sold += 1;
    b.gross += Number(r.calculated_price_paid) || 0;

    if (!isPastEvent_(r.event_date, cutoff)) { skippedFutureEvent++; return; }
    b.units_sold_past_events += 1;
    if (!redeemedCodes.has(r.ticket_code)) {
      b.unredeemed += 1;
      b.unredeemed_gross += Number(r.calculated_price_paid) || 0;
    }
  });

  Logger.log('getBreakageData: purch_rows_total=%s skipped_not_token_line=%s skipped_no_package_match=%s skipped_future_event=%s packages_found=%s',
    purchRows.length, skippedNotTokenLine, skippedNoPackageMatch, skippedFutureEvent, Object.keys(byPackage).length);

  const packages = Object.values(byPackage).map(b => ({
    package_name: b.package_name,
    package_type: b.package_type,
    package_size: b.package_size,
    units_sold: b.units_sold,
    units_sold_past_events: b.units_sold_past_events,
    unredeemed: b.unredeemed,
    pct_unredeemed: b.units_sold_past_events ? round2_((b.unredeemed / b.units_sold_past_events) * 100) : 0,
    unredeemed_gross: round2_(b.unredeemed_gross),
    unredeemed_net: round2_(b.unredeemed_gross / (1 + VAT_RATE_WEBAPP)),
    gross: round2_(b.gross),
    net: round2_(b.gross / (1 + VAT_RATE_WEBAPP))
  }));

  packages.sort((a, b) => b.units_sold - a.units_sold);
  return packages;
}

// PARKED — not wired into the UI (see "Not included yet" in the Packages tab).
// Returns null in production despite validated source code; deployment,
// /dev vs /exec, duplicate functions and the browser console were all
// checked without finding the root cause. The scorecard's "% Unredeemed"
// column already covers the core signal per package, so this isn't blocking
// anything — revisit only if a time-series view of breakage becomes a
// specific ask, at which point re-check root cause before re-wiring it.
// Breakage trend by event_date — mirrors the PDF's "rate is trending up" note
function getBreakageTrendData() {
  const purchRows = getCachedSheetRows_('Purch_raw');
  const rdmpRows = getCachedSheetRows_('Rdmp_raw');
  const redeemedCodes = new Set(rdmpRows.map(r => r.ticket_code));

  const byDate = {};
  let skippedNotTokenLine = 0;
  purchRows.forEach(r => {
    if (!isTokenLine_(r)) { skippedNotTokenLine++; return; }
    const date = r.event_date;
    if (!byDate[date]) byDate[date] = { date: date, units_sold: 0, unredeemed: 0 };
    byDate[date].units_sold += 1;
    if (!redeemedCodes.has(r.ticket_code)) byDate[date].unredeemed += 1;
  });

  Logger.log('getBreakageTrendData: purch_rows_total=%s skipped_not_token_line=%s dates_found=%s',
    purchRows.length, skippedNotTokenLine, Object.keys(byDate).length);

  return Object.values(byDate)
    .map(d => ({
      date: d.date,
      units_sold: d.units_sold,
      unredeemed: d.unredeemed,
      pct_unredeemed: d.units_sold ? round2_((d.unredeemed / d.units_sold) * 100) : 0
    }))
    .sort((a, b) => parseDMY_(a.date) - parseDMY_(b.date));
}

// ---------------------------------------------------------------------------
// Behavior — purchase window, cart abandonment, redemption hour-of-day
// ---------------------------------------------------------------------------

function getPurchaseWindowData() {
  const purchRows = getCachedSheetRows_('Purch_raw');
  const byPackage = {};
  let skippedNotParent = 0, skippedNoPackageMatch = 0, skippedNotCompleted = 0, skippedBadDate = 0;

  purchRows.forEach(r => {
    if (Number(r.total_redemptions) <= 1) { skippedNotParent++; return; }
    const parsed = parsePackageName_(r.ticket_type_name);
    if (!parsed) { skippedNoPackageMatch++; return; }
    if (r.payment_status !== 'COMPLETED') { skippedNotCompleted++; return; }

    const purchaseDate = parseDMY_(r.purchase_date);
    const eventDate = parseDMY_(r.event_date);
    if (isNaN(purchaseDate) || isNaN(eventDate)) { skippedBadDate++; return; }

    const days = Math.round((eventDate - purchaseDate) / 86400000);
    const key = parsed.package_name;
    if (!byPackage[key]) byPackage[key] = [];
    byPackage[key].push(days);
  });

  // Diagnostic — check Executions log if this tab renders empty. A high
  // skippedNoPackageMatch means ticket_type_name doesn't match the expected
  // pattern; a high skippedBadDate means purchase_date/event_date aren't
  // parsing (check their actual format in Purch_raw).
  Logger.log('getPurchaseWindowData: parent_rows_total=%s not_parent=%s no_package_match=%s not_completed=%s bad_date=%s used=%s',
    purchRows.length, skippedNotParent, skippedNoPackageMatch, skippedNotCompleted, skippedBadDate,
    Object.values(byPackage).reduce((s, a) => s + a.length, 0));

  return Object.keys(byPackage).map(key => {
    const arr = byPackage[key];
    const avg = arr.reduce((a, b) => a + b, 0) / arr.length;
    return {
      package_name: key,
      avg_days_before_event: round2_(avg),
      min_days: Math.min.apply(null, arr),
      max_days: Math.max.apply(null, arr),
      n: arr.length
    };
  }).sort((a, b) => b.avg_days_before_event - a.avg_days_before_event);
}

function getCartAbandonmentData() {
  const purchRows = getCachedSheetRows_('Purch_raw');
  const carts = purchRows.filter(r => Number(r.total_redemptions) > 1); // one row per purchase attempt
  const total = carts.length;
  const abandoned = carts.filter(r => r.payment_status !== 'COMPLETED').length;

  return {
    total_carts: total,
    abandoned_carts: abandoned,
    abandonment_pct: total ? round2_((abandoned / total) * 100) : 0
  };
}

// ---------------------------------------------------------------------------
// Packages — breakage by purchase window (does buying last-minute predict a
// lower redemption rate?). Crosses two sources that otherwise live in
// separate tabs: the purchase window comes from the parent row (purchase_date
// vs event_date), the redemption outcome from the child token row — joined
// via square_order_id since only the parent carries both dates.
// ---------------------------------------------------------------------------

const WINDOW_BUCKETS_ = [
  { label: '0-1 days', maxDays: 1 },
  { label: '2-7 days', maxDays: 7 },
  { label: '8+ days', maxDays: Infinity }
];

function bucketDays_(days) {
  for (let i = 0; i < WINDOW_BUCKETS_.length; i++) {
    if (days <= WINDOW_BUCKETS_[i].maxDays) return WINDOW_BUCKETS_[i].label;
  }
  return WINDOW_BUCKETS_[WINDOW_BUCKETS_.length - 1].label;
}

function getBreakageByPurchaseWindowData() {
  const purchRows = getCachedSheetRows_('Purch_raw');
  const rdmpRows = getCachedSheetRows_('Rdmp_raw');
  const redeemedCodes = new Set(rdmpRows.map(r => r.ticket_code));
  const cutoff = getPastEventCutoff_();

  // square_order_id -> { package_name, bucket }, from parent (cart) rows only
  const orderInfo = {};
  let skippedNotCompleted = 0, skippedNoPackageMatch = 0, skippedBadDate = 0;
  purchRows.forEach(r => {
    if (Number(r.total_redemptions) <= 1) return; // only parents carry purchase_date/event_date
    if (r.payment_status !== 'COMPLETED') { skippedNotCompleted++; return; }
    const parsed = parsePackageName_(r.ticket_type_name);
    if (!parsed) { skippedNoPackageMatch++; return; }
    const purchaseDate = parseDMY_(r.purchase_date);
    const eventDate = parseDMY_(r.event_date);
    if (isNaN(purchaseDate) || isNaN(eventDate)) { skippedBadDate++; return; }
    const days = Math.round((eventDate - purchaseDate) / 86400000);
    orderInfo[r.square_order_id] = { package_name: parsed.package_name, bucket: bucketDays_(days) };
  });

  // Same "no future events" rule as getBreakageData: a token whose event
  // hasn't happened yet isn't breakage, it hasn't had its chance.
  const byKey = {};
  let skippedNotTokenLine = 0, skippedNoOrderInfo = 0, skippedFutureEvent = 0;
  purchRows.forEach(r => {
    if (!isTokenLine_(r)) { skippedNotTokenLine++; return; }
    const info = orderInfo[r.square_order_id];
    if (!info) { skippedNoOrderInfo++; return; }
    if (!isPastEvent_(r.event_date, cutoff)) { skippedFutureEvent++; return; }
    const key = info.package_name + '||' + info.bucket;
    if (!byKey[key]) byKey[key] = { package_name: info.package_name, bucket: info.bucket, units_sold: 0, unredeemed: 0 };
    byKey[key].units_sold += 1;
    if (!redeemedCodes.has(r.ticket_code)) byKey[key].unredeemed += 1;
  });

  Logger.log('getBreakageByPurchaseWindowData: not_completed=%s no_package_match=%s bad_date=%s not_token_line=%s no_order_info=%s skipped_future_event=%s',
    skippedNotCompleted, skippedNoPackageMatch, skippedBadDate, skippedNotTokenLine, skippedNoOrderInfo, skippedFutureEvent);

  return Object.values(byKey).map(b => ({
    package_name: b.package_name,
    bucket: b.bucket,
    units_sold: b.units_sold,
    unredeemed: b.unredeemed,
    pct_unredeemed: b.units_sold ? round2_((b.unredeemed / b.units_sold) * 100) : 0
  }));
}

// ---------------------------------------------------------------------------
// Same question as above ("does buying last-minute predict breakage?") but
// at exact-day granularity instead of 3 buckets, all packages pooled — the
// 3-bucket view can hide the actual shape of the curve (e.g. a cliff at a
// specific day rather than a smooth slope). Days beyond DAY_CAP_ are folded
// into a single "DAY_CAP_+" point since the tail gets thin and noisy.
// ---------------------------------------------------------------------------

const WINDOW_DAY_CAP_ = 20;

function getBreakageByExactDaysData() {
  const purchRows = getCachedSheetRows_('Purch_raw');
  const rdmpRows = getCachedSheetRows_('Rdmp_raw');
  const redeemedCodes = new Set(rdmpRows.map(r => r.ticket_code));
  const cutoff = getPastEventCutoff_();

  const daysByOrder = {};
  purchRows.forEach(r => {
    if (Number(r.total_redemptions) <= 1) return;
    if (r.payment_status !== 'COMPLETED') return;
    const purchaseDate = parseDMY_(r.purchase_date);
    const eventDate = parseDMY_(r.event_date);
    if (isNaN(purchaseDate) || isNaN(eventDate)) return;
    daysByOrder[r.square_order_id] = Math.round((eventDate - purchaseDate) / 86400000);
  });

  const byDay = {};
  purchRows.forEach(r => {
    if (!isTokenLine_(r)) return;
    const days = daysByOrder[r.square_order_id];
    if (days === undefined || days < 0) return;
    if (!isPastEvent_(r.event_date, cutoff)) return;

    const bucketDay = Math.min(days, WINDOW_DAY_CAP_);
    if (!byDay[bucketDay]) byDay[bucketDay] = { days: bucketDay, units_sold: 0, unredeemed: 0 };
    byDay[bucketDay].units_sold += 1;
    if (!redeemedCodes.has(r.ticket_code)) byDay[bucketDay].unredeemed += 1;
  });

  return Object.values(byDay)
    .map(b => ({
      days: b.days,
      units_sold: b.units_sold,
      unredeemed: b.unredeemed,
      pct_unredeemed: b.units_sold ? round2_((b.unredeemed / b.units_sold) * 100) : 0
    }))
    .sort((a, b) => a.days - b.days);
}

// ---------------------------------------------------------------------------
// Packages — package mix by location (% share of each location's
// redemptions coming from each package). Location doesn't appear anywhere
// in the Packages tab today; this crosses it in for staffing/inventory
// reads by point of service.
// ---------------------------------------------------------------------------

function getLocationPackageData() {
  const rows = getConsolidatedData_();
  const byLocation = {};
  const packageSet = new Set();

  rows.forEach(r => {
    const loc = r.location_name || 'Unknown';
    const pkg = r.package_name || 'Unknown package';
    packageSet.add(pkg);
    if (!byLocation[loc]) byLocation[loc] = { location_name: loc, total: 0, packages: {} };
    byLocation[loc].total += 1;
    byLocation[loc].packages[pkg] = (byLocation[loc].packages[pkg] || 0) + 1;
  });

  const packages = Array.from(packageSet).sort();
  const locations = Object.values(byLocation).map(l => {
    const shares = {};
    packages.forEach(p => {
      shares[p] = l.total ? round2_(((l.packages[p] || 0) / l.total) * 100) : 0;
    });
    return { location_name: l.location_name, total: l.total, shares: shares };
  }).sort((a, b) => b.total - a.total);

  return { packages: packages, locations: locations };
}

// ---------------------------------------------------------------------------
// Redemptions by hour of day, day of week AND package type — all crossed.
// The old version only returned a flat 24-bucket total, hiding whether
// Tuesdays behave like Saturdays or whether water follows the same hourly
// curve as beer. Returns one row per (hour, day_of_week, package_type)
// combination — small (well under a thousand rows) — so the client can
// filter/compare interactively without extra server round-trips.
// ---------------------------------------------------------------------------

function getRedemptionHourBreakdown() {
  const rdmpRows = getCachedSheetRows_('Rdmp_raw');
  const purchRows = getCachedSheetRows_('Purch_raw');
  const ticketPackageMap = buildTicketToPackageMap_(purchRows);

  const byKey = {};
  rdmpRows.forEach(r => {
    const hour = extractHour_(r.redeemed_at);
    // Day of week comes from operational_date, not the calendar date of
    // redeemed_at — an event starting at 17:00 can run into 4-5am the next
    // calendar day, and operational_date already encodes which event-night
    // those early hours belong to (same field used across Consolidated).
    const dayOfWeek = dayOfWeekFromOperationalDate_(r.operational_date);
    if (hour === null || dayOfWeek === null) return;
    const pkg = ticketPackageMap[String(r.ticket_code || '')];
    const packageType = pkg ? pkg.package_type : 'UNKNOWN';

    const key = hour + '|' + dayOfWeek + '|' + packageType;
    byKey[key] = (byKey[key] || 0) + 1;
  });

  return Object.keys(byKey).map(k => {
    const parts = k.split('|');
    return { hour: Number(parts[0]), day_of_week: Number(parts[1]), package_type: parts[2], count: byKey[k] };
  });
}

// ticket_code (child token) -> { package_type, package_size, package_name },
// mirrors buildPackageMap_ in Consolidate.gs but scoped here since Code.gs
// already has buildPackageByOrder_ (order -> package, from parent rows).
function buildTicketToPackageMap_(purchRows) {
  const packageByOrder = buildPackageByOrder_(purchRows);
  const map = {};
  purchRows.forEach(r => {
    if (isTokenLine_(r)) map[String(r.ticket_code || '')] = packageByOrder[r.square_order_id] || null;
  });
  return map;
}

function extractHour_(value) {
  if (value instanceof Date) return value.getHours();
  const str = String(value || '');
  const parts = str.split(' ');
  const timePart = parts[1];
  if (!timePart) return null;
  const hour = Number(timePart.split(':')[0]);
  return isNaN(hour) ? null : hour;
}

// 0 = Sunday ... 6 = Saturday, matches JS Date#getDay(). Reads operational_date
// directly (already a business-day-aware date from the source data), not the
// calendar date of redeemed_at.
function dayOfWeekFromOperationalDate_(value) {
  const d = value instanceof Date ? value : new Date(value);
  return isNaN(d.getTime()) ? null : d.getDay();
}

// ---------------------------------------------------------------------------
// Data Quality — always read fresh, never cached
// ---------------------------------------------------------------------------

function getDataQualityData() {
  const rows = getConsolidatedData_();
  const byLocationSource = {};
  const byDateMatchStatus = {};
  const byDatePriceStatus = {};
  const byDatePackageStatus = {};
  const byCategory = {};
  const locations = new Set();
  const sources = new Set();
  const matchStatuses = new Set();
  const priceStatuses = new Set();
  const packageStatuses = new Set();

  rows.forEach(r => {
    locations.add(r.location_name);
    sources.add(r.cost_source);
    matchStatuses.add(r.match_status);
    priceStatuses.add(r.price_status);
    packageStatuses.add(r.package_status);

    byLocationSource[r.location_name] = byLocationSource[r.location_name] || {};
    byLocationSource[r.location_name][r.cost_source] = (byLocationSource[r.location_name][r.cost_source] || 0) + 1;

    byDateMatchStatus[r.operational_date] = byDateMatchStatus[r.operational_date] || {};
    byDateMatchStatus[r.operational_date][r.match_status] = (byDateMatchStatus[r.operational_date][r.match_status] || 0) + 1;

    byDatePriceStatus[r.operational_date] = byDatePriceStatus[r.operational_date] || {};
    byDatePriceStatus[r.operational_date][r.price_status] = (byDatePriceStatus[r.operational_date][r.price_status] || 0) + 1;

    byDatePackageStatus[r.operational_date] = byDatePackageStatus[r.operational_date] || {};
    byDatePackageStatus[r.operational_date][r.package_status] = (byDatePackageStatus[r.operational_date][r.package_status] || 0) + 1;

    // item_category_bucket = what was actually served (SPIRIT_COCKTAIL / BEER /
    // WATER_SOFT / OTHER / REVIEW) — a quality check, not a package view: if
    // BEER's average COS% suddenly jumps or goes negative, that's a sign
    // something broke in cost/price resolution for that category, not a
    // real business change.
    const cat = r.item_category_bucket || 'UNKNOWN';
    if (!byCategory[cat]) byCategory[cat] = { category: cat, count: 0, cos: 0, net_revenue: 0 };
    byCategory[cat].count += 1;
    byCategory[cat].cos += r.cos;
    byCategory[cat].net_revenue += r.net_revenue;
  });

  const cosPctByCategory = Object.values(byCategory).map(c => ({
    category: c.category,
    count: c.count,
    net_revenue: round2_(c.net_revenue),
    cos: round2_(c.cos),
    cos_pct: c.net_revenue ? round2_((c.cos / c.net_revenue) * 100) : 0
  })).sort((a, b) => b.count - a.count);

  return {
    locations: Array.from(locations),
    sources: Array.from(sources),
    match_statuses: Array.from(matchStatuses),
    price_statuses: Array.from(priceStatuses),
    package_statuses: Array.from(packageStatuses),
    cost_source_by_location: byLocationSource,
    match_status_by_date: byDateMatchStatus,
    price_status_by_date: byDatePriceStatus,
    package_status_by_date: byDatePackageStatus,
    cos_pct_by_category: cosPctByCategory
  };
}

// ---------------------------------------------------------------------------
// Phase 2 (placeholder) — replace with a real call to the Anthropic API
// ---------------------------------------------------------------------------

function getAiCommentary(sectionId, payload) {
  // TODO Phase 2: UrlFetchApp to api.anthropic.com/v1/messages with the
  // prompt defined earlier, key stored in Script Properties (same pattern
  // as SquareApi.gs / TspoonApi.gs). Returns a visible placeholder for now.
  return {
    text: 'AI-generated commentary — pending Phase 2.',
    pending: true
  };
}

// ---------------------------------------------------------------------------
// Shared helpers (some also used by / mirrored in Consolidate.gs)
// ---------------------------------------------------------------------------

function isTokenLine_(r) {
  const totalRedemptions = Number(r.total_redemptions);
  const ticketCode = String(r.ticket_code || '');
  return totalRedemptions === 1 && ticketCode !== 'PENDING' && ticketCode !== '';
}

function buildPackageByOrder_(purchRows) {
  const map = {};
  purchRows.forEach(r => {
    if (Number(r.total_redemptions) > 1) {
      const parsed = parsePackageName_(r.ticket_type_name);
      if (parsed) map[r.square_order_id] = parsed;
    }
  });
  return map;
}

// dd/mm/yyyy -> epoch ms. Accepts either a text string (as in the raw CSV
// export) or a Date object (if Sheets/BigQuery auto-converts the column) —
// don't assume which one you'll get.
function parseDMY_(value) {
  if (value instanceof Date) return value.getTime();
  const parts = String(value).split('/');
  if (parts.length !== 3) return NaN;
  const day = Number(parts[0]), month = Number(parts[1]), year = Number(parts[2]);
  return new Date(year, month - 1, day).getTime();
}

function sum_(rows, field) {
  return rows.reduce((s, r) => s + (r[field] || 0), 0);
}

function round2_(n) {
  return Math.round((n + Number.EPSILON) * 100) / 100;
}