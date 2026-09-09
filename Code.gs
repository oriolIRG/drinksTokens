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
  const sheet = SpreadsheetApp.getActive().getSheetByName(SHEET_NAME);
  if (!sheet) throw new Error('Sheet "' + SHEET_NAME + '" not found');

  const values = sheet.getDataRange().getValues();
  const headers = values.shift();
  const idx = {};
  headers.forEach((h, i) => idx[h] = i);

  const required = ['redemption_id', 'ticket_id', 'ticket_code', 'item_name', 'catalog_object_id',
    'location_name', 'operational_date', 'package_type', 'package_size', 'package_name',
    'calculated_price_net', 'cost_price_resolved', 'margin_resolved', 'cost_source',
    'price_status', 'package_status', 'match_status'];
  required.forEach(col => {
    if (!(col in idx)) throw new Error('Missing column "' + col + '" in ' + SHEET_NAME);
  });

  const tz = Session.getScriptTimeZone();

  return values.map(r => ({
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

  return {
    total_redemptions: total,
    net_revenue: round2_(netRevenue),
    cos: round2_(cos),
    cos_pct: netRevenue ? round2_((cos / netRevenue) * 100) : 0,
    contribution: round2_(contribution),
    margin_pct: netRevenue ? round2_((contribution / netRevenue) * 100) : 0,
    match_pct: total ? round2_((matched / total) * 100) : 0,
    price_match_pct: total ? round2_((priceMatched / total) * 100) : 0,
    top_products: getProductData().slice(0, 10)
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
      byProduct[key] = { catalog_object_id: key, item_name: r.item_name, count: 0, cos: 0, costs: [] };
    }
    const p = byProduct[key];
    p.count += 1;
    p.cos += r.cos;
    p.costs.push(r.cos);
  });

  let products = Object.values(byProduct).map(p => ({
    catalog_object_id: p.catalog_object_id,
    item_name: p.item_name,
    count: p.count,
    cos: round2_(p.cos),
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
// Packages — breakage (Purch_raw + Rdmp_raw directly, NOT Consolidated:
// an unredeemed token never generates a Consolidated row by definition)
// ---------------------------------------------------------------------------

function getBreakageData() {
  const purchRows = sheetToObjects_('Purch_raw');
  const rdmpRows = sheetToObjects_('Rdmp_raw');
  const redeemedCodes = new Set(rdmpRows.map(r => r.ticket_code));

  const packageByOrder = buildPackageByOrder_(purchRows);
  const byPackage = {};

  purchRows.forEach(r => {
    if (!isTokenLine_(r)) return; // skip the aggregation/parent line
    const pkg = packageByOrder[r.square_order_id];
    if (!pkg) return; // shouldn't happen — only 6 packages exist

    const key = pkg.package_name;
    if (!byPackage[key]) {
      byPackage[key] = {
        package_name: key, package_type: pkg.package_type, package_size: pkg.package_size,
        units_sold: 0, unredeemed: 0, gross: 0
      };
    }
    const b = byPackage[key];
    b.units_sold += 1;
    b.gross += Number(r.calculated_price_paid) || 0;
    if (!redeemedCodes.has(r.ticket_code)) b.unredeemed += 1;
  });

  const packages = Object.values(byPackage).map(b => ({
    package_name: b.package_name,
    package_type: b.package_type,
    package_size: b.package_size,
    units_sold: b.units_sold,
    unredeemed: b.unredeemed,
    pct_unredeemed: b.units_sold ? round2_((b.unredeemed / b.units_sold) * 100) : 0,
    gross: round2_(b.gross),
    net: round2_(b.gross / (1 + VAT_RATE_WEBAPP))
  }));

  packages.sort((a, b) => b.units_sold - a.units_sold);
  return packages;
}

// Breakage trend by event_date — mirrors the PDF's "rate is trending up" note
function getBreakageTrendData() {
  const purchRows = sheetToObjects_('Purch_raw');
  const rdmpRows = sheetToObjects_('Rdmp_raw');
  const redeemedCodes = new Set(rdmpRows.map(r => r.ticket_code));

  const byDate = {};
  purchRows.forEach(r => {
    if (!isTokenLine_(r)) return;
    const date = r.event_date;
    if (!byDate[date]) byDate[date] = { date: date, units_sold: 0, unredeemed: 0 };
    byDate[date].units_sold += 1;
    if (!redeemedCodes.has(r.ticket_code)) byDate[date].unredeemed += 1;
  });

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
  const purchRows = sheetToObjects_('Purch_raw');
  const byPackage = {};

  purchRows.forEach(r => {
    // Use PARENT (aggregation) rows only — purchase_date/event_date are the
    // same for every token in that purchase, counting children would just
    // duplicate the same window N times per package size.
    if (Number(r.total_redemptions) <= 1) return;
    const parsed = parsePackageName_(r.ticket_type_name);
    if (!parsed) return;
    if (r.payment_status !== 'COMPLETED') return; // abandoned carts have no real purchase window

    const purchaseDate = parseDMY_(r.purchase_date);
    const eventDate = parseDMY_(r.event_date);
    const days = Math.round((eventDate - purchaseDate) / 86400000);

    const key = parsed.package_name;
    if (!byPackage[key]) byPackage[key] = [];
    byPackage[key].push(days);
  });

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
  const purchRows = sheetToObjects_('Purch_raw');
  const carts = purchRows.filter(r => Number(r.total_redemptions) > 1); // one row per purchase attempt
  const total = carts.length;
  const abandoned = carts.filter(r => r.payment_status !== 'COMPLETED').length;

  return {
    total_carts: total,
    abandoned_carts: abandoned,
    abandonment_pct: total ? round2_((abandoned / total) * 100) : 0
  };
}

// Redemptions by hour of day — helps read bar throughput at peak times
function getRedemptionHourData() {
  const rdmpRows = sheetToObjects_('Rdmp_raw');
  const byHour = {};
  for (let h = 0; h < 24; h++) byHour[h] = 0;

  rdmpRows.forEach(r => {
    const hour = extractHour_(r.redeemed_at);
    if (hour !== null) byHour[hour] += 1;
  });

  return Object.keys(byHour).map(h => ({ hour: Number(h), count: byHour[h] }));
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

// ---------------------------------------------------------------------------
// Data Quality — always read fresh, never cached
// ---------------------------------------------------------------------------

function getDataQualityData() {
  const rows = getConsolidatedData_();
  const byLocationSource = {};
  const byDateMatchStatus = {};
  const byDatePriceStatus = {};
  const byDatePackageStatus = {};
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
  });

  return {
    locations: Array.from(locations),
    sources: Array.from(sources),
    match_statuses: Array.from(matchStatuses),
    price_statuses: Array.from(priceStatuses),
    package_statuses: Array.from(packageStatuses),
    cost_source_by_location: byLocationSource,
    match_status_by_date: byDateMatchStatus,
    price_status_by_date: byDatePriceStatus,
    package_status_by_date: byDatePackageStatus
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

// dd/mm/yyyy -> epoch ms, for sorting/subtracting date-only fields from Purch_raw
function parseDMY_(str) {
  const parts = String(str).split('/');
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