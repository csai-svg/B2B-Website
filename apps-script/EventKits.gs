/**
 * Event Kits: curated product collections for a named occasion (New Joinee
 * Program, Employee Recognition & Rewards, New Mom & Baby Kit, Sustainability,
 * Festive Gift Kits, Personal Milestone, CXO Gifting, Executive Gifting).
 *
 * Unlike kit.html's budget/headcount generator, a kit here is HAND-CURATED:
 * an admin (or this script) sets its product list and, once supplied, its
 * own hero photo. The public storefront never talks to this file directly —
 * it reads the event_kits block inside the published assets/products.json
 * (see buildEventKitsJson() in Admin.gs). This file is the admin-side CRUD
 * on the EventKits sheet tab, plus a couple of one-shot seeding helpers.
 */

function fnAdminEventKitSave(req) {
  requireAdmin(req);
  var k = req.kit || {};
  var slug = String(k.slug || '').trim().toLowerCase().replace(/[^a-z0-9]+/g, '-').replace(/^-+|-+$/g, '');
  var label = String(k.label || '').trim();
  if (!slug) throw new Error('An event kit needs a name to derive its slug from.');
  if (!label) throw new Error('Event kit label is required.');

  var skus = (Array.isArray(k.product_skus) ? k.product_skus : String(k.product_skus || '').split(','))
    .map(function (s) { return String(s).trim().toUpperCase(); })
    .filter(String);

  return withLock(function () {
    var hit = null;
    readTab(SHEETS.EVENT_KITS).forEach(function (r) {
      if (String(r.slug).trim().toLowerCase() === slug) hit = r;
    });

    var row = {
      slug: slug, label: label,
      tagline: String(k.tagline || '').trim(),
      hero_image: String(k.hero_image || '').trim(),
      product_skus: skus.join(','),
      sort_order: Number(k.sort_order || 0),
      active: k.active === false ? 'FALSE' : 'TRUE'
    };

    if (hit) {
      updateRow(SHEETS.EVENT_KITS, hit._row, row);
      audit('admin', 'event_kit_updated', 'event_kit', slug, null, row);
    } else {
      appendRow(SHEETS.EVENT_KITS, row);
      audit('admin', 'event_kit_created', 'event_kit', slug, null, row);
    }
    return { ok: true, slug: slug };
  });
}

function fnAdminEventKitDelete(req) {
  requireAdmin(req);
  var slug = String(req.slug || '').trim().toLowerCase();
  return withLock(function () {
    var hit = null;
    readTab(SHEETS.EVENT_KITS).forEach(function (r) {
      if (String(r.slug).trim().toLowerCase() === slug) hit = r;
    });
    if (!hit) throw new Error('No event kit "' + slug + '".');
    updateRow(SHEETS.EVENT_KITS, hit._row, { active: 'FALSE' });
    audit('admin', 'event_kit_deactivated', 'event_kit', slug, null, null);
    return { ok: true };
  });
}

/**
 * One-time seed: creates the eight standard event kits with an empty product
 * list if the tab is otherwise empty, so the admin console and the nav have
 * something to point at on day one. Safe to run more than once — it never
 * overwrites a kit that already exists (matched by slug).
 */
function seedEventKits() {
  var defaults = [
    ['new-joinee-program', 'New Joinee Program', 'Welcome kits for people joining the company'],
    ['employee-recognition-rewards', 'Employee Recognition & Rewards', 'Milestones, work anniversaries, and thank-yous'],
    ['new-mom-baby-kit', 'New Mom & Baby Kit', 'For a colleague starting parental leave'],
    ['sustainability', 'Sustainability', 'Eco-friendly and reusable branded merchandise'],
    ['festive-gift-kits', 'Festive Gift Kits', 'Seasonal and holiday gifting'],
    ['personal-milestone', 'Personal Milestone', 'Birthdays, promotions, and personal wins'],
    ['cxo-gifting', 'CXO Gifting', 'Premium gifting for C-suite recipients'],
    ['executive-gifting', 'Executive Gifting', 'Curated gifting for senior leadership'],
  ];
  var existing = {};
  readTab(SHEETS.EVENT_KITS).forEach(function (r) { existing[String(r.slug).trim().toLowerCase()] = 1; });

  var n = 0;
  defaults.forEach(function (d, i) {
    if (existing[d[0]]) return;
    appendRow(SHEETS.EVENT_KITS, {
      slug: d[0], label: d[1], tagline: d[2], hero_image: '',
      product_skus: '', sort_order: i, active: 'TRUE'
    });
    n++;
  });
  console.log('Seeded ' + n + ' event kit' + (n === 1 ? '' : 's') + ' (skipped ' + (defaults.length - n) + ' already present).');
}
