/**
 * Maintenance.gs — one-shot data repairs, run by hand from the Apps Script
 * editor. Nothing here is wired into the web app. Each job is idempotent and
 * has a dry run, so you can read the plan before anything is written:
 *
 *   showProductHeaders()          <- what columns this spreadsheet actually has
 *   showSourceHeaders()           <- what columns the Event Based Catalog has
 *   applySellingPrice_dryRun()    <- pull column I "Selling Price" from the catalogue
 *   applySellingPrice()           <- then apply
 *   blankUnpricedProducts()       <- clear stale prices where there is no Selling Price
 *   exportCatalogueToDrive()      <- publish without a GitHub token
 *   mergeStyleVariants_dryRun()   <- look first
 *   mergeStyleVariants()          <- then apply
 *   runImageMapOnce()             <- bulk image fixes (edit the map inside it)
 *
 * ---------------------------------------------------------------------------
 * NEVER put a bare call like `applyImageMap({...});` at the top level of this
 * file. Apps Script evaluates every top-level statement on EVERY execution of
 * the project — including every web app request from the storefront — so a
 * stray call there would reopen the spreadsheet and rewrite those cells on
 * every single API call. Put the arguments inside a zero-argument function
 * (see runImageMapOnce below) and pick THAT from the Run dropdown.
 * ---------------------------------------------------------------------------
 *
 * WHY mergeStyleVariants EXISTS
 * The catalogue export gives every SIZE of a garment its own row. The importer
 * only collapsed underscore-delimited size suffixes (`_XL`), so rows whose size
 * code is hyphen-delimited (CSUN-0002-XL), glued onto the style code
 * (760038AXL) or outside its size list (RARE-071_4XL) leaked onto the
 * storefront as separate products — the same t-shirt listed five times.
 *
 * Rows count as one product only when NAME, IMAGE and DESCRIPTION are all
 * identical: hard evidence, rather than a guess at SKU shapes. The surviving
 * parent SKU is the group's longest common prefix (the style code); every
 * original per-size SKU stays on as a Variants row, so what a buyer actually
 * orders never changes.
 */

function mergeStyleVariants_dryRun() { return mergeStyleVariants_(true); }
function mergeStyleVariants()        { return mergeStyleVariants_(false); }


function mergeStyleVariants_(dryRun) {
  var products = readTab(SHEETS.PRODUCTS);
  var groups = {}, order = [];

  products.forEach(function (p) {
    var key = [norm_(p.name), String(p.image || '').trim(), norm_(p.description)].join('|~|');
    if (!groups[key]) { groups[key] = []; order.push(key); }
    groups[key].push(p);
  });

  var remap = {}, doomedRows = [], report = [];
  order.forEach(function (key) {
    var g = groups[key];
    if (g.length < 2) return;
    var skus = g.map(function (p) { return String(p.sku); });
    var parent = commonPrefix_(skus).replace(/[-_ ]+$/, '');
    if (parent.length < 4) {
      report.push('SKIP (no shared style code): ' + skus.join(', '));
      return;
    }
    skus.forEach(function (s) { remap[s] = parent; });
    g.slice(1).forEach(function (p) { doomedRows.push(p._row); });
    report.push(parent + '  <- ' + skus.length + ' rows [' + skus.join(', ') + ']  ' + g[0].name);
  });

  if (!report.length) {
    Logger.log('Nothing to merge — catalogue is already clean.');
    return 'clean';
  }
  Logger.log((dryRun ? 'DRY RUN — nothing written\n' : 'APPLYING\n') + report.join('\n'));
  Logger.log('parent rows to remove: ' + doomedRows.length);
  if (dryRun) return report.join('\n');

  // --- Products: rename survivors, then delete the extras bottom-up --------
  var psh = sheet(SHEETS.PRODUCTS);
  var pHead = psh.getRange(1, 1, 1, psh.getLastColumn()).getValues()[0];
  var skuCol = colOf_(pHead, 'sku', 'Products');
  var sizesCol = colOf_(pHead, 'has_sizes', 'Products');
  products.forEach(function (p) {
    var to = remap[String(p.sku)];
    if (to && doomedRows.indexOf(p._row) === -1 && String(p.sku) !== to) {
      psh.getRange(p._row, skuCol).setValue(to);
      if (sizesCol > 0) psh.getRange(p._row, sizesCol).setValue(true);
    }
  });
  doomedRows.sort(function (a, b) { return b - a; })
            .forEach(function (r) { psh.deleteRow(r); });

  // --- Variants / PriceTiers: repoint parent_sku ---------------------------
  [SHEETS.VARIANTS, SHEETS.TIERS].forEach(function (tab) {
    var sh = sheet(tab);
    var head = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
    var col = head.indexOf('parent_sku') + 1;
    if (col < 1) return;
    var n = sh.getLastRow() - 1;
    if (n < 1) return;
    var rng = sh.getRange(2, col, n, 1);
    var vals = rng.getValues();
    var touched = 0;
    for (var i = 0; i < vals.length; i++) {
      var to = remap[String(vals[i][0])];
      if (to && String(vals[i][0]) !== to) { vals[i][0] = to; touched++; }
    }
    rng.setValues(vals);
    Logger.log(tab + ': repointed ' + touched + ' rows');
  });

  // --- EventKits: repoint + dedupe the SKU lists ---------------------------
  var ksh = sheet(SHEETS.EVENT_KITS);
  var kHead = ksh.getRange(1, 1, 1, ksh.getLastColumn()).getValues()[0];
  var kCol = kHead.indexOf('product_skus') + 1;
  if (kCol > 0 && ksh.getLastRow() > 1) {
    var kRng = ksh.getRange(2, kCol, ksh.getLastRow() - 1, 1);
    var kVals = kRng.getValues();
    for (var j = 0; j < kVals.length; j++) {
      var seen = {}, out = [];
      String(kVals[j][0] || '').split(',').forEach(function (s) {
        s = s.trim(); if (!s) return;
        s = remap[s] || s;
        if (seen[s]) return;
        seen[s] = 1; out.push(s);
      });
      kVals[j][0] = out.join(',');
    }
    kRng.setValues(kVals);
  }

  Logger.log('Done. Re-run the publish step to regenerate assets/products.json.');
  return report.join('\n');
}


/**
 * applyImageMap(map) — point products at new image URLs in bulk.
 *
 * Keys may be a SKU, or the exact product name (useful for the rows whose
 * "SKU" is really a product name — "Smoor Cookies", "Bib  Gucci"):
 *
 *   applyImageMap({
 *     'CSUN-3494'    : 'https://.../tesla-classic-ball-pen.jpg',
 *     'Smoor Cookies': 'https://.../smoor-cookies.jpg'
 *   });
 *
 * Every key it cannot place is reported rather than skipped silently.
 */
function applyImageMap(map) {
  if (!map || !Object.keys(map).length) {
    throw new Error('applyImageMap needs a { sku: url } object.');
  }
  var psh = sheet(SHEETS.PRODUCTS);
  var head = psh.getRange(1, 1, 1, psh.getLastColumn()).getValues()[0];
  var imgCol = colOf_(head, 'image', 'Products');

  var bySku = {}, byName = {};
  readTab(SHEETS.PRODUCTS).forEach(function (r) {
    bySku[String(r.sku).trim()] = r;
    byName[norm_(r.name)] = r;
  });

  var done = [], missed = [];
  Object.keys(map).forEach(function (k) {
    var r = bySku[String(k).trim()] || byName[norm_(k)];
    if (!r) { missed.push(k); return; }
    psh.getRange(r._row, imgCol).setValue(map[k]);
    done.push(r.sku + '  <- ' + map[k]);
  });

  Logger.log('updated ' + done.length + ' image(s)\n' + done.join('\n'));
  if (missed.length) Logger.log('\nNO MATCHING PRODUCT for:\n  ' + missed.join('\n  '));
  Logger.log('\nRe-run the publish step to regenerate assets/products.json.');
  return { updated: done.length, missed: missed };
}


/**
 * Edit the map below, pick runImageMapOnce from the Run dropdown, press Run.
 * Re-running is harmless — it just rewrites the same cells.
 *
 * Keys may be a SKU, or the exact product name for rows whose "SKU" is really
 * a name (e.g. the wipes, whose sku column literally contains "NA").
 *
 * Applied 2026-09-04 — leave as a record, or replace with the next batch.
 */
function runImageMapOnce() {
  return applyImageMap({
    'Himalayan Intimate Wipes': 'https://lh3.googleusercontent.com/d/1vCCuDUNdIUIyMsnbqiI92ba3sCZGrrMp=w1200',
    'CSUN-3494':                'https://lh3.googleusercontent.com/d/1G8k7vluU6aEljnp8ZpzEi-DzxF9Qz0RF=w1200',
    'CSUN-3652':                'https://lh3.googleusercontent.com/d/1Kswy5-NZU879cXF-5WnS3hRkmOXEKA1T=w1200',
    'URBAN-210':                'https://lh3.googleusercontent.com/d/1Nqx2YomGtrXfEzZcaZITJqSatJ9YVIXK=w1200'
  });
}


/**
 * Find a column by header name, tolerantly.
 *
 * readTab() trims header cells but a direct getRange() read does not, so a
 * header typed as "list_price " or "List_Price" made a plain indexOf() miss and
 * the job died claiming the column did not exist. Match on trimmed+lowercased
 * text, ignore spaces vs underscores, and when it really is absent say what IS
 * there instead of just asserting a negative.
 */
function colOf_(head, name, tabName) {
  var want = String(name).trim().toLowerCase().replace(/[\s_]+/g, '');
  for (var i = 0; i < head.length; i++) {
    if (String(head[i]).trim().toLowerCase().replace(/[\s_]+/g, '') === want) return i + 1;
  }
  throw new Error(
    'The ' + (tabName || 'sheet') + ' tab has no "' + name + '" column.\n' +
    'Headers actually present: ' + head.map(function (h) { return JSON.stringify(String(h)); }).join(', '));
}

/** Print the Products header row exactly as stored, quotes and all. */
function showProductHeaders() {
  var sh = sheet(SHEETS.PRODUCTS);
  var head = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  var out = head.map(function (h, i) {
    return '  ' + String.fromCharCode(65 + i) + '  ' + JSON.stringify(String(h));
  });
  Logger.log('Products headers (' + head.length + ' columns):\n' + out.join('\n'));
  return head;
}

/* ---------------------------------------------------------------- helpers */

function norm_(s) {
  return String(s == null ? '' : s).replace(/\s+/g, ' ').trim().toLowerCase();
}

function commonPrefix_(arr) {
  if (!arr.length) return '';
  var p = arr[0];
  for (var i = 1; i < arr.length; i++) {
    var j = 0;
    while (j < p.length && j < arr[i].length && p.charAt(j) === arr[i].charAt(j)) j++;
    p = p.substring(0, j);
    if (!p) break;
  }
  return p;
}


/**
 * applySellingPrice — make column I ("Selling Price") of the Event Based
 * Catalog the price the store shows and charges.
 *
 *   applySellingPrice_dryRun()   <- read the plan, writes nothing
 *   applySellingPrice()          <- then apply
 *
 * WHY IT READS THE SOURCE SHEET DIRECTLY
 * The catalogue is 2,528 rows and lives outside this spreadsheet. Exporting it
 * to move the numbers by hand invites transcription errors on money, so this
 * opens it by id and reads the column itself. It also means the job can be
 * re-run whenever Selling Price changes.
 *
 * WHY IT WRITES THE SHEET RATHER THAN THE STOREFRONT
 * priceOrder() in Orders.gs re-prices every submitted order from this
 * spreadsheet and deliberately does not trust the browser. Changing only what
 * the site displays would show a customer one figure and invoice another.
 *
 * WHAT IT TOUCHES
 * Products.base_price, and the product's PriceTiers row. Nothing else. A tier
 * whose unit_price differs from base_price is a genuine negotiated slab and is
 * left alone and reported.
 */

/** The Event Based Catalog. Override with Script Property SOURCE_SHEET_ID. */
var SOURCE_SHEET_ID_DEFAULT = '1miHFd3YxqCzOdgSbXBajq9PjQh080O741GoacMpmqJw';
/* The tab, not the workbook — the user's link ends #gid=242498202. Reading
   getSheets()[0] would silently price off whichever tab happens to sit first.
   Override with Script Property SOURCE_SHEET_GID. */
var SOURCE_SHEET_GID_DEFAULT = 242498202;
var SELLING_PRICE_HEADER = 'Selling Price';

function applySellingPrice_dryRun() { return applySellingPrice_(true); }
function applySellingPrice()        { return applySellingPrice_(false); }


/** Read the source catalogue into { skuOrName -> selling price }. */
function readSellingPrices_() {
  var id = prop('SOURCE_SHEET_ID', SOURCE_SHEET_ID_DEFAULT);
  var ss;
  try {
    ss = SpreadsheetApp.openById(id);
  } catch (e) {
    throw new Error('Cannot open the source catalogue (' + id + '). Make sure the ' +
                    'account running this script can at least view it.\n' + e.message);
  }
  var gid = Number(prop('SOURCE_SHEET_GID', String(SOURCE_SHEET_GID_DEFAULT)));
  var sh = null, names = [];
  ss.getSheets().forEach(function (s) {
    names.push(s.getName() + ' (gid ' + s.getSheetId() + ')');
    if (s.getSheetId() === gid) sh = s;
  });
  if (!sh) {
    throw new Error('No tab with gid ' + gid + ' in "' + ss.getName() + '".\n' +
                    'Tabs present: ' + names.join(', '));
  }
  Logger.log('reading tab "' + sh.getName() + '" (gid ' + sh.getSheetId() + ') of "' + ss.getName() + '"');
  var values = sh.getDataRange().getValues();
  if (values.length < 2) throw new Error('Source sheet "' + sh.getName() + '" is empty.');

  var head = values[0];
  var priceCol = colOf_(head, SELLING_PRICE_HEADER, 'source catalogue') - 1;

  /* SKU lives under one of these, in preference order — the export has used
     both spellings. Name is the last resort, for rows whose SKU cell holds a
     product name instead of a code. */
  var skuCol = -1;
  ['Final SKU Code', 'SKU'].forEach(function (n) {
    if (skuCol >= 0) return;
    try { skuCol = colOf_(head, n, 'source catalogue') - 1; } catch (e) {}
  });
  if (skuCol < 0) {
    throw new Error('No "Final SKU Code" or "SKU" column in the source catalogue.\n' +
                    'Headers: ' + head.join(', '));
  }
  var nameCol = -1;
  ['CS Catalog Final Product Name', 'Name'].forEach(function (n) {
    if (nameCol >= 0) return;
    try { nameCol = colOf_(head, n, 'source catalogue') - 1; } catch (e) {}
  });

  var bySku = {}, byName = {}, conflicts = [], seenConflict = {}, blanks = 0, junkSku = 0;
  for (var i = 1; i < values.length; i++) {
    var raw = String(values[i][skuCol] || '').trim();
    if (!raw) continue;
    var price = Number(String(values[i][priceCol]).replace(/[^0-9.\-]/g, ''));
    if (!isFinite(price) || price <= 0) { blanks++; continue; }

    /* Every SIZE of a garment is its own row here, all at the same price, so
       record the size-stripped style code as well as the literal SKU. */
    if (!isJunkSku_(raw)) {
      var keys = raw === styleKey_(raw) ? [raw] : [raw, styleKey_(raw)];
      keys.forEach(function (k) {
        if (bySku[k] === undefined) bySku[k] = price;
        else if (bySku[k] !== price && seenConflict[k] === undefined) {
          seenConflict[k] = 1;
          conflicts.push(k + ': keeping ' + bySku[k] + ', also saw ' + price);
        }
      });
    } else { junkSku++; }
    if (nameCol >= 0) {
      var nm = norm_(values[i][nameCol]);
      if (nm && byName[nm] === undefined) byName[nm] = price;
    }
  }
  Logger.log('source catalogue: ' + (values.length - 1) + ' rows, ' +
             Object.keys(bySku).length + ' price keys, ' + blanks + ' row(s) with no usable price, ' +
             junkSku + ' row(s) with a placeholder SKU (matched by name only)');
  if (conflicts.length) {
    Logger.log('\nSAME SKU, DIFFERENT SELLING PRICE (first wins — check these):\n  ' +
               conflicts.slice(0, 20).join('\n  '));
  }
  return { bySku: bySku, byName: byName };
}


/**
 * Placeholder SKU cells. The catalogue has many DIFFERENT products whose SKU
 * cell literally reads "NA", so keying prices by that string collided them all
 * — the dry run priced Himalayan Intimate Wipes at 49, which belonged to some
 * unrelated "NA" row. Rows like this are never a usable price key; they can
 * still be matched by product name.
 */
function isJunkSku_(s) {
  return /^(na|n\/a|nil|tbd|none|null|-{1,3}|\.)?$/i.test(String(s == null ? '' : s).trim());
}


/** Strip a trailing size code so every size row folds onto its style code. */
function styleKey_(sku) {
  var s = String(sku).trim();
  s = s.replace(/[-_ ](XXXL|XXL|[2-6]XL|XS|XL|S|M|L)$/i, '');
  s = s.replace(/([A-Za-z0-9]{4,})(SM|MD|LG|XL|[2-6]X)$/, '$1');
  return s;
}


function applySellingPrice_(dryRun) {
  var src = readSellingPrices_();
  var products = readTab(SHEETS.PRODUCTS);
  var psh = sheet(SHEETS.PRODUCTS);
  var pHead = psh.getRange(1, 1, 1, psh.getLastColumn()).getValues()[0];
  var baseCol = colOf_(pHead, 'base_price', 'Products');

  var plan = [], same = 0, unmatched = [], oldBase = {};
  products.forEach(function (p) {
    var sku = String(p.sku).trim();
    var price;
    if (!isJunkSku_(sku)) {
      price = src.bySku[sku];
      if (price === undefined) price = src.bySku[styleKey_(sku)];
    }
    /* Name is the only safe key for a product whose SKU column holds a
       placeholder or a product name. */
    if (price === undefined) price = src.byName[norm_(p.name)];
    if (price === undefined) { unmatched.push(sku + '  (' + p.name + ')'); return; }

    oldBase[sku] = Number(p.base_price);
    if (Number(p.base_price) === price) { same++; return; }
    plan.push({ row: p._row, sku: sku, from: Number(p.base_price), to: price, name: p.name });
  });

  Logger.log('\n' + (dryRun ? 'DRY RUN — nothing written' : 'APPLYING') +
             '\n  ' + plan.length + ' to change, ' + same + ' already correct, ' +
             unmatched.length + ' not found in the catalogue');
  plan.slice(0, 50).forEach(function (c) {
    Logger.log('  ' + c.sku + '  ' + c.from + ' -> ' + c.to + '   ' + String(c.name).slice(0, 44));
  });
  if (plan.length > 50) Logger.log('  ... and ' + (plan.length - 50) + ' more');
  if (unmatched.length) {
    Logger.log('\nNO SELLING PRICE FOUND — these keep their current price:\n  ' + unmatched.join('\n  '));
  }
  if (dryRun) return plan.length + ' would change, ' + unmatched.length + ' unmatched';
  if (!plan.length) { Logger.log('\nNothing to do.'); return 'clean'; }

  // --- Products.base_price -------------------------------------------------
  plan.forEach(function (c) { psh.getRange(c.row, baseCol).setValue(c.to); });

  // --- PriceTiers: move the duplicate tier, leave real slabs alone ---------
  var newBySku = {};
  plan.forEach(function (c) { newBySku[c.sku] = c.to; });

  var tsh = sheet(SHEETS.TIERS);
  var tHead = tsh.getRange(1, 1, 1, tsh.getLastColumn()).getValues()[0];
  var tSkuCol = colOf_(tHead, 'parent_sku', 'PriceTiers');
  var tUnitCol = colOf_(tHead, 'unit_price', 'PriceTiers');
  var moved = 0, left = [];
  if (tsh.getLastRow() > 1) {
    var n = tsh.getLastRow() - 1;
    var skus = tsh.getRange(2, tSkuCol, n, 1).getValues();
    var rng = tsh.getRange(2, tUnitCol, n, 1);
    var units = rng.getValues();
    for (var i = 0; i < n; i++) {
      var sku = String(skus[i][0]).trim();
      var to = newBySku[sku];
      if (to === undefined) continue;
      if (Number(units[i][0]) === oldBase[sku]) { units[i][0] = to; moved++; }
      else if (Number(units[i][0]) !== to) { left.push(sku + ' @ ' + units[i][0]); }
    }
    rng.setValues(units);
  }
  Logger.log('\nPriceTiers: ' + moved + ' tier(s) moved to the selling price');
  if (left.length) Logger.log('LEFT ALONE (genuine volume slabs):\n  ' + left.join('\n  '));

  Logger.log('\nDone. Re-run the publish step to regenerate assets/products.json.');
  return plan.length + ' repriced, ' + unmatched.length + ' unmatched';
}


/** Print the source catalogue tab's header row, with column letters. */
function showSourceHeaders() {
  var id = prop('SOURCE_SHEET_ID', SOURCE_SHEET_ID_DEFAULT);
  var gid = Number(prop('SOURCE_SHEET_GID', String(SOURCE_SHEET_GID_DEFAULT)));
  var ss = SpreadsheetApp.openById(id);
  var sh = null, names = [];
  ss.getSheets().forEach(function (s) {
    names.push(s.getName() + ' (gid ' + s.getSheetId() + ')');
    if (s.getSheetId() === gid) sh = s;
  });
  if (!sh) { Logger.log('No tab with gid ' + gid + '. Tabs: ' + names.join(', ')); return names; }
  var head = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  var out = head.map(function (h, i) {
    var letter = i < 26 ? String.fromCharCode(65 + i)
                        : String.fromCharCode(64 + Math.floor(i / 26)) + String.fromCharCode(65 + (i % 26));
    return '  ' + letter + '  ' + JSON.stringify(String(h));
  });
  Logger.log('"' + sh.getName() + '" — ' + head.length + ' columns, ' +
             (sh.getLastRow() - 1) + ' data rows:\n' + out.join('\n'));
  return head;
}


/**
 * blankUnpricedProducts — clear the price of anything the catalogue has no
 * Selling Price for, so the storefront shows a blank rather than a stale
 * figure left over from the old pricing.
 *
 *   blankUnpricedProducts_dryRun()
 *   blankUnpricedProducts()
 *
 * Sets base_price to 0 and the product's PriceTiers rows to 0. The storefront
 * treats "no usable price" as price-on-request: no price text, no volume
 * table, no quantity boxes, no Add to cart. priceOrder() in Orders.gs refuses
 * to book such a line unless a negotiated price is supplied, so a zero here
 * cannot become a free order.
 *
 * Run AFTER applySellingPrice(). Idempotent.
 */
function blankUnpricedProducts_dryRun() { return blankUnpriced_(true); }
function blankUnpricedProducts()        { return blankUnpriced_(false); }


function blankUnpriced_(dryRun) {
  var src = readSellingPrices_();
  var products = readTab(SHEETS.PRODUCTS);
  var psh = sheet(SHEETS.PRODUCTS);
  var pHead = psh.getRange(1, 1, 1, psh.getLastColumn()).getValues()[0];
  var baseCol = colOf_(pHead, 'base_price', 'Products');

  var doomed = [];
  products.forEach(function (p) {
    var sku = String(p.sku).trim();
    var price;
    if (!isJunkSku_(sku)) {
      price = src.bySku[sku];
      if (price === undefined) price = src.bySku[styleKey_(sku)];
    }
    if (price === undefined) price = src.byName[norm_(p.name)];
    if (price !== undefined) return;                 // has a selling price
    if (!(Number(p.base_price) > 0)) return;         // already blank
    doomed.push({ row: p._row, sku: sku, was: Number(p.base_price), name: p.name });
  });

  Logger.log('\n' + (dryRun ? 'DRY RUN — nothing written' : 'APPLYING') +
             '\n  ' + doomed.length + ' product(s) to blank');
  doomed.forEach(function (d) {
    Logger.log('  ' + d.sku + '  was ' + d.was + '  ->  (blank)   ' + String(d.name).slice(0, 44));
  });
  if (dryRun) return doomed.length + ' would be blanked';
  if (!doomed.length) { Logger.log('\nNothing to do.'); return 'clean'; }

  var blankSkus = {};
  doomed.forEach(function (d) {
    psh.getRange(d.row, baseCol).setValue(0);
    blankSkus[d.sku] = 1;
  });

  var tsh = sheet(SHEETS.TIERS);
  var tHead = tsh.getRange(1, 1, 1, tsh.getLastColumn()).getValues()[0];
  var tSkuCol = colOf_(tHead, 'parent_sku', 'PriceTiers');
  var tUnitCol = colOf_(tHead, 'unit_price', 'PriceTiers');
  var cleared = 0;
  if (tsh.getLastRow() > 1) {
    var n = tsh.getLastRow() - 1;
    var skus = tsh.getRange(2, tSkuCol, n, 1).getValues();
    var rng = tsh.getRange(2, tUnitCol, n, 1);
    var units = rng.getValues();
    for (var i = 0; i < n; i++) {
      if (blankSkus[String(skus[i][0]).trim()] && Number(units[i][0]) !== 0) {
        units[i][0] = 0; cleared++;
      }
    }
    rng.setValues(units);
  }
  Logger.log('\nPriceTiers: ' + cleared + ' row(s) cleared');
  Logger.log('Done. Re-run the publish step to regenerate assets/products.json.');
  return doomed.length + ' blanked';
}


/**
 * exportCatalogueToDrive — write the published products.json and site.json
 * into the evidence folder, so they can be picked up without a GitHub token.
 *
 * fnAdminPublish() commits straight to the repo, but only once a fine-grained
 * GITHUB_TOKEN is set in Script Properties. This is the no-token route: run it,
 * and the two files appear in Drive ready to drop into the repo.
 *
 * Overwrites the previous export rather than piling up copies.
 */
function exportCatalogueToDrive() {
  var folderId = prop('FOLDER_ID', '');
  if (!folderId) throw new Error('Script Property FOLDER_ID is not set.');
  var folder = DriveApp.getFolderById(folderId);

  var files = [
    ['products.json', JSON.stringify(buildCatalogueJson(), null, 1)],
    ['site.json', JSON.stringify(buildSiteJson(), null, 1)]
  ];
  var out = [];
  files.forEach(function (pair) {
    var name = pair[0], body = pair[1];
    var existing = folder.getFilesByName(name);
    while (existing.hasNext()) existing.next().setTrashed(true);
    var f = folder.createFile(name, body, 'application/json');
    out.push(name + '  ' + body.length + ' bytes  id=' + f.getId());
  });
  Logger.log('exported to Drive folder ' + folder.getName() + ':\n  ' + out.join('\n  '));
  return out;
}
