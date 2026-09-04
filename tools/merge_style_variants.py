#!/usr/bin/env python3
"""Collapse size-split parent products back into one style-code parent.

The catalogue export gives every size its own row. build_catalog_cs.py's
parent_key() only stripped an underscore-delimited size suffix (`_XL`), so
rows whose size code is hyphen-delimited (`CSUN-0002-XL`), glued to the style
code (`760038AXL`) or outside its size list (`_4XL`) escaped grouping and
surfaced as separate products on the storefront.

Rather than guess at SKU shapes, this groups on hard evidence: rows are the
same product when their NAME, IMAGE and DESCRIPTION are all identical. Sizes
differ; nothing else does. The surviving parent SKU is the group's longest
common prefix (the style code — CSUN-0002, 760038A, RARE-071); each original
per-size SKU is preserved as a variant_sku, so what a buyer actually orders is
unchanged.

Rewrites in place: sheet-seed/{Products,Variants,PriceTiers,EventKits}.csv and
assets/products.json. Idempotent.
"""
import csv, json, os, re, sys
from collections import OrderedDict, defaultdict

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SEED = os.path.join(ROOT, 'sheet-seed')

SIZE_ORDER = ['XS','S','SM','M','MD','L','LG','XL','XXL','2XL','3XL','XXXL',
              '4XL','5XL','6XL','2X','3X','4X','5X','6X','OS']
def size_rank(s):
    u = (s or '').strip().upper()
    return (SIZE_ORDER.index(u) if u in SIZE_ORDER else len(SIZE_ORDER), u)

def norm(s):
    return re.sub(r'\s+', ' ', (s or '').strip().lower())

def read(name):
    with open(os.path.join(SEED, name), encoding='utf-8') as f:
        r = csv.DictReader(f)
        return list(r), r.fieldnames

def write(name, rows, cols):
    with open(os.path.join(SEED, name), 'w', encoding='utf-8', newline='') as f:
        w = csv.DictWriter(f, fieldnames=cols, lineterminator='\n')
        w.writeheader(); w.writerows(rows)

def common_prefix(skus):
    p = os.path.commonprefix(skus)
    return re.sub(r'[-_ ]+$', '', p)


def main():
    products, pcols = read('Products.csv')
    variants, vcols = read('Variants.csv')
    tiers,    tcols = read('PriceTiers.csv')
    kits,     kcols = read('EventKits.csv')

    # ---- group on (name, image, description) -----------------------------
    groups = OrderedDict()
    for p in products:
        key = (norm(p['name']), (p['image'] or '').strip(), norm(p['description']))
        groups.setdefault(key, []).append(p)

    remap = {}            # old parent sku -> surviving parent sku
    survivors = []
    merged_report = []

    for members in groups.values():
        if len(members) == 1:
            survivors.append(members[0]); continue

        skus = [m['sku'] for m in members]
        parent = common_prefix(skus)
        if len(parent) < 4:                      # no shared style code — leave alone
            survivors.extend(members)
            print(f'  SKIP (no common style code): {skus}', file=sys.stderr)
            continue

        keep = dict(members[0])
        keep['sku'] = parent
        keep['has_sizes'] = 'TRUE'
        # widest MOQ/pricing in the group is the same by construction; keep row 0's
        survivors.append(keep)
        for m in members:
            remap[m['sku']] = parent
        merged_report.append((parent, skus, keep['name']))

    # ---- variants: repoint parents, keep the real per-size SKUs -----------
    by_parent = defaultdict(list)
    for v in variants:
        v['parent_sku'] = remap.get(v['parent_sku'], v['parent_sku'])
        by_parent[v['parent_sku']].append(v)

    new_variants = []
    for parent, vs in by_parent.items():
        seen, out = set(), []
        for v in vs:
            k = (v['variant_sku'], v['size'].strip().upper())
            if k in seen: continue
            seen.add(k); out.append(v)
        out.sort(key=lambda v: size_rank(v['size']))
        new_variants.extend(out)

    # ---- price tiers: one set per surviving parent ------------------------
    seen_t, new_tiers = set(), []
    for t in tiers:
        t['parent_sku'] = remap.get(t['parent_sku'], t['parent_sku'])
        k = (t['parent_sku'], t['min_qty'], t['max_qty'])
        if k in seen_t: continue
        seen_t.add(k); new_tiers.append(t)

    # ---- event kits + related SKUs: repoint, dedupe, keep order ----------
    def fix_list(csv_str):
        out, seen = [], set()
        for s in [x.strip() for x in (csv_str or '').split(',') if x.strip()]:
            s = remap.get(s, s)
            if s in seen: continue
            seen.add(s); out.append(s)
        return ','.join(out)

    for k in kits:
        k['product_skus'] = fix_list(k['product_skus'])
    for p in survivors:
        for col in ('related_skus', 'auto_related_skus'):
            if col in p:
                p[col] = ','.join(s for s in fix_list(p[col]).split(',')
                                  if s and s != p['sku'])

    write('Products.csv',  survivors,    pcols)
    write('Variants.csv',  new_variants, vcols)
    write('PriceTiers.csv',new_tiers,    tcols)
    write('EventKits.csv', kits,         kcols)

    # ---- assets/products.json --------------------------------------------
    jpath = os.path.join(ROOT, 'assets', 'products.json')
    data = json.load(open(jpath, encoding='utf-8'))
    vsizes = defaultdict(list)
    for v in new_variants:
        s = v['size'].strip()
        if s and s.upper() != 'OS' and s not in vsizes[v['parent_sku']]:
            vsizes[v['parent_sku']].append(s)

    keep_json, seen_json = [], set()
    for p in data['products']:
        sku = remap.get(p['sku'], p['sku'])
        if sku in seen_json: continue
        seen_json.add(sku)
        p['sku'] = sku
        if sku in set(remap.values()):
            slug = re.sub(r'[^a-z0-9]+', '-', p['name'].lower()).strip('-')
            p['url_key'] = f'{slug}-{sku.lower()}'
        if sku in set(remap.values()) and vsizes.get(sku):
            p['sizes'] = sorted(vsizes[sku], key=size_rank)
            p['has_sizes'] = True
        if 'related_skus' in p:
            p['related_skus'] = [remap.get(s, s) for s in p['related_skus']]
        keep_json.append(p)
    data['products'] = keep_json
    for k in data.get('event_kits', []):
        seen_k, out = set(), []
        for s in k.get('product_skus', []):
            s = remap.get(s, s)
            if s in seen_k: continue
            seen_k.add(s); out.append(s)
        k['product_skus'] = out
        k['product_count'] = len(out)
    json.dump(data, open(jpath, 'w', encoding='utf-8'), indent=1, ensure_ascii=False)

    print(f'merged {len(merged_report)} style groups')
    for parent, skus, name in merged_report:
        print(f'  {parent:12} <- {len(skus)} rows  {skus}')
        print(f'               {name[:60]}')
    print(f'products: {len(products)} -> {len(survivors)}')
    print(f'variants: {len(variants)} -> {len(new_variants)}')
    print(f'tiers:    {len(tiers)} -> {len(new_tiers)}')

if __name__ == '__main__':
    main()
