#!/usr/bin/env python3
"""Build assets/products.json + sheet-seed CSVs from the real CompanyStore
"Event Based Catalog" export (2,528 rows, real tiered/discounted pricing —
no placeholder pricing here, unlike the original RSM build script).

Input : /home/claude/data/catalog.csv
Output: assets/products.json
        sheet-seed/Products.csv, Variants.csv, PriceTiers.csv, Categories.csv,
        sheet-seed/EventKits.csv (new — event-tag -> product mapping)
"""
import csv, json, os, re, html as htmllib
from collections import defaultdict, OrderedDict

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SRC = "/home/claude/data/catalog.csv"

EVENT_TAG_MAP = {
    "new joinee kits": "New Joinee Program",
    "new joinee kit": "New Joinee Program",
    "employee recognition": "Employee Recognition & Rewards",
    "recognition": "Employee Recognition & Rewards",
    "new mom": "New Mom & Baby Kit",
    "baby": "New Mom & Baby Kit",
    "sustainability": "Sustainability",
    "sustainable": "Sustainability",
    "festive": "Festive Gift Kits",
    "milestone": "Personal Milestone",
    "cxo": "CXO Gifting",
    "executive": "Executive Gifting",
}
EVENT_KIT_ORDER = [
    "New Joinee Program", "Employee Recognition & Rewards",
    "New Mom & Baby Kit", "Sustainability", "Festive Gift Kits",
    "Personal Milestone",
]


def clean(s):
    return htmllib.unescape((s or "").strip())


def num(s, default=0.0):
    s = (s or "").strip().replace(",", "")
    try:
        return float(s)
    except ValueError:
        return default


def slugify(s):
    s = re.sub(r"[^a-zA-Z0-9]+", "-", (s or "").strip().lower()).strip("-")
    return s or "misc"


def map_event_tags(raw):
    out = OrderedDict()
    for part in (raw or "").split(","):
        p = part.strip().lower()
        if not p:
            continue
        for key, label in EVENT_TAG_MAP.items():
            if key in p:
                out[label] = True
    return list(out.keys())


def best_image(row):
    for col in ("Thumbnail", "Image1", "Product Image", "Image2"):
        v = clean(row.get(col, ""))
        if v and v.lower().startswith("http"):
            return v
    return ""


def parent_key(row):
    fsc = clean(row.get("Final SKU Code", ""))
    sku = clean(row.get("SKU", ""))
    name = clean(row.get("CS Catalog Final Product Name") or row.get("Name", ""))
    base = fsc or sku or name
    # Strip a trailing size code so every size of a style lands on one parent.
    # Three shapes occur in the export, and missing any of them leaks each size
    # onto the storefront as its own product:
    #   CSUN-0002-XL / RARE-071_4XL  separator-delimited
    #   760038AXL / 760038A2X        glued two-char vendor code (SM MD LG XL 2X)
    # Longest alternatives come first so "2XL" is not eaten as "XL".
    SZ = r"XXXL|XXL|[2-6]XL|XS|XL|S|M|L"
    base = re.sub(rf"[-_ ]({SZ})$", "", base, flags=re.I)
    base = re.sub(r"(?<=[A-Za-z0-9]{4})(SM|MD|LG|XL|[2-6]X)$", "", base)
    return base or name


def main():
    rows = list(csv.DictReader(open(SRC, encoding="utf-8")))
    print(f"read {len(rows)} rows")

    groups = OrderedDict()
    for r in rows:
        if clean(r.get("Availability", "")).lower() != "show":
            continue
        key = parent_key(r)
        groups.setdefault(key, []).append(r)

    cats = defaultdict(set)
    products = []
    variants = []
    tiers = []
    kit_map = defaultdict(list)

    for key, group in groups.items():
        r0 = group[0]
        name = clean(r0.get("CS Catalog Final Product Name") or r0.get("Name", ""))
        if not name:
            continue

        cat_raw = clean(r0.get("Category", ""))
        cat_parts = [c.strip() for c in cat_raw.split(",") if c.strip() and c.strip().lower() != "all products"]
        top = cat_parts[0] if cat_parts else "Miscellaneous"
        sub = cat_parts[1] if len(cat_parts) > 1 else top
        cats[top].add(sub)

        moq = int(num(r0.get("Minimum Order Quantity (Calculates Per Unit Price)"), 25) or 25)
        lead = int(num(r0.get("Lead Time"), 14) or 14)
        list_price = num(r0.get("Price"))
        disc_price = num(r0.get("Discounted Price")) or list_price
        base_price = round(disc_price) if disc_price else round(list_price)
        if base_price <= 0:
            continue

        sizes = sorted({clean(r.get("Size", "")) for r in group if clean(r.get("Size", "")) and clean(r.get("Size","")).upper() != "NA"})
        colors = sorted({clean(r.get("Color", "")) for r in group if clean(r.get("Color", ""))})
        has_sizes = bool(sizes)

        for r in group:
            vsize = clean(r.get("Size", "")) or "OS"
            vsku = clean(r.get("Final SKU Code") or r.get("SKU", "")) or f"{key}-{vsize}"
            stock = int(num(r.get("Stock"), 0))
            variants.append({
                "variant_sku": vsku, "parent_sku": key,
                "size": vsize, "stock_qty": stock, "active": "TRUE",
            })

        # single real price point at MOQ (this catalog gives one negotiated
        # price per SKU, not a break ladder) — surfaced as a one-row tier so
        # the existing pricing engine still has something to key off.
        tiers.append({"parent_sku": key, "min_qty": moq, "max_qty": "", "unit_price": base_price})

        event_tags = map_event_tags(r0.get("Event Tags", ""))
        for et in event_tags:
            kit_map[et].append(key)

        desc = clean(r0.get("Description", ""))
        products.append({
            "sku": key,
            "name": name,
            "url_key": slugify(name) + "-" + slugify(key),
            "category": top,
            "subcategory": sub,
            "brand": clean(r0.get("Brand", "")),
            "description": desc,
            "specs": [],
            "moq": moq,
            "list_price": round(list_price) if list_price else base_price,
            "base_price": base_price,
            "tiers": [{"min_qty": moq, "max_qty": "", "unit_price": base_price}],
            "sizes": sizes,
            "colors": colors,
            "has_sizes": has_sizes,
            "image": best_image(r0),
            "decoration_method": clean(r0.get("Default decoration Method", "")),
            "lead_time_days": lead,
            "event_tags": event_tags,
            "active": True,
        })

    products.sort(key=lambda p: (p["category"], p["subcategory"], p["name"]))

    EVENT_KIT_META = {
        "New Joinee Program": "Welcome kits for people joining the company",
        "Employee Recognition & Rewards": "Milestones, work anniversaries, and thank-yous",
        "New Mom & Baby Kit": "For a colleague starting parental leave",
        "Sustainability": "Eco-friendly and reusable branded merchandise",
        "Festive Gift Kits": "Seasonal and holiday gifting",
        "Personal Milestone": "Birthdays, promotions, and personal wins",
        "CXO Gifting": "Premium gifting for C-suite recipients",
        "Executive Gifting": "Curated gifting for senior leadership",
    }
    all_kit_names = EVENT_KIT_ORDER + ["CXO Gifting", "Executive Gifting"]

    os.makedirs(f"{ROOT}/sheet-seed", exist_ok=True)
    with open(f"{ROOT}/assets/products.json", "w") as f:
        json.dump({
            "generated_from": "Event Based Catalog.csv (real CompanyStore pricing)",
            "pricing_status": "real - Discounted Price column used as the per-unit price at MOQ",
            # NOTE: slug is the raw category string, not a URL-slugified one.
            # The storefront's category.html and Filters.matches() compare
            # against product.category by exact string equality (this is how
            # the original RSM build worked too), so slug must equal it.
            "categories": [
                {"slug": k, "label": k, "subcategories": sorted(v)}
                for k, v in sorted(cats.items())
            ],
            "event_kits": [
                {
                    "slug": slugify(name), "label": name,
                    "tagline": EVENT_KIT_META.get(name, ""),
                    "hero_image": "",
                    "product_skus": sorted(set(kit_map.get(name, []))),
                    "product_count": len(set(kit_map.get(name, []))),
                }
                for name in all_kit_names
            ],
            "products": products,
        }, f, indent=1)

    def dump(name, rowlist, cols):
        with open(f"{ROOT}/sheet-seed/{name}.csv", "w", newline="") as f:
            w = csv.DictWriter(f, fieldnames=cols)
            w.writeheader()
            for row in rowlist:
                w.writerow({c: row.get(c, "") for c in cols})

    dump("Products", [{
        "sku": p["sku"], "name": p["name"], "category": p["category"],
        "subcategory": p["subcategory"], "description": p["description"],
        "moq": p["moq"], "base_price": p["base_price"], "list_price": p["list_price"],
        "has_sizes": "TRUE" if p["has_sizes"] else "FALSE",
        "image": p["image"], "lead_time_days": p["lead_time_days"], "active": "TRUE",
        "event_tags": ",".join(p["event_tags"]), "sort_order": i, "gst_rate": 18,
    } for i, p in enumerate(products)],
        ["sku", "name", "category", "subcategory", "description", "moq",
         "gst_rate", "base_price", "list_price", "has_sizes", "image", "lead_time_days",
         "active", "sort_order", "related_skus", "auto_related_skus", "event_tags"])

    dump("Variants", variants,
         ["variant_sku", "parent_sku", "size", "stock_qty", "active"])
    dump("PriceTiers", tiers,
         ["parent_sku", "min_qty", "max_qty", "unit_price"])
    # slug/parent_slug are raw strings, matching Products.category exactly —
    # see the note above buildCatalogueJson()'s categories block.
    dump("Categories", [
        {"slug": sub, "parent_slug": top, "label": sub,
         "sort_order": i, "active": "TRUE"}
        for i, (top, subs) in enumerate(sorted(cats.items()))
        for sub in sorted(subs)
    ], ["slug", "parent_slug", "label", "sort_order", "active"])
    dump("EventKits", [
        {
            "slug": slugify(name), "label": name,
            "tagline": EVENT_KIT_META.get(name, ""),
            "hero_image": "",
            "product_skus": ",".join(sorted(set(kit_map.get(name, [])))),
            "sort_order": i, "active": "TRUE",
        }
        for i, name in enumerate(all_kit_names)
    ], ["slug", "label", "tagline", "hero_image", "product_skus", "sort_order", "active"])

    print(f"products      {len(products)}")
    print(f"variants      {len(variants)}")
    print(f"categories    {{ {', '.join(f'{k}: {len(v)}' for k,v in cats.items())} }}")
    print(f"with image    {sum(1 for p in products if p['image'])}")
    print(f"event kits    {[ (k, len(v)) for k,v in kit_map.items() ]}")
    print(f"price range   {min(p['base_price'] for p in products)} - {max(p['base_price'] for p in products)}")


if __name__ == "__main__":
    main()
