import json
import openpyxl
import urllib.request
import os

URL = "https://www.ema.europa.eu/en/documents/other/article-57-product-data_en.xlsx"
OUTPUT = "article57.json"

print("Downloading Article 57 Excel from EMA...")
urllib.request.urlretrieve(URL, "article57.xlsx")
print(f"Downloaded: {os.path.getsize('article57.xlsx') / (1024*1024):.1f} MB")

print("Parsing Excel...")
wb = openpyxl.load_workbook("article57.xlsx", read_only=True)
ws = wb["Art57 product data"]

# Raw extraction — headers at row 20 (1-indexed), data from row 21
raw = []
for row in ws.iter_rows(min_row=21, values_only=True):
    vals = list(row)
    name = vals[0]
    if not name or not str(name).strip():
        continue
    raw.append({
        "name": str(name).strip(),
        "substance": str(vals[1] or "").strip(),
        "route": str(vals[2] or "").strip(),
        "country": str(vals[3] or "").strip(),
        "manufacturer": str(vals[4] or "").strip(),
    })

wb.close()
os.remove("article57.xlsx")
print(f"Extracted {len(raw)} records")

# Compact format: deduplicate countries, routes, and manufacturers into lookup tables
country_list = sorted(set(d["country"] for d in raw))
route_list = sorted(set(d["route"] for d in raw))
mfr_list = sorted(set(d["manufacturer"] for d in raw))

country_map = {c: i for i, c in enumerate(country_list)}
route_map = {r: i for i, r in enumerate(route_list)}
mfr_map = {m: i for i, m in enumerate(mfr_list)}

# Each record: [product_name, active_substance, route_idx, country_idx, manufacturer_idx]
compact = {
    "updated": __import__("datetime").date.today().isoformat(),
    "countries": country_list,
    "routes": route_list,
    "manufacturers": mfr_list,
    "data": [
        [d["name"], d["substance"], route_map[d["route"]], country_map[d["country"]], mfr_map[d["manufacturer"]]]
        for d in raw
    ]
}

print(f"Writing {OUTPUT}...")
with open(OUTPUT, "w", encoding="utf-8") as f:
    json.dump(compact, f, ensure_ascii=False, separators=(",", ":"))

size_mb = os.path.getsize(OUTPUT) / (1024 * 1024)
print(f"Done! {len(raw)} records, {size_mb:.1f} MB")
print(f"Unique: {len(country_list)} countries, {len(route_list)} routes, {len(mfr_list)} manufacturers")
