import pandas as pd
from ti_scrape import load_excel_data
df = load_excel_data("ti_specs.xlsx", header_row=10)
pn  = next(c for c in df.columns if "product or part number" in c.lower())
pkg = next(c for c in df.columns if "package name" in c.lower())
pin = next(c for c in df.columns if "pin count" in c.lower())
for p in ["TPD4S009","TPD2E001","TPD3E001","TPD4E001","TPD4EUSB30"]:
    r = df[df[pn].astype(str).str.strip() == p]
    if not r.empty:
        print(f"{p:12} pkg={r.iloc[0][pkg]!r:35} pins={r.iloc[0][pin]!r}")