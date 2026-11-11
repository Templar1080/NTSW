# Smart ID Finder (NTSW) — v6.1

Bulk lookup of official Iranian product/service ID codes (شناسه کالا و خدمات)
for e-invoicing under سامانه مودیان (the Taxpayers System).

## What it does

Takes an Excel list of products (columns: Store, Barcode, InnerShenasehCode,
ItemCode, LongDescription) and finds each product's official ID code by
querying the public thdorsan.com lookup. Persian text normalisation and brand
extraction from ~200 built-in Iranian brands improve matching.

Confidence rules:

- **≥ 78%** — auto-accepted
- **60–78%** — held for manual review (the code is not inserted until a person confirms)
- **< 60%** — rejected

The top-5 candidates can be picked per row, and goods that are not registrable
(fresh produce, mobile phones, stationery, …) are marked غیرمشمول.
Output is a colour-coded Excel file (green / yellow / red). Result caching,
pause / resume / stop are supported.

## Run

```bash
pip install -r requirements.txt
python Smart_ID_Finder_v6.1.py
```

## Build the Windows .exe

```bat
build_exe.bat
```

See `Smart_ID_Finder_v6.spec` and `راهنمای_ساخت_exe.txt` (Persian build guide).
The built single-file executable is produced in `dist/` and is not kept in
this repository — build it locally.

## Notes

- The lookup uses thdorsan.com's undocumented `/Stuff` endpoint, so it can
  break if that site changes. There is no official API or fallback.
- VAT rates returned by the lookup are not written to the output.
- Stack: Python 3.10/3.12, PyQt5 (Persian GUI), openpyxl,
  requests + BeautifulSoup4.
