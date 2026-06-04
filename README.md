# Amazon FBA Manifest Generator

Streamlit app that reads a RAW upload Excel file and generates one Amazon FBA manifest file per pack group, bundled in a ZIP download.

## Usage

```bash
pip install -r requirements.txt
streamlit run app.py
```

1. Upload your RAW file (`.xlsx`)
2. Select the shipment date for output filenames
3. Click **Download All Manifests** to get a ZIP of files named like `CBN PG2 6.4.26.xlsx`

## Output filename format

`{Vendor} PG{pack_group} {M}.{D}.{YY}.xlsx`

Example: `CBN PG2 6.4.26.xlsx` — vendor CBN, pack group 2, June 4, 2026.
