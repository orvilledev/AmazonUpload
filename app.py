import streamlit as st
import pandas as pd
from copy import copy
from openpyxl import load_workbook
import io
import zipfile
from datetime import datetime
from pathlib import Path

st.set_page_config(
    page_title="Amazon FBA Manifest Generator",
    page_icon="📦",
    layout="centered",
)

st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&display=swap');

*, *::before, *::after { box-sizing: border-box; }

html, body, .stApp {
    font-family: 'Inter', sans-serif !important;
    background: #F0F2F6 !important;
}

/* ── Hide Streamlit chrome ── */
#MainMenu, footer, header { visibility: hidden; }
[data-testid="stToolbar"] { display: none; }

/* ── Page wrapper ── */
.block-container {
    max-width: 680px !important;
    padding: 40px 20px 60px !important;
}

/* ── Hero banner ── */
.hero {
    background: linear-gradient(135deg, #FF9900 0%, #FF6D00 100%);
    border-radius: 20px;
    padding: 36px 40px;
    margin-bottom: 28px;
    color: white;
    text-align: center;
    box-shadow: 0 8px 32px rgba(255,153,0,.35);
}
.hero .icon { font-size: 2.6rem; margin-bottom: 10px; }
.hero h1 {
    margin: 0 0 8px;
    font-size: 1.75rem;
    font-weight: 800;
    letter-spacing: -0.5px;
}
.hero p {
    margin: 0;
    font-size: 0.97rem;
    opacity: .88;
    font-weight: 400;
}

/* ── Section card ── */
.card {
    background: white;
    border-radius: 16px;
    padding: 24px 28px;
    margin-bottom: 18px;
    box-shadow: 0 2px 14px rgba(0,0,0,.06);
}
.card-label {
    font-size: 0.72rem;
    font-weight: 700;
    text-transform: uppercase;
    letter-spacing: 0.08em;
    color: #888;
    margin-bottom: 10px;
}
.card-title {
    font-size: 1rem;
    font-weight: 600;
    color: #1a1a1a;
    margin-bottom: 4px;
}

/* ── Pack group grid ── */
.pg-grid {
    display: grid;
    grid-template-columns: repeat(5, 1fr);
    gap: 10px;
    margin-top: 4px;
}
.pg-tile {
    background: #FFF8EE;
    border: 1.5px solid #FFE0A0;
    border-radius: 12px;
    padding: 12px 8px;
    text-align: center;
}
.pg-tile .pg-label {
    font-size: 0.68rem;
    font-weight: 700;
    color: #FF9900;
    text-transform: uppercase;
    letter-spacing: 0.06em;
    margin-bottom: 4px;
}
.pg-tile .pg-count {
    font-size: 1.35rem;
    font-weight: 800;
    color: #1a1a1a;
    line-height: 1;
}
.pg-tile .pg-unit {
    font-size: 0.68rem;
    color: #888;
    font-weight: 500;
    margin-top: 2px;
}

/* ── Date display pill ── */
.date-pill {
    display: inline-block;
    background: #FFF3E0;
    border: 1.5px solid #FFE0A0;
    color: #FF6D00;
    font-weight: 700;
    font-size: 1.1rem;
    border-radius: 30px;
    padding: 6px 20px;
    margin-top: 6px;
}

/* ── File name display ── */
.file-chip {
    display: inline-flex;
    align-items: center;
    gap: 8px;
    background: #F0FFF4;
    border: 1.5px solid #86EFAC;
    color: #166534;
    border-radius: 8px;
    padding: 8px 14px;
    font-size: 0.88rem;
    font-weight: 600;
    margin-top: 6px;
}
.file-chip .dot {
    width: 8px; height: 8px;
    background: #22C55E;
    border-radius: 50%;
    flex-shrink: 0;
}

/* ── Download button ── */
div[data-testid="stDownloadButton"] {
    margin-top: 4px;
}
div[data-testid="stDownloadButton"] > button {
    width: 100% !important;
    background: linear-gradient(135deg, #FF9900 0%, #FF6D00 100%) !important;
    color: white !important;
    border: none !important;
    border-radius: 12px !important;
    font-size: 1.05rem !important;
    font-weight: 700 !important;
    padding: 16px 0 !important;
    box-shadow: 0 4px 18px rgba(255,153,0,.4) !important;
    letter-spacing: 0.01em !important;
    transition: opacity .15s;
}
div[data-testid="stDownloadButton"] > button:hover {
    opacity: .92 !important;
}

/* ── Uploader overrides ── */
[data-testid="stFileUploader"] section {
    border: 2px dashed #E0E0E0 !important;
    border-radius: 12px !important;
    background: #FAFAFA !important;
    padding: 20px !important;
}
[data-testid="stFileUploader"] section:hover {
    border-color: #FF9900 !important;
    background: #FFFAF0 !important;
}

/* ── Date input ── */
[data-testid="stDateInput"] input {
    border-radius: 10px !important;
    border: 1.5px solid #E0E0E0 !important;
    font-size: 1rem !important;
    font-weight: 600 !important;
    padding: 10px 14px !important;
}
[data-testid="stDateInput"] input:focus {
    border-color: #FF9900 !important;
    box-shadow: 0 0 0 3px rgba(255,153,0,.15) !important;
}

/* ── Info box ── */
.info-box {
    background: #EFF6FF;
    border: 1.5px solid #BFDBFE;
    border-radius: 10px;
    padding: 14px 18px;
    color: #1E40AF;
    font-size: 0.9rem;
    font-weight: 500;
    text-align: center;
    margin-top: 4px;
}

/* ── Divider ── */
.divider {
    border: none;
    border-top: 1.5px solid #F0F0F0;
    margin: 20px 0;
}
</style>
""", unsafe_allow_html=True)

# ── Constants ─────────────────────────────────────────────────────────────────
TEMPLATE_PATH = Path(__file__).parent / "template.xlsx"

# ── Helpers ───────────────────────────────────────────────────────────────────

def parse_raw(file) -> pd.DataFrame:
    df = pd.read_excel(file)
    df.columns = df.columns.str.strip()
    df["Pack Group #"] = df["Pack Group #"].ffill().astype(int)
    return df


def find_template_sheet(wb) -> str:
    for name in wb.sheetnames:
        n = name.lower()
        if "template" in n and "example" not in n:
            return name
    raise ValueError("Template sheet not found.")


def find_header_row(ws) -> int:
    for row in ws.iter_rows():
        for cell in row:
            if cell.value and str(cell.value).strip() == "Merchant SKU":
                return cell.row
    return 6


def format_upc_code(value) -> str:
    text = str(value).strip()
    if "-FNSKU" in text:
        return text
    return str(int(value))


def aggregate_pg_rows(pg_rows: pd.DataFrame) -> pd.DataFrame:
    work = pg_rows.copy()
    work["UPC Code"] = work["UPC Code"].map(format_upc_code)
    work["_order"] = range(len(work))
    aggregated = (
        work.groupby("UPC Code", as_index=False)
        .agg(Total_QTY=("Total QTY", "sum"), _order=("_order", "min"))
        .rename(columns={"Total_QTY": "Total QTY"})
        .sort_values("_order")
        .drop(columns="_order")
    )
    aggregated["Total QTY"] = aggregated["Total QTY"].astype(int)
    return aggregated.reset_index(drop=True)


def build_manifest(template_bytes: bytes, pg_rows: pd.DataFrame) -> bytes:
    pg_rows = aggregate_pg_rows(pg_rows)
    wb = load_workbook(io.BytesIO(template_bytes))
    ws = wb[find_template_sheet(wb)]
    hr = find_header_row(ws)
    style_ref_sku = ws.cell(row=hr + 1, column=1)
    style_ref_qty = ws.cell(row=hr + 1, column=2)
    sku_font = copy(style_ref_sku.font)
    sku_alignment = copy(style_ref_sku.alignment)
    qty_font = copy(style_ref_qty.font)
    qty_alignment = copy(style_ref_qty.alignment)
    qty_number_format = style_ref_qty.number_format

    for r in range(hr + 1, ws.max_row + 1):
        for c in range(1, ws.max_column + 1):
            ws.cell(row=r, column=c).value = None

    for i, (_, row) in enumerate(pg_rows.iterrows()):
        r = hr + 1 + i
        sku_cell = ws.cell(row=r, column=1)
        qty_cell = ws.cell(row=r, column=2)
        sku_cell.value = row["UPC Code"]
        sku_cell.number_format = "@"
        sku_cell.font = copy(sku_font)
        sku_cell.alignment = copy(sku_alignment)
        qty_cell.value = int(row["Total QTY"])
        qty_cell.number_format = qty_number_format
        qty_cell.font = copy(qty_font)
        qty_cell.alignment = copy(qty_alignment)

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


# ── Hero ──────────────────────────────────────────────────────────────────────
st.markdown("""
<div class="hero">
    <div class="icon">📦</div>
    <h1>Amazon FBA Manifest Generator</h1>
    <p>Upload your RAW file, pick a date, and download all manifests in one click.</p>
</div>
""", unsafe_allow_html=True)

# ── Step 1 — Upload ───────────────────────────────────────────────────────────
st.markdown('<div class="card-label">Step 1 — Source Data</div>', unsafe_allow_html=True)
raw_file = st.file_uploader(
    "Upload RAW File",
    type=["xlsx", "xls"],
    label_visibility="collapsed",
)

# ── Step 2 — Date ─────────────────────────────────────────────────────────────
st.markdown("<div style='height:14px'></div>", unsafe_allow_html=True)
st.markdown('<div class="card-label">Step 2 — Shipment Date</div>', unsafe_allow_html=True)
selected_date = st.date_input("Date", value=datetime.today(), label_visibility="collapsed")
date_str = f"{selected_date.month}.{selected_date.day}.{str(selected_date.year)[-2:]}"
st.markdown(f'<div class="date-pill">📅 &nbsp;{date_str}</div>', unsafe_allow_html=True)

# ── Step 3 — Generate ─────────────────────────────────────────────────────────
st.markdown("<div style='height:20px'></div>", unsafe_allow_html=True)

if raw_file:
    try:
        raw_file.seek(0)
        df = parse_raw(raw_file)
        vendor = str(df["Vendor"].iloc[0]).strip()
        pack_groups = sorted(df["Pack Group #"].unique())
        template_bytes = TEMPLATE_PATH.read_bytes()

        # ── File chip ────────────────────────────────────────────────────────
        st.markdown(
            f'<div class="file-chip"><div class="dot"></div>{raw_file.name} &nbsp;·&nbsp; '
            f'<span style="font-weight:400;color:#166534;">{vendor} &nbsp;·&nbsp; '
            f'{len(pack_groups)} pack groups &nbsp;·&nbsp; {len(df)} SKUs</span></div>',
            unsafe_allow_html=True,
        )

        st.markdown("<div style='height:20px'></div>", unsafe_allow_html=True)

        # ── Pack group grid ──────────────────────────────────────────────────
        st.markdown('<div class="card-label">Pack Groups Detected</div>', unsafe_allow_html=True)

        tiles_html = '<div class="pg-grid">'
        for pg in pack_groups:
            n_skus = len(df[df["Pack Group #"] == pg])
            tiles_html += f"""
            <div class="pg-tile">
                <div class="pg-label">PG {pg}</div>
                <div class="pg-count">{n_skus}</div>
                <div class="pg-unit">SKUs</div>
            </div>"""
        tiles_html += "</div>"
        st.markdown(tiles_html, unsafe_allow_html=True)

        st.markdown("<div style='height:24px'></div>", unsafe_allow_html=True)

        # ── Build ZIP ────────────────────────────────────────────────────────
        st.markdown('<div class="card-label">Step 3 — Download</div>', unsafe_allow_html=True)

        zip_buf = io.BytesIO()
        with zipfile.ZipFile(zip_buf, "w", zipfile.ZIP_DEFLATED) as zf:
            for pg in pack_groups:
                pg_rows = df[df["Pack Group #"] == pg][["UPC Code", "Total QTY"]]
                zf.writestr(
                    f"{vendor} PG{pg} {date_str}.xlsx",
                    build_manifest(template_bytes, pg_rows),
                )
        zip_buf.seek(0)

        st.download_button(
            label=f"⬇️   Download All {len(pack_groups)} Manifests  ({vendor} · {date_str})",
            data=zip_buf.getvalue(),
            file_name=f"{vendor} FBA Manifests {date_str}.zip",
            mime="application/zip",
        )

        # ── File list preview ────────────────────────────────────────────────
        st.markdown("<div style='height:16px'></div>", unsafe_allow_html=True)
        st.markdown('<div class="card-label">Files Included in ZIP</div>', unsafe_allow_html=True)

        rows_html = ""
        for pg in pack_groups:
            n_skus = len(df[df["Pack Group #"] == pg])
            fname = f"{vendor} PG{pg} {date_str}.xlsx"
            rows_html += f"""
            <div style="display:flex;justify-content:space-between;align-items:center;
                        padding:10px 0;border-bottom:1px solid #F3F3F3;font-size:0.88rem;">
                <span style="font-weight:600;color:#1a1a1a;">📄 &nbsp;{fname}</span>
                <span style="color:#888;font-weight:500;">{n_skus} SKUs</span>
            </div>"""
        st.markdown(f'<div style="background:white;border-radius:12px;padding:4px 20px;box-shadow:0 2px 14px rgba(0,0,0,.06)">{rows_html}</div>', unsafe_allow_html=True)

    except Exception as exc:
        st.error(f"Error: {exc}")
        st.exception(exc)

else:
    st.markdown(
        '<div class="info-box">📂 &nbsp; Upload your RAW file above to get started.</div>',
        unsafe_allow_html=True,
    )
