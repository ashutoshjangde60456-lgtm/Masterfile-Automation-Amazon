import io
import json
import re
import time
from textwrap import dedent
from pathlib import Path

import pandas as pd
import streamlit as st
from openpyxl import load_workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.worksheet.table import Table
from openpyxl.worksheet.worksheet import Worksheet

# ─────────────────────────────────────────────────────────────────────
# Page meta + theming
# ─────────────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Masterfile Automation - Amazon",
    page_icon="🧾",
    layout="wide"
)

st.markdown("""
<style>
:root{
  --bg1:#f6f9fc; --bg2:#ffffff;
  --card:#ffffff; --card-border:#e8eef6;
  --ink:#0f172a; --muted:#64748b; --accent:#2563eb;
  --badge-ok:#ecfdf5; --badge-ok-ink:#065f46;
  --badge-warn:#fff7ed; --badge-warn-ink:#9a3412;
  --badge-info:#eef2ff; --badge-info-ink:#1e40af;
}
.stApp { background: linear-gradient(180deg, var(--bg1) 0%, var(--bg2) 70%); }
.block-container { padding-top: 0.75rem; }
.section {
  border: 1px solid var(--card-border);
  background: var(--card);
  border-radius: 16px;
  padding: 18px 20px;
  box-shadow: 0 6px 24px rgba(2, 6, 23, 0.05);
  margin-bottom: 18px;
}
h1, h2, h3 { color: var(--ink); }
hr { border-color: #eef2f7; }
.badge { display:inline-block;padding:4px 10px;border-radius:999px;
  font-size:0.82rem;font-weight:600;letter-spacing:.2px;margin-right:.25rem }
.badge-info { background:var(--badge-info); color:var(--badge-info-ink); }
.badge-ok { background:var(--badge-ok); color:var(--badge-ok-ink); }
.badge-warn { background:var(--badge-warn); color:var(--badge-warn-ink); }
.small-note{ color:var(--muted); font-size:0.92rem; }
div.stButton>button, .stDownloadButton>button {
  background: var(--accent) !important; color:#fff !important;
  border-radius: 10px !important; border:0 !important;
  box-shadow: 0 8px 18px rgba(37,99,235,.18);
}
div.stButton>button:hover, .stDownloadButton>button:hover{ filter: brightness(0.95); }
.stTextArea, .stFileUploader, .stTabs { border-radius: 12px !important; }
</style>
""", unsafe_allow_html=True)

# ─────────────────────────────────────────────────────────────────────
# Constants (masterfile template layout)
# ─────────────────────────────────────────────────────────────────────
MASTER_TEMPLATE_SHEET = "Template"   # write only here
MASTER_DISPLAY_ROW    = 2            # mapping row in master (normal headers)
MASTER_SECONDARY_ROW  = 3            # ONLY used to disambiguate Key Product Features bullets
MASTER_DATA_START_ROW = 4            # first data row in master

# ─────────────────────────────────────────────────────────────────────
# Helpers
# ─────────────────────────────────────────────────────────────────────
_INVALID_XML_CHARS = re.compile(r"[\x00-\x08\x0B\x0C\x0E-\x1F\uD800-\uDFFF]")

def sanitize_xml_text(s) -> str:
    if s is None: return ""
    return _INVALID_XML_CHARS.sub("", str(s))

def norm(s: str) -> str:
    if s is None: return ""
    x = str(s).strip().lower()
    x = re.sub(r"\s*-\s*en\s*[-_ ]\s*us\s*$", "", x)
    x = x.replace("–","-").replace("—","-").replace("−","-")
    x = re.sub(r"[._/\\-]+", " ", x)
    x = re.sub(r"[^0-9a-z\s]+", " ", x)
    return re.sub(r"\s+", " ", x).strip()

def nonempty_rows(df: pd.DataFrame) -> int:
    if df.empty: return 0
    return df.replace("", pd.NA).dropna(how="all").shape[0]

def worksheet_used_cols(ws: Worksheet, header_rows=(1,), hard_cap=2048, empty_streak_stop=8):
    # Only scan header rows for speed
    max_try = min(ws.max_column, hard_cap)
    last_nonempty, streak = 0, 0
    for c in range(1, max_try + 1):
        any_val = any((ws.cell(row=r, column=c).value not in (None, "")) for r in header_rows)
        if any_val:
            last_nonempty, streak = c, 0
        else:
            streak += 1
            if streak >= empty_streak_stop:
                break
    return max(last_nonempty, 1)

def pick_best_onboarding_sheet(uploaded_file, mapping_aliases_by_master):
    uploaded_file.seek(0)
    xl = pd.ExcelFile(uploaded_file)
    best, best_score, best_info = None, -1, ""
    for sheet in xl.sheet_names:
        try:
            df = xl.parse(sheet_name=sheet, header=0, dtype=str).fillna("")
            df.columns = [str(c).strip() for c in df.columns]
        except Exception:
            continue
        header_set = {norm(c) for c in df.columns}
        matches = sum(any(norm(a) in header_set for a in aliases)
                      for aliases in mapping_aliases_by_master.values())
        rows = nonempty_rows(df)
        score = matches + (0.01 if rows > 0 else 0.0)
        if score > best_score:
            best, best_score = (df, sheet), score
            best_info = f"matched headers: {matches}, non-empty rows: {rows}"
    if best is None:
        raise ValueError("No readable onboarding sheet found.")
    return best[0], best[1], best_info

def expand_or_fix_tables_on_template(ws: Worksheet, header_row: int, start_row: int, used_cols: int, n_rows: int):
    """
    Make sure any Excel Table / AutoFilter on the Template sheet covers the data block.
    This prevents Excel 'repaired records' popups.
    """
    # A1-style right corner
    from openpyxl.utils import get_column_letter
    last_row = max(header_row, start_row + max(0, n_rows) - 1)
    last_col_letter = get_column_letter(used_cols)
    new_ref = f"A{header_row}:{last_col_letter}{last_row}"

    # AutoFilter
    if ws.auto_filter and ws.auto_filter.ref:
        ws.auto_filter.ref = new_ref
    else:
        ws.auto_filter.ref = new_ref

    # Tables
    if hasattr(ws, "tables") and ws.tables:
        # Repoint every table to the same region (common in master files)
        for tname, t in list(ws.tables.items()):
            t.ref = new_ref

def remove_calc_chain_if_present(wb):
    """
    If calcChain exists, remove it to avoid 'Excel found unreadable content' after bulk updates.
    openpyxl exposes it as wb.calculation, but the safe way is:
    """
    try:
        # openpyxl >= 3.1 keeps calc chain state on wb.calculation.calcId; clearing resets it
        if hasattr(wb, "calculation") and hasattr(wb.calculation, "calcId"):
            wb.calculation.calcId = None
    except Exception:
        pass

# ─────────────────────────────────────────────────────────────────────
# UI
# ─────────────────────────────────────────────────────────────────────
st.title("🧾 Masterfile Automation – Amazon")
st.caption("Fills only the **Template** sheet and preserves all other sheets, styles, formulas, and macros (if any).")
st.markdown(
    "<div class='section'><span class='badge badge-info'>Template-only writer</span>"
    " <span class='badge badge-ok'>Excel-safe (no repair dialogs)</span></div>",
    unsafe_allow_html=True
)

st.markdown("<div class='section'>", unsafe_allow_html=True)
c1, c2 = st.columns([1, 1])
with c1:
    masterfile_file = st.file_uploader("📄 Masterfile Template (.xlsx / .xlsm)", type=["xlsx", "xlsm"])
with c2:
    onboarding_file = st.file_uploader("🧾 Onboarding (.xlsx)", type=["xlsx"])

st.markdown("#### 🔗 Mapping JSON")
tab1, tab2 = st.tabs(["Paste JSON", "Upload JSON"])
mapping_json_text, mapping_json_file = "", None
with tab1:
    mapping_json_text = st.text_area(
        "Paste mapping JSON", height=200,
        placeholder='\n{\n  "Partner SKU": ["Seller SKU", "item_sku"]\n}\n'
    )
with tab2:
    mapping_json_file = st.file_uploader("Or upload mapping.json", type=["json"], key="mapping_file")
st.markdown("</div>", unsafe_allow_html=True)

st.divider()
go = st.button("🚀 Generate Final Masterfile", type="primary")

# ─────────────────────────────────────────────────────────────────────
# Main
# ─────────────────────────────────────────────────────────────────────
SENTINEL_LIST = object()

if go:
    st.markdown("<div class='section'>", unsafe_allow_html=True)
    st.markdown("### 📝 Log")
    log = st.empty()
    def slog(msg): log.markdown(msg)

    if not masterfile_file or not onboarding_file:
        st.error("Please upload both **Masterfile Template** and **Onboarding**.")
        st.markdown("</div>", unsafe_allow_html=True)
        st.stop()

    # original extension
    ext = (Path(masterfile_file.name).suffix or ".xlsx").lower()
    mime_map = {
        ".xlsx": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        ".xlsm": "application/vnd.ms-excel.sheet.macroEnabled.12",
    }
    out_mime = mime_map.get(ext, mime_map[".xlsx"])

    # Parse mapping JSON
    try:
        mapping_raw = json.loads(mapping_json_text) if mapping_json_text.strip() else json.load(mapping_json_file)
    except Exception as e:
        st.error(f"Mapping JSON parse error: {e}")
        st.markdown("</div>", unsafe_allow_html=True)
        st.stop()
    if not isinstance(mapping_raw, dict):
        st.error("Mapping JSON must be an object: {\"Master header\": [aliases...]}.")
        st.markdown("</div>", unsafe_allow_html=True)
        st.stop()

    # Normalize mapping { master_norm: [alias1, alias2, ..., fallback master display] }
    mapping_aliases = {}
    for k, v in mapping_raw.items():
        aliases = v[:] if isinstance(v, list) else [v]
        if k not in aliases:
            aliases.append(k)
        mapping_aliases[norm(k)] = aliases

    # Read master headers quickly
    masterfile_file.seek(0)
    master_bytes = masterfile_file.read()
    slog("⏳ Reading Template headers…")
    t0 = time.time()

    wb_ro = load_workbook(io.BytesIO(master_bytes), read_only=True, data_only=True, keep_links=True)
    if MASTER_TEMPLATE_SHEET not in wb_ro.sheetnames:
        st.error(f"Sheet **'{MASTER_TEMPLATE_SHEET}'** not found in the masterfile.")
        st.markdown("</div>", unsafe_allow_html=True)
        st.stop()
    ws_ro = wb_ro[MASTER_TEMPLATE_SHEET]
    used_cols = worksheet_used_cols(ws_ro, header_rows=(MASTER_DISPLAY_ROW, MASTER_SECONDARY_ROW), hard_cap=2048, empty_streak_stop=8)
    display_headers   = [ws_ro.cell(row=MASTER_DISPLAY_ROW,   column=c).value or "" for c in range(1, used_cols+1)]
    secondary_headers = [ws_ro.cell(row=MASTER_SECONDARY_ROW, column=c).value or "" for c in range(1, used_cols+1)]
    wb_ro.close()
    slog(f"✅ Template headers loaded (cols={used_cols}) in {time.time()-t0:.2f}s")

    # Pick best onboarding sheet
    try:
        best_df, best_sheet, info = pick_best_onboarding_sheet(onboarding_file, mapping_aliases)
    except Exception as e:
        st.error(f"Onboarding error: {e}")
        st.markdown("</div>", unsafe_allow_html=True)
        st.stop()
    on_df = best_df.fillna("")
    on_df.columns = [str(c).strip() for c in on_df.columns]
    on_headers = list(on_df.columns)
    st.success(f"Using onboarding sheet: **{best_sheet}** ({info})")

    # Build mapping: master col -> source Series (or SENTINEL_LIST)
    series_by_alias = {norm(h): on_df[h] for h in on_headers}
    report_lines, unmatched = [], []
    report_lines.append("#### 🔎 Mapping Summary (Template)")

    BULLET_DISP_N = norm("Key Product Features")
    master_to_source = {}

    for c, (disp, sec) in enumerate(zip(display_headers, secondary_headers), start=1):
        disp_norm = norm(disp); sec_norm = norm(sec)
        if disp_norm == BULLET_DISP_N and sec_norm:
            effective_header = sec          # e.g. bullet_point1
            label_for_log = f"{disp} ({sec})"
        else:
            effective_header = disp
            label_for_log = disp

        eff_norm = norm(effective_header)
        if not eff_norm: continue

        aliases = mapping_aliases.get(eff_norm, [effective_header])
        resolved = None
        for a in aliases:
            s = series_by_alias.get(norm(a))
            if s is not None:
                resolved = s
                report_lines.append(f"- ✅ **{label_for_log}** ← `{a}`")
                break
        if resolved is not None:
            master_to_source[c] = resolved
        else:
            if disp_norm == norm("Listing Action (List or Unlist)"):
                master_to_source[c] = SENTINEL_LIST
                report_lines.append(f"- 🟨 **{label_for_log}** ← (will fill `'List'`)")
            else:
                unmatched.append(label_for_log or f"Col {c}")
                report_lines.append(f"- ❌ **{label_for_log}** ← *no match*")

    st.markdown("\n".join(report_lines))

    n_rows = len(on_df)

    # Build a sanitized 2D block for fast append
    block = [[""] * used_cols for _ in range(n_rows)]
    for col, src in master_to_source.items():
        if src is SENTINEL_LIST:
            for i in range(n_rows):
                block[i][col-1] = "List"
        else:
            vals = src.astype(str).tolist()
            m = min(len(vals), n_rows)
            for i in range(m):
                v = sanitize_xml_text(vals[i].strip())
                if v and v.lower() not in ("nan", "none", ""):
                    block[i][col-1] = v

    # ─────────────────────────────────────────────────────────────
    # SAFE WRITER (openpyxl, macro-preserving, table-safe)
    # ─────────────────────────────────────────────────────────────
    slog("🚀 Writing values into Template (Excel-safe)…")
    t_write = time.time()

    wb = load_workbook(io.BytesIO(master_bytes), read_only=False, keep_vba=(ext == ".xlsm"))
    ws = wb[MASTER_TEMPLATE_SHEET]

    # Remove old data rows (keep headers)
    if ws.max_row >= MASTER_DATA_START_ROW:
        n_to_delete = ws.max_row - MASTER_DATA_START_ROW + 1
        if n_to_delete > 0:
            ws.delete_rows(MASTER_DATA_START_ROW, n_to_delete)

    # Append new rows fast
    for r in block:
        ws.append(r)

    # Update Table and AutoFilter ranges to cover header_row..last_row, all used columns
    expand_or_fix_tables_on_template(
        ws=ws,
        header_row=MASTER_DISPLAY_ROW,
        start_row=MASTER_DATA_START_ROW,
        used_cols=used_cols,
        n_rows=n_rows
    )

    # Clear calcChain to prevent 'Excel repaired...' after bulk changes
    remove_calc_chain_if_present(wb)

    # Save to bytes
    out_bio = io.BytesIO()
    wb.save(out_bio)
    wb.close()
    out_bio.seek(0)
    out_bytes = out_bio.getvalue()

    slog(f"✅ Finished in {time.time()-t_write:.2f}s")
    st.download_button(
        "⬇️ Download Final Masterfile",
        data=out_bytes,
        file_name=f"final_masterfile{ext}",
        mime=out_mime,
        key="dl_excel_safe",
    )

    st.markdown("</div>", unsafe_allow_html=True)

# ─────────────────────────────────────────────────────────────────────
# Friendly Instructions (bottom)
# ─────────────────────────────────────────────────────────────────────
with st.expander("📘 How to use (step-by-step)", expanded=False):
    st.markdown(dedent(f"""
    **What this tool does**
    - It only writes data into the **`{MASTER_TEMPLATE_SHEET}`** sheet of your Masterfile.
    - All other tabs, formulas, formatting, and macros (.xlsm) are preserved.
    - For **Key Product Features**, we read the little labels in **Row {MASTER_SECONDARY_ROW}** (e.g. `bullet_point1..5`).
      For everything else, we use the column names in **Row {MASTER_DISPLAY_ROW}``.
    - Your product rows start from **Row {MASTER_DATA_START_ROW}**.

    **How to run**
    1. Upload the **Masterfile** (.xlsx / .xlsm) and the **Onboarding** (.xlsx) files above.
    2. Paste or upload the **Mapping JSON**.
    3. Click **Generate Final Masterfile**.

    **Notes**
    - Invalid XML control characters in inputs are auto-removed (prevents Excel repair prompts).
    - Table & AutoFilter ranges are automatically synchronized to avoid any "repaired records" dialogs.
    """))

st.markdown(
    "<div class='section small-note'>This build avoids Linux XML patching entirely to eliminate Excel repair popups while keeping speed and macro preservation.</div>",
    unsafe_allow_html=True
)
