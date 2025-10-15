import io
import json
import re
from pathlib import Path

import pandas as pd
import streamlit as st
import xlsxwriter
from textwrap import dedent

# ---------------- UI & Styling ----------------
st.set_page_config(page_title="Masterfile Automation – Fast (XlsxWriter)", page_icon="🧾", layout="wide")
st.markdown("""
<style>
.section{border:1px solid #e8eef6;background:#fff;border-radius:16px;padding:18px;margin:12px 0;box-shadow:0 6px 24px rgba(2,6,23,.05)}
div.stButton>button,.stDownloadButton>button{background:#2563eb!important;color:#fff!important;border-radius:10px!important;border:0!important}
</style>
""", unsafe_allow_html=True)

MASTER_TEMPLATE_SHEET = "Template"
MASTER_DISPLAY_ROW    = 2
MASTER_SECONDARY_ROW  = 3
MASTER_DATA_START_ROW = 4

# ---------------- Helpers ----------------
_INVALID_XML_CHARS = re.compile(r"[\x00-\x08\x0B\x0C\x0E-\x1F\uD800-\uDFFF]")
def clean(s):
    if s is None: return ""
    return _INVALID_XML_CHARS.sub("", str(s))

def norm(s: str) -> str:
    if s is None: return ""
    x = str(s).strip().lower()
    x = re.sub(r"[._/\\-]+"," ",x)
    return re.sub(r"[^0-9a-z\s]+"," ",x).strip()

def nonempty_rows(df: pd.DataFrame) -> int:
    if df.empty: return 0
    return df.replace("", pd.NA).dropna(how="all").shape[0]

def pick_best_onboarding_sheet(uploaded_file, mapping_aliases_by_master):
    uploaded_file.seek(0)
    xl = pd.ExcelFile(uploaded_file, engine="openpyxl")
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

# ---------------- UI ----------------
st.title("🧾 Masterfile Automation – FAST mode")
st.caption("No Node, no COM, no XML patch. Writes with XlsxWriter only.")

with st.container():
    c1, c2 = st.columns([1,1])
    with c1:
        # The template is read ONLY to fetch headers; we don’t write back to it.
        masterfile_file = st.file_uploader("📄 Masterfile Template (read headers only)", type=["xlsx","xlsm"])
    with c2:
        onboarding_file = st.file_uploader("🧾 Onboarding (.xlsx)", type=["xlsx"])

st.markdown("#### 🔗 Mapping JSON")
tab1, tab2 = st.tabs(["Paste JSON", "Upload JSON"])
mapping_json_text, mapping_json_file = "", None
with tab1:
    mapping_json_text = st.text_area("Paste mapping JSON", height=200,
                                     placeholder='{\n  "Partner SKU": ["Seller SKU","item_sku"]\n}')
with tab2:
    mapping_json_file = st.file_uploader("Or upload mapping.json", type=["json"], key="mapping_file")

go = st.button("🚀 Generate Final Masterfile", type="primary")

# ---------------- Main ----------------
if go:
    st.markdown("<div class='section'>", unsafe_allow_html=True)
    log = st.empty(); slog = lambda m: log.markdown(m)

    if not masterfile_file or not onboarding_file:
        st.error("Please upload both files."); st.stop()

    # Parse mapping
    try:
        mapping_raw = json.loads(mapping_json_text) if mapping_json_text.strip() else json.load(mapping_json_file)
    except Exception as e:
        st.error(f"Mapping JSON parse error: {e}"); st.stop()
    mapping_aliases = {}
    for k, v in mapping_raw.items():
        aliases = v[:] if isinstance(v, list) else [v]
        if k not in aliases: aliases.append(k)
        mapping_aliases[norm(k)] = aliases

    # Read headers from template (read-only, openpyxl engine via pandas/openpyxl)
    import openpyxl
    wb_ro = openpyxl.load_workbook(masterfile_file, read_only=True, data_only=True)
    if MASTER_TEMPLATE_SHEET not in wb_ro.sheetnames:
        st.error(f"Sheet '{MASTER_TEMPLATE_SHEET}' not found in template."); st.stop()
    ws_ro = wb_ro[MASTER_TEMPLATE_SHEET]
    used_cols = ws_ro.max_column
    display_headers   = [ws_ro.cell(row=MASTER_DISPLAY_ROW,   column=c).value or "" for c in range(1, used_cols+1)]
    secondary_headers = [ws_ro.cell(row=MASTER_SECONDARY_ROW, column=c).value or "" for c in range(1, used_cols+1)]
    wb_ro.close()
    slog(f"✅ Template headers loaded ({used_cols} columns)")

    # Pick best onboarding sheet
    try:
        best_df, best_sheet, info = pick_best_onboarding_sheet(onboarding_file, mapping_aliases)
    except Exception as e:
        st.error(f"Onboarding error: {e}"); st.stop()
    on_df = best_df.fillna("")
    on_headers = list(on_df.columns)
    st.success(f"Using onboarding sheet: **{best_sheet}** ({info})")

    # Build mapping master->source
    from difflib import SequenceMatcher
    def top_matches(query, candidates, k=3):
        q = norm(query)
        scored = [(SequenceMatcher(None, q, norm(c)).ratio(), c) for c in candidates]
        scored.sort(key=lambda t: t[0], reverse=True)
        return scored[:k]

    SENTINEL_LIST = object()
    series_by_alias = {norm(h): on_df[h] for h in on_headers}
    master_to_source, report_lines = {}, []
    BULLET_DISP_N = norm("Key Product Features")

    for c, (disp, sec) in enumerate(zip(display_headers, secondary_headers), start=1):
        disp_norm, sec_norm = norm(disp), norm(sec)
        if disp_norm == BULLET_DISP_N and sec_norm:
            effective, label = sec, f"{disp} ({sec})"
        else:
            effective, label = disp, disp
        eff_norm = norm(effective)
        if not eff_norm:
            continue
        aliases = mapping_aliases.get(eff_norm, [effective])
        resolved = None
        for a in aliases:
            s = series_by_alias.get(norm(a))
            if s is not None:
                resolved = s; report_lines.append(f"- ✅ **{label}** ← `{a}`"); break
        if resolved is not None:
            master_to_source[c] = resolved
        else:
            if disp_norm == norm("Listing Action (List or Unlist)"):
                master_to_source[c] = SENTINEL_LIST
                report_lines.append(f"- 🟨 **{label}** ← default `'List'`")
            else:
                sugg = top_matches(effective, on_headers, 3)
                sug_txt = ", ".join(f"`{name}` ({round(sc*100,1)}%)" for sc, name in sugg) if sugg else "none"
                report_lines.append(f"- ❌ **{label}** ← no match. Suggestions: {sug_txt}")

    st.markdown("\n".join(report_lines))

    # Build dense 2D block
    n_rows = len(on_df)
    block = [[""] * used_cols for _ in range(n_rows)]
    for col, src in master_to_source.items():
        if src is SENTINEL_LIST:
            for i in range(n_rows):
                block[i][col-1] = "List"
        else:
            vals = src.astype(str).tolist()
            m = min(len(vals), n_rows)
            for i in range(m):
                v = clean(vals[i].strip())
                if v and v.lower() not in ("nan", "none", ""):
                    block[i][col-1] = v

    # -------- FAST WRITE with XlsxWriter (no other writers referenced) --------
    out_io = io.BytesIO()
    wb = xlsxwriter.Workbook(out_io, {"in_memory": True})
    ws = wb.add_worksheet(MASTER_TEMPLATE_SHEET)

    # Simple, fast formatting (customize as you like)
    hdr = wb.add_format({"bold": True, "bg_color": "#D9E1F2", "border": 1})
    body = wb.add_format({"border": 1})

    # Write headers
    for c, val in enumerate(display_headers, start=0):
        ws.write(MASTER_DISPLAY_ROW - 1, c, val, hdr)
    for c, val in enumerate(secondary_headers, start=0):
        ws.write(MASTER_SECONDARY_ROW - 1, c, val, hdr)

    # Write data block quickly
    start_row = MASTER_DATA_START_ROW - 1
    for r, row in enumerate(block, start=start_row):
        ws.write_row(r, 0, row, body)

    # UX niceties (still fast)
    ws.freeze_panes(start_row, 0)
    ws.autofilter(0, 0, start_row + n_rows, max(0, used_cols - 1))

    wb.close()
    out_io.seek(0)

    st.download_button(
        "⬇️ Download Final Masterfile",
        data=out_io.getvalue(),
        file_name="final_masterfile.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

    st.markdown("</div>", unsafe_allow_html=True)
