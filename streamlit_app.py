# app_masterfile.py
import io
import json
import re
import time
from pathlib import Path
from textwrap import dedent

import pandas as pd
import streamlit as st
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter

# ─────────────────────────────────────────────────────────────────────
# App setup
# ─────────────────────────────────────────────────────────────────────
st.set_page_config(page_title="Masterfile Automation – Preserve Others, Write Template Fast", page_icon="🧾", layout="wide")

st.markdown("""
<style>
.section{border:1px solid #e8eef6;background:#fff;border-radius:16px;padding:18px;margin:12px 0;box-shadow:0 6px 24px rgba(2,6,23,.05)}
div.stButton>button,.stDownloadButton>button{background:#2563eb!important;color:#fff!important;border-radius:10px!important;border:0!important}
</style>
""", unsafe_allow_html=True)

# ─────────────────────────────────────────────────────────────────────
# Constants
# ─────────────────────────────────────────────────────────────────────
MASTER_TEMPLATE_SHEET = "Template"  # the only sheet we write
MASTER_DISPLAY_ROW    = 2           # headers row
MASTER_SECONDARY_ROW  = 3           # subheaders row (e.g., bullet_point labels)
MASTER_DATA_START_ROW = 4           # first data row

# ─────────────────────────────────────────────────────────────────────
# Helpers (fast, cloud-safe; NO XML patching / NO COM / NO Node)
# ─────────────────────────────────────────────────────────────────────
_INVALID_XML_CHARS = re.compile(r"[\x00-\x08\x0B\x0C\x0E-\x1F\uD800-\uDFFF]")
def clean(s):
    if s is None: return ""
    return _INVALID_XML_CHARS.sub("", str(s))

def norm(s: str) -> str:
    if s is None: return ""
    x = str(s).strip().lower()
    x = x.replace("–","-").replace("—","-").replace("−","-")
    x = re.sub(r"[._/\\-]+"," ",x)
    x = re.sub(r"[^0-9a-z\s]+"," ",x)
    return re.sub(r"\s+", " ", x).strip()

def nonempty_rows(df: pd.DataFrame) -> int:
    if df.empty: return 0
    return df.replace("", pd.NA).dropna(how="all").shape[0]

def worksheet_used_cols(ws, header_rows=(1,), hard_cap=4096, empty_streak_stop=8):
    max_try = min(ws.max_column or 1, hard_cap)
    last_nonempty, streak = 0, 0
    for c in range(1, max_try + 1):
        any_val = False
        for r in header_rows:
            v = ws.cell(row=r, column=c).value
            if v not in (None, ""):
                any_val = True; break
        if any_val:
            last_nonempty, streak = c, 0
        else:
            streak += 1
            if streak >= empty_streak_stop: break
    return max(last_nonempty, 1)

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

def clear_data_region(ws, start_row: int):
    max_row = ws.max_row or start_row
    if max_row >= start_row:
        ws.delete_rows(idx=start_row, amount=max_row - start_row + 1)

def append_block(ws, start_row: int, block_2d):
    # Ensure header rows exist; then append fast (values only)
    cur_max = ws.max_row or 0
    need = (start_row - 1) - cur_max
    if need > 0:
        ws.insert_rows(idx=cur_max + 1, amount=need)
    for row in block_2d:
        ws.append([clean(v) for v in row])

def update_tables_to_new_ref(ws, header_row: int, start_row: int, n_cols: int, n_rows: int):
    if not getattr(ws, "tables", None):
        return
    last_row = max(header_row, start_row + max(0, n_rows - 1))
    new_ref = f"A{header_row}:{get_column_letter(n_cols)}{last_row}"
    for _, tbl in list(ws.tables.items()):
        tbl.ref = new_ref

# ─────────────────────────────────────────────────────────────────────
# UI
# ─────────────────────────────────────────────────────────────────────
st.title("🧾 Masterfile Automation – Fast (Write Template only, keep other sheets)")
st.caption("Cloud-safe: only updates the Template sheet; preserves all other tabs, styles, formulas, and macros.")

with st.container():
    c1, c2 = st.columns([1,1])
    with c1:
        masterfile_file = st.file_uploader("📄 Masterfile Template (.xlsx / .xlsm)", type=["xlsx","xlsm"])
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

# ─────────────────────────────────────────────────────────────────────
# Main
# ─────────────────────────────────────────────────────────────────────
download_placeholder = st.empty()   # we’ll place the download button here after write finishes

if go:
    if not masterfile_file or not onboarding_file:
        st.error("Please upload both files.")
        st.stop()

    # Progress panel (Streamlit status keeps messages grouped under the button)
    with st.status("Starting…", expanded=True) as status:
        try:
            # Step 1: parse mapping
            status.update(label="Parsing mapping JSON…")
            try:
                mapping_raw = json.loads(mapping_json_text) if mapping_json_text.strip() else json.load(mapping_json_file)
            except Exception as e:
                st.error(f"Mapping JSON parse error: {e}")
                status.update(state="error")
                st.stop()

            mapping_aliases = {}
            for k, v in mapping_raw.items():
                aliases = v[:] if isinstance(v, list) else [v]
                if k not in aliases: aliases.append(k)
                mapping_aliases[norm(k)] = aliases

            # Step 2: read template headers (read-only)
            status.write("Reading Template headers…")
            masterfile_file.seek(0)
            master_bytes = masterfile_file.read()

            t0 = time.time()
            wb_ro = load_workbook(io.BytesIO(master_bytes), read_only=True, data_only=True, keep_links=True)
            if MASTER_TEMPLATE_SHEET not in wb_ro.sheetnames:
                st.error(f"Sheet '{MASTER_TEMPLATE_SHEET}' not found in template.")
                status.update(state="error"); st.stop()
            ws_ro = wb_ro[MASTER_TEMPLATE_SHEET]
            used_cols = worksheet_used_cols(ws_ro, header_rows=(MASTER_DISPLAY_ROW, MASTER_SECONDARY_ROW))
            display_headers   = [ws_ro.cell(row=MASTER_DISPLAY_ROW,   column=c).value or "" for c in range(1, used_cols+1)]
            secondary_headers = [ws_ro.cell(row=MASTER_SECONDARY_ROW, column=c).value or "" for c in range(1, used_cols+1)]
            wb_ro.close()
            status.write(f"✅ Template headers loaded ({used_cols} columns) in {time.time()-t0:.2f}s")

            # Step 3: pick onboarding sheet
            status.update(label="Selecting best onboarding sheet…")
            try:
                onboarding_file.seek(0)
                best_df, best_sheet, info = pick_best_onboarding_sheet(onboarding_file, mapping_aliases)
            except Exception as e:
                st.error(f"Onboarding error: {e}")
                status.update(state="error"); st.stop()

            on_df = best_df.fillna("")
            on_df.columns = [str(c).strip() for c in on_df.columns]
            on_headers = list(on_df.columns)
            status.write(f"✅ Using onboarding sheet: **{best_sheet}** ({info})")

            # Step 4: build mapping master->source
            status.update(label="Resolving column mapping…")
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
                if not eff_norm: continue

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

            status.write("**Mapping Summary**")
            for line in report_lines: status.write(line)

            # Step 5: build dense data block (fast)
            status.update(label="Building data block…")
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

            # Step 6: write only Template sheet, keep others intact (fast openpyxl path)
            status.update(label="Writing Template sheet (preserving other sheets)…")
            t_write = time.time()
            ext = (Path(masterfile_file.name).suffix or ".xlsx").lower()
            keep_vba = (ext == ".xlsm")

            wb = load_workbook(io.BytesIO(master_bytes), keep_vba=keep_vba, data_only=False, keep_links=True)
            ws = wb[MASTER_TEMPLATE_SHEET]

            # Slight speed knobs: defer Excel calc, value-only appends
            try:
                wb.calculation_properties.fullCalcOnLoad = True
            except Exception:
                pass

            clear_data_region(ws, start_row=MASTER_DATA_START_ROW)
            append_block(ws, start_row=MASTER_DATA_START_ROW, block_2d=block)
            update_tables_to_new_ref(ws, header_row=MASTER_DISPLAY_ROW, start_row=MASTER_DATA_START_ROW,
                                     n_cols=used_cols, n_rows=n_rows)

            out_io = io.BytesIO()
            wb.save(out_io)
            out_io.seek(0)
            status.write(f"✅ Wrote & saved in {time.time()-t_write:.2f}s")

            # Step 7: finish + show download button BELOW the status panel
            status.update(label="Finished", state="complete")

            mime_map = {
                ".xlsx": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                ".xlsm": "application/vnd.ms-excel.sheet.macroEnabled.12",
            }
            out_mime = mime_map.get(ext, mime_map[".xlsx"])

            download_placeholder.download_button(
                "⬇️ Download Final Masterfile",
                data=out_io.getvalue(),
                file_name=f"final_masterfile{ext}",
                mime=out_mime
            )

        except Exception as e:
            status.update(label="Error", state="error")
            st.exception(e)

# ─────────────────────────────────────────────────────────────────────
# Help
# ─────────────────────────────────────────────────────────────────────
with st.expander("📘 Notes", expanded=False):
    st.markdown(dedent(f"""
    - **Only the `{MASTER_TEMPLATE_SHEET}` sheet is modified**; all other sheets/macros/formatting remain intact.
    - Fast path: value-only bulk appends (no style writes), single table resize, and Excel recalculates on open.
    - Works on Streamlit Cloud / Cloud Run / HF Spaces (pure Python).
    """))
