# app_masterfile.py
import io
import json
import re
import time
import hashlib
from pathlib import Path
from textwrap import dedent

import numpy as np
import pandas as pd
import streamlit as st
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter

# ─────────────────────────────────────────────────────────────────────
# App setup
# ─────────────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Masterfile Automation – Preserve Others, Write Template Fast",
    page_icon="🧾",
    layout="wide"
)

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
def clean_vec(arr: np.ndarray) -> np.ndarray:
    """Vectorized clean: remove invalid XML chars and trim common 'nan/none' strings."""
    # Convert to str (vectorized) once
    arr = arr.astype(object)
    # Normalize None/NaN → ""
    mask_none = pd.isna(arr)
    if mask_none.any():
        arr[mask_none] = ""
    # Strip and drop 'nan'/'none'
    def _norm_one(x):
        s = str(x).strip()
        if not s:
            return ""
        sl = s.lower()
        if sl in ("nan", "none"):
            return ""
        # remove invalid xml chars
        return _INVALID_XML_CHARS.sub("", s)
    vfunc = np.vectorize(_norm_one, otypes=[object])
    return vfunc(arr)

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

def clear_data_region_fast(ws, start_row: int):
    """Delete a trailing data region only if it exists (avoids expensive shifts)."""
    max_row = ws.max_row or start_row
    if max_row >= start_row:
        ws.delete_rows(idx=start_row, amount=max_row - start_row + 1)

def append_block_fast(ws, start_row: int, block_2d: np.ndarray):
    """Append using openpyxl's append in a tight loop; pre-pad to start_row-1."""
    cur_max = ws.max_row or 0
    need = (start_row - 1) - cur_max
    if need > 0:
        ws.insert_rows(idx=cur_max + 1, amount=need)
    # Append row-by-row (openpyxl's internal row builder is efficient when styles aren't touched)
    for i in range(block_2d.shape[0]):
        ws.append(block_2d[i, :].tolist())

def update_tables_to_new_ref(ws, header_row: int, start_row: int, n_cols: int, n_rows: int):
    if not getattr(ws, "tables", None):
        return
    last_row = max(header_row, start_row + max(0, n_rows - 1))
    new_ref = f"A{header_row}:{get_column_letter(n_cols)}{last_row}"
    for _, tbl in list(ws.tables.items()):
        tbl.ref = new_ref

def sha1_bytes(b: bytes) -> str:
    h = hashlib.sha1(); h.update(b); return h.hexdigest()

# ─────────────────────────────────────────────────────────────────────
# CACHED OPERATIONS (major speed-ups on repeat runs)
# ─────────────────────────────────────────────────────────────────────
@st.cache_data(show_spinner=False)
def cached_parse_mapping(mapping_text: str, mapping_file_bytes: bytes | None):
    if mapping_text.strip():
        mapping_raw = json.loads(mapping_text)
    else:
        mapping_raw = json.loads(mapping_file_bytes.decode("utf-8"))
    mapping_aliases = {}
    for k, v in mapping_raw.items():
        aliases = v[:] if isinstance(v, list) else [v]
        if k not in aliases: aliases.append(k)
        mapping_aliases[norm(k)] = aliases
    return mapping_aliases

@st.cache_data(show_spinner=False)
def cached_template_headers(master_bytes: bytes, sheet_name: str):
    wb_ro = load_workbook(io.BytesIO(master_bytes), read_only=True, data_only=True, keep_links=True)
    if sheet_name not in wb_ro.sheetnames:
        wb_ro.close()
        raise ValueError(f"Sheet '{sheet_name}' not found in template.")
    ws_ro = wb_ro[sheet_name]
    used_cols = worksheet_used_cols(ws_ro, header_rows=(MASTER_DISPLAY_ROW, MASTER_SECONDARY_ROW))
    display_headers   = [ws_ro.cell(row=MASTER_DISPLAY_ROW,   column=c).value or "" for c in range(1, used_cols+1)]
    secondary_headers = [ws_ro.cell(row=MASTER_SECONDARY_ROW, column=c).value or "" for c in range(1, used_cols+1)]
    wb_ro.close()
    return used_cols, display_headers, secondary_headers

@st.cache_data(show_spinner=False)
def cached_pick_onboarding(onboarding_bytes: bytes, mapping_aliases_by_master: dict):
    bio = io.BytesIO(onboarding_bytes)
    xl = pd.ExcelFile(bio, engine="openpyxl")
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
    df = best[0].fillna("")
    df.columns = [str(c).strip() for c in df.columns]
    return df, best[1], best_info

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
    mapping_json_upload = st.file_uploader("Or upload mapping.json", type=["json"], key="mapping_file")

go = st.button("🚀 Generate Final Masterfile", type="primary")
download_placeholder = st.empty()   # Download appears here after write completes

# ─────────────────────────────────────────────────────────────────────
# Main
# ─────────────────────────────────────────────────────────────────────
if go:
    if not masterfile_file or not onboarding_file:
        st.error("Please upload both files.")
        st.stop()

    with st.status("Starting…", expanded=True) as status:
        try:
            # Read bytes once (for caching keys)
            masterfile_file.seek(0); master_bytes = masterfile_file.read()
            onboarding_file.seek(0); onboarding_bytes = onboarding_file.read()
            map_bytes = mapping_json_upload.read() if mapping_json_upload is not None else None

            # 1) Parse mapping (cached)
            status.update(label="Parsing mapping JSON…")
            mapping_aliases = cached_parse_mapping(mapping_json_text, map_bytes)
            status.write("✅ Mapping parsed & cached")

            # 2) Read template headers (cached by file hash)
            status.update(label="Reading Template headers…")
            used_cols, display_headers, secondary_headers = cached_template_headers(master_bytes, MASTER_TEMPLATE_SHEET)
            status.write(f"✅ Template headers loaded ({used_cols} columns)")

            # 3) Pick onboarding sheet (cached)
            status.update(label="Selecting best onboarding sheet…")
            on_df, best_sheet, info = cached_pick_onboarding(onboarding_bytes, mapping_aliases)
            on_headers = list(on_df.columns)
            status.write(f"✅ Using onboarding sheet: **{best_sheet}** ({info})")

            # 4) Resolve mapping master->source (fast)
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

            # 5) Build dense data block (NUMPY-ACCELERATED, by column)
            status.update(label="Building data block…")
            n_rows = len(on_df)
            # Preallocate empty matrix
            block = np.empty((n_rows, used_cols), dtype=object)
            block[:] = ""

            for col, src in master_to_source.items():
                if src is SENTINEL_LIST:
                    block[:, col-1] = "List"
                else:
                    arr = src.to_numpy(dtype=object, copy=False)
                    arr = clean_vec(arr)  # vectorized cleaning
                    # Truncate or pad to n_rows
                    if arr.shape[0] < n_rows:
                        colvec = np.empty((n_rows,), dtype=object); colvec[:] = ""
                        colvec[:arr.shape[0]] = arr
                        block[:, col-1] = colvec
                    else:
                        block[:, col-1] = arr[:n_rows]

            # 6) Write only Template sheet, keep others intact (openpyxl, optimized)
            status.update(label="Writing Template sheet (preserving other sheets)…")
            t_write = time.time()
            ext = (Path(masterfile_file.name).suffix or ".xlsx").lower()
            keep_vba = (ext == ".xlsm")

            wb = load_workbook(io.BytesIO(master_bytes), keep_vba=keep_vba, data_only=False, keep_links=True)
            ws = wb[MASTER_TEMPLATE_SHEET]

            # Speed knob: Excel recalculates on open (single hit)
            try:
                wb.calculation_properties.fullCalcOnLoad = True
            except Exception:
                pass

            # Clear old data once (skip if nothing to clear)
            if (ws.max_row or 0) >= MASTER_DATA_START_ROW:
                clear_data_region_fast(ws, start_row=MASTER_DATA_START_ROW)

            # Append fast
            append_block_fast(ws, start_row=MASTER_DATA_START_ROW, block_2d=block)

            # Single table resize
            update_tables_to_new_ref(ws, header_row=MASTER_DISPLAY_ROW,
                                     start_row=MASTER_DATA_START_ROW,
                                     n_cols=used_cols, n_rows=n_rows)

            out_io = io.BytesIO()
            wb.save(out_io)
            out_io.seek(0)
            status.write(f"✅ Wrote & saved in {time.time()-t_write:.2f}s")

            # Finish + show download button BELOW the status panel
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
    - Speed-ups added:
        - Cached mapping + header detection + onboarding sheet selection.
        - NumPy/pandas vectorized column fills (minimal Python loops).
        - Skip row deletion if nothing to clear; otherwise single `delete_rows`.
        - Single table resize and `fullCalcOnLoad=True` to defer calc to Excel.
    """))
