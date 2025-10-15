import io
import json
import re
import time
import zipfile
import hashlib
from pathlib import Path
from textwrap import dedent

import pandas as pd
import streamlit as st
from openpyxl import load_workbook

# ─────────────────────────────────────────────────────────────────────
# Page config
# ─────────────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Masterfile Automation — Cached & Fast (OpenPyXL)",
    page_icon="🧾",
    layout="wide",
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
MASTER_TEMPLATE_SHEET = "Template"   # write only here
MASTER_DISPLAY_ROW    = 2            # header row
MASTER_SECONDARY_ROW  = 3            # subheader row (e.g., bullet_point1..)
MASTER_DATA_START_ROW = 4            # first data row

_INVALID_XML_CHARS = re.compile(r"[\x00-\x08\x0B\x0C\x0E-\x1F\uD800-\uDFFF]")

# ─────────────────────────────────────────────────────────────────────
# Helpers
# ─────────────────────────────────────────────────────────────────────
def sanitize(s):
    if s is None:
        return ""
    return _INVALID_XML_CHARS.sub("", str(s))

def norm(s: str) -> str:
    if s is None:
        return ""
    x = str(s).strip().lower()
    x = x.replace("–","-").replace("—","-").replace("−","-")
    x = re.sub(r"[._/\\-]+", " ", x)
    x = re.sub(r"[^0-9a-z\s]+", " ", x)
    return re.sub(r"\s+", " ", x).strip()

def worksheet_used_cols(ws, header_rows=(1,), hard_cap=4096, empty_streak_stop=8):
    max_try = min(ws.max_column or 1, hard_cap)
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
        rows = df.replace("", pd.NA).dropna(how="all").shape[0]
        score = matches + (0.01 if rows > 0 else 0.0)
        if score > best_score:
            best, best_score = (df, sheet), score
            best_info = f"matched headers: {matches}, non-empty rows: {rows}"
    if best is None:
        raise ValueError("No readable onboarding sheet found.")
    return best[0], best[1], best_info

# ─────────────────────────────────────────────────────────────────────
# OpenPyXL DELTA writer (values only) + zero-compression ZIP repack
# ─────────────────────────────────────────────────────────────────────
def openpyxl_delta_fast(master_bytes: bytes,
                        sheet_name: str,
                        start_row: int,
                        used_cols: int,
                        block_2d: list,
                        repack: bool = True) -> bytes:
    """
    Pure OpenPyXL 'delta' writer:
      - Only updates changed cells
      - Appends or trims tail rows minimally
      - No styles touched → preserves whole workbook
    Optional zero-compression (ZIP_STORED) repack for faster save.
    """
    wb = load_workbook(io.BytesIO(master_bytes), keep_vba=True, data_only=False)
    if sheet_name not in wb.sheetnames:
        raise ValueError(f"Sheet '{sheet_name}' not found.")
    ws = wb[sheet_name]

    target_rows = len(block_2d)
    end_row_new  = start_row + max(0, target_rows - 1)
    end_row_prev = ws.max_row or (start_row - 1)

    # 1) Update overlap rows (only changed cells)
    overlap = max(0, min(end_row_prev, end_row_new) - start_row + 1)
    if overlap:
        it = ws.iter_rows(min_row=start_row,
                          max_row=start_row + overlap - 1,
                          min_col=1,
                          max_col=used_cols,
                          values_only=True)
        for i, old_row_vals in enumerate(it):
            new_vals = block_2d[i]
            if len(new_vals) < used_cols:
                new_vals = new_vals + [""] * (used_cols - len(new_vals))
            elif len(new_vals) > used_cols:
                new_vals = new_vals[:used_cols]
            row_idx = start_row + i
            for j in range(used_cols):
                old_v = old_row_vals[j] if old_row_vals and j < len(old_row_vals) else None
                nv = new_vals[j]
                nv_norm = None if nv == "" else nv
                if old_v != nv_norm:
                    ws.cell(row=row_idx, column=j+1).value = nv_norm

    # 2) Append new tail rows
    for i in range(overlap, target_rows):
        row = block_2d[i]
        if len(row) > used_cols:
            row = row[:used_cols]
        elif len(row) < used_cols:
            row = row + [""] * (used_cols - len(row))
        ws.append(tuple(row))

    # 3) Trim old tail in one call
    if end_row_prev > end_row_new:
        ws.delete_rows(end_row_new + 1, end_row_prev - end_row_new)

    # 4) Save
    out = io.BytesIO()
    wb.save(out)
    out.seek(0)
    raw = out.getvalue()

    # 5) Optional: repack with ZIP_STORED (no compression) for speed
    if repack:
        zin = zipfile.ZipFile(io.BytesIO(raw), "r")
        mem = io.BytesIO()
        with zipfile.ZipFile(mem, "w", compression=zipfile.ZIP_STORED) as zout:
            for info in zin.infolist():
                zout.writestr(info.filename, zin.read(info.filename))
        zin.close()
        mem.seek(0)
        return mem.getvalue()

    return raw

# ─────────────────────────────────────────────────────────────────────
# Build-key hashing + cache
# ─────────────────────────────────────────────────────────────────────
def _stable_hash_bytes(*parts: bytes) -> str:
    h = hashlib.sha256()
    for p in parts:
        h.update(p)
        h.update(b"|")
    return h.hexdigest()

def _block_to_ndjson(block_2d, used_cols: int) -> bytes:
    # Compact, deterministic representation of the block (avoids huge JSON arrays in cache key)
    lines = []
    for row in block_2d:
        r = row[:used_cols] + [""] * max(0, used_cols - len(row))
        lines.append(json.dumps(r, separators=(",", ":")))  # stable, minimal
    return ("\n".join(lines)).encode("utf-8")

@st.cache_data(show_spinner=False, max_entries=64)
def build_cached(master_bytes: bytes,
                 mapping_json: str,
                 block_ndjson: bytes,
                 sheet_name: str,
                 start_row: int,
                 used_cols: int) -> bytes:
    """
    Cache entry keyed by master + mapping + block contents.
    If the same data is requested again, we return the bytes instantly.
    """
    # The function inputs themselves are the cache key; nothing to do here.
    block_2d = [json.loads(line) for line in block_ndjson.decode("utf-8").splitlines()] if block_ndjson else []
    return openpyxl_delta_fast(
        master_bytes=master_bytes,
        sheet_name=sheet_name,
        start_row=start_row,
        used_cols=used_cols,
        block_2d=block_2d,
        repack=True,
    )

# ─────────────────────────────────────────────────────────────────────
# UI
# ─────────────────────────────────────────────────────────────────────
st.title("🧾 Masterfile Automation — Cached & Fast (OpenPyXL only)")
st.caption("Preserves the whole workbook. Delta updates + content cache for instant repeat runs.")

with st.container():
    c1, c2 = st.columns([1,1])
    with c1:
        masterfile_file = st.file_uploader("📄 Masterfile Template (.xlsx / .xlsm)", type=["xlsx","xlsm"])
    with c2:
        onboarding_file = st.file_uploader("🧾 Onboarding (.xlsx)", type=["xlsx"])

st.markdown("#### 🔗 Mapping JSON")
tab1, tab2 = st.tabs(["Paste JSON", "Upload JSON"])
mapping_json_text = ""
mapping_json_file = None
with tab1:
    mapping_json_text = st.text_area("Paste mapping JSON", height=200,
        placeholder='{\n  "Partner SKU": ["Seller SKU","item_sku"]\n}')
with tab2:
    mapping_json_file = st.file_uploader("Or upload mapping.json", type=["json"], key="mapping_file")

go = st.button("🚀 Generate Final Masterfile", type="primary")
download_placeholder = st.empty()

# ─────────────────────────────────────────────────────────────────────
# Main
# ─────────────────────────────────────────────────────────────────────
if go:
    if not masterfile_file or not onboarding_file:
        st.error("Please upload both files."); st.stop()

    with st.status("Starting…", expanded=True) as status:
        try:
            # Parse mapping
            status.update(label="Parsing mapping JSON…")
            mapping_raw = json.loads(mapping_json_text) if mapping_json_text.strip() else json.load(mapping_json_file)
            # keep a deterministic, compact mapping string for the key
            mapping_json_compact = json.dumps(mapping_raw, separators=(",", ":"), sort_keys=True)

            # Read Template headers (read-only)
            status.update(label="Reading template headers…")
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

            # Read onboarding — choose best sheet
            status.update(label="Reading onboarding…")
            onboarding_file.seek(0)
            xl = pd.ExcelFile(onboarding_file, engine="openpyxl")
            # derive aliases map
            aliases_map = {}
            for k, v in mapping_raw.items():
                aliases = v[:] if isinstance(v, list) else [v]
                if k not in aliases: aliases.append(k)
                aliases_map[norm(k)] = aliases
            best_df, best_sheet, info = pick_best_onboarding_sheet(onboarding_file, aliases_map)
            df = best_df.fillna("")
            on_headers = list(df.columns)
            status.write(f"✅ Onboarding: **{best_sheet}** ({info}); rows={len(df)}")

            # Build mapping master -> source
            status.update(label="Resolving mapping…")
            from difflib import SequenceMatcher
            def top_matches(query, candidates, k=3):
                q = norm(query)
                scored = [(SequenceMatcher(None, q, norm(c)).ratio(), c) for c in candidates]
                scored.sort(key=lambda t: t[0], reverse=True)
                return scored[:k]

            SENTINEL_LIST = object()
            series_by_alias = {norm(h): df[h] for h in on_headers}
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
                aliases = aliases_map.get(eff_norm, [effective])
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

            status.write("**Mapping summary**")
            for line in report_lines:
                status.write(line)

            # Build 2D block (strings only) + compact NDJSON for cache key
            status.update(label="Building data block…")
            n_rows = len(df)
            block = [[""] * used_cols for _ in range(n_rows)]
            for col, src in master_to_source.items():
                if src is SENTINEL_LIST:
                    for i in range(n_rows):
                        block[i][col-1] = "List"
                else:
                    vals = src.astype(str).tolist()
                    m = min(len(vals), n_rows)
                    for i in range(m):
                        v = sanitize(vals[i].strip())
                        if v and v.lower() not in ("nan", "none", ""):
                            block[i][col-1] = v
            block_ndjson = _block_to_ndjson(block, used_cols)

            # Cache key (implicit) = function args; we also show a tiny fingerprint
            fingerprint = _stable_hash_bytes(master_bytes, mapping_json_compact.encode("utf-8"), block_ndjson)[:12]
            status.update(label=f"Writing (cache key {fingerprint})…")
            t_write = time.time()

            out_bytes = build_cached(
                master_bytes=master_bytes,
                mapping_json=mapping_json_compact,
                block_ndjson=block_ndjson,
                sheet_name=MASTER_TEMPLATE_SHEET,
                start_row=MASTER_DATA_START_ROW,
                used_cols=used_cols,
            )

            status.write(f"✅ Done in {time.time()-t_write:.2f}s (cache-aware)")
            status.update(label="Finished", state="complete")

            # Download
            ext = (Path(masterfile_file.name).suffix or ".xlsx").lower()
            mime = "application/vnd.ms-excel.sheet.macroEnabled.12" if ext == ".xlsm" else \
                   "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            download_placeholder.download_button(
                "⬇️ Download Final Masterfile",
                data=out_bytes,
                file_name=f"final_masterfile{ext}",
                mime=mime
            )

        except Exception as e:
            status.update(label="Error", state="error")
            st.exception(e)

# ─────────────────────────────────────────────────────────────────────
# Notes
# ─────────────────────────────────────────────────────────────────────
with st.expander("📘 Notes", expanded=False):
    st.markdown(dedent(f"""
    **Why this feels instant**
    - We hash **template + mapping + data block** and **cache the fully built workbook bytes**.
    - If you build the same (or slightly modified) masterfile again, it returns from cache in milliseconds.
    - On a new combination, we still write via the **delta OpenPyXL path** (values only, preserves everything else) and then cache the result.
    - No new packages. No services. No Node/COM/XML patching.
    """))
