import io
import json
import re
import time
import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path
from textwrap import dedent

import pandas as pd
import streamlit as st
from openpyxl import load_workbook

# Try optional speed libs (read acceleration)
try:
    import pyarrow as pa
    import pyarrow.csv as pacsv
    _HAS_ARROW = True
except Exception:
    _HAS_ARROW = False

try:
    import pyxlsb  # noqa: F401 (engine used by pandas)
    _HAS_PYXLSB = True
except Exception:
    _HAS_PYXLSB = False

# ─────────────────────────────────────────────────────────────────────
# Page config
# ─────────────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Masterfile Automation – Fast XML Writer + Read Acceleration",
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
MASTER_SECONDARY_ROW  = 3            # subheader row
MASTER_DATA_START_ROW = 4            # first data row

XL_NS_MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
XL_NS_REL  = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
ET.register_namespace("", XL_NS_MAIN)
ET.register_namespace("r", XL_NS_REL)
ET.register_namespace("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006")
ET.register_namespace("x14ac", "http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac")

_INVALID_XML_CHARS = re.compile(r"[\x00-\x08\x0B\x0C\x0E-\x1F\uD800-\uDFFF]")

# ─────────────────────────────────────────────────────────────────────
# Helpers (general)
# ─────────────────────────────────────────────────────────────────────
def sanitize_xml_text(s):
    if s is None:
        return ""
    return _INVALID_XML_CHARS.sub("", str(s))

def _col_letter(n: int) -> str:
    s = ""
    while n:
        n, r = divmod(n-1, 26)
        s = chr(65+r) + s
    return s

def _col_number(letters: str) -> int:
    n = 0
    for ch in letters:
        if not ch.isalpha(): break
        n = n * 26 + (ord(ch.upper()) - 64)
    return n

def norm(s: str) -> str:
    if s is None: return ""
    x = str(s).strip().lower()
    x = x.replace("–","-").replace("—","-").replace("−","-")
    x = re.sub(r"[._/\\-]+"," ",x)
    x = re.sub(r"[^0-9a-z\s]+"," ",x)
    return re.sub(r"\s+"," ",x).strip()

def nonempty_rows(df: pd.DataFrame) -> int:
    if df.empty: return 0
    return df.replace("", pd.NA).dropna(how="all").shape[0]

def pick_best_onboarding_sheet(uploaded_file, mapping_aliases_by_master):
    uploaded_file.seek(0)
    # pandas engine chosen by file type; this only enumerates sheet names when Excel
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

# ─────────────────────────────────────────────────────────────────────
# OOXML path helpers
# ─────────────────────────────────────────────────────────────────────
def _find_sheet_part_path(z: zipfile.ZipFile, sheet_name: str) -> str:
    wb_xml = ET.fromstring(z.read("xl/workbook.xml"))
    rels_xml = ET.fromstring(z.read("xl/_rels/workbook.xml.rels"))
    rid = None
    for sh in wb_xml.find(f"{{{XL_NS_MAIN}}}sheets"):
        if sh.attrib.get("name") == sheet_name:
            rid = sh.attrib.get(f"{{{XL_NS_REL}}}id")
            break
    if not rid:
        raise ValueError(f"Sheet '{sheet_name}' not found.")
    target = None
    for rel in rels_xml:
        if rel.attrib.get("Id") == rid:
            target = rel.attrib.get("Target"); break
    if not target:
        raise ValueError(f"Relationship for sheet '{sheet_name}' not found.")
    target = target.replace("\\", "/")
    if target.startswith("../"): target = target[3:]
    if not target.startswith("xl/"): target = "xl/" + target
    return target

def _get_table_paths_for_sheet(z: zipfile.ZipFile, sheet_path: str) -> list:
    rels_path = sheet_path.replace("worksheets/", "worksheets/_rels/").replace(".xml", ".xml.rels")
    if rels_path not in z.namelist(): return []
    root = ET.fromstring(z.read(rels_path))
    out = []
    for rel in root:
        t = rel.attrib.get("Type", "")
        if t.endswith("/table"):
            target = rel.attrib.get("Target", "").replace("\\", "/")
            if target.startswith("../"): target = target[3:]
            if not target.startswith("xl/"): target = "xl/" + target
            out.append(target)
    return out

def _read_table_cols_count(table_xml_bytes: bytes) -> int:
    try:
        root = ET.fromstring(table_xml_bytes)
        tcols = root.find(f"{{{XL_NS_MAIN}}}tableColumns")
        if tcols is None: return 0
        count_attr = tcols.attrib.get("count")
        try:
            count = int(count_attr) if count_attr is not None else 0
        except Exception:
            count = 0
        child_count = sum(1 for _ in tcols)
        return max(count, child_count)
    except Exception:
        return 0

def _union_dimension(orig_dim_ref: str, used_cols: int, last_row: int) -> str:
    try:
        _, right = orig_dim_ref.split(":", 1)
        m = re.match(r"([A-Z]+)(\d+)", right)
        if m:
            orig_last_col = _col_number(m.group(1))
            orig_last_row = int(m.group(2))
        else:
            orig_last_col, orig_last_row = used_cols, last_row
    except Exception:
        orig_last_col, orig_last_row = used_cols, last_row
    u_last_col = max(orig_last_col, used_cols)
    u_last_row = max(orig_last_row, last_row)
    return f"A1:{_col_letter(u_last_col)}{u_last_row}"

# ─────────────────────────────────────────────────────────────────────
# Sheet XML patch (rows + cleanup)
# ─────────────────────────────────────────────────────────────────────
def _patch_sheet_xml(sheet_xml_bytes, header_row, start_row, used_cols_final, block_2d):
    root = ET.fromstring(sheet_xml_bytes)
    ns = XL_NS_MAIN

    def in_replaced_block(ref):
        def _any_row(rng):
            rng = rng.strip()
            if not rng: return False
            parts = rng.split(":")
            def _row(addr):
                m = re.match(r"([A-Z]+)(\d+)$", addr)
                return int(m.group(2)) if m else 10**9
            if len(parts) == 1:
                return _row(parts[0]) >= start_row
            r1, r2 = _row(parts[0]), _row(parts[1])
            return max(r1, r2) >= start_row
        return any(_any_row(tok) for tok in str(ref).split())

    sheetData = root.find(f"{{{ns}}}sheetData")
    if sheetData is None:
        sheetData = ET.SubElement(root, f"{{{ns}}}sheetData")

    for row in list(sheetData):
        try: r = int(row.attrib.get("r", "0") or "0")
        except Exception: r = 0
        if r >= start_row:
            sheetData.remove(row)

    row_span = f"1:{used_cols_final}" if used_cols_final > 0 else "1:1"
    for i, row_vals in enumerate(block_2d):
        r = start_row + i
        row_el = ET.Element(f"{{{ns}}}row", r=str(r))
        row_el.set("spans", row_span)
        row_el.set("{http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac}dyDescent", "0.25")
        any_val = False
        for j in range(used_cols_final):
            v = row_vals[j] if j < len(row_vals) else ""
            if not v: continue
            txt = sanitize_xml_text(v)
            if not txt: continue
            any_val = True
            col = _col_letter(j+1)
            c = ET.Element(f"{{{ns}}}c", r=f"{col}{r}", t="inlineStr")
            is_el = ET.SubElement(c, f"{{{ns}}}is")
            t_el = ET.SubElement(is_el, f"{{{ns}}}t")
            t_el.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
            t_el.text = txt
            row_el.append(c)
        if any_val:
            sheetData.append(row_el)

    dim = root.find(f"{{{ns}}}dimension")
    if dim is None:
        dim = ET.SubElement(root, f"{{{ns}}}dimension")
        dim.set("ref", "A1:A1")
    last_row = start_row + max(0, len(block_2d) - 1)
    dim.set("ref", _union_dimension(dim.attrib.get("ref", "A1:A1"), used_cols_final, last_row))

    # Clean stale structures
    mcs = root.find(f"{{{ns}}}mergeCells")
    if mcs is not None:
        for mc in list(mcs):
            if in_replaced_block(mc.attrib.get("ref", "")):
                mcs.remove(mc)
        cnt = len(list(mcs))
        if cnt: mcs.set("count", str(cnt))
        else: root.remove(mcs)

    dvs = root.find(f"{{{ns}}}dataValidations")
    if dvs is not None:
        for dv in list(dvs):
            if in_replaced_block(dv.attrib.get("sqref", "")):
                dvs.remove(dv)
        cnt = len(list(dvs))
        if cnt: dvs.set("count", str(cnt))
        else: root.remove(dvs)

    for cf in list(root.findall(f"{{{ns}}}conditionalFormatting")):
        sqref = cf.attrib.get("sqref", "")
        if not sqref: continue
        kept = " ".join(tok for tok in sqref.split() if not in_replaced_block(tok))
        if kept: cf.set("sqref", kept)
        else: root.remove(cf)

    kept_hl_ids = set()
    hls = root.find(f"{{{ns}}}hyperlinks")
    if hls is not None:
        for hl in list(hls):
            ref = hl.attrib.get("ref", "")
            rid = hl.attrib.get(f"{{{XL_NS_REL}}}id")
            if in_replaced_block(ref):
                hls.remove(hl)
            else:
                if rid: kept_hl_ids.add(rid)
        if not list(hls):
            root.remove(hls)

    rbr = root.find(f"{{{ns}}}rowBreaks")
    if rbr is not None:
        for brk in list(rbr):
            try: r = int(brk.attrib.get("id", "0"))
            except Exception: r = 0
            if r >= start_row:
                rbr.remove(brk)
        if not list(rbr):
            root.remove(rbr)

    return ET.tostring(root, encoding="utf-8", xml_declaration=True), kept_hl_ids

def _patch_table_xml(table_xml_bytes, header_row, last_row, last_col_n):
    root = ET.fromstring(table_xml_bytes)
    ns = XL_NS_MAIN

    new_ref = f"A{header_row}:{_col_letter(last_col_n)}{last_row}"
    root.set("ref", new_ref)
    af = root.find(f"{{{ns}}}autoFilter")
    if af is None:
        af = ET.SubElement(root, f"{{{ns}}}autoFilter")
    af.set("ref", new_ref)

    tcols = root.find(f"{{{ns}}}tableColumns")
    if tcols is None:
        tcols = ET.SubElement(root, f"{{{ns}}}tableColumns")

    existing = list(tcols)
    child_count = len(existing)
    max_id = 0
    for tc in existing:
        try:
            max_id = max(max_id, int(tc.attrib.get("id", "0")))
        except Exception:
            pass

    for i in range(child_count + 1, last_col_n + 1):
        max_id += 1
        tc = ET.SubElement(tcols, f"{{{ns}}}tableColumn")
        tc.set("id", str(max_id))
        tc.set("name", f"Column{i}")

    tcols.set("count", str(len(list(tcols))))
    return ET.tostring(root, encoding="utf-8", xml_declaration=True)

def _patch_sheet_rels(rels_xml_bytes, kept_hyperlink_ids):
    try:
        ns_pkg = "http://schemas.openxmlformats.org/package/2006/relationships"
        ET.register_namespace("", ns_pkg)
        root = ET.fromstring(rels_xml_bytes)
        changed = False
        for rel in list(root.findall(f"{{{ns_pkg}}}Relationship")):
            rel_type = rel.attrib.get("Type", "")
            rid = rel.attrib.get("Id")
            if rel_type.endswith("/hyperlink") and rid and rid not in kept_hyperlink_ids:
                root.remove(rel); changed = True
        if changed:
            return ET.tostring(root, encoding="utf-8", xml_declaration=True)
        return rels_xml_bytes
    except Exception:
        return rels_xml_bytes

def _range_intersects_rows(ref_text: str, start_row: int) -> bool:
    if not ref_text:
        return False
    def _row(addr: str) -> int:
        addr = addr.split("!")[-1].replace("$", "")
        m = re.match(r"([A-Z]+)(\d+)$", addr)
        return int(m.group(2)) if m else 10**9
    for tok in str(ref_text).split():
        parts = tok.split(":")
        if len(parts) == 1:
            if _row(parts[0]) >= start_row:
                return True
        else:
            r1, r2 = _row(parts[0]), _row(parts[1])
            if max(r1, r2) >= start_row:
                return True
    return False

def _patch_defined_names_wide(wb_root: ET.Element, sheet_name: str,
                              local_sheet_id: int,
                              header_row: int, last_row: int, last_col_n: int) -> ET.Element:
    ns = XL_NS_MAIN
    dnames = wb_root.find(f"{{{ns}}}definedNames")
    if dnames is None:
        return wb_root
    new_ref_abs = f"'{sheet_name}'!$A${header_row}:${_col_letter(last_col_n)}${last_row}"
    to_remove = []
    for dn in list(dnames):
        lsid = dn.attrib.get("localSheetId")
        if lsid is None or str(lsid) != str(local_sheet_id):
            continue
        name = dn.attrib.get("name")
        if name == "_xlnm._FilterDatabase":
            dn.text = new_ref_abs; dn.set("hidden", "1"); continue
        ref_txt = (dn.text or "").strip()
        if _range_intersects_rows(ref_txt, header_row):
            to_remove.append(dn)
    for dn in to_remove:
        dnames.remove(dn)
    return wb_root

def fast_patch_template(master_bytes, sheet_name, header_row, start_row, used_cols, block_2d):
    zin = zipfile.ZipFile(io.BytesIO(master_bytes), "r")

    wb_xml = ET.fromstring(zin.read("xl/workbook.xml"))
    sheets_el = wb_xml.find(f"{{{XL_NS_MAIN}}}sheets")
    sheet_elems = list(sheets_el) if sheets_el is not None else []
    local_id = None
    for idx, sh in enumerate(sheet_elems):
        if sh.attrib.get("name") == sheet_name:
            local_id = idx; break
    if local_id is None:
        raise ValueError(f"Sheet '{sheet_name}' not found in workbook.xml")

    sheet_path = _find_sheet_part_path(zin, sheet_name)
    sheet_rels_path = sheet_path.replace("worksheets/", "worksheets/_rels/").replace(".xml", ".xml.rels")
    table_paths = _get_table_paths_for_sheet(zin, sheet_path)

    max_cols = used_cols
    for tp in table_paths:
        try:
            cnt = _read_table_cols_count(zin.read(tp))
            if cnt and cnt > max_cols:
                max_cols = cnt
        except Exception:
            pass

    original_sheet_xml = zin.read(sheet_path)
    new_sheet_xml, kept_hl_ids = _patch_sheet_xml(original_sheet_xml, header_row, start_row, max_cols, block_2d)

    last_row = start_row + max(0, len(block_2d) - 1)
    if last_row < header_row:
        last_row = header_row

    patched_tables = {}
    for tp in table_paths:
        try:
            patched_tables[tp] = _patch_table_xml(zin.read(tp), header_row, last_row, max_cols)
        except Exception:
            pass

    wb_xml_patched = _patch_defined_names_wide(
        wb_root=wb_xml,
        sheet_name=sheet_name,
        local_sheet_id=local_id,
        header_row=header_row,
        last_row=last_row,
        last_col_n=max_cols
    )
    wb_xml_bytes = ET.tostring(wb_xml_patched, encoding="utf-8", xml_declaration=True)

    out_bio = io.BytesIO()
    with zipfile.ZipFile(out_bio, "w", zipfile.ZIP_DEFLATED) as zout:
        content_types_bytes = None
        for item in zin.infolist():
            name = item.filename
            if name == sheet_path:
                zout.writestr(item, new_sheet_xml); continue
            if name in patched_tables:
                zout.writestr(item, patched_tables[name]); continue
            if name == "xl/workbook.xml":
                zout.writestr(item, wb_xml_bytes); continue
            if name == "xl/calcChain.xml":
                continue
            if name == "[Content_Types].xml":
                content_types_bytes = zin.read(name); continue
            if name == sheet_rels_path:
                orig_rels = zin.read(name)
                patched_rels = _patch_sheet_rels(orig_rels, kept_hl_ids)
                zout.writestr(item, patched_rels); continue
            zout.writestr(item, zin.read(name))
        if content_types_bytes is not None:
            try:
                ns_pkg = "http://schemas.openxmlformats.org/package/2006/content-types"
                ET.register_namespace("", ns_pkg)
                ct_root = ET.fromstring(content_types_bytes)
                changed = False
                for ov in list(ct_root.findall(f"{{{ns_pkg}}}Override")):
                    if ov.attrib.get("PartName", "").replace("\\", "/") == "/xl/calcChain.xml":
                        ct_root.remove(ov); changed = True
                ct_bytes = ET.tostring(ct_root, encoding="utf-8", xml_declaration=True) if changed else content_types_bytes
                zout.writestr("[Content_Types].xml", ct_bytes)
            except Exception:
                zout.writestr("[Content_Types].xml", content_types_bytes)

    zin.close()
    out_bio.seek(0)
    return out_bio.getvalue()

# ─────────────────────────────────────────────────────────────────────
# Read acceleration: onboarding loader
# ─────────────────────────────────────────────────────────────────────
def load_onboarding_fast(uploaded_file) -> pd.DataFrame:
    """
    Fast path for onboarding:
      - .csv  -> pyarrow.csv (if available) else pandas.read_csv
      - .xlsb -> pandas.read_excel(engine='pyxlsb') (if available)
      - .xlsx/.xlsm -> pandas.read_excel(engine='openpyxl')
    Returns a dataframe with all values as strings (no NaNs).
    """
    name = uploaded_file.name if hasattr(uploaded_file, "name") else "data"
    suffix = Path(name).suffix.lower()

    uploaded_file.seek(0)
    if suffix == ".csv":
        if _HAS_ARROW:
            # Read via pyarrow, then to pandas (fast & low memory)
            table = pacsv.read_csv(uploaded_file)
            df = table.to_pandas(types_mapper=pd.ArrowDtype).astype(str).fillna("")
        else:
            df = pd.read_csv(uploaded_file, dtype=str, keep_default_na=False).fillna("")
        return df

    # Excel paths
    uploaded_file.seek(0)
    if suffix == ".xlsb" and _HAS_PYXLSB:
        # Binary Excel: much faster than openpyxl
        df = pd.read_excel(uploaded_file, engine="pyxlsb", dtype=str).fillna("")
        return df

    # Fallback for .xlsx/.xlsm and anything else readable by openpyxl
    uploaded_file.seek(0)
    df = pd.read_excel(uploaded_file, engine="openpyxl", dtype=str).fillna("")
    return df

# ─────────────────────────────────────────────────────────────────────
# UI
# ─────────────────────────────────────────────────────────────────────
st.title("🧾 Masterfile Automation – Fast XML (Template-only) + Read Acceleration")
st.caption("Reads large onboarding files faster (.xlsb via pyxlsb, .csv via pyarrow). Writes Template only; preserves other sheets/macros/styles.")

with st.container():
    c1, c2 = st.columns([1,1])
    with c1:
        masterfile_file = st.file_uploader("📄 Masterfile Template (.xlsx / .xlsm)", type=["xlsx","xlsm"])
    with c2:
        onboarding_file = st.file_uploader("🧾 Onboarding (.xlsx / .xlsm / .xlsb / .csv)", type=["xlsx","xlsm","xlsb","csv"])

st.markdown("#### 🔗 Mapping JSON")
tab1, tab2 = st.tabs(["Paste JSON", "Upload JSON"])
mapping_json_text = ""
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
            mapping_aliases = {}
            for k, v in mapping_raw.items():
                aliases = v[:] if isinstance(v, list) else [v]
                if k not in aliases: aliases.append(k)
                mapping_aliases[norm(k)] = aliases

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

            def worksheet_used_cols(ws, header_rows=(1,), hard_cap=4096, empty_streak_stop=8):
                max_try = min(ws.max_column or 1, hard_cap)
                last_nonempty, streak = 0, 0
                for c in range(1, max_try + 1):
                    any_val = any((ws.cell(row=r, column=c).value not in (None, "")) for r in header_rows)
                    if any_val:
                        last_nonempty, streak = c, 0
                    else:
                        streak += 1
                        if streak >= empty_streak_stop: break
                return max(last_nonempty, 1)

            used_cols = worksheet_used_cols(ws_ro, header_rows=(MASTER_DISPLAY_ROW, MASTER_SECONDARY_ROW))
            display_headers   = [ws_ro.cell(row=MASTER_DISPLAY_ROW,   column=c).value or "" for c in range(1, used_cols+1)]
            secondary_headers = [ws_ro.cell(row=MASTER_SECONDARY_ROW, column=c).value or "" for c in range(1, used_cols+1)]
            wb_ro.close()
            status.write(f"✅ Template headers loaded ({used_cols} columns) in {time.time()-t0:.2f}s")

            # Read onboarding with acceleration
            status.update(label="Reading onboarding (accelerated)…")
            t_read = time.time()
            df = load_onboarding_fast(onboarding_file)
            on_headers = list(df.columns)
            status.write(f"✅ Onboarding loaded ({len(df)} rows, {len(on_headers)} columns) in {time.time()-t_read:.2f}s")

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

            status.write("**Mapping summary**")
            for line in report_lines: status.write(line)

            # Build 2D block (string values)
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
                        v = vals[i].strip()
                        if v and v.lower() not in ("nan", "none", ""):
                            block[i][col-1] = v

            # ⚡ XML fast patch write (preserve other sheets/styles/macros)
            status.update(label="Writing (XML fast patch)…")
            t_write = time.time()
            out_bytes = fast_patch_template(
                master_bytes=master_bytes,
                sheet_name=MASTER_TEMPLATE_SHEET,
                header_row=MASTER_DISPLAY_ROW,
                start_row=MASTER_DATA_START_ROW,
                used_cols=used_cols,
                block_2d=block
            )
            status.write(f"✅ Wrote & saved in {time.time()-t_write:.2f}s")
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
    **Read acceleration**
    - `.xlsb` uses **pyxlsb** for faster Excel binary reads.
    - `.csv` uses **pyarrow.csv** (if installed) for very fast ingest.
    - `.xlsx/.xlsm` continue to use **openpyxl** (reliable & widely supported).

    **Writer**
    - Fast OOXML patch writes **only** the `{MASTER_TEMPLATE_SHEET}` sheet.
    - Preserves all other tabs, styles, macros, drawings, etc.
    - Cleans stale merges/validations/CF/hyperlinks/page breaks and syncs tables to avoid Excel repair prompts.

    **Tip**
    - If your onboarding files are huge, convert to CSV and upload — with **pyarrow** this is the fastest path.
    """))
