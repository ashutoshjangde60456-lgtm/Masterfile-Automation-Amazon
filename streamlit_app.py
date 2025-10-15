import io, json, re, time, zipfile, xml.etree.ElementTree as ET
from pathlib import Path
import pandas as pd, streamlit as st
from openpyxl import load_workbook

# ─────────────────────────────────────────────────────────────────────
# Page config
# ─────────────────────────────────────────────────────────────────────
st.set_page_config(page_title="Masterfile Automation – Hybrid Smart Writer", page_icon="🧾", layout="wide")
st.markdown("""
<style>
.section{border:1px solid #e8eef6;background:#fff;border-radius:16px;padding:18px;margin:12px 0;box-shadow:0 6px 24px rgba(2,6,23,.05)}
div.stButton>button,.stDownloadButton>button{background:#2563eb!important;color:#fff!important;border-radius:10px!important;border:0!important}
</style>
""", unsafe_allow_html=True)

# ─────────────────────────────────────────────────────────────────────
# Constants
# ─────────────────────────────────────────────────────────────────────
MASTER_TEMPLATE_SHEET = "Template"
MASTER_DISPLAY_ROW    = 2
MASTER_SECONDARY_ROW  = 3
MASTER_DATA_START_ROW = 4
XL_NS_MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
XL_NS_REL  = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
ET.register_namespace("", XL_NS_MAIN)
ET.register_namespace("r", XL_NS_REL)
ET.register_namespace("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006")
ET.register_namespace("x14ac", "http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac")
_INVALID_XML_CHARS = re.compile(r"[\x00-\x08\x0B\x0C\x0E-\x1F\uD800-\uDFFF]")

# ─────────────────────────────────────────────────────────────────────
# Helper utilities
# ─────────────────────────────────────────────────────────────────────
def sanitize_xml_text(s):
    if s is None: return ""
    return _INVALID_XML_CHARS.sub("", str(s))

def _col_letter(n):
    s = ""
    while n:
        n, r = divmod(n-1, 26)
        s = chr(65+r) + s
    return s

def norm(s):
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

# ─────────────────────────────────────────────────────────────────────
# XML fast patch (lightweight version)
# ─────────────────────────────────────────────────────────────────────
def fast_patch_template(master_bytes, sheet_name, header_row, start_row, used_cols, block_2d):
    zin = zipfile.ZipFile(io.BytesIO(master_bytes), "r")
    # find target sheet xml path
    wb_xml = ET.fromstring(zin.read("xl/workbook.xml"))
    rels_xml = ET.fromstring(zin.read("xl/_rels/workbook.xml.rels"))
    rid = None
    for sh in wb_xml.find(f"{{{XL_NS_MAIN}}}sheets"):
        if sh.attrib.get("name") == sheet_name:
            rid = sh.attrib.get(f"{{{XL_NS_REL}}}id"); break
    if not rid:
        raise ValueError("Sheet not found")
    target = None
    for rel in rels_xml:
        if rel.attrib.get("Id") == rid:
            target = rel.attrib.get("Target"); break
    target = target.replace("\\","/")
    if target.startswith("../"): target = target[3:]
    if not target.startswith("xl/"): target = "xl/"+target
    sheet_path = target

    # patch rows
    sheet_xml = zin.read(sheet_path)
    root = ET.fromstring(sheet_xml)
    ns = XL_NS_MAIN
    sheetData = root.find(f"{{{ns}}}sheetData")
    if sheetData is None:
        sheetData = ET.SubElement(root, f"{{{ns}}}sheetData")
    # delete old rows
    for r in list(sheetData):
        try: rnum = int(r.attrib.get("r","0"))
        except Exception: rnum = 0
        if rnum >= start_row: sheetData.remove(r)
    # add new rows
    for i,row_vals in enumerate(block_2d):
        r = start_row + i
        row_el = ET.Element(f"{{{ns}}}row", r=str(r))
        for j,v in enumerate(row_vals[:used_cols]):
            if not v: continue
            col = _col_letter(j+1)
            c = ET.Element(f"{{{ns}}}c", r=f"{col}{r}", t="inlineStr")
            is_el = ET.SubElement(c,f"{{{ns}}}is")
            t_el = ET.SubElement(is_el,f"{{{ns}}}t")
            t_el.set("{http://www.w3.org/XML/1998/namespace}space","preserve")
            t_el.text = sanitize_xml_text(v)
            row_el.append(c)
        sheetData.append(row_el)

    # replace sheet xml and drop calcChain
    new_sheet = ET.tostring(root, encoding="utf-8", xml_declaration=True)
    out_bio = io.BytesIO()
    with zipfile.ZipFile(out_bio, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename == sheet_path:
                data = new_sheet
            if item.filename == "xl/calcChain.xml":
                continue
            zout.writestr(item, data)
    zin.close(); out_bio.seek(0)
    return out_bio.getvalue()

# ─────────────────────────────────────────────────────────────────────
# Validation + fallback
# ─────────────────────────────────────────────────────────────────────
def validate_excel_zip(xlsx_bytes: bytes) -> bool:
    """Light validation: ensure required parts exist and parse cleanly."""
    try:
        z = zipfile.ZipFile(io.BytesIO(xlsx_bytes))
        if "[Content_Types].xml" not in z.namelist() or "xl/workbook.xml" not in z.namelist():
            return False
        ET.fromstring(z.read("xl/workbook.xml"))
        for name in z.namelist():
            if name.startswith("xl/worksheets/sheet") and name.endswith(".xml"):
                ET.fromstring(z.read(name))
        return True
    except Exception:
        return False

def safe_openpyxl_write(master_bytes, sheet_name, header_row, start_row, used_cols, block_2d):
    wb = load_workbook(io.BytesIO(master_bytes), keep_vba=True)
    ws = wb[sheet_name]
    maxr = ws.max_row or start_row
    if maxr >= start_row:
        ws.delete_rows(start_row, maxr - start_row + 1)
    for row in block_2d:
        ws.append(row[:used_cols])
    out = io.BytesIO(); wb.save(out); out.seek(0)
    return out.getvalue()

# ─────────────────────────────────────────────────────────────────────
# UI
# ─────────────────────────────────────────────────────────────────────
st.title("🚦 Hybrid Smart Writer – Fast + Safe")
st.caption("XML Fast patch with auto fallback to OpenPyXL if any issue detected.")

with st.container():
    c1,c2 = st.columns([1,1])
    with c1:
        masterfile_file = st.file_uploader("📄 Masterfile Template (.xlsx / .xlsm)", type=["xlsx","xlsm"])
    with c2:
        onboarding_file = st.file_uploader("🧾 Onboarding (.xlsx)", type=["xlsx"])

st.markdown("#### 🔗 Mapping JSON")
tab1,tab2 = st.tabs(["Paste JSON","Upload JSON"])
mapping_json_text,mapping_json_file="",""
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
            # parse mapping
            status.update(label="Parsing mapping JSON…")
            mapping_raw = json.loads(mapping_json_text) if mapping_json_text.strip() else json.load(mapping_json_file)
            mapping_aliases = {norm(k): (v if isinstance(v,list) else [v]) + [k] for k,v in mapping_raw.items()}

            # read headers
            status.update(label="Reading template headers…")
            masterfile_file.seek(0); master_bytes = masterfile_file.read()
            wb_ro = load_workbook(io.BytesIO(master_bytes), read_only=True, data_only=True, keep_links=True)
            ws = wb_ro[MASTER_TEMPLATE_SHEET]
            used_cols = ws.max_column
            headers = [ws.cell(row=MASTER_DISPLAY_ROW, column=c).value or "" for c in range(1,used_cols+1)]
            wb_ro.close(); status.write(f"✅ Template headers loaded ({used_cols} cols)")

            # onboarding
            status.update(label="Reading onboarding…")
            onboarding_file.seek(0)
            df = pd.read_excel(onboarding_file, engine="openpyxl").fillna("")
            on_headers = list(df.columns)
            status.write(f"✅ Onboarding sheet read ({len(df)} rows)")

            # build block
            block = df.astype(str).fillna("").values.tolist()

            # ⚡ Hybrid write
            status.update(label="Writing via Hybrid Smart Writer…")
            t0 = time.time()
            xml_bytes = fast_patch_template(
                master_bytes, MASTER_TEMPLATE_SHEET,
                MASTER_DISPLAY_ROW, MASTER_DATA_START_ROW,
                used_cols, block)
            valid = validate_excel_zip(xml_bytes)
            if valid:
                out_bytes = xml_bytes
                status.write(f"✅ Fast XML patch succeeded in {time.time()-t0:.2f}s (validated clean)")
            else:
                status.write("⚠️ XML output failed validation – switching to safe OpenPyXL mode…")
                t1 = time.time()
                out_bytes = safe_openpyxl_write(
                    master_bytes, MASTER_TEMPLATE_SHEET,
                    MASTER_DISPLAY_ROW, MASTER_DATA_START_ROW,
                    used_cols, block)
                status.write(f"✅ Safe OpenPyXL rewrite finished in {time.time()-t1:.2f}s")
            status.update(label="Finished", state="complete")

            # download
            ext = Path(masterfile_file.name).suffix.lower()
            mime = "application/vnd.ms-excel.sheet.macroEnabled.12" if ext==".xlsm" else \
                   "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            download_placeholder.download_button(
                "⬇️ Download Final Masterfile",
                data=out_bytes,
                file_name=f"final_masterfile{ext}",
                mime=mime)
        except Exception as e:
            status.update(label="Error", state="error")
            st.exception(e)
