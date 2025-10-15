import io, json, re, time, zipfile, xml.etree.ElementTree as ET
from pathlib import Path
import pandas as pd, streamlit as st
from openpyxl import load_workbook

# ───────────────────────────────
# Basic constants
# ───────────────────────────────
MASTER_TEMPLATE_SHEET = "Template"
MASTER_DISPLAY_ROW    = 2
MASTER_SECONDARY_ROW  = 3
MASTER_DATA_START_ROW = 4
XL_NS_MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
XL_NS_REL  = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
ET.register_namespace("", XL_NS_MAIN)
ET.register_namespace("r", XL_NS_REL)

# ───────────────────────────────
# Utilities (shared with XML fast writer)
# ───────────────────────────────
def sanitize_xml_text(s):
    if s is None: return ""
    return re.sub(r"[\x00-\x08\x0B\x0C\x0E-\x1F\uD800-\uDFFF]", "", str(s))

def _col_letter(n):
    s = ""
    while n:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s

def norm(s):  # simplified normalization
    if s is None: return ""
    x = str(s).strip().lower()
    x = re.sub(r"[._/\\-]+", " ", x)
    return re.sub(r"[^0-9a-z\s]+", " ", x).strip()

# ───────────────────────────────
# XML fast writer (shortened)
# ───────────────────────────────
def xml_fast_write(master_bytes, sheet_name, header_row, start_row, used_cols, block_2d):
    """Simplified XML patch: rewrite only the Template sheet."""
    zin = zipfile.ZipFile(io.BytesIO(master_bytes), "r")
    # find sheet
    wb_xml = ET.fromstring(zin.read("xl/workbook.xml"))
    rels_xml = ET.fromstring(zin.read("xl/_rels/workbook.xml.rels"))
    rid = None
    for sh in wb_xml.find(f"{{{XL_NS_MAIN}}}sheets"):
        if sh.attrib.get("name") == sheet_name:
            rid = sh.attrib.get(f"{{{XL_NS_REL}}}id")
            break
    if not rid:
        raise ValueError("Sheet not found")
    target = None
    for rel in rels_xml:
        if rel.attrib.get("Id") == rid:
            target = rel.attrib.get("Target"); break
    target = target.replace("\\","/"); 
    if target.startswith("../"): target = target[3:]
    if not target.startswith("xl/"): target = "xl/" + target
    sheet_path = target

    sheet_xml = zin.read(sheet_path)
    root = ET.fromstring(sheet_xml)
    ns = XL_NS_MAIN
    sheetData = root.find(f"{{{ns}}}sheetData")
    if sheetData is None:
        sheetData = ET.SubElement(root, f"{{{ns}}}sheetData")
    for r in list(sheetData):
        rnum = int(r.attrib.get("r","0"))
        if rnum >= start_row: sheetData.remove(r)
    for i,row in enumerate(block_2d):
        rr = start_row + i
        row_el = ET.Element(f"{{{ns}}}row", r=str(rr))
        for j,v in enumerate(row[:used_cols]):
            if not v: continue
            c = ET.Element(f"{{{ns}}}c", r=f"{_col_letter(j+1)}{rr}", t="inlineStr")
            is_el = ET.SubElement(c,f"{{{ns}}}is")
            t_el = ET.SubElement(is_el,f"{{{ns}}}t")
            t_el.set("{http://www.w3.org/XML/1998/namespace}space","preserve")
            t_el.text = sanitize_xml_text(v)
            row_el.append(c)
        sheetData.append(row_el)
    new_xml = ET.tostring(root, encoding="utf-8", xml_declaration=True)
    out_bio = io.BytesIO()
    with zipfile.ZipFile(out_bio, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename == sheet_path:
                data = new_xml
            if item.filename == "xl/calcChain.xml":
                continue
            zout.writestr(item, data)
    zin.close(); out_bio.seek(0)
    return out_bio.getvalue()

# ───────────────────────────────
# Validation (simple structural)
# ───────────────────────────────
def validate_excel_zip(xlsx_bytes):
    try:
        z = zipfile.ZipFile(io.BytesIO(xlsx_bytes))
        must_have = ["[Content_Types].xml","xl/workbook.xml"]
        for m in must_have:
            if m not in z.namelist(): return False
        # try parse workbook
        ET.fromstring(z.read("xl/workbook.xml"))
        return True
    except Exception:
        return False

# ───────────────────────────────
# Safe fallback: openpyxl writer
# ───────────────────────────────
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

# ───────────────────────────────
# Streamlit UI
# ───────────────────────────────
st.title("🚦 Hybrid Smart Writer – Fast + Safe")
st.caption("Tries XML Fast Write; falls back to OpenPyXL if invalid.")
tmpl = st.file_uploader("📄 Masterfile Template", type=["xlsx","xlsm"])
data = st.file_uploader("🧾 Onboarding", type=["xlsx"])
go = st.button("Generate")

if go and tmpl and data:
    st.info("Reading files…")
    tbytes = tmpl.read(); df = pd.read_excel(data)
    used_cols = len(df.columns)
    block = df.astype(str).fillna("").values.tolist()
    st.info("⚡ Trying XML fast path…")
    xml_bytes = xml_fast_write(tbytes, MASTER_TEMPLATE_SHEET,
                               MASTER_DISPLAY_ROW, MASTER_DATA_START_ROW,
                               used_cols, block)
    if validate_excel_zip(xml_bytes):
        out_bytes = xml_bytes
        st.success("✅ XML fast path successful (no repair expected).")
    else:
        st.warning("⚠️ XML fast path invalid – using safe OpenPyXL path.")
        out_bytes = safe_openpyxl_write(tbytes, MASTER_TEMPLATE_SHEET,
                                        MASTER_DISPLAY_ROW, MASTER_DATA_START_ROW,
                                        used_cols, block)
    st.download_button("⬇️ Download Final File",
                       data=out_bytes,
                       file_name="final_masterfile.xlsx",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
