import io
import json
import re
import time
import zipfile
import xml.etree.ElementTree as ET
from difflib import SequenceMatcher
from pathlib import Path
import pandas as pd
import streamlit as st
from openpyxl import load_workbook

# ─────────────────────────────────────────────────────────────────────
# Constants and XML Namespaces
# ─────────────────────────────────────────────────────────────────────
XL_NS_MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
XL_NS_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
ET.register_namespace("", XL_NS_MAIN)
ET.register_namespace("r", XL_NS_REL)

_INVALID_XML_CHARS = re.compile(r"[\x00-\x08\x0B\x0C\x0E-\x1F\uD800-\uDFFF]")

def sanitize_xml_text(s) -> str:
    if s is None:
        return ""
    s = str(s)
    return _INVALID_XML_CHARS.sub("", s)

def _col_letter(n: int) -> str:
    s = ""
    while n:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s

def _find_sheet_part_path(z: zipfile.ZipFile, sheet_name: str) -> str:
    wb_xml = ET.fromstring(z.read("xl/workbook.xml"))
    for sh in wb_xml.find(f"{{{XL_NS_MAIN}}}sheets"):
        if sh.attrib.get("name") == sheet_name:
            rid = sh.attrib.get(f"{{{XL_NS_REL}}}id")
            break
    else:
        raise ValueError(f"Sheet '{sheet_name}' not found in workbook.xml")

    rels_xml = ET.fromstring(z.read("xl/_rels/workbook.xml.rels"))
    for rel in rels_xml:
        if rel.attrib.get("Id") == rid:
            target = rel.attrib.get("Target").replace("\\", "/")
            return target if target.startswith("xl/") else "xl/" + target
    raise ValueError(f"Relationship for sheet '{sheet_name}' not found.")

def _patch_sheet_xml(sheet_xml_bytes: bytes, start_row: int, used_cols_final: int, block_2d: list) -> bytes:
    root = ET.fromstring(sheet_xml_bytes)
    sheetData = root.find(f"{{{XL_NS_MAIN}}}sheetData")
    if sheetData is None:
        sheetData = ET.SubElement(root, f"{{{XL_NS_MAIN}}}sheetData")

    # Clear existing rows
    for row in list(sheetData):
        sheetData.remove(row)

    # Append new rows
    for i, row_vals in enumerate(block_2d):
        r = start_row + i
        row_el = ET.Element(f"{{{XL_NS_MAIN}}}row", r=str(r))
        for j in range(used_cols_final):
            v = row_vals[j] if j < len(row_vals) else ""
            if v:
                txt = sanitize_xml_text(v)
                col = _col_letter(j + 1)
                c = ET.Element(f"{{{XL_NS_MAIN}}}c", r=f"{col}{r}", t="inlineStr")
                is_el = ET.SubElement(c, f"{{{XL_NS_MAIN}}}is")
                t_el = ET.SubElement(is_el, f"{{{XL_NS_MAIN}}}t")
                t_el.text = txt
                row_el.append(c)
        sheetData.append(row_el)

    return ET.tostring(root, encoding="utf-8", xml_declaration=True)

def fast_patch_template(master_bytes: bytes, sheet_name: str, start_row: int, used_cols: int, block_2d: list) -> bytes:
    zin = zipfile.ZipFile(io.BytesIO(master_bytes), "r")
    sheet_path = _find_sheet_part_path(zin, sheet_name)
    original_sheet_xml = zin.read(sheet_path)
    new_sheet_xml = _patch_sheet_xml(original_sheet_xml, start_row, used_cols, block_2d)

    out_bio = io.BytesIO()
    with zipfile.ZipFile(out_bio, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename == sheet_path:
                data = new_sheet_xml
            zout.writestr(item, data)
    zin.close()
    out_bio.seek(0)
    return out_bio.getvalue()

# ─────────────────────────────────────────────────────────────────────
# Streamlit UI
# ─────────────────────────────────────────────────────────────────────
st.title("Masterfile Automation – Amazon")
st.caption("Fills **only** the Template sheet and preserves all other sheets/styles.")

# File Uploads
masterfile_file = st.file_uploader("📄 Masterfile Template (.xlsx / .xlsm)", type=["xlsx", "xlsm"])
onboarding_file = st.file_uploader("🧾 Onboarding (.xlsx)", type=["xlsx"])

if st.button("🚀 Generate Final Masterfile"):
    if not masterfile_file or not onboarding_file:
        st.error("Please upload both **Masterfile Template** and **Onboarding**.")
    else:
        try:
            # Read master file
            master_bytes = masterfile_file.read()
            # Load workbook to get headers
            wb_ro = load_workbook(io.BytesIO(master_bytes), read_only=True, data_only=True)
            sheet_name = "Template"  # Change this if your sheet name is different
            if sheet_name not in wb_ro.sheetnames:
                st.error(f"Sheet **'{sheet_name}'** not found in the masterfile.")
            else:
                ws_ro = wb_ro[sheet_name]
                used_cols = ws_ro.max_column  # Get the number of used columns
                block_2d = [["Sample Data"] * used_cols]  # Replace with actual data processing logic

                # Generate final file
                out_bytes = fast_patch_template(master_bytes, sheet_name, start_row=4, used_cols=used_cols, block_2d=block_2d)
                st.download_button("⬇️ Download Final Masterfile", data=out_bytes, file_name="final_masterfile.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        except Exception as e:
            st.error(f"An error occurred: {e}")
