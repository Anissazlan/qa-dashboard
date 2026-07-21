import os
import streamlit as st
import pandas as pd
from datetime import datetime
from PIL import Image
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.drawing.image import Image as OpenpyxlImage
from openpyxl.drawing.spreadsheet_drawing import TwoCellAnchor, AnchorMarker

# ==========================================================
# PAGE CONFIGURATION
# ==========================================================
st.set_page_config(
    page_title="New Incoming Check In",
    page_icon="📦",
    layout="centered"
)

# Custom Styling (Warm Background, Bold Labels, Custom Title)
st.markdown("""
    <style>
    /* Warm Soft Background for Light Mode */
    .stApp {
        background-color: #FAF6EE !important; /* Soft warm ivory / cream color */
    }
    
    /* Input Box Backgrounds to match cleanly */
    div[data-baseweb="input"] > div, 
    div[data-baseweb="select"] > div,
    textarea {
        background-color: #FFFFFF !important;
    }

    /* Force input labels to be Bold and Dark Blue */
    div[data-widget="text_input"] label, 
    div[data-widget="number_input"] label, 
    div[data-widget="selectbox"] label, 
    div[data-widget="date_input"] label, 
    div[data-widget="file_uploader"] label,
    div[data-widget="textarea"] label {
        font-weight: bold !important;
        font-size: 16px !important;
        color: #1F4E79 !important;
    }
    
    /* Main Heading Styling */
    .main-title {
        color: #1F4E79;
        font-weight: bold;
        font-size: 32px;
        margin-bottom: 0px;
        text-align: center;
    }
    .sub-title {
        color: #555555;
        font-style: italic;
        font-size: 14px;
        margin-bottom: 20px;
        text-align: center;
    }
    </style>
""", unsafe_allow_html=True)

EXCEL_FILE = "New_Incoming_Check_In.xlsx"
LISTS_SHEET_NAME = "Lists"

DEFAULT_SUPPLIER = [
    "AVNET ASIA", "DIGIKEY ELECTRONICS", "ELEMENT 14", "FUTURE ELECTRONICS",
    "MOUSER ELECTRONICS", "GPV ASIA THAILAND", "HONG KONG YIHAO", "PHYTECH",
    "WIN SOURCE ELECTRONICS", "HONENG ELEC", "UNI BETTER", "LCSC",
    "NELSON MILLER GROUP", "E-STAR", "WURTH ELEKTRONIK"
]

DEFAULT_COMMODITY = [
    "PCB", "CAPACITOR", "RESISTOR", "DIODE", "FERRITE", "TRANSISTOR", "IC"
]

HEADERS = [
    "Store Receive Date", "IQA Receive Date", "Supplier", 
    "MPN No.", "Part No.", "DC/Lot No.", "Quantity", 
    "Commodity", "Status", "Picture", "Remark"
]

# ==========================================================
# EXCEL HELPER FUNCTIONS
# ==========================================================
def apply_excel_formatting(ws):
    thin_side = Side(style="thin", color="000000")
    cell_border = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)
    
    ws.row_dimensions[1].height = 25
    for col in range(1, len(HEADERS) + 1):
        cell = ws.cell(row=1, column=col)
        cell.font = Font(name="Calibri", size=11, bold=True)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = cell_border
        ws.column_dimensions[get_column_letter(col)].width = 22

    data_font = Font(name="Calibri", size=11)
    for row in range(2, ws.max_row + 1):
        ws.row_dimensions[row].height = 65
        for col in range(1, len(HEADERS) + 1):
            cell = ws.cell(row=row, column=col)
            cell.font = data_font
            cell.border = cell_border
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

def initialize_excel_database():
    if not os.path.exists(EXCEL_FILE):
        wb = Workbook()
        wb.remove(wb.active)
        
        ws_lists = wb.create_sheet(title=LISTS_SHEET_NAME)
        ws_lists.cell(row=1, column=1, value="Supplier").font = Font(name="Calibri", size=11, bold=True)
        ws_lists.cell(row=1, column=2, value="Commodity").font = Font(name="Calibri", size=11, bold=True)
        
        for idx, item in enumerate(DEFAULT_SUPPLIER, start=2):
            ws_lists.cell(row=idx, column=1, value=item)
            
        for idx, item in enumerate(DEFAULT_COMMODITY, start=2):
            ws_lists.cell(row=idx, column=2, value=item)

        wb.save(EXCEL_FILE)
        wb.close()

def load_lists_from_excel():
    initialize_excel_database()
    suppliers = []
    commodities = []
    
    try:
        wb = load_workbook(EXCEL_FILE, data_only=True)
        if LISTS_SHEET_NAME in wb.sheetnames:
            ws = wb[LISTS_SHEET_NAME]
            for row in range(2, ws.max_row + 1):
                val = ws.cell(row=row, column=1).value
                if val and str(val).strip() not in suppliers:
                    suppliers.append(str(val).strip())
            for row in range(2, ws.max_row + 1):
                val = ws.cell(row=row, column=2).value
                if val and str(val).strip() not in commodities:
                    commodities.append(str(val).strip())
        wb.close()
    except Exception as e:
        st.error(f"Error loading lists: {e}")

    if not suppliers:
        suppliers = DEFAULT_SUPPLIER.copy()
    if not commodities:
        commodities = DEFAULT_COMMODITY.copy()

    return suppliers, commodities

def save_list_to_excel(column_idx, items):
    wb = load_workbook(EXCEL_FILE)
    if LISTS_SHEET_NAME not in wb.sheetnames:
        ws = wb.create_sheet(title=LISTS_SHEET_NAME)
    else:
        ws = wb[LISTS_SHEET_NAME]
    
    for r in range(2, ws.max_row + 2):
        ws.cell(row=r, column=column_idx).value = None
        
    for idx, item in enumerate(items, start=2):
        ws.cell(row=idx, column=column_idx, value=item)
        
    wb.save(EXCEL_FILE)
    wb.close()

if "suppliers" not in st.session_state or "commodities" not in st.session_state:
    s_list, c_list = load_lists_from_excel()
    st.session_state.suppliers = s_list
    st.session_state.commodities = c_list

# ==========================================================
# UI LAYOUT
# ==========================================================

# Display Logo centered if it exists
if os.path.exists("Kaltech Logo.png"):
    l_col1, l_col2, l_col3 = st.columns([1, 1, 1])
    with l_col2:
        st.image("Kaltech Logo.png", width=150)

# Styled Colored Title & Subtitle
st.markdown('<div class="main-title">NEW INCOMING CHECK IN</div>', unsafe_allow_html=True)
st.markdown('<div class="sub-title">Incoming Material Inspection System</div>', unsafe_allow_html=True)

st.divider()

# Form Inputs
col1, col2 = st.columns(2)

with col1:
    store_date = st.date_input("Store Receive Date", value=datetime.now())
with col2:
    iqa_date = st.date_input("IQA Receive Date", value=datetime.now())

# Supplier Section
sup_col1, sup_col2 = st.columns([3, 1])
with sup_col1:
    selected_supplier = st.selectbox("Supplier", [""] + st.session_state.suppliers)
with sup_col2:
    st.write(" ")
    st.write(" ")
    if st.popover("⚙️ Manage"):
        new_sup = st.text_input("Add Supplier").strip().upper()
        if st.button("Add Supplier") and new_sup:
            if new_sup not in st.session_state.suppliers:
                st.session_state.suppliers.append(new_sup)
                save_list_to_excel(1, st.session_state.suppliers)
                st.success(f"Added {new_sup}")
                st.rerun()
        
        if selected_supplier and st.button("Delete Selected Supplier"):
            st.session_state.suppliers.remove(selected_supplier)
            save_list_to_excel(1, st.session_state.suppliers)
            st.success(f"Deleted {selected_supplier}")
            st.rerun()

mpn = st.text_input("MPN No.")
part_no = st.text_input("Part No.")
lot_no = st.text_input("DC/Lot No.")
quantity = st.number_input("Quantity", min_value=1, step=1, value=1)

# Commodity Section
com_col1, com_col2 = st.columns([3, 1])
with com_col1:
    selected_commodity = st.selectbox("Commodity", [""] + st.session_state.commodities)
with com_col2:
    st.write(" ")
    st.write(" ")
    if st.popover("⚙️ Manage "):
        new_com = st.text_input("Add Commodity").strip().upper()
        if st.button("Add Commodity") and new_com:
            if new_com not in st.session_state.commodities:
                st.session_state.commodities.append(new_com)
                save_list_to_excel(2, st.session_state.commodities)
                st.success(f"Added {new_com}")
                st.rerun()
                
        if selected_commodity and st.button("Delete Selected Commodity"):
            st.session_state.commodities.remove(selected_commodity)
            save_list_to_excel(2, st.session_state.commodities)
            st.success(f"Deleted {selected_commodity}")
            st.rerun()

status = st.selectbox("Status", ["", "ACCEPT", "REJECT", "WAIVER", "QA PASS", "ON HOLD"])
remark = st.text_area("Remark", height=70)

uploaded_image = st.file_uploader("📷 Upload Picture", type=["png", "jpg", "jpeg", "bmp"])
if uploaded_image:
    st.image(uploaded_image, caption="Preview Image", width=150)

# ==========================================================
# SAVE DATA PIPELINE
# ==========================================================
st.divider()

if st.button("📥 CHECK IN", type="primary", use_container_width=True):
    if not selected_supplier or not mpn or not part_no or not status:
        st.warning("⚠️ Please fill in all required fields (Supplier, MPN, Part No, Status).")
    else:
        try:
            sheet_name = store_date.strftime("%b %Y")
            temp_img_path = None

            if uploaded_image:
                temp_img_path = f"temp_{uploaded_image.name}"
                with open(temp_img_path, "wb") as f:
                    f.write(uploaded_image.getbuffer())

            initialize_excel_database()
            wb = load_workbook(EXCEL_FILE)

            if sheet_name not in wb.sheetnames:
                ws = wb.create_sheet(title=sheet_name, index=0)
                for col, text in enumerate(HEADERS, start=1):
                    ws.cell(row=1, column=col).value = text
                ws.auto_filter.ref = f"A1:{get_column_letter(len(HEADERS))}1"
            else:
                ws = wb[sheet_name]

            row = ws.max_row + 1

            data = [
                store_date.strftime("%d/%m/%Y"),
                iqa_date.strftime("%d/%m/%Y"),
                selected_supplier,
                mpn.upper(),
                part_no.upper(),
                lot_no.upper(),
                quantity,
                selected_commodity,
                status,
                "",
                remark
            ]

            for col, value in enumerate(data, start=1):
                ws.cell(row=row, column=col, value=value)

            colors = {
                "ACCEPT": "009900", "REJECT": "FF0000", "WAIVER": "00B0F0",
                "QA PASS": "66FF33", "ON HOLD": "FFFF00"
            }
            if status in colors:
                ws.cell(row=row, column=9).fill = PatternFill("solid", fgColor=colors[status])

            if temp_img_path and os.path.exists(temp_img_path):
                try:
                    img_to_insert = OpenpyxlImage(temp_img_path)
                    col_idx = 9
                    row_idx = row - 1

                    cell_w_px = 154
                    cell_h_px = 86

                    max_w, max_h = 135, 72
                    w, h = img_to_insert.width, img_to_insert.height
                    ratio = min(max_w / float(w), max_h / float(h))
                    new_w = int(w * ratio)
                    new_h = int(h * ratio)

                    EMU_PER_PX = 9525
                    offset_x_emu = int((cell_w_px - new_w) / 2) * EMU_PER_PX
                    offset_y_emu = int((cell_h_px - new_h) / 2) * EMU_PER_PX

                    marker_from = AnchorMarker(
                        col=col_idx, colOff=offset_x_emu,
                        row=row_idx, rowOff=offset_y_emu
                    )
                    marker_to = AnchorMarker(
                        col=col_idx, colOff=offset_x_emu + (new_w * EMU_PER_PX),
                        row=row_idx, rowOff=offset_y_emu + (new_h * EMU_PER_PX)
                    )

                    img_to_insert.anchor = TwoCellAnchor(_from=marker_from, to=marker_to, editAs="oneCell")
                    ws.add_image(img_to_insert)
                except Exception as img_err:
                    st.error(f"Image error: {img_err}")

            apply_excel_formatting(ws)
            wb.save(EXCEL_FILE)
            wb.close()

            if temp_img_path and os.path.exists(temp_img_path):
                os.remove(temp_img_path)

            st.success(f"✅ Material checked in successfully under tab [{sheet_name}]!")

        except Exception as e:
            st.error(f"Error saving to Excel: {e}")

# Download Link
if os.path.exists(EXCEL_FILE):
    with open(EXCEL_FILE, "rb") as file:
        st.download_button(
            label="📊 Download Excel Log",
            data=file,
            file_name=EXCEL_FILE,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )
