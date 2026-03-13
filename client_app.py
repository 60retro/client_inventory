import sys
from types import ModuleType

# --- 🛠️ ส่วนแก้บั๊ก Python 3.13 (Mock imghdr module) ---
if sys.version_info >= (3, 13):
    m = ModuleType("imghdr")
    m.what = lambda *args: None  
    sys.modules["imghdr"] = m
# ----------------------------------------------------

import streamlit as st
import pandas as pd
import gspread
from google.oauth2.service_account import Credentials
import time
import os

# --- Config ---
SHEET_NAME = "inventory_data"
CREDENTIALS_FILE = "credentials.json"

# ตั้งค่าหน้าเว็บ
st.set_page_config(page_title="Nami Stock Client", page_icon="📱")

# --- 1. ระบบภาษา (Translation System) ---
TRANSLATIONS = {
    "th": {
        "title": "📱 Nami Stock Check",
        "caption": "ระบบจัดการออเดอร์และรับสินค้า (Procurement System)",
        "select_category": "📂 เลือกหมวดหมู่",
        "mode_select": "สลับโหมดการทำงาน:",
        "mode_order": "📝 1. ตรวจนับและสั่งของ",
        "mode_receive": "📦 2. ตรวจรับสินค้า",
        "no_items": "⚠️ ไม่มีสินค้าในหมวดหมู่นี้",
        "no_pending": "🎉 ไม่มีรายการรอรับสินค้าในหมวดนี้",
        "instruction_order": "📝 กรอกยอด **'คงเหลือ'** หรือ **'สั่งเพิ่ม'**",
        "instruction_receive": "🔎 ตรวจสอบยอดที่สั่งเทียบกับ **'ของที่มาส่งจริง'**",
        "col_name": "รายการ",
        "col_remain": "📦 คงเหลือ",
        "col_order": "🛒 สั่งเพิ่ม",
        "col_ordered": "📋 ยอดที่สั่ง",
        "col_actual": "✅ รับจริง",
        "submit_btn": "🚀 ส่งข้อมูล (Submit)",
        "no_changes": "⚠️ ไม่มีการแก้ไขข้อมูล",
        "sending": "กำลังส่งข้อมูล...",
        "success": "✅ ส่งข้อมูลเรียบร้อย!",
        "error": "❌ เกิดข้อผิดพลาด: ",
        "conn_error": "❌ ไม่สามารถเชื่อมต่อ Google API ได้",
        "sheet_error": "❌ หาไฟล์ Google Sheet ไม่เจอ: "
    },
    "en": {
        "title": "📱 Nami Stock Check",
        "caption": "Order & Receive Management System",
        "select_category": "📂 Select Category",
        "mode_select": "Select Mode:",
        "mode_order": "📝 1. Count & Order",
        "mode_receive": "📦 2. Receive Goods",
        "no_items": "⚠️ No items found",
        "no_pending": "🎉 No pending orders to receive",
        "instruction_order": "📝 Enter **'Remaining'** or **'Order'**",
        "instruction_receive": "🔎 Verify ordered vs **'Actual Received'**",
        "col_name": "Item Name",
        "col_remain": "📦 Remaining",
        "col_order": "🛒 Order Qty",
        "col_ordered": "📋 Ordered",
        "col_actual": "✅ Actual Recv",
        "submit_btn": "🚀 Submit Data",
        "no_changes": "⚠️ No changes detected",
        "sending": "Sending data...",
        "success": "✅ Data sent successfully!",
        "error": "❌ Error occurred: ",
        "conn_error": "❌ Cannot connect to Google API",
        "sheet_error": "❌ Google Sheet not found: "
    }
}

st.sidebar.title("Language / ภาษา")
lang_option = st.sidebar.radio("Select Language:", ("ภาษาไทย (Thai)", "English"))
current_lang = "th" if "Thai" in lang_option else "en"

def t(key): return TRANSLATIONS[current_lang][key]

# --- 2. Function เชื่อมต่อ Google Sheet ---
@st.cache_resource
def get_google_sheet_client():
    scopes = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]
    try:
        if "gcp_json" in st.secrets:
            info = st.secrets["gcp_json"]
            creds = Credentials.from_service_account_info(info, scopes=scopes)
        elif os.path.exists(CREDENTIALS_FILE):
            creds = Credentials.from_service_account_file(CREDENTIALS_FILE, scopes=scopes)
        else:
            return None
        return gspread.authorize(creds)
    except Exception as e:
        st.error(f"Connect Error: {e}")
        return None

# --- 3. Main App Logic ---
st.title(t("title"))
st.caption(t("caption"))

# 🟢 เพิ่ม Radio Button สำหรับเลือกโหมด
app_mode = st.radio(t("mode_select"), [t("mode_order"), t("mode_receive")], horizontal=True)
st.markdown("---")

client = get_google_sheet_client()
if not client:
    st.error(t("conn_error"))
    st.stop()

try: sh = client.open(SHEET_NAME)
except gspread.exceptions.SpreadsheetNotFound:
    st.error(f"{t('sheet_error')} '{SHEET_NAME}'")
    st.stop()

# Load Data
all_worksheets = [ws.title for ws in sh.worksheets()]
selected_tab = st.selectbox(t("select_category"), all_worksheets)

if selected_tab:
    ws = sh.worksheet(selected_tab)
    try:
        data = ws.get_all_records()
        df = pd.DataFrame(data)
    except Exception as e:
        st.error(f"Load Error: {e}")
        st.stop()

    if df.empty:
        st.warning(t("no_items"))
        st.stop()

    with st.form("stock_entry_form"):
        updates = {} 
        
        # ==========================================
        # 📝 โหมด 1: ตรวจนับและสั่งของ (Order Mode)
        # ==========================================
        if app_mode == t("mode_order"):
            st.info(t("instruction_order"))
            for i, row in df.iterrows():
                st.markdown("---") 
                cols = st.columns([3, 1.5, 1.5])
                cols[0].markdown(f"**{row['Name']}**")
                
                try: curr_val = int(row['Current']) if str(row['Current']).strip() != '' else 0
                except: curr_val = 0
                try: order_val = int(row['Order']) if str(row['Order']).strip() != '' else 0
                except: order_val = 0
                
                new_curr = cols[1].number_input(t("col_remain"), min_value=0, value=curr_val, key=f"c_{i}")
                new_order = cols[2].number_input(t("col_order"), min_value=0, value=order_val, key=f"o_{i}")
                
                if new_curr != curr_val or new_order != order_val:
                    updates[i + 2] = {"Current": new_curr, "Order": new_order, "Status": "Order_Submitted"}

        # ==========================================
        # 📦 โหมด 2: ตรวจรับสินค้า (Receive Mode)
        # ==========================================
        else:
            st.info(t("instruction_receive"))
            
            # กรองเฉพาะรายการที่สั่งไป (Order > 0 หรือ Status เป็น Order_Submitted)
            pending_items = 0
            
            for i, row in df.iterrows():
                try: order_val = int(row.get('Order', 0)) if str(row.get('Order', '')).strip() != '' else 0
                except: order_val = 0
                status_val = str(row.get('Status', '')).strip()
                
                if order_val > 0 or status_val == "Order_Submitted":
                    pending_items += 1
                    st.markdown("---") 
                    cols = st.columns([3, 1.5, 1.5])
                    cols[0].markdown(f"**{row['Name']}**")
                    
                    # โชว์ยอดที่สั่ง (อ่านอย่างเดียว)
                    cols[1].markdown(f"<div style='text-align:center;'><small>{t('col_ordered')}</small><br><b>{order_val}</b></div>", unsafe_allow_html=True)
                    
                    # ช่องกรอกยอดรับจริง (Default เท่ากับยอดที่สั่ง)
                    actual_recv = cols[2].number_input(t("col_actual"), min_value=0, value=order_val, key=f"r_{i}")
                    
                    updates[i + 2] = {"Actual_Recv": actual_recv, "Status": "Received"}
            
            if pending_items == 0:
                st.success(t("no_pending"))

        # ==========================================
        # 🚀 ปุ่มส่งข้อมูลและอัปเดต Google Sheet
        # ==========================================
        st.markdown("---")
        if st.form_submit_button(t("submit_btn"), type="primary"):
            if not updates:
                st.warning(t("no_changes"))
            else:
                try:
                    with st.spinner(t("sending")):
                        cells_to_update = []
                        for r_idx, vals in updates.items():
                            if app_mode == t("mode_order"):
                                cells_to_update.append(gspread.Cell(r_idx, 4, vals['Current'])) # Col D
                                cells_to_update.append(gspread.Cell(r_idx, 5, vals['Order']))   # Col E
                                cells_to_update.append(gspread.Cell(r_idx, 7, vals['Status']))  # Col G
                            else:
                                cells_to_update.append(gspread.Cell(r_idx, 8, vals['Actual_Recv'])) # Col H (คอลัมน์ใหม่สำหรับรับจริง)
                                cells_to_update.append(gspread.Cell(r_idx, 7, vals['Status']))      # Col G

                        ws.update_cells(cells_to_update)
                        
                    st.success(f"{t('success')} ({len(updates)} items)")
                    st.balloons()
                    time.sleep(1.5)
                    st.rerun()
                    
                except Exception as e:
                    st.error(f"{t('error')} {e}")
