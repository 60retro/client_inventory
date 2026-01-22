import sys
from types import ModuleType

# --- 🛠️ ส่วนแก้บั๊ก Python 3.13 (Mock imghdr module) ---
# ต้องใส่ไว้บนสุด ก่อน import streamlit เสมอ
if sys.version_info >= (3, 13):
    m = ModuleType("imghdr")
    m.what = lambda *args: None  # สร้างฟังก์ชันปลอมๆ ขึ้นมาหลอก
    sys.modules["imghdr"] = m
# ----------------------------------------------------

import streamlit as st
import pandas as pd
import gspread
from google.oauth2.service_account import Credentials
import time
import os
import json

# --- Config ---
SHEET_NAME = "inventory_data"
CREDENTIALS_FILE = "credentials.json"

# ตั้งค่าหน้าเว็บ
st.set_page_config(page_title="Nami Stock Client", page_icon="📱")

# --- 1. ระบบภาษา (Translation System) ---
TRANSLATIONS = {
    "th": {
        "title": "📱 Nami Stock Check",
        "caption": "ระบบตรวจนับสต๊อกและสั่งของ (Client)",
        "select_category": "📂 เลือกหมวดหมู่",
        "no_items": "⚠️ ไม่มีสินค้าในหมวดหมู่นี้",
        "instruction": "📝 กรอกยอด **'คงเหลือ'** หรือ **'สั่งเพิ่ม'**",
        "col_name": "รายการ",
        "col_remain": "📦 คงเหลือ (Remaining)",
        "col_order": "🛒 สั่งเพิ่ม (Order)",
        "submit_btn": "🚀 ส่งข้อมูล (Submit)",
        "no_changes": "⚠️ ไม่มีการแก้ไขข้อมูล",
        "sending": "กำลังส่งข้อมูล... (Sending)",
        "success": "✅ ส่งข้อมูลเรียบร้อย! (Success)",
        "error": "❌ เกิดข้อผิดพลาด: ",
        "conn_error": "❌ ไม่สามารถเชื่อมต่อ Google API ได้",
        "sheet_error": "❌ หาไฟล์ Google Sheet ไม่เจอ: "
    },
    "en": {
        "title": "📱 Nami Stock Check",
        "caption": "Inventory Counting & Ordering System",
        "select_category": "📂 Select Category",
        "no_items": "⚠️ No items found in this category",
        "instruction": "📝 Enter **'Remaining'** stock or **'Order'** quantity",
        "col_name": "Item Name",
        "col_remain": "📦 Remaining",
        "col_order": "🛒 Order Qty",
        "submit_btn": "🚀 Submit Data",
        "no_changes": "⚠️ No changes detected",
        "sending": "Sending data...",
        "success": "✅ Data sent successfully!",
        "error": "❌ Error occurred: ",
        "conn_error": "❌ Cannot connect to Google API",
        "sheet_error": "❌ Google Sheet not found: "
    },
    "mm": { 
        "title": "📱 Nami Stock Check",
        "caption": "ကုန်ပစ္စည်းစာရင်း စစ်ဆေးခြင်းနှင့် မှာယူခြင်းစနစ်",
        "select_category": "📂 အမျိုးအစား ရွေးပါ (Category)",
        "no_items": "⚠️ ဤအမျိုးအစားတွင် ပစ္စည်းမရှိပါ",
        "instruction": "📝 **'လက်ကျန်'** သို့မဟုတ် **'မှာယူမည့်အရေအတွက်'** ကို ထည့်ပါ",
        "col_name": "ပစ္စည်းအမည်",
        "col_remain": "📦 လက်ကျန် (Remaining)",
        "col_order": "🛒 မှာယူမည် (Order)",
        "submit_btn": "🚀 ပေးပို့ပါ (Submit)",
        "no_changes": "⚠️ ပြင်ဆင်ထားသော အချက်အလက် မရှိပါ",
        "sending": "ပေးပို့နေသည်... (Sending)",
        "success": "✅ ပေးပို့ပြီးပါပြီ! (Success)",
        "error": "❌ မှားယွင်းမှုရှိသည်: ",
        "conn_error": "❌ Google API နှင့် ချိတ်ဆက်၍ မရပါ",
        "sheet_error": "❌ Google Sheet ဖိုင်ကို ရှာမတွေ့ပါ: "
    }
}

# Sidebar Language Selection
st.sidebar.title("Language / ภาษา / ဘာသာစကား")
lang_option = st.sidebar.radio(
    "Select Language:",
    ("ภาษาไทย (Thai)", "English", "မြန်မာ (Burmese)")
)

if "Thai" in lang_option: current_lang = "th"
elif "Burmese" in lang_option: current_lang = "mm"
else: current_lang = "en"

def t(key):
    return TRANSLATIONS[current_lang][key]

# --- 2. Function เชื่อมต่อ Google Sheet ---
@st.cache_resource
def get_google_sheet_client():
    scopes = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive"
    ]
    try:
        # อ่านจาก Secrets (Cloud)
        if "gcp_json" in st.secrets:
            info = st.secrets["gcp_json"]
            creds = Credentials.from_service_account_info(info, scopes=scopes)
        # อ่านจาก Local File (PC)
        elif os.path.exists(CREDENTIALS_FILE):
            creds = Credentials.from_service_account_file(CREDENTIALS_FILE, scopes=scopes)
        else:
            return None
        
        client = gspread.authorize(creds)
        return client
    except Exception as e:
        st.error(f"Connect Error: {e}")
        return None

# --- 3. Main App Logic ---
st.title(t("title"))
st.caption(t("caption"))

client = get_google_sheet_client()

if not client:
    st.error(t("conn_error"))
    st.stop()

try:
    sh = client.open(SHEET_NAME)
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
    else:
        st.info(t("instruction"))
        
        with st.form("stock_entry_form"):
            updates = {} 
            
            for i, row in df.iterrows():
                st.markdown(f"---") 
                cols = st.columns([3, 1.5, 1.5])
                
                cols[0].markdown(f"**{row['Name']}**")
                
                try: curr_val = int(row['Current']) if row['Current'] != '' else 0
                except: curr_val = 0
                try: order_val = int(row['Order']) if row['Order'] != '' else 0
                except: order_val = 0
                
                new_curr = cols[1].number_input(t("col_remain"), min_value=0, value=curr_val, key=f"c_{i}")
                new_order = cols[2].number_input(t("col_order"), min_value=0, value=order_val, key=f"o_{i}")
                
                if new_curr != curr_val or new_order != order_val:
                    updates[i + 2] = {"Current": new_curr, "Order": new_order}

            st.markdown("---")
            if st.form_submit_button(t("submit_btn"), type="primary"):
                if not updates:
                    st.warning(t("no_changes"))
                else:
                    try:
                        with st.spinner(t("sending")):
                            cells_to_update = []
                            for r_idx, vals in updates.items():
                                cells_to_update.append(gspread.Cell(r_idx, 4, vals['Current'])) 
                                cells_to_update.append(gspread.Cell(r_idx, 5, vals['Order']))   
                                cells_to_update.append(gspread.Cell(r_idx, 7, 'Pending'))       
                            
                            ws.update_cells(cells_to_update)
                            
                        st.success(f"{t('success')} ({len(updates)} items)")
                        st.balloons()
                        time.sleep(1)
                        st.rerun()
                        
                    except Exception as e:
                        st.error(f"{t('error')} {e}")
