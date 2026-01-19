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

# --- Function เชื่อมต่อ Google Sheet ---
@st.cache_resource
def get_google_sheet_client():
    scopes = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive"
    ]
    try:
        # 1. อ่านจาก Streamlit Secrets (Cloud)
        if "gcp_json" in st.secrets:
            info = st.secrets["gcp_json"]
            creds = Credentials.from_service_account_info(info, scopes=scopes)
        
        # 2. อ่านจากไฟล์ Local (PC)
        elif os.path.exists(CREDENTIALS_FILE):
            creds = Credentials.from_service_account_file(CREDENTIALS_FILE, scopes=scopes)
        else:
            return None

        client = gspread.authorize(creds)
        return client
    except Exception as e:
        st.error(f"Connect Error: {e}")
        return None

# --- Main App ---
st.title("📱 Nami Stock Check")
st.caption("ระบบตรวจนับสต๊อกและสั่งของ (Client)")

# 1. เชื่อมต่อ
client = get_google_sheet_client()

if not client:
    st.error("❌ ไม่สามารถเชื่อมต่อ Google API ได้")
    st.warning("Cloud: กรุณาตั้งค่า Secrets [gcp_json]\nPC: เช็คไฟล์ credentials.json")
    st.stop()

try:
    sh = client.open(SHEET_NAME)
except gspread.exceptions.SpreadsheetNotFound:
    st.error(f"❌ หาไฟล์ Google Sheet ชื่อ '{SHEET_NAME}' ไม่เจอ")
    st.stop()

# 2. ดึงรายชื่อ Tab
all_worksheets = [ws.title for ws in sh.worksheets()]
selected_tab = st.selectbox("📂 เลือกหมวดหมู่", all_worksheets)

if selected_tab:
    ws = sh.worksheet(selected_tab)
    try:
        data = ws.get_all_records()
        df = pd.DataFrame(data)
    except Exception as e:
        st.error(f"Load Error: {e}")
        st.stop()

    if df.empty:
        st.warning("ไม่มีสินค้าในหมวดหมู่นี้")
    else:
        st.info("📝 กรอกยอด **'คงเหลือ'** หรือ **'สั่งเพิ่ม'**")
        
        with st.form("stock_entry_form"):
            updates = {} 
            # สร้างตัวแปรเก็บ cell object เพื่อรอ update ทีเดียว
            batch_cells = []
            
            for i, row in df.iterrows():
                st.markdown(f"---") 
                cols = st.columns([3, 1.5, 1.5])
                cols[0].markdown(f"**{row['Name']}**")
                
                try: curr_val = int(row['Current']) if row['Current'] != '' else 0
                except: curr_val = 0
                try: order_val = int(row['Order']) if row['Order'] != '' else 0
                except: order_val = 0
                
                new_curr = cols[1].number_input("คงเหลือ", min_value=0, value=curr_val, key=f"c_{i}")
                new_order = cols[2].number_input("สั่งเพิ่ม", min_value=0, value=order_val, key=f"o_{i}")
                
                # เช็คว่ามีการแก้ไขหรือไม่
                if new_curr != curr_val or new_order != order_val:
                    # เก็บข้อมูลตำแหน่ง Row และค่าที่จะแก้ไว้ก่อน
                    # (Row เริ่มที่ 2 เพราะ header=1, i เริ่ม 0)
                    row_num = i + 2
                    updates[row_num] = {"Current": new_curr, "Order": new_order}

            st.markdown("---")
            if st.form_submit_button("🚀 ส่งข้อมูล (Submit)", type="primary"):
                if not updates:
                    st.warning("⚠️ ไม่มีการแก้ไขข้อมูล")
                else:
                    try:
                        with st.spinner("กำลังส่งข้อมูลแบบ Batch..."):
                            # เตรียม List ของ Cell ทั้งหมดที่จะแก้
                            cells_to_update = []
                            for r_idx, vals in updates.items():
                                # Column 4 = Current, 5 = Order, 7 = Status
                                cells_to_update.append(gspread.Cell(r_idx, 4, vals['Current']))
                                cells_to_update.append(gspread.Cell(r_idx, 5, vals['Order']))
                                cells_to_update.append(gspread.Cell(r_idx, 7, 'Pending'))
                            
                            # ยิง API ครั้งเดียวจบ (Batch Update) แก้ปัญหา Quota Exceeded
                            ws.update_cells(cells_to_update)
                            
                        st.success(f"✅ ส่งข้อมูลเรียบร้อย! (อัปเดต {len(updates)} รายการ)")
                        st.balloons()
                        time.sleep(1)
                        st.rerun()
                        
                    except Exception as e:
                        st.error(f"❌ Error: {e}")
