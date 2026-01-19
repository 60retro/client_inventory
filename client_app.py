import streamlit as st
import pandas as pd
import gspread
from google.oauth2.service_account import Credentials
import time

# --- Config ---
SHEET_NAME = "Nami_Inventory_DB"
CREDENTIALS_FILE = "credentials.json"

# ตั้งค่าหน้าเว็บ
st.set_page_config(page_title="Nami Stock Client", page_icon="📱")

# --- Function เชื่อมต่อ Google Sheet (ใช้ Cache เพื่อความเร็ว) ---
@st.cache_resource
def get_google_sheet_client():
    scopes = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive"
    ]
    try:
        creds = Credentials.from_service_account_file(CREDENTIALS_FILE, scopes=scopes)
        client = gspread.authorize(creds)
        return client
    except Exception as e:
        return None

# --- Main App ---
st.title("📱 Nami Stock Check")
st.caption("ระบบตรวจนับสต๊อกและสั่งของ (Client)")

# 1. เชื่อมต่อ
client = get_google_sheet_client()

if not client:
    st.error("❌ ไม่สามารถเชื่อมต่อ Google API ได้ (เช็คไฟล์ credentials.json)")
    st.stop()

try:
    sh = client.open(SHEET_NAME)
except gspread.exceptions.SpreadsheetNotFound:
    st.error(f"❌ หาไฟล์ Google Sheet ชื่อ '{SHEET_NAME}' ไม่เจอ")
    st.stop()

# 2. ดึงรายชื่อ Tab (Category)
all_worksheets = [ws.title for ws in sh.worksheets()]
selected_tab = st.selectbox("📂 เลือกหมวดหมู่ (Select Category)", all_worksheets)

if selected_tab:
    # โหลดข้อมูลจาก Tab ที่เลือก
    ws = sh.worksheet(selected_tab)
    
    # ใช้ pandas ดึงข้อมูลจะจัดการง่ายกว่า
    try:
        data = ws.get_all_records()
        df = pd.DataFrame(data)
    except Exception as e:
        st.error(f"Error loading data: {e}")
        st.stop()

    if df.empty:
        st.warning("ไม่มีสินค้าในหมวดหมู่นี้")
    else:
        st.info("📝 กรอกยอด **'คงเหลือ'** หรือ **'สั่งเพิ่ม'** แล้วกดปุ่มส่งด้านล่าง")
        
        # --- Form สำหรับกรอกข้อมูล ---
        with st.form("stock_entry_form"):
            # ตัวแปรเก็บค่าที่แก้ไข {row_index: {col: val}}
            updates = {} 
            
            # วนลูปสร้าง Input ตามจำนวนสินค้า
            # ใช้ columns เพื่อจัดหน้าตาให้ดูง่าย (ชื่อสินค้า | คงเหลือ | สั่ง)
            for i, row in df.iterrows():
                st.markdown(f"---") 
                cols = st.columns([3, 1.5, 1.5])
                
                # ชื่อสินค้า
                cols[0].markdown(f"**{row['Name']}**")
                
                # แปลงค่าเดิมเป็น int เพื่อใส่ในช่อง Input (ถ้าว่างให้เป็น 0)
                try: curr_val = int(row['Current']) if row['Current'] != '' else 0
                except: curr_val = 0
                
                try: order_val = int(row['Order']) if row['Order'] != '' else 0
                except: order_val = 0
                
                # ช่องกรอก Current (คงเหลือ)
                new_curr = cols[1].number_input(
                    "📦 คงเหลือ", 
                    min_value=0, 
                    value=curr_val, 
                    key=f"curr_{i}"
                )
                
                # ช่องกรอก Order (สั่ง)
                new_order = cols[2].number_input(
                    "🛒 สั่งเพิ่ม", 
                    min_value=0, 
                    value=order_val, 
                    key=f"order_{i}"
                )
                
                # เช็คว่ามีการเปลี่ยนแปลงหรือไม่
                if new_curr != curr_val or new_order != order_val:
                    # เก็บ row index (Google Sheet เริ่มที่ 1, Header คือแถว 1, ดังนั้น data เริ่มแถว 2)
                    # i เริ่ม 0 ดังนั้น row จริงคือ i + 2
                    updates[i + 2] = {
                        "Current": new_curr,
                        "Order": new_order
                    }

            st.markdown("---")
            submitted = st.form_submit_button("🚀 ส่งข้อมูลไปที่ Host (Submit)", type="primary")
            
            if submitted:
                if not updates:
                    st.warning("⚠️ คุณยังไม่ได้แก้ไขข้อมูลใดๆ")
                else:
                    # --- Process Update ---
                    progress_bar = st.progress(0)
                    status_text = st.empty()
                    
                    try:
                        total_upd = len(updates)
                        count = 0
                        
                        # คอลัมน์อ้างอิงตาม Header: No, Name, Prev, Current, Order, Price, Status
                        # Current = Col 4 (D)
                        # Order   = Col 5 (E)
                        # Status  = Col 7 (G)
                        
                        for row_idx, vals in updates.items():
                            status_text.text(f"Updating row {row_idx}...")
                            
                            # อัปเดตทีละ Cell (ช้าหน่อยแต่ชัวร์)
                            # ถ้าข้อมูลเยอะมากแนะนำให้อัปเกรดเป็น batch_update ในอนาคต
                            ws.update_cell(row_idx, 4, vals['Current']) # Update Current
                            ws.update_cell(row_idx, 5, vals['Order'])   # Update Order
                            ws.update_cell(row_idx, 7, 'Pending')       # Update Status -> ให้ Host รู้ว่าต้อง Sync
                            
                            count += 1
                            progress_bar.progress(count / total_upd)
                            
                        st.success("✅ ส่งข้อมูลเรียบร้อยแล้ว! (Data sent successfully)")
                        st.balloons()
                        
                        # รอ 2 วินาทีแล้ว Refresh หน้าจอเพื่อโหลดค่าใหม่
                        time.sleep(2)
                        st.rerun()
                        
                    except Exception as e:
                        st.error(f"❌ เกิดข้อผิดพลาดในการส่งข้อมูล: {e}")