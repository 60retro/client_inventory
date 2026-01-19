import tkinter as tk
from tkinter import ttk, messagebox, simpledialog, filedialog
import pandas as pd
from datetime import datetime
import os
import json
import sys
import shutil
import time
import warnings
import threading
import requests  # <--- เพิ่ม library สำหรับยิง LINE API

# --- เพิ่ม Library สำหรับ Google Sheets ---
import gspread
from oauth2client.service_account import ServiceAccountCredentials

# ปิดคำเตือน
warnings.filterwarnings("ignore")

# --- Matplotlib Check ---
try:
    import matplotlib
    matplotlib.use('Agg') 
    import matplotlib.pyplot as plt
    from pandas.plotting import table
    import matplotlib.font_manager as fm
    MATPLOTLIB_AVAILABLE = True
except ImportError:
    MATPLOTLIB_AVAILABLE = False

# --- Gemini AI Check ---
GEMINI_AVAILABLE = False
try:
    import google.generativeai as genai
    GEMINI_AVAILABLE = True
except ImportError: pass

# --- Configuration ---
PRICE_DATA_FILE = "item_prices.json"
INVENTORY_STATE_FILE = "inventory_state.json"
CONFIG_FILE = "config.json"
HISTORY_FILE = "summary_history.json"
TODAY_LOG_FILE = "today_log.json"
CUSTOM_FONT_FILE = "custom_font.ttf" 
# --- เพิ่ม Config สำหรับ Cloud ---
CREDENTIALS_FILE = "credentials.json"
SHEET_NAME = "inventory_data"

class InventoryApp:
    def __init__(self, root):
        self.root = root
        self.root.title("ระบบจัดการสต๊อก (Inventory System V.Host Hybrid - LINE/Google Sheet)")
        self.root.geometry("1500x850")
        
        # Style
        self.style = ttk.Style()
        self.style.theme_use("clam")
        self.style.configure("Treeview", rowheight=30, font=('Arial', 10))
        self.style.configure("Treeview.Heading", font=('Arial', 11, 'bold'))
        
        # Data containers
        self.data = [] 
        self.categories = [] 
        self.treeviews = {}
        self.today_logs = [] 
        
        # --- ส่วนที่เพิ่ม: ตัวแปร Search ---
        self.search_var = tk.StringVar()
        
        # State
        self.show_prices = False
        self.current_date = datetime.now().strftime("%Y-%m-%d")
        self.current_filename = "ยังไม่ได้เลือกไฟล์"
        self.target_tab_index = None 
        
        # Helpers
        self.price_history = self.load_price_history() 
        self.summary_history = self.load_summary_history() 
        self.load_today_log() 
        
        # Settings
        self.admin_password = "admin"        
        self.dashboard_password = "1234"
        self.drive_path = "" 
        self.gemini_api_key = "AIzaSyDCcmdHAkdPht3vvRmcDeIRDW6yhHXZYV4" 
        
        # --- LINE CONFIG ---
        self.line_channel_token = "" 
        self.line_user_id = "" 
        
        self.load_config() 

        # Setup AI
        if GEMINI_AVAILABLE and self.gemini_api_key:
            genai.configure(api_key=self.gemini_api_key)

        # UI Setup
        self.setup_menu()
        self.setup_top_panel()
        self.setup_notebook()
        self.setup_bottom_panel() 
        
        # Load Data
        if not self.load_state():
            self.show_welcome_screen() 

        root.protocol("WM_DELETE_WINDOW", self.on_closing)
        self.editor = None

        # --- ส่วนที่เพิ่ม: เชื่อมต่อ Google Sheet แบบแยก Thread ---
        self.gc = None
        self.sh = None
        threading.Thread(target=self.init_google_sheet, daemon=True).start()

    # --- ฟังก์ชันใหม่: เชื่อมต่อ Google Sheet ---
    def init_google_sheet(self):
        try:
            if os.path.exists(CREDENTIALS_FILE):
                scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
                creds = ServiceAccountCredentials.from_json_keyfile_name(CREDENTIALS_FILE, scope)
                self.gc = gspread.authorize(creds)
                self.sh = self.gc.open(SHEET_NAME)
                print(f"✅ Connected to Google Sheet: {SHEET_NAME}")
            else:
                print("⚠️ Credentials file not found.")
        except Exception as e:
            print(f"❌ GSheet Connection Error: {e}")

    # --- Config & Data Handlers ---
    def load_config(self):
        if os.path.exists(CONFIG_FILE):
            try:
                with open(CONFIG_FILE, 'r', encoding='utf-8') as f:
                    config = json.load(f)
                    self.admin_password = config.get('admin_password', 'admin') 
                    self.dashboard_password = config.get('dashboard_password', '1234')
                    self.drive_path = config.get('drive_path', '')
                    self.gemini_api_key = config.get('gemini_api_key', '')
                    # Load LINE Config
                    self.line_channel_token = config.get('line_channel_token', '')
                    self.line_user_id = config.get('line_user_id', '')
            except: pass
        self.save_config()

    def save_config(self):
        try:
            with open(CONFIG_FILE, 'w', encoding='utf-8') as f:
                json.dump({
                    'admin_password': self.admin_password, 
                    'dashboard_password': self.dashboard_password,
                    'drive_path': self.drive_path,
                    'gemini_api_key': self.gemini_api_key,
                    'line_channel_token': self.line_channel_token, # Save LINE Token
                    'line_user_id': self.line_user_id             # Save User ID
                }, f, indent=4)
        except: pass
            
    def load_price_history(self):
        if os.path.exists(PRICE_DATA_FILE):
            try:
                with open(PRICE_DATA_FILE, 'r', encoding='utf-8') as f: return json.load(f)
            except: return {}
        return {}
    
    def load_summary_history(self):
        if os.path.exists(HISTORY_FILE):
            try:
                with open(HISTORY_FILE, 'r', encoding='utf-8') as f: return json.load(f)
            except: return []
        return []

    def save_summary_history(self):
        try:
            with open(HISTORY_FILE, 'w', encoding='utf-8') as f: json.dump(self.summary_history, f, indent=4, ensure_ascii=False)
        except: pass

    def save_price_history(self):
        for item in self.data:
            key = str((item['category'], item['name'])) 
            self.price_history[key] = item['price']
        try:
            with open(PRICE_DATA_FILE, 'w', encoding='utf-8') as f: json.dump(self.price_history, f, indent=4, ensure_ascii=False)
        except: pass

    # --- TODAY LOG LOGIC ---
    def load_today_log(self):
        if os.path.exists(TODAY_LOG_FILE):
            try:
                with open(TODAY_LOG_FILE, 'r', encoding='utf-8') as f:
                    data = json.load(f)
                    if data.get('date') == self.current_date:
                        self.today_logs = data.get('logs', [])
                    else:
                        self.today_logs = [] 
            except: self.today_logs = []
        else: self.today_logs = []

    def save_today_log(self):
        try:
            with open(TODAY_LOG_FILE, 'w', encoding='utf-8') as f:
                json.dump({'date': self.current_date, 'logs': self.today_logs}, f, indent=4)
        except: pass

    # --- LOAD STATE LOGIC ---
    def load_state(self):
        if os.path.exists(INVENTORY_STATE_FILE):
            try:
                with open(INVENTORY_STATE_FILE, 'r', encoding='utf-8') as f:
                    loaded_json = json.load(f)
                    if isinstance(loaded_json, list):
                        self.data = loaded_json
                        self.categories = sorted(list(set(d['category'] for d in self.data)))
                    elif isinstance(loaded_json, dict):
                        self.data = loaded_json.get('items', [])
                        self.categories = loaded_json.get('categories', [])
                        if not self.categories and self.data:
                            self.categories = sorted(list(set(d['category'] for d in self.data)))
                    
                    for item in self.data:
                        if 'item_no' not in item: item['item_no'] = "-"
                        if 'prev_stock' not in item: item['prev_stock'] = 0
                        if 'stock_remaining' not in item: item['stock_remaining'] = 0
                        if 'min_stock_target' not in item: item['min_stock_target'] = 5
                        if 'order_qty' not in item: item['order_qty'] = 0
                        if 'price' not in item: item['price'] = 0.0
                        if 'last_received_qty' not in item: item['last_received_qty'] = 0
                    
                    self.current_filename = "Auto-loaded State"
                    if hasattr(self, 'file_label'):
                        self.file_label.config(text=f"File: {self.current_filename}", foreground="blue")
                    self.refresh_tabs()
                    return True
            except Exception as e:
                print(f"Load Error: {e}")
                return False
        return False

    # --- SAVE STATE LOGIC ---
    def save_state(self):
        save_data = {"items": self.data, "categories": self.categories}
        try:
            with open(INVENTORY_STATE_FILE, 'w', encoding='utf-8') as f:
                json.dump(save_data, f, indent=4, ensure_ascii=False)
        except: pass
    
    def on_closing(self):
        self.save_state()
        self.save_price_history()
        self.root.destroy()

    # --- Folder Sync Logic ---
    def select_drive_folder(self):
        folder_selected = filedialog.askdirectory(title="เลือกโฟลเดอร์ Google Drive")
        if folder_selected:
            self.drive_path = folder_selected
            self.save_config()
            messagebox.showinfo("Success", f"ตั้งค่าโฟลเดอร์เรียบร้อย:\n{self.drive_path}")

    def copy_to_drive_folder(self, source_file):
        if not self.drive_path or not os.path.exists(self.drive_path): return False
        filename = os.path.basename(source_file)
        destination = os.path.join(self.drive_path, filename)
        if os.path.abspath(source_file) == os.path.abspath(destination): return True
        for i in range(3):
            try:
                shutil.copyfile(source_file, destination)
                return True
            except (PermissionError, shutil.SameFileError): time.sleep(1.0)
            except Exception: return False
        return False

    # --- LINE Messaging API Feature ---
    def send_line_push(self, message):
        if not self.line_channel_token or not self.line_user_id:
            print("LINE Config missing: Token or User ID not set.")
            return

        url = 'https://api.line.me/v2/bot/message/push'
        headers = {
            'Content-Type': 'application/json',
            'Authorization': f'Bearer {self.line_channel_token}'
        }
        payload = {
            'to': self.line_user_id,
            'messages': [
                {
                    'type': 'text',
                    'text': message
                }
            ]
        }

        try:
            response = requests.post(url, headers=headers, json=payload)
            if response.status_code != 200:
                print(f"LINE Error: {response.text}")
            else:
                print("LINE sent successfully.")
        except Exception as e:
            print(f"LINE Connection Error: {e}")
            
    def setting_line_dialog(self):
        d = tk.Toplevel(self.root); d.title("LINE API Settings")
        d.geometry("500x300")
        
        ttk.Label(d, text="Channel Access Token:", font=('bold')).pack(pady=5)
        e_token = ttk.Entry(d, width=60); e_token.pack(pady=5)
        e_token.insert(0, self.line_channel_token)
        
        ttk.Label(d, text="Your User ID (U...):", font=('bold')).pack(pady=5)
        e_uid = ttk.Entry(d, width=60); e_uid.pack(pady=5)
        e_uid.insert(0, self.line_user_id)
        
        def save():
            self.line_channel_token = e_token.get().strip()
            self.line_user_id = e_uid.get().strip()
            self.save_config()
            messagebox.showinfo("Success", "บันทึกค่า LINE เรียบร้อย\nลองกดปิดรอบสต๊อกเพื่อทดสอบ")
            d.destroy()
            
        ttk.Button(d, text="Save Settings", command=save).pack(pady=20)

    # --- ฟังก์ชันใหม่: Sync จาก Cloud ---
    def sync_from_cloud(self):
        if not self.sh:
            messagebox.showerror("Error", "ยังไม่ได้เชื่อมต่อ Google Sheet หรือ Key ผิด\nกรุณาตรวจสอบไฟล์ credentials.json")
            return
            
        loading = tk.Toplevel(self.root)
        tk.Label(loading, text="Syncing data from Client...", padx=20, pady=20).pack()
        loading.geometry("300x100")
        loading.update()

        try:
            update_count = 0
            for cat in self.categories:
                try:
                    worksheet = self.sh.worksheet(cat)
                except:
                    print(f"Sheet '{cat}' not found on cloud. Skipping.")
                    continue
                
                records = worksheet.get_all_records()
                # คอลัมน์ใน Sheet ต้องมี: Name, Current, Order, Status
                
                for row in records:
                    name = str(row.get('Name', '')).strip()
                    status = str(row.get('Status', '')).strip()
                    
                    # ถ้าสถานะเป็น Pending (Client ส่งมา) หรือมี Order/Current
                    has_data = (row.get('Order') != '' and row.get('Order') != 0) or (row.get('Current') != '' and row.get('Current') != 0)
                    
                    if status == 'Pending' or has_data:
                        # หา Item ใน Local
                        local_item = next((x for x in self.data if x['category'] == cat and x['name'] == name), None)
                        if local_item:
                            try:
                                client_current = row.get('Current')
                                if client_current != '' and client_current is not None:
                                    local_item['stock_remaining'] = int(client_current)
                                
                                client_order = row.get('Order')
                                if client_order != '' and client_order is not None:
                                    local_item['order_qty'] = int(client_order)
                                
                                update_count += 1
                            except: pass
    # --- ฟังก์ชันใหม่: ส่งรายชื่อสินค้าทั้งหมดขึ้น Cloud ---
    def push_items_to_cloud(self):
        if not self.sh:
            messagebox.showerror("Error", "ยังไม่ได้เชื่อมต่อ Google Sheet")
            return

        if not messagebox.askyesno("Confirm Upload", "คุณต้องการ 'ล้างข้อมูลบน Cloud' และอัปโหลดรายชื่อสินค้าจากเครื่องนี้ขึ้นไปใหม่ใช่หรือไม่?\n\n(ทำเมื่อมีการเพิ่มเมนูใหม่ หรือเปลี่ยนราคา)"):
            return

        loading = tk.Toplevel(self.root)
        tk.Label(loading, text="Uploading all items to Cloud...", padx=20, pady=20).pack()
        loading.geometry("300x100")
        loading.update()

        try:
            # วนลูปทุกหมวดหมู่ที่มีในเครื่อง
            for cat in self.categories:
                try:
                    # พยายามเปิด Sheet ถ้าไม่มีให้สร้างใหม่ (ถ้าทำได้)
                    try:
                        worksheet = self.sh.worksheet(cat)
                    except:
                        # ถ้าหา Tab ไม่เจอ ให้ข้าม หรือแจ้งเตือน (Google Sheet API สร้าง Tab ยากถ้าสิทธิ์ไม่ถึง)
                        print(f"Warning: Sheet '{cat}' not found. Please create it manually.")
                        continue

                    # เตรียมข้อมูล Header
                    new_rows = [['No', 'Name', 'Prev', 'Current', 'Order', 'Price', 'Status']]
                    
                    # ดึงสินค้าเฉพาะหมวดนี้จากเครื่อง
                    items_in_cat = [x for x in self.data if x['category'] == cat]
                    
                    for item in items_in_cat:
                        # สร้างแถวข้อมูล (เซ็ตค่า Order/Current เป็น 0 เพื่อเริ่มรอบใหม่)
                        new_rows.append([
                            item.get('item_no', '-'),
                            item['name'],
                            item['prev_stock'], # เอาสต๊อกล่าสุดในเครื่องขึ้นไปเป็น Prev
                            0,                  # Current (ว่างไว้รอพนักงานกรอก)
                            0,                  # Order (ว่างไว้รอพนักงานกรอก)
                            item['price'],
                            'Clean'             # สถานะเริ่มต้น
                        ])
                    
                    # ล้างข้อมูลเก่าแล้วเขียนทับ
                    worksheet.clear()
                    worksheet.update(new_rows)
                    
                except Exception as e:
                    print(f"Error uploading category {cat}: {e}")

            loading.destroy()
            messagebox.showinfo("Success", "อัปโหลดรายชื่อสินค้าขึ้น Cloud เรียบร้อยแล้ว!\nClient สามารถกด Refresh เพื่อเห็นเมนูใหม่ได้ทันที")

        except Exception as e:
            loading.destroy()
            messagebox.showerror("Upload Error", f"{e}")
                            
            self.save_state()
            self.refresh_tabs()
            loading.destroy()
            messagebox.showinfo("Sync Complete", f"ดึงข้อมูลล่าสุดจาก Client เรียบร้อย\nอัปเดตทั้งหมด: {update_count} รายการ")
            
        except Exception as e:
            loading.destroy()
            messagebox.showerror("Sync Error", f"{e}")

    # --- HTML Generator ---
    def generate_html_receipt(self, df, filename):
        try:
            items_df = df[df['Item Name'] != 'GRAND TOTAL']
            total_items = len(items_df)
            
            html_content = f"""
            <!DOCTYPE html><html lang="th"><head><meta charset="UTF-8">
            <style>
                body {{ font-family: 'Sarabun', 'Tahoma', sans-serif; background: #eee; margin: 0; padding: 0; }}
                @page {{ size: A4 landscape; margin: 10mm; }}
                .page {{ background: white; width: 100%; margin: 0 auto; box-sizing: border-box; }}
                h2 {{ text-align: center; margin: 0 0 5px 0; font-size: 18px; }}
                .header-info {{ text-align: center; margin-bottom: 10px; font-weight: bold; font-size: 14px; }}
                .multi-col-container {{ column-count: 3; column-gap: 10px; column-rule: 1px solid #ddd; width: 100%; }}
                .item-row {{ break-inside: avoid-column; page-break-inside: avoid; border-bottom: 1px solid #ccc; display: flex; font-size: 11px; }}
                .item-row:last-child {{ border-bottom: none; }}
                .item-table {{ width: 100%; break-inside: avoid; margin-bottom: -1px; }}
                table {{ width: 100%; border-collapse: collapse; font-size: 11px; margin-bottom: 5px; }}
                th, td {{ border: 1px solid #000; padding: 3px; }}
                th {{ background-color: #FFFF00; color: black; text-align: center; }}
            </style>
            </head>
            <body>
            <div class="page">
                <h2>ใบรายการเบิกสินค้า (Order Request)</h2>
                <div class="header-info">วันที่: {self.current_date} (Total: {total_items} items)</div>
                <div class="multi-col-container">
            """
            
            for idx, row in items_df.iterrows():
                if idx == 0:
                      html_content += """
                      <table class="item-table" style="margin-bottom: 0;">
                        <thead>
                            <tr>
                                <th width="10%">No.</th>
                                <th width="60%">รายการ</th>
                                <th width="15%">เบิก</th>
                                <th width="15%">รับ</th>
                            </tr>
                        </thead>
                      </table>
                      """

                html_content += f"""
                <table class="item-table">
                    <tbody>
                        <tr>
                            <td width="10%" align="center">{row.get('No.', '-')}</td>
                            <td width="60%">{row.get('Item Name', '')}</td>
                            <td width="15%" align="center"><strong>{row.get('Order Qty', 0)}</strong></td>
                            <td width="15%"></td>
                        </tr>
                    </tbody>
                </table>
                """
            
            html_content += """</div></div></body></html>"""
            
            html_path = filename.replace(".xlsx", ".html")
            with open(html_path, "w", encoding="utf-8") as f: f.write(html_content)
            return html_path
            
        except Exception as e:
            print(f"HTML Error: {e}")
            return None

    # --- AI Feature ---
    def open_smart_add_dialog(self):
        if not GEMINI_AVAILABLE:
            messagebox.showerror("Error", "ไม่พบโมดูล google-generativeai\nกรุณาติดตั้ง: pip install google-generativeai")
            return

        dialog = tk.Toplevel(self.root)
        dialog.title("✨ AI Smart Add (Gemini)")
        dialog.geometry("600x500")
        
        api_frame = ttk.Frame(dialog, padding=10); api_frame.pack(fill='x')
        ttk.Label(api_frame, text="API Key:").pack(side='left')
        api_var = tk.StringVar(value=self.gemini_api_key)
        ttk.Entry(api_frame, textvariable=api_var, show="*").pack(side='left', fill='x', expand=True, padx=5)
        def save_key():
            self.gemini_api_key = api_var.get().strip()
            self.save_config()
            try:
                genai.configure(api_key=self.gemini_api_key)
                messagebox.showinfo("OK", "Saved")
            except: pass
        ttk.Button(api_frame, text="Save", command=save_key).pack(side='right')

        cat_frame = ttk.Frame(dialog, padding=10); cat_frame.pack(fill='x')
        ttk.Label(cat_frame, text="Target Tab:").pack(anchor='w')
        current = self.notebook.tab(self.notebook.select(), "text") if self.notebook.select() else ""
        cat_combo = ttk.Combobox(cat_frame, values=self.categories, state="readonly")
        if current in self.categories: cat_combo.set(current)
        elif self.categories: cat_combo.set(self.categories[0])
        cat_combo.pack(fill='x')

        ttk.Label(dialog, text="Items Text:", padding=10).pack(anchor='w')
        text_input = tk.Text(dialog, height=10); text_input.pack(fill='both', expand=True, padx=10)

        btn_frame = ttk.Frame(dialog, padding=10); btn_frame.pack(fill='x', side='bottom')
        lbl = ttk.Label(btn_frame, text="", foreground="blue"); lbl.pack(side='left')
        
        def on_process():
            if not self.gemini_api_key: messagebox.showwarning("Warning", "No API Key"); return
            txt = text_input.get("1.0", tk.END).strip()
            cat = cat_combo.get()
            if not txt or not cat: return
            lbl.config(text="Processing...")
            threading.Thread(target=lambda: self.run_gemini(txt, cat, dialog, lbl)).start()
            
        ttk.Button(btn_frame, text="Run AI", command=on_process).pack(side='right')

    def run_gemini(self, text, cat, dialog, lbl):
        try:
            model_name = 'gemini-1.5-flash' 
            try:
                for m in genai.list_models():
                    if 'generateContent' in m.supported_generation_methods:
                        if 'flash' in m.name: model_name = m.name; break
                        elif 'pro' in m.name: model_name = m.name
            except: pass

            model = genai.GenerativeModel(model_name)
            prompt = f"Extract items from: '{text}'. Return JSON array: [{{\"name\":\"ItemName\",\"qty\":Number}}]. No markdown. Only JSON."
            resp = model.generate_content(prompt)
            data = json.loads(resp.text.replace('```json','').replace('```','').strip())
            self.root.after(0, lambda: self.process_ai(data, cat, dialog))
        except Exception as e:
            self.root.after(0, lambda: lbl.config(text=f"Error: {e}", foreground="red"))

    def process_ai(self, items, cat, dialog):
        c_add = 0; c_upd = 0
        for item in items:
            raw_name = item.get('name', 'Unknown').strip()
            qty = int(item.get('qty', 1))
            existing = next((x for x in self.data if x['category'] == cat and x['name'] == raw_name), None)
            if existing: existing['order_qty'] += qty; c_upd += 1
            else:
                self.data.append({"item_no":"-", "category":cat, "name":raw_name, "stock_remaining":0, "prev_stock":0, "order_qty":qty, "min_stock_target":5, "price":0.0})
                c_add += 1
        self.save_state(); self.refresh_tabs(); dialog.destroy()
        messagebox.showinfo("Done", f"Added: {c_add}, Merged: {c_upd}")

    # --- UI Setup ---
    def setup_menu(self):
        menubar = tk.Menu(self.root)
        file_menu = tk.Menu(menubar, tearoff=0)
        file_menu.add_command(label="Import New Excel File", command=self.import_file)
        file_menu.add_command(label="Export Backup (Separated Tabs)", command=self.export_selective_backup)
        file_menu.add_separator()
        file_menu.add_command(label="Exit", command=self.root.quit)
        menubar.add_cascade(label="File", menu=file_menu)
        
        view_menu = tk.Menu(menubar, tearoff=0)
        view_menu.add_command(label="Toggle Price Edit", command=self.toggle_price_view)
        menubar.add_cascade(label="View", menu=view_menu)

        edit_menu = tk.Menu(menubar, tearoff=0)
        edit_menu.add_command(label="Add New Item", command=self.add_new_item_dialog)
        edit_menu.add_command(label="Add New Category (Tab)", command=self.add_new_category_dialog)
        edit_menu.add_separator()
        edit_menu.add_command(label="⚠️ Reset Stock to 0 (ล้างสต๊อก)", command=self.reset_all_stock)
        edit_menu.add_separator()
        edit_menu.add_command(label="Change Admin Password...", command=self.change_admin_password)
        edit_menu.add_command(label="Change Dashboard Password...", command=self.change_dashboard_password)
        edit_menu.add_separator()
        edit_menu.add_command(label="📂 Set Google Drive Folder...", command=self.select_drive_folder)
        # เพิ่มเมนูตั้งค่า LINE
        edit_menu.add_command(label="💬 Set LINE API...", command=self.setting_line_dialog)
        menubar.add_cascade(label="Settings", menu=edit_menu)
        self.root.config(menu=menubar)

    def setup_top_panel(self):
        frame = ttk.Frame(self.root, padding=10)
        frame.pack(fill='x', side='top')
        
        ttk.Button(frame, text="📂 Import", command=self.import_file).pack(side='left', padx=5)
        ttk.Button(frame, text="✨ AI Smart Add", command=self.open_smart_add_dialog).pack(side='left', padx=5)
        
        # --- ส่วนที่เพิ่ม: Search Box ---
        search_frame = ttk.Frame(frame)
        search_frame.pack(side='left', padx=5)
        ttk.Label(search_frame, text="🔍 Search:").pack(side='left')
        self.search_var.trace("w", lambda name, index, mode, sv=self.search_var: self.on_search_change())
        ttk.Entry(search_frame, textvariable=self.search_var, width=20).pack(side='left', padx=5)
        # -----------------------------

        # --- ส่วนที่เพิ่ม: Sync Button ---
        ttk.Button(frame, text="☁️ Sync Client", command=self.sync_from_cloud, style='Accent.TButton').pack(side='left', padx=20)
        # -----------------------------
        
        # +++ เพิ่มปุ่มนี้เข้าไป +++
        ttk.Button(frame, text="📤 Upload Items", command=self.push_items_to_cloud).pack(side='left', padx=5)
        # +++++++++++++++++++++++

        self.file_label = ttk.Label(frame, text="File: None", foreground="gray")
        self.file_label.pack(side='left', padx=10)
        
        ttk.Button(frame, text="🖨️ Print", command=self.print_order_slip_dialog).pack(side='right', padx=5)
        ttk.Button(frame, text="📦 Receive & Close", command=self.open_receive_goods_dialog, style='Accent.TButton').pack(side='right', padx=5)
        ttk.Separator(frame, orient='vertical').pack(side='right', fill='y', padx=10)
        ttk.Button(frame, text="📊 Dashboard", command=self.check_password_for_summary).pack(side='right', padx=5)
        ttk.Button(frame, text="💾 Save", command=self.save_state).pack(side='right', padx=5)

    # --- ฟังก์ชันใหม่ Search Listener ---
    def on_search_change(self):
        if not self.notebook.select(): return
        current_tab_name = self.notebook.tab(self.notebook.select(), "text")
        self.refresh_treeview_data(current_tab_name)

    def setup_notebook(self):
        self.notebook = ttk.Notebook(self.root)
        self.notebook.pack(fill='both', expand=True, padx=10, pady=5)
        self.notebook.bind("<Button-3>", self.show_tab_context_menu)
        
    def show_welcome_screen(self):
        f = ttk.Frame(self.notebook); self.notebook.add(f, text="Welcome")
        ttk.Label(f, text="Inventory System Ready (Hybrid Host)", font=('Arial', 14)).place(relx=0.5, rely=0.5, anchor='center')

    def setup_bottom_panel(self):
        panel = ttk.Frame(self.root, padding=10)
        panel.pack(fill='x', side='bottom')
        self.status_lbl = ttk.Label(panel, text=f"Today: {self.current_date}")
        self.status_lbl.pack(side='left')
        self.total_value_lbl = ttk.Label(panel, text="", font=('Arial', 11, 'bold'), foreground="blue")
        self.total_value_lbl.pack(side='right')

    # --- Dashboard Features ---
    def check_password_for_summary(self):
        pwd = simpledialog.askstring("Pass", "Dash Pass:", show='*')
        if pwd == self.dashboard_password: self.show_summary_dashboard()
        elif pwd is not None: messagebox.showerror("Error", "Wrong Pass")

    def show_summary_dashboard(self):
        self.perform_save_history(0) 

        win = tk.Toplevel(self.root); win.title("Dashboard"); win.geometry("1000x700")
        tabs = ttk.Notebook(win); tabs.pack(fill='both', expand=True, padx=10, pady=10)
        self.build_today_tab(ttk.Frame(tabs), win); tabs.add(tabs.winfo_children()[0], text="Today Status")
        self.build_log_tab(ttk.Frame(tabs)); tabs.add(tabs.winfo_children()[1], text="Today Receipt Log") 
        self.build_history_tab(ttk.Frame(tabs)); tabs.add(tabs.winfo_children()[2], text="History")

    def build_today_tab(self, p, w):
        stock_v = sum(d['prev_stock']*d['price'] for d in self.data)
        order_v = sum(d['order_qty']*d['price'] for d in self.data)
        usage_v = 0
        cats = {}
        for d in self.data:
            u = max(0, d['prev_stock'] - d['stock_remaining']) if d['stock_remaining'] > 0 else 0
            usage_val = u * d['price']
            usage_v += usage_val
            
            c = d['category']
            if c not in cats: cats[c] = {'s':0,'o':0,'u':0}
            cats[c]['s'] += d['prev_stock']*d['price']
            cats[c]['o'] += d['order_qty']*d['price']
            cats[c]['u'] += usage_val
        
        tf = ttk.Frame(p, padding=10); tf.pack(fill='x')
        for i, (t, v) in enumerate([("Total Stock Value",stock_v), ("Total Usage Value",usage_v), ("Pending Order",order_v)]):
            f = tk.LabelFrame(tf, text=t, font=('Bold',10)); f.grid(row=0, column=i, padx=10, sticky='nsew')
            tk.Label(f, text=f"{v:,.2f} B", font=('Bold',16)).pack()
        
        tv = ttk.Treeview(p, columns=("Cat","Stock Val","Usage Val","Pending Order"), show='headings'); tv.pack(fill='both', expand=True)
        for c in ("Cat","Stock Val","Usage Val","Pending Order"): tv.heading(c, text=c)
        for c, v in cats.items(): 
            tv.insert("", "end", values=(c, f"{v['s']:,.2f}", f"{v['u']:,.2f}", f"{v['o']:,.2f}"))
            
        ttk.Button(p, text="Export Detailed Report", command=lambda: self.export_summary_report(w)).pack(side='bottom', pady=5)

    def build_log_tab(self, p):
        columns = ("Time", "Category", "Item", "Ordered", "Actual Recv", "Actual Pay")
        tv = ttk.Treeview(p, columns=columns, show='headings')
        tv.pack(fill='both', expand=True)
        
        tv.heading("Time", text="Time")
        tv.heading("Category", text="Category")
        tv.heading("Item", text="Item Name")
        tv.heading("Ordered", text="สั่ง (Qty)")
        tv.heading("Actual Recv", text="รับจริง (Qty)")
        tv.heading("Actual Pay", text="จ่ายจริง (Baht)")
        
        tv.column("Ordered", width=80, anchor='center')
        tv.column("Actual Recv", width=80, anchor='center')
        tv.column("Actual Pay", width=100, anchor='e')

        total_recv_val = 0
        for log in self.today_logs:
            o_qty = log.get('order_qty', '-') 
            r_qty = log.get('recv_qty', log.get('qty', 0))
            val = log.get('val', 0)
            
            tv.insert("", "end", values=(
                log.get('time'), 
                log.get('cat'), 
                log.get('name'), 
                o_qty, 
                r_qty, 
                f"{val:,.2f}"
            ))
            total_recv_val += val
            
        ttk.Label(p, text=f"Total Actual Pay Today: {total_recv_val:,.2f} THB", font=('Bold', 14), foreground="blue").pack(pady=10)

    def build_history_tab(self, p):
        tv = ttk.Treeview(p, columns=("Date","Stock Val","Order Val"), show='headings'); tv.pack(fill='both', expand=True)
        for c in ("Date","Stock Val","Order Val"): tv.heading(c, text=c)
        for r in reversed(self.summary_history): 
            tv.insert("", "end", values=(r.get('date'), f"{r.get('total_stock_val',0):,.2f}", f"{r.get('total_order_val',0):,.2f}"), tags=(r.get('date'),))
        
        tv.bind("<Double-1>", lambda e: self.on_history_double_click(e, tv))

    def on_history_double_click(self, event, tree):
        item = tree.selection()
        if not item: return
        vals = tree.item(item, "values")
        date_str = vals[0]
        
        record = next((r for r in self.summary_history if r['date'] == date_str), None)
        if not record or 'details' not in record:
            messagebox.showinfo("Info", "No detailed data for this date.")
            return
            
        pop = tk.Toplevel(self.root)
        pop.title(f"Details: {date_str}")
        pop.geometry("600x400")
        
        tv = ttk.Treeview(pop, columns=("Category", "Stock Val", "Order Val"), show='headings')
        tv.pack(fill='both', expand=True)
        for c in ("Category", "Stock Val", "Order Val"): tv.heading(c, text=c)
        
        details = record['details']
        for cat, val in details.items():
            tv.insert("", "end", values=(cat, f"{val.get('stock',0):,.2f}", f"{val.get('order',0):,.2f}"))

    def export_summary_report(self, window):
        f = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel Files", "*.xlsx")], initialfile=f"Full_Report_{self.current_date}.xlsx")
        if not f: return
        
        try:
            with pd.ExcelWriter(f) as writer:
                if self.summary_history:
                    hist_data = []
                    for h in self.summary_history:
                        hist_data.append({
                            "Date": h.get('date'),
                            "Total Stock Value": h.get('total_stock_val',0),
                            "Actual Paid Value": h.get('total_order_val',0)
                        })
                    pd.DataFrame(hist_data).to_excel(writer, sheet_name="Overview History", index=False)
                
                stock_data = []
                for d in self.data:
                    stock_data.append({
                        "Date": self.current_date,
                        "Category": d['category'],
                        "Item Name": d['name'],
                        "Current Stock": d['prev_stock'],
                        "Unit Price": d['price'],
                        "Stock Value": d['prev_stock'] * d['price']
                    })
                pd.DataFrame(stock_data).to_excel(writer, sheet_name="Current Stock", index=False)
                
                cat_summary = {}
                for d in self.data:
                    c = d['category']
                    if c not in cat_summary: cat_summary[c] = {'Stock Val': 0, 'Pending Order Val': 0}
                    cat_summary[c]['Stock Val'] += d['prev_stock'] * d['price']
                    cat_summary[c]['Pending Order Val'] += d['order_qty'] * d['price']
                
                cat_list = [{"Date": self.current_date, "Category": k, "Stock Value": v['Stock Val'], "Pending Order Value": v['Pending Order Val']} for k,v in cat_summary.items()]
                pd.DataFrame(cat_list).to_excel(writer, sheet_name="Category Summary", index=False)
                
                if self.today_logs:
                    log_data = []
                    for l in self.today_logs:
                        o_qty = l.get('order_qty', 0)
                        r_qty = l.get('recv_qty', l.get('qty', 0))
                        try: diff = r_qty - int(o_qty)
                        except: diff = 0
                        
                        log_data.append({
                            "Date": self.current_date,
                            "Time": l['time'],
                            "Category": l['cat'],
                            "Item": l['name'],
                            "Ordered Qty": o_qty,
                            "Actual Received": r_qty,
                            "Diff (Recv-Order)": diff,
                            "Actual Pay (Baht)": l['val']
                        })
                    pd.DataFrame(log_data).to_excel(writer, sheet_name="Receipt Log Today", index=False)

            if self.drive_path: self.copy_to_drive_folder(f)
            messagebox.showinfo("Success", "Export Complete (4 Sheets)\nรวมถึงข้อมูล Actual Recv vs Ordered")
            if sys.platform == "win32": os.startfile(f)
            
        except Exception as e:
            messagebox.showerror("Error", f"{e}")

    # --- Print Function (Re-added) ---
    def print_order_slip_dialog(self):
        if not any(d.get('order_qty', 0) > 0 for d in self.data):
            messagebox.showinfo("No Orders", "ไม่มีรายการสินค้าที่มียอดเบิก")
            return
        dialog = tk.Toplevel(self.root); dialog.title("Select Categories"); dialog.geometry("400x500")
        
        btn_frame = ttk.Frame(dialog, padding=10); btn_frame.pack(side='bottom', fill='x')
        ttk.Label(dialog, text="เลือก Tab ที่ต้องการพิมพ์:", font=('Arial', 11, 'bold')).pack(side='top', pady=10)
        
        container = ttk.Frame(dialog); container.pack(side='top', fill='both', expand=True, padx=10)
        canvas = tk.Canvas(container); sb = ttk.Scrollbar(container, orient="vertical", command=canvas.yview)
        scroll_frame = ttk.Frame(canvas)
        scroll_frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0,0), window=scroll_frame, anchor="nw"); canvas.configure(yscrollcommand=sb.set)
        canvas.pack(side="left", fill="both", expand=True); sb.pack(side="right", fill="y")
        
        cat_vars = {}
        for cat in self.categories:
            var = tk.BooleanVar(value=True); cat_vars[cat] = var
            ttk.Checkbutton(scroll_frame, text=cat, variable=var).pack(anchor='w', padx=5, pady=2)
            
        def toggle(s): 
            for v in cat_vars.values(): v.set(s)
        ttk.Button(btn_frame, text="All", command=lambda: toggle(True)).pack(side='left')
        ttk.Button(btn_frame, text="None", command=lambda: toggle(False)).pack(side='left', padx=5)
        ttk.Button(btn_frame, text="Print", command=lambda: self.process_print(cat_vars, dialog)).pack(side='right')

    def process_print(self, cat_vars, dialog):
        selected = [c for c, v in cat_vars.items() if v.get()]
        dialog.destroy()
        if not selected: return
        
        slip_data = []
        total_qty = 0
        for item in self.data:
            qty = item.get('order_qty', 0)
            if qty > 0 and item['category'] in selected:
                slip_data.append({"No.": item.get('item_no', '-'), "Category": item['category'], "Item Name": item['name'], "Order Qty": item['order_qty']})
                total_qty += qty
        
        if not slip_data: messagebox.showinfo("Info", "Empty"); return
        
        df = pd.DataFrame(slip_data)
        filename = f"Request_{self.current_date}_{datetime.now().strftime('%H%M')}.xlsx"
        filepath = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel Files", "*.xlsx")], initialfile=filename)
        if not filepath: return
        
        try:
            df_ex = df.copy(); df_ex = pd.concat([df_ex, pd.DataFrame([{'No.':'', 'Category':'', 'Item Name':'GRAND TOTAL', 'Order Qty': total_qty}])], ignore_index=True)
            with pd.ExcelWriter(filepath, engine='openpyxl') as writer: df_ex.to_excel(writer, sheet_name='Request', index=False)
            html_path = self.generate_html_receipt(df, filepath)
            if sys.platform == "win32": os.startfile(filepath)
            
            if self.drive_path:
                self.copy_to_drive_folder(filepath)
                if html_path: self.copy_to_drive_folder(html_path)
            
            messagebox.showinfo("Success", "Saved")
            if html_path and messagebox.askyesno("Open", "View HTML?"): os.startfile(html_path)
        except Exception as e: messagebox.showerror("Error", f"{e}")

    # *** FIX: UPDATE ALL STOCKS (Even if no order) ***
    # Mod: Update Cloud Logic added inside confirm
    def open_receive_goods_dialog(self):
        orders = [d for d in self.data if d.get('order_qty', 0) > 0]
        has_counts = any(d.get('stock_remaining', 0) > 0 for d in self.data)
        
        if not orders and not has_counts:
            messagebox.showinfo("Info", "ไม่มีรายการที่มีการเคลื่อนไหว (No Orders or Counts)")
            return
            
        d = tk.Toplevel(self.root); d.title("Receive & Close Cycle (ปิดยอดและอัปเดต Cloud)"); d.geometry("900x600")
        
        c_vals = {i['name']: str(i['order_qty']) for i in orders}
        
        lbl_info = ttk.Label(d, text="เมื่อกด Confirm ระบบจะตัดสต๊อกในเครื่อง และ 'ส่งยอดคงเหลือใหม่' ไปทับที่ Client (Reset Round)", foreground="blue", font=('Arial', 11))
        lbl_info.pack(pady=10)

        cv = tk.Canvas(d); sb = ttk.Scrollbar(d, orient="vertical", command=cv.yview)
        sf = ttk.Frame(cv); sf.bind("<Configure>", lambda e: cv.configure(scrollregion=cv.bbox("all")))
        cv.create_window((0,0), window=sf, anchor="nw"); cv.configure(yscrollcommand=sb.set)
        
        if orders:
            cv.pack(side="left", fill="both", expand=True, padx=10); sb.pack(side="right", fill="y")
            ttk.Label(sf, text="รายการ", font=('bold')).grid(row=0, column=0, padx=5)
            ttk.Label(sf, text="สั่ง (Order)", font=('bold')).grid(row=0, column=1, padx=5)
            ttk.Label(sf, text="รับจริง (Actual)", font=('bold')).grid(row=0, column=2, padx=5)

            for idx, item in enumerate(orders):
                r = idx + 1
                ttk.Label(sf, text=item['name']).grid(row=r, column=0, sticky='w')
                ttk.Label(sf, text=str(item['order_qty']), foreground="gray").grid(row=r, column=1)
                e = ttk.Entry(sf, width=10); e.insert(0, c_vals[item['name']])
                e.grid(row=r, column=2, padx=5, pady=2)
                e.bind("<KeyRelease>", lambda ev, n=item['name']: c_vals.update({n: ev.widget.get()}))
        else:
            ttk.Label(d, text="--- ไม่มีรายการสั่งซื้อ (No Orders) ---", font=('bold')).pack(pady=20)
            ttk.Label(d, text="กด Confirm เพื่อนำยอดที่นับ (Current) ไปเป็นยอดยกมา (Prev) ของรอบถัดไป", foreground="green").pack()

        def confirm():
            if not messagebox.askyesno("Confirm", "ยืนยันการปิดยอดสต๊อกและรับของ? \n(ระบบจะ Reset ข้อมูลบน Cloud ด้วย)"): return
            total_actual_pay = 0
            updated_count = 0
            total_usage_val_today = 0
            
            # --- LOOP 1: รายการที่มี Order ---
            for item in orders:
                ordered_qty = item['order_qty']
                try: actual_recv_qty = int(float(c_vals.get(item['name'], 0)))
                except: actual_recv_qty = 0
                
                actual_val = actual_recv_qty * item['price']
                current_count = item['stock_remaining']
                
                # Usage Calculation Logic
                if current_count > 0:
                    usage_qty = max(0, item['prev_stock'] - current_count)
                    total_usage_val_today += usage_qty * item['price']
                    item['prev_stock'] = current_count + actual_recv_qty
                else:
                    if item['prev_stock'] == 0: 
                        item['prev_stock'] = actual_recv_qty
                    else:
                        item['prev_stock'] = item['prev_stock'] + actual_recv_qty 

                item['stock_remaining'] = 0 
                item['order_qty'] = 0        
                
                total_actual_pay += actual_val
                
                self.today_logs.append({
                    'time': datetime.now().strftime("%H:%M:%S"), 
                    'cat': item['category'], 'name': item['name'], 
                    'order_qty': ordered_qty, 'recv_qty': actual_recv_qty, 'val': actual_val
                })
                updated_count += 1

            # --- LOOP 2: รายการไม่มี Order แต่มีการนับ ---
            for item in self.data:
                if item in orders: continue
                if item.get('stock_remaining', 0) > 0:
                    usage_qty = max(0, item['prev_stock'] - item['stock_remaining'])
                    total_usage_val_today += usage_qty * item['price']
                    
                    item['prev_stock'] = item['stock_remaining']
                    item['stock_remaining'] = 0
                    item['order_qty'] = 0
                    updated_count += 1
            
            # Calculate Total Stock Value (End of Cycle)
            total_stock_val_now = sum(d['prev_stock'] * d['price'] for d in self.data)
            
            # --- Upload NEW 'Prev' to Cloud (Reset Cycle) ---
            if self.sh:
                try:
                    upload_status = tk.Toplevel(self.root)
                    tk.Label(upload_status, text="Updating Cloud for next cycle...", padx=20, pady=20).pack()
                    upload_status.geometry("300x100")
                    upload_status.update()
                    
                    for cat in self.categories:
                        try:
                            worksheet = self.sh.worksheet(cat)
                            records = worksheet.get_all_records()
                            new_rows = []
                            # Header must match Google Sheet
                            header = ['No', 'Name', 'Prev', 'Current', 'Order', 'Price', 'Status']
                            new_rows.append(header)
                            
                            for row in records:
                                name = str(row.get('Name', '')).strip()
                                local_item = next((x for x in self.data if x['category'] == cat and x['name'] == name), None)
                                
                                if local_item:
                                    new_prev = local_item['prev_stock']
                                    price = local_item['price']
                                else:
                                    new_prev = row.get('Prev', 0)
                                    price = row.get('Price', 0)
                                    
                                new_rows.append([
                                    row.get('No', '-'),
                                    name,
                                    new_prev, # New Start Stock (from Local)
                                    0,        # Reset Count
                                    0,        # Reset Order
                                    price,
                                    'Clean'   # Reset Status
                                ])
                            
                            worksheet.clear()
                            worksheet.update(new_rows)
                            
                        except Exception as e:
                            print(f"Cloud Update Error {cat}: {e}")
                    
                    upload_status.destroy()
                except Exception as e:
                    print(f"Upload Critical Error: {e}")
            # -----------------------------------------------------------

            self.perform_save_history(total_actual_pay)
            self.save_today_log()
            self.save_state(); self.refresh_tabs()
            
            # --- SEND LINE PUSH MESSAGE ---
            try:
                msg = (
                    f"📊 Nami Stock Summary\n"
                    f"📅 Date: {self.current_date}\n"
                    f"----------------------------\n"
                    f"💰 Total Order: {total_actual_pay:,.2f} THB\n"
                    f"📉 Total Usage: {total_usage_val_today:,.2f} THB\n"
                    f"📦 Total Stock: {total_stock_val_now:,.2f} THB\n"
                    f"----------------------------\n"
                    f"✅ Cycle Closed Successfully"
                )
                threading.Thread(target=lambda: self.send_line_push(msg)).start()
            except Exception as e:
                print(f"Failed to trigger LINE: {e}")
            # ------------------------------

            d.destroy()
            messagebox.showinfo("Success", f"ปิดรอบสต๊อกเรียบร้อย\nอัปเดตข้อมูลสินค้าไปทั้งหมด {updated_count} รายการ\nและ Reset ข้อมูลบน Cloud แล้ว")
            
        ttk.Button(d, text="Confirm Update & Close Cycle", command=confirm).pack(pady=10, side='bottom')

    def perform_save_history(self, _):
        # 1. Snapshot Stock Value
        total_stock = sum(d['prev_stock']*d['price'] for d in self.data)
        
        # 2. Daily Total Order from Log
        daily_total_order = sum(log.get('val', 0) for log in self.today_logs)

        # 3. Snapshot by Category
        cat_snapshot = {}
        for d in self.data:
            c = d['category']
            if c not in cat_snapshot: cat_snapshot[c] = {'stock':0, 'order':0}
            cat_snapshot[c]['stock'] += d['prev_stock'] * d['price']
            
        for log in self.today_logs:
            c = log.get('cat', 'Unknown')
            val = log.get('val', 0)
            if c not in cat_snapshot: cat_snapshot[c] = {'stock':0, 'order':0}
            cat_snapshot[c]['order'] += val

        # 4. Save to History
        history_record = {
            "date": self.current_date, 
            "total_stock_val": total_stock, 
            "total_order_val": daily_total_order, 
            "details": cat_snapshot 
        }
        
        existing_idx = next((i for i, item in enumerate(self.summary_history) if item["date"] == self.current_date), -1)
        if existing_idx != -1: self.summary_history[existing_idx] = history_record
        else: self.summary_history.append(history_record)
        self.save_summary_history()

    # --- ฟังก์ชัน Grand Total ที่หายไป ---
    def update_grand_total(self): 
        try:
            if self.show_prices: 
                total = sum(d['price']*(d['stock_remaining'] if d['stock_remaining']>0 else d['prev_stock']) for d in self.data)
                self.total_value_lbl.config(text=f"Total: {total:,.2f} THB")
            else: 
                self.total_value_lbl.config(text="")
        except: 
            pass

    def reset_all_stock(self):
        pwd = simpledialog.askstring("Admin Reset", "ใส่รหัส Admin เพื่อยืนยันการล้างสต๊อก:", show='*')
        if pwd != self.admin_password:
            if pwd is not None: messagebox.showerror("Error", "รหัสผ่านผิด")
            return

        if not messagebox.askyesno("Confirm Reset", "⚠️ คำเตือน: คุณต้องการปรับยอดสต๊อก 'ทุกรายการ' เป็น 0 ใช่หรือไม่?\n(ชื่อเมนูจะยังอยู่ แต่จำนวนคงเหลือจะหายหมด)"): 
            return

        for item in self.data:
            item['prev_stock'] = 0      
            item['stock_remaining'] = 0 
            item['order_qty'] = 0        
        
        self.save_state()
        self.refresh_tabs()
        messagebox.showinfo("Success", "ล้างค่าสต๊อกเป็น 0 เรียบร้อยแล้ว")

    def toggle_price_view(self):
        if not self.show_prices:
            if simpledialog.askstring("Admin", "Password:", show='*') == self.admin_password: self.show_prices = True
        else: self.show_prices = False
        # เรียก refresh_treeview_data ของ Tab ปัจจุบัน เพื่ออัปเดตการแสดงผล
        if self.notebook.select():
            current_tab = self.notebook.tab(self.notebook.select(), "text")
            self.refresh_treeview_data(current_tab)
        self.update_grand_total()

    def change_admin_password(self): self.change_pass('admin')
    def change_dashboard_password(self): self.change_pass('dashboard')
    def change_pass(self, role):
        p = simpledialog.askstring("Old", "Old Pass:", show='*')
        check = self.admin_password if role=='admin' else self.dashboard_password
        if p == check:
            n = simpledialog.askstring("New", "New Pass:", show='*')
            if n: 
                if role=='admin': self.admin_password = n
                else: self.dashboard_password = n
                self.save_config(); messagebox.showinfo("Success", "Changed")

    # --- NEW EXPORT FUNCTION (BACKUP) WITH MULTI-TABS ---
    def export_selective_backup(self):
        dialog = tk.Toplevel(self.root); dialog.title("Export Backup (Multi-Tab)")
        dialog.geometry("400x500")
        
        btn_frame = ttk.Frame(dialog, padding=10); btn_frame.pack(side='bottom', fill='x')
        ttk.Label(dialog, text="เลือก Tab ที่ต้องการ Export (Backup):", font=('Arial', 11, 'bold')).pack(side='top', pady=10)
        
        container = ttk.Frame(dialog); container.pack(side='top', fill='both', expand=True, padx=10)
        canvas = tk.Canvas(container); sb = ttk.Scrollbar(container, orient="vertical", command=canvas.yview)
        scroll_frame = ttk.Frame(canvas)
        scroll_frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0,0), window=scroll_frame, anchor="nw"); canvas.configure(yscrollcommand=sb.set)
        canvas.pack(side="left", fill="both", expand=True); sb.pack(side="right", fill="y")
        
        cat_vars = {}
        for cat in self.categories:
            var = tk.BooleanVar(value=True); cat_vars[cat] = var
            ttk.Checkbutton(scroll_frame, text=cat, variable=var).pack(anchor='w', padx=5, pady=2)
            
        def toggle(s): 
            for v in cat_vars.values(): v.set(s)
        ttk.Button(btn_frame, text="All", command=lambda: toggle(True)).pack(side='left')
        ttk.Button(btn_frame, text="None", command=lambda: toggle(False)).pack(side='left', padx=5)
        
        def do_export():
            selected = [c for c, v in cat_vars.items() if v.get()]
            if not selected: return
            
            f = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel Files", "*.xlsx")], initialfile=f"Backup_{self.current_date}.xlsx")
            if not f: return

            try:
                # ใช้ ExcelWriter เพื่อเขียนหลายๆ Sheet ลงไฟล์เดียว
                with pd.ExcelWriter(f) as writer:
                    exported_count = 0
                    for cat in selected:
                        cat_data = []
                        for item in self.data:
                            if item['category'] == cat:
                                cat_data.append({
                                    "No": item.get('item_no', '-'),
                                    "Name": item.get('name', ''),
                                    "Price": item.get('price', 0.0)
                                })
                        
                        if cat_data:
                            df = pd.DataFrame(cat_data)
                            safe_sheet_name = str(cat)[:31].replace(":", "").replace("/", "")
                            df.to_excel(writer, sheet_name=safe_sheet_name, index=False)
                            exported_count += len(cat_data)
                    
                    if exported_count > 0:
                        messagebox.showinfo("Success", f"Exported {exported_count} items across {len(selected)} sheets.")
                        dialog.destroy()
                    else:
                        messagebox.showinfo("Info", "No data found in selected categories.")
            
            except Exception as e:
                messagebox.showerror("Error", f"{e}")

        ttk.Button(btn_frame, text="Export", command=do_export).pack(side='right')
    # ----------------------------------------------------

    def import_file(self): 
        pwd = simpledialog.askstring("Admin Access", "กรุณาใส่รหัสผ่าน Admin เพื่อ Import:", show='*')
        if pwd != self.admin_password: 
            if pwd is not None: 
                messagebox.showerror("Access Denied", "รหัสผ่านไม่ถูกต้อง")
            return 

        path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx"), ("CSV Files", "*.csv")])
        if not path: return
        try:
            if path.endswith('.csv'): 
                dfs = {'Sheet1': pd.read_csv(path)}
            else: 
                dfs = pd.read_excel(path, sheet_name=None)
            
            total_imported = 0
            existing = {item['name']: item for item in self.data}
            new_data = []

            for sheet_name, df in dfs.items():
                df.columns = [str(c).strip() for c in df.columns]
                
                col_cat = next((c for c in df.columns if any(k in str(c) for k in ['แผนก', 'Dept', 'Category', 'Cat'])), None)
                col_name = next((c for c in df.columns if any(k in str(c) for k in ['รายการ', 'ชื่อ', 'Item', 'Name'])), None)
                col_price = next((c for c in df.columns if any(k in str(c).lower() for k in ['price', 'ราคา', 'cost'])), None)
                
                if not col_name: continue 
                
                for _, row in df.iterrows():
                    name = str(row[col_name]).strip()
                    if not name or name == "nan": continue
                    
                    if col_cat:
                        cat = str(row[col_cat]).strip()
                    else:
                        cat = str(sheet_name).strip() 
                    
                    price = 0.0
                    if col_price:
                        try: price = float(row[col_price])
                        except: price = 0.0

                    if name in existing:
                        item = existing[name]
                        item['category'] = cat 
                        if price > 0: item['price'] = price 
                        new_data.append(item)
                    else:
                        new_data.append({
                            "item_no": "-", 
                            "category": cat, 
                            "name": name, 
                            "stock_remaining": 0, 
                            "prev_stock": 0, 
                            "order_qty": 0, 
                            "min_stock_target": 5, 
                            "price": price
                        })
                    total_imported += 1
            
            if new_data:
                self.data = new_data
                self.categories = sorted(list(set(d['category'] for d in self.data)))
                self.refresh_tabs()
                messagebox.showinfo("Success", f"Imported {len(self.data)} items from all sheets.")
            else:
                messagebox.showinfo("Info", "No valid data found in file.")

        except Exception as e: messagebox.showerror("Error", f"{e}")

    # --- Helper methods ---
    def delete_tab(self):
        # Fix: เช็คก่อนว่ามี tab ให้ลบไหม
        if self.target_tab_index is None: return
        cat = self.notebook.tab(self.target_tab_index, "text")
        if messagebox.askyesno("Confirm", f"Delete '{cat}'?"):
            self.data = [d for d in self.data if d['category']!=cat]
            if cat in self.categories: self.categories.remove(cat)
            self.save_state(); self.refresh_tabs()

    def show_tab_context_menu(self, event):
        try:
            self.target_tab_index = self.notebook.index(f"@{event.x},{event.y}")
            m = tk.Menu(self.root, tearoff=0)
            m.add_command(label="Delete Tab", command=self.delete_tab)
            m.post(event.x_root, event.y_root)
        except: pass

    def edit_cell(self, tv, row, key, col):
        x, y, w, h = tv.bbox(row, col)
        name = tv.item(row, 'values')[1]
        target_cat = next((cat for cat, tree in self.treeviews.items() if tree == tv), None)
        item = next((x for x in self.data if x['name']==name and x['category']==target_cat), None)
        if not item: return
        self.editor = ttk.Entry(tv, justify='center')
        self.editor.insert(0, str(item.get(key, '')))
        self.editor.place(x=x, y=y, width=w, height=h); self.editor.focus_set()
        self.editor.item = item; self.editor.key = key
        self.editor.bind('<Return>', lambda e: self.save_inline_edit())
        self.editor.bind('<FocusOut>', lambda e: self.save_inline_edit())

    def save_inline_edit(self):
        if not self.editor: return
        try:
            val = self.editor.get().strip()
            if self.editor.key == 'price': val = float(val)
            elif self.editor.key not in ['item_no', 'name']: val = int(float(val))
            self.editor.item[self.editor.key] = val
            self.save_state()
            self.refresh_treeview_data(self.editor.item['category'])
        except: pass
        self.editor.destroy(); self.editor = None

    def add_new_item_dialog(self):
        if not self.notebook.tabs(): return
        cat = self.notebook.tab(self.notebook.select(), "text")
        self.data.append({"item_no": "-", "category": cat, "name": "New Item", "stock_remaining": 0, "prev_stock": 0, "order_qty": 0, "min_stock_target": 5, "price": 0.0})
        self.save_state(); self.refresh_treeview_data(cat)
        tv = self.treeviews[cat]; ch = tv.get_children()
        if ch: self.root.after(100, lambda: self.edit_cell(tv, ch[-1], 'name', '#2'))

    def add_new_category_dialog(self):
        n = simpledialog.askstring("New Tab", "Name:")
        if n and n not in self.categories:
            self.categories.append(n); self.save_state(); self.refresh_tabs()

    def refresh_tabs(self):
        for i in self.notebook.winfo_children(): i.destroy()
        self.treeviews = {}
        for cat in self.categories:
            f = ttk.Frame(self.notebook); self.notebook.add(f, text=cat)
            cols = ("No", "Name", "Prev", "Current", "Usage", "Min", "Order", "Price", "Total")
            tv = ttk.Treeview(f, columns=cols, show='headings')
            for c in cols: tv.heading(c, text=c); tv.column(c, width=60 if c!='Name' else 200, anchor='center' if c!='Name' else 'w')
            tv.pack(fill='both', expand=True); tv.bind("<Button-1>", self.on_tree_click)
            self.treeviews[cat] = tv
            self.refresh_treeview_data(cat)

    def refresh_treeview_data(self, cat):
        if cat not in self.treeviews: return
        tv = self.treeviews[cat]
        for i in tv.get_children(): tv.delete(i)
        
        # --- Logic Search ที่เพิ่มเข้ามา ---
        search_keyword = self.search_var.get().lower().strip()
        items_to_show = [x for x in self.data if x['category'] == cat]
        
        for i, item in enumerate(items_to_show):
            # ถ้ามีคำค้นหา และชื่อสินค้าไม่ตรง ให้ข้ามไป
            if search_keyword and search_keyword not in item['name'].lower():
                continue
            
            cur = item['stock_remaining']; prev = item['prev_stock']
            usage = max(0, prev - cur) if cur > 0 else 0
            p = f"{item['price']:.2f}" if self.show_prices else "****"
            t = f"{item['price']*(cur if cur>0 else prev):.2f}" if self.show_prices else "****"
            tv.insert("", "end", values=(item.get('item_no','-'), item['name'], prev, cur, usage, item['min_stock_target'], item['order_qty'], p, t))
        self.update_grand_total()

    def on_tree_click(self, event):
        if self.editor: self.save_inline_edit()
        tv = event.widget
        if tv.identify("region", event.x, event.y) != "cell": return
        col = tv.identify_column(event.x); row = tv.identify_row(event.y)
        col_map = {'#1':'item_no', '#2':'name', '#4':'stock_remaining', '#6':'min_stock_target', '#7':'order_qty', '#8':'price'}
        key = col_map.get(col)
        if not key or (key=='price' and not self.show_prices): return
        self.edit_cell(tv, row, key, col)

if __name__ == "__main__":
    root = tk.Tk()
    app = InventoryApp(root)
    root.mainloop()
