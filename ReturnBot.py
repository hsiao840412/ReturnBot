import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
import os
import platform
import subprocess
import threading
import sys
import json
import urllib.request
import webbrowser
import ssl  # 新增：用於處理 SSL 憑證問題
from datetime import datetime
from pathlib import Path

# 引用 xlwings
try:
    import xlwings as xw
except ImportError:
    print("請安裝 xlwings: pip3 install xlwings")

class ReturnBotV1_2:
    def __init__(self, root):
        # === 版本與 GitHub 設定 ===
        self.current_version = "2.2"
        self.github_repo = "hsiao840412/ReturnBot" 
        
        self.root = root
        self.root.title(f"退料機器人 v{self.current_version}") 
        self.root.geometry("520x680")
        self.root.resizable(False, False)

        self.unit_price = 50.00

        # === 模板對照表 ===
        self.template_map = {
            "Mail in": "mail-in template.xlsx",
            "Mail in Battery": "mail-in swollen template.xlsx",
            "KBB": "kbb template.xlsx",
            "KBB Battery": "battery kbb template.xlsx"
        }

        # === 設定 Logo 與路徑 ===
        try:
            if getattr(sys, 'frozen', False):
                base_folder = sys._MEIPASS
            else:
                base_folder = os.path.dirname(os.path.abspath(__file__))
            self.base_folder = base_folder
            
            icon_path = os.path.join(base_folder, "ipsw_logo_200.png") 
            if os.path.exists(icon_path):
                logo_img = tk.PhotoImage(file=icon_path)
                self.root.iconphoto(True, logo_img)
        except:
            self.base_folder = os.getcwd()
            pass

        if platform.system() == "Darwin":
            try:
                root.tk.call('set', '::tk::mac::useCustomTheme', '1') 
            except:
                pass

        self.epacking_path = None
        self.setup_ui()
        
        # 啟動後自動檢查更新
        threading.Thread(target=self.check_for_updates, daemon=True).start()

    def check_for_updates(self):
        """檢查 GitHub 最新 Release，並忽略 SSL 憑證驗證"""
        api_url = f"https://api.github.com/repos/{self.github_repo}/releases/latest"
        try:
            # 建立忽略 SSL 驗證的 context，解決 macOS 常見的證書錯誤
            ctx = ssl.create_default_context()
            ctx.check_hostname = False
            ctx.verify_mode = ssl.CERT_NONE

            # GitHub API 需要 User-Agent 標頭
            req = urllib.request.Request(api_url, headers={'User-Agent': 'Python-ReturnBot-Checker'})
            
            with urllib.request.urlopen(req, timeout=5, context=ctx) as response:
                data = json.loads(response.read().decode())
                # 取得 Tag Name 並去除 'v' 字元
                latest_version = data['tag_name'].replace('v', '').strip()
                download_url = data['html_url']

                # 比對邏輯：若最新版本較新則提示
                if latest_version > self.current_version:
                    self.root.after(0, lambda: self.show_update_dialog(latest_version, download_url))
        except Exception as e:
            # 靜默處理錯誤，僅供 debug
            print(f"Update check failed: {e}")

    def show_update_dialog(self, new_ver, url):
        """跳出更新視窗"""
        msg = f"發現新版本 v{new_ver}！\n(目前版本: v{self.current_version})\n\n是否要前往 GitHub 下載更新？"
        if messagebox.askyesno("更新提醒", msg):
            webbrowser.open(url)

    def setup_ui(self):
        system = platform.system()
        font_main = "PingFang TC" if system == "Darwin" else "Microsoft JhengHei"
        
        style = ttk.Style()
        style.configure("Title.TLabel", font=(font_main, 14, "bold"))
        style.configure("Big.TRadiobutton", font=(font_main, 12))
        style.configure("TButton", font=(font_main, 12))
        style.configure("Green.Horizontal.TProgressbar", foreground='#28CD41', background='#28CD41')
        style.configure("Green.TLabel", font=(font_main, 10), foreground="#008000")
        style.configure("TLabelframe.Label", font=(font_main, 12, "bold"), foreground="white")

        main_frame = ttk.Frame(self.root, padding=20)
        main_frame.pack(fill="both", expand=True)

        # 1. 選擇類型
        type_frame = ttk.LabelFrame(main_frame, text="步驟 1: 選擇退料類型", padding=15)
        type_frame.pack(fill="x", pady=(0, 20))

        self.return_type = tk.StringVar(value="Mail in")
        
        options = [
            ("Mail-in KBB", "Mail in"),
            ("Mail-in 電池膨脹", "Mail in Battery"),
            ("一般 KBB", "KBB"),
            ("單獨鋰電池 KBB", "KBB Battery")
        ]

        for text, val in options:
            ttk.Radiobutton(type_frame, text=text, value=val, variable=self.return_type, style="Big.TRadiobutton").pack(anchor="w", pady=5)

        # 2. 匯入檔案
        file_frame = ttk.LabelFrame(main_frame, text="步驟 2: 匯入 ePacking List", padding=15)
        file_frame.pack(fill="x", pady=(0, 20))

        self.file_label = ttk.Label(file_frame, text="尚未選擇檔案...", foreground="#AAAAAA", font=(font_main, 10))
        self.file_label.pack(side="left", fill="x", expand=True)
        
        ttk.Button(file_frame, text="選擇 CSV", command=self.select_file).pack(side="right")

        # 3. 生成按鈕
        self.gen_btn = ttk.Button(main_frame, text="✨ 啟動 Excel 生成", command=self.start_generation, state="disabled")
        self.gen_btn.pack(fill="x", ipady=15)

        # 4. 進度條
        self.progress = ttk.Progressbar(main_frame, mode='indeterminate', length=400, style="Green.Horizontal.TProgressbar")

        # 5. 狀態標籤
        self.status_label = ttk.Label(main_frame, text="需安裝 Microsoft Excel", anchor="center", style="Green.TLabel")
        self.status_label.pack(pady=20)

    def select_file(self):
        path = filedialog.askopenfilename(filetypes=[("CSV files", "*.csv"), ("All files", "*.*")])
        if path:
            self.epacking_path = path
            self.file_label.config(text=os.path.basename(path), foreground="white")
            self.gen_btn.config(state="normal")

    def start_generation(self):
        if not self.epacking_path:
            return

        self.gen_btn.config(state="disabled")
        self.progress.pack(pady=(20, 5))
        self.progress.start(10)
        self.status_label.config(text="正在呼叫 Excel 編輯中，請稍候...", foreground="#008000")
        
        thread = threading.Thread(target=self.run_excel_task)
        thread.start()

    def get_country_code(self, country_str):
        if pd.isna(country_str): return "CN"
        name = str(country_str).strip()
        mapping = {
            "中國大陸": "CN", "CHINA": "CN", "PRC": "CN", "中国": "CN",
            "台灣": "TW", "TAIWAN": "TW", "ROC": "TW",
            "香港": "HK", "HONG KONG": "HK",
            "澳門": "MO", "MACAU": "MO",
            "新加坡": "SG", "SINGAPORE": "SG",
            "越南": "VN", "VIETNAM": "VN",
            "日本": "JP", "JAPAN": "JP",
            "韓國": "KR", "SOUTH KOREA": "KR", "KOREA": "KR",
            "泰國": "TH", "THAILAND": "TH",
            "馬來西亞": "MY", "MALAYSIA": "MY",
            "菲律賓": "PH", "PHILIPPINES": "PH",
            "印尼": "ID", "INDONESIA": "ID",
            "印度": "IN", "INDIA": "IN",
            "美國": "US", "UNITED STATES": "US", "USA": "US",
            "加拿大": "CA", "CANADA": "CA",
            "英國": "GB", "UNITED KINGDOM": "GB", "UK": "GB",
            "德國": "DE", "GERMANY": "DE",
            "法國": "FR", "FRANCE": "FR",
            "義大利": "IT", "ITALY": "IT",
            "荷蘭": "NL", "NETHERLANDS": "NL",
            "西班牙": "ES", "SPAIN": "ES",
            "瑞士": "CH", "SWITZERLAND": "CH",
            "澳洲": "AU", "澳大利亞": "AU", "AUSTRALIA": "AU",
            "紐西蘭": "NZ", "NEW ZEALAND": "NZ",
            "巴西": "BR", "BRAZIL": "BR",
            "俄羅斯": "RU", "RUSSIA": "RU",
        }
        for key, val in mapping.items():
            if key in name: return val
        return "CN"

    def get_weight(self, row):
        text_to_check = str(row.get('產品名稱', '')) + str(row.get('零件說明', ''))
        if "iPad" in text_to_check or "IPAD" in text_to_check:
            return "0.5"
        return "0.2"

    def generate_dhl_csv(self, df, folder, invoice_no):
        try:
            dhl_data = []
            for i, row in df.iterrows():
                dhl_row = {
                    'A': 1, 'B': 'INV_ITEM',
                    'C': str(row.get('零件說明', '')),
                    'D': '', 'E': 1, 'F': 'PCS', 'G': 50, 'H': 'USD',
                    'I': self.get_weight(row),
                    'J': '',
                    'K': self.get_country_code(row.get('來源國家/地區', ''))
                }
                dhl_data.append(dhl_row)
            
            df_dhl = pd.DataFrame(dhl_data)
            df_dhl = df_dhl[['A','B','C','D','E','F','G','H','I','J','K']]
            
            safe_inv = invoice_no.replace("/", "-").replace("#", "").replace(" ", "_")
            filename = f"DHL_Upload_{safe_inv}.csv"
            save_path = os.path.join(folder, filename)
            
            df_dhl.to_csv(save_path, index=False, header=False, encoding='utf-8-sig')
            return True, filename
        except Exception as e:
            return False, str(e)

    def run_excel_task(self):
        try:
            return_val = self.return_type.get()
            template_filename = self.template_map.get(return_val)
            template_path = os.path.join(self.base_folder, template_filename)

            if not os.path.exists(template_path):
                error_msg = f"找不到模板檔案：{template_filename}\n路徑：{self.base_folder}"
                self.root.after(0, lambda: self.finish_generation(False, error_msg))
                return

            try:
                df = pd.read_csv(self.epacking_path)
            except UnicodeDecodeError:
                df = pd.read_csv(self.epacking_path, encoding='cp950')

            df = df.fillna('')
            
            now = datetime.now()
            today_str = now.strftime("%Y%m%d")
            date_slash = now.strftime("%Y/%m/%d")
            year_dash_month = now.strftime("%Y-%m")

            if return_val == "Mail in":
                invoice_no = f"800935_{today_str}"
                output_filename = f"800935 + HAWB#：Mail in KBB({today_str}).xlsx"
            elif return_val == "Mail in Battery":
                invoice_no = f"SRR#{year_dash_month}T935(電膨)"
                output_filename = f"{invoice_no}.xlsx"
            elif return_val == "KBB":
                invoice_no = f"SRR#{year_dash_month}T935(KBB)"
                output_filename = f"{invoice_no}.xlsx"
            elif return_val == "KBB Battery":
                invoice_no = f"SRR#{year_dash_month}T935(單獨鋰電池)"
                output_filename = f"{invoice_no}.xlsx"
            
            downloads_path = str(Path.home() / "Downloads")
            safe_filename = output_filename.replace("/", "-").replace("\\", "-")
            output_path = os.path.join(downloads_path, safe_filename)

            dhl_generated = False
            if return_val in ["Mail in", "KBB"]:
                dhl_success, dhl_msg = self.generate_dhl_csv(df, downloads_path, invoice_no)
                dhl_generated = dhl_success

            with xw.App(visible=False) as app:
                wb = app.books.open(template_path)
                
                # --- Sheet 1: KBB&KGB invoice ---
                try:
                    sht_inv = wb.sheets['KBB&KGB invoice']
                    sht_inv.range('K1').value = invoice_no
                    sht_inv.range('K2').value = date_slash

                    start_row = 13
                    default_rows = 4
                    target_rows = len(df)
                    diff = target_rows - default_rows
                    
                    if diff > 0:
                        sht_inv.range(f'{start_row + default_rows}:{start_row + default_rows + diff - 1}').insert('down')
                        sht_inv.range(f'{start_row}:{start_row}').copy()
                        sht_inv.range(f'{start_row + 1}:{start_row + target_rows - 1}').paste(paste='formats')
                    elif diff < 0:
                        delete_start = start_row + target_rows
                        delete_end = start_row + default_rows - 1
                        sht_inv.range(f'{delete_start}:{delete_end}').delete()

                    data_to_write = []
                    for i, row in df.iterrows():
                        returns_cell = "KBB"
                        if "KBB" in return_val and "Mail in" not in return_val:
                            raw_returns = row.get('預期退回', 'KBB')
                            returns_cell = str(raw_returns) if raw_returns else 'KBB'

                        data_row = [
                            i + 1,
                            str(row.get('零件', '')),
                            str(row.get('維修', '')),
                            str(row.get('零件說明', '')),
                            None, None, None,
                            1,
                            returns_cell,
                            self.unit_price,
                            self.unit_price,
                            None
                        ]
                        data_to_write.append(data_row)
                    
                    if data_to_write:
                        sht_inv.range(f'A{start_row}').value = data_to_write

                    footer_total_row = 17 + diff
                    footer_qty_row = 19 + diff
                    sht_inv.range(f'J{footer_total_row}').value = "Total:"
                    sum_end_row = 12 + target_rows
                    sht_inv.range(f'K{footer_total_row}').formula = f"=SUM(K13:K{sum_end_row})"
                    sht_inv.range(f'K{footer_qty_row}').value = target_rows
                except: pass

                # --- Sheet 3: ePacking List ---
                try:
                    sht_pack = wb.sheets['ePacking List']
                    sht_pack.range('B1:AD200').value = None
                    sht_pack.range('A2:A200').value = None
                    csv_cols = df.columns.tolist()
                    if csv_cols and "no" in str(csv_cols[0]).lower():
                        final_headers = csv_cols[1:]
                        final_data = df.iloc[:, 1:].fillna('').values.tolist()
                    else:
                        final_headers = csv_cols
                        final_data = df.fillna('').values.tolist()
                    sht_pack.range('B1').value = final_headers
                    sht_pack.range('B2').value = final_data
                    sht_pack.range('A2').value = [[i + 1] for i in range(len(df))]
                except: pass

                # --- Sheet: 條碼 ---
                if return_val == "KBB Battery":
                    try:
                        sht_barcode = wb.sheets['條碼']
                        row_count = len(df)
                        start_row = 4
                        if row_count > 1:
                            insert_range = f'{start_row + 1}:{start_row + row_count - 1}'
                            sht_barcode.range(insert_range).insert('down')
                            sht_barcode.range(f'{start_row}:{start_row}').copy()
                            sht_barcode.range(insert_range).paste()
                        
                        sht_barcode.range(f'A{start_row}').options(transpose=True).value = [i+1 for i in range(row_count)]
                        if '維修' in df.columns:
                            sht_barcode.range(f'B{start_row}').options(transpose=True).value = df['維修'].astype(str).tolist()
                        if '退回訂單' in df.columns:
                            sht_barcode.range(f'D{start_row}').options(transpose=True).value = df['退回訂單'].astype(str).tolist()
                        if '零件' in df.columns:
                            sht_barcode.range(f'E{start_row}').options(transpose=True).value = df['零件'].astype(str).tolist()
                        if '零件說明' in df.columns:
                            sht_barcode.range(f'F{start_row}').options(transpose=True).value = df['零件說明'].astype(str).tolist()
                    except: pass

                wb.save(output_path)
                wb.close()

            final_msg = output_path
            if dhl_generated:
                final_msg += "\n(+ DHL CSV)"
            self.root.after(0, lambda: self.finish_generation(True, final_msg))

        except Exception as e:
            self.root.after(0, lambda: self.finish_generation(False, str(e)))

    def finish_generation(self, success, result_msg):
        self.progress.stop()
        self.progress.pack_forget()
        self.gen_btn.config(state="normal")
        if success:
            lines = result_msg.split('\n')
            filename = os.path.basename(lines[0])
            msg_text = f"檔案已生成：\n{filename}"
            if len(lines) > 1:
                msg_text += "\n(已產生 DHL 上傳檔)"
            self.status_label.config(text="✅ 生成成功！", foreground="#008000")
            if messagebox.askyesno("成功", f"{msg_text}\n\n是否立即打開 Excel？"):
                self.open_file(lines[0])
        else:
            self.status_label.config(text="❌ 發生錯誤", foreground="#FF3B30")
            messagebox.showerror("錯誤", f"發生錯誤：\n{result_msg}")

    def open_file(self, file_path):
        try:
            if platform.system() == "Darwin":
                subprocess.run(["open", file_path], check=True)
            elif platform.system() == "Windows":
                os.startfile(file_path)
            else:
                subprocess.run(["xdg-open", file_path], check=True)
        except: pass

if __name__ == "__main__":
    root = tk.Tk()
    app = ReturnBotV1_2(root)
    root.mainloop()