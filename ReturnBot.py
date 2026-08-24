import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
import os
import platform
import subprocess
import threading
import queue
import sys
import json
import urllib.request
import webbrowser
from datetime import datetime
from pathlib import Path

# 引用 xlwings
try:
    import xlwings as xw
except ImportError:
    xw = None
    print("請安裝 xlwings: pip install xlwings")

class ReturnBotV1_2:
    def __init__(self, root):
        # === 版本與 GitHub 設定 ===
        self.current_version = "2.5"
        self.github_repo = "hsiao840412/ReturnBot"
        
        self.root = root
        self.root.title(f"退料機器人 v{self.current_version}")
        self.root.geometry("520x720") 
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
        self.task_queue = queue.Queue()
        self.setup_ui()
        
        # 啟動後自動檢查更新
        threading.Thread(target=self.check_for_updates, daemon=True).start()

    def check_for_updates(self):
        """檢查 GitHub 最新 Release"""
        api_url = f"https://api.github.com/repos/{self.github_repo}/releases/latest"
        try:
            req = urllib.request.Request(api_url, headers={'User-Agent': 'Python-ReturnBot-Checker'})
            
            with urllib.request.urlopen(req, timeout=5) as response:
                data = json.loads(response.read().decode())
                latest_version = data['tag_name'].removeprefix('v').strip()
                download_url = data['html_url']
                if self.version_tuple(latest_version) > self.version_tuple(self.current_version):
                    self.root.after(0, lambda: self.show_update_dialog(latest_version, download_url))
        except Exception as e:
            print(f"Update check failed: {e}")

    @staticmethod
    def version_tuple(version):
        """將 2.10.1 轉成可正確比較的數字 tuple。"""
        try:
            return tuple(int(part) for part in version.split('.'))
        except (TypeError, ValueError):
            raise ValueError(f"無法辨識版本號：{version}")

    def show_update_dialog(self, new_ver, url):
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
        style.configure("Hint.TLabel", font=(font_main, 10), foreground="#888888") 
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
        self.status_label.pack(pady=(20, 5))

        # 6. 新增：儲存路徑提示
        self.path_hint = ttk.Label(main_frame, text="💡 提示：檔案會自動保存在「下載項目」中", anchor="center", style="Hint.TLabel")
        self.path_hint.pack(pady=(0, 10))

    def select_file(self):
        path = filedialog.askopenfilename(filetypes=[("CSV files", "*.csv"), ("All files", "*.*")])
        if path:
            self.epacking_path = path
            self.file_label.config(text=os.path.basename(path), foreground="white")
            self.gen_btn.config(state="normal")

    def start_generation(self):
        if not self.epacking_path: return
        if xw is None:
            messagebox.showerror("錯誤", "缺少 xlwings，請先安裝後再試。")
            return
        self.gen_btn.config(state="disabled")
        self.progress.pack(pady=(20, 5))
        self.progress.start(10)
        self.status_label.config(text="正在讀取 CSV...", foreground="#008000")
        return_val = self.return_type.get()
        epacking_path = self.epacking_path
        self.clear_task_queue()
        threading.Thread(
            target=self.run_excel_task,
            args=(return_val, epacking_path),
            daemon=True
        ).start()
        self.root.after(100, self.poll_task_queue)

    def clear_task_queue(self):
        """清除上一次任務可能殘留的訊息。"""
        while True:
            try:
                self.task_queue.get_nowait()
            except queue.Empty:
                break

    def report_status(self, message):
        """由背景執行緒安全地回報處理進度。"""
        self.task_queue.put(("status", message))

    def report_result(self, success, message, warnings=None):
        """由背景執行緒安全地回報最終結果。"""
        self.task_queue.put(("result", success, message, warnings or []))

    def poll_task_queue(self):
        """只在 Tk 主執行緒中更新介面。"""
        finished = False
        try:
            while True:
                item = self.task_queue.get_nowait()
                if item[0] == "status":
                    self.status_label.config(text=item[1], foreground="#008000")
                elif item[0] == "result":
                    finished = True
                    self.finish_generation(item[1], item[2], item[3])
        except queue.Empty:
            pass

        if not finished:
            self.root.after(100, self.poll_task_queue)

    def get_country_code(self, country_str):
        return self.resolve_country_code(country_str)[0]

    def resolve_country_code(self, country_str):
        """返回 (DHL 國家代碼, 未辨識的原始國家)。"""
        if pd.isna(country_str):
            return "CN", "(空白)"

        original_name = str(country_str).strip()
        if not original_name:
            return "CN", "(空白)"
        name = original_name.upper()
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
            if key.upper() in name:
                return val, None
        return "CN", original_name

    def get_weight(self, row):
        text_to_check = str(row.get('產品名稱', '')) + str(row.get('零件說明', ''))
        return "0.5" if "IPAD" in text_to_check.upper() else "0.2"

    def validate_dataframe(self, df, return_val):
        """驗證 CSV 是否具備該退料類型所需的資料。"""
        if df.empty:
            raise ValueError("CSV 沒有任何可處理的資料。")

        required_columns = {'零件', '零件說明'}
        if return_val in {'KBB', 'KBB Battery'}:
            required_columns.add('退回訂單')
        else:
            required_columns.add('維修')
        if return_val in {'Mail in', 'KBB'}:
            required_columns.add('來源國家/地區')

        missing = sorted(required_columns - set(df.columns))
        if missing:
            raise ValueError(f"CSV 缺少必要欄位：{'、'.join(missing)}")

        key_column = '退回訂單' if return_val in {'KBB', 'KBB Battery'} else '維修'
        empty_columns = [
            column for column in ('零件', '零件說明', key_column)
            if df[column].astype(str).str.strip().eq('').any()
        ]
        if empty_columns:
            raise ValueError(f"CSV 的必要欄位含有空白值：{'、'.join(empty_columns)}")

    def generate_dhl_csv(self, df, folder, invoice_no):
        try:
            dhl_data = []
            unknown_countries = set()
            for i, row in df.iterrows():
                country_code, unknown_country = self.resolve_country_code(row.get('來源國家/地區', ''))
                if unknown_country:
                    unknown_countries.add(unknown_country)
                dhl_row = {
                    'A': 1, 'B': 'INV_ITEM',
                    'C': str(row.get('零件說明', '')),
                    'D': '', 'E': 1, 'F': 'PCS', 'G': 50, 'H': 'USD',
                    'I': self.get_weight(row),
                    'J': f"未辨識國家：{unknown_country}" if unknown_country else '',
                    'K': country_code
                }
                dhl_data.append(dhl_row)
            df_dhl = pd.DataFrame(dhl_data)[['A','B','C','D','E','F','G','H','I','J','K']]
            safe_inv = invoice_no.replace("/", "-").replace("#", "").replace(" ", "_")
            requested_path = os.path.join(folder, f"DHL_Upload_{safe_inv}.csv")
            output_path = self.get_unique_path(requested_path)
            df_dhl.to_csv(output_path, index=False, header=False, encoding='utf-8-sig')
            return True, os.path.basename(output_path), sorted(unknown_countries)
        except Exception as e:
            return False, str(e), []

    @staticmethod
    def get_unique_path(path):
        """若檔案已存在，自動加上 (2)、(3) 以避免覆蓋。"""
        if not os.path.exists(path):
            return path

        folder, filename = os.path.split(path)
        stem, suffix = os.path.splitext(filename)
        number = 2
        while True:
            candidate = os.path.join(folder, f"{stem} ({number}){suffix}")
            if not os.path.exists(candidate):
                return candidate
            number += 1

    def run_excel_task(self, return_val, epacking_path):
        try:
            self.report_status("正在讀取 CSV...")
            template_filename = self.template_map.get(return_val)
            if not template_filename:
                raise ValueError(f"不支援的退料類型：{return_val}")
            template_path = os.path.join(self.base_folder, template_filename)

            if not os.path.exists(template_path):
                raise FileNotFoundError(f"找不到模板：{template_filename}")

            try:
                df = pd.read_csv(epacking_path)
            except UnicodeDecodeError:
                df = pd.read_csv(epacking_path, encoding='cp950')
            df = df.fillna('')
            self.report_status("正在驗證資料...")
            self.validate_dataframe(df, return_val)
            
            now = datetime.now()
            today_str, date_slash, year_dash_month = now.strftime("%Y%m%d"), now.strftime("%Y/%m/%d"), now.strftime("%Y-%m")

            if return_val == "Mail in":
                invoice_no, output_filename = f"800935_{today_str}", f"800935 + HAWB#：Mail in KBB({today_str}).xlsx"
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
            requested_path = os.path.join(downloads_path, output_filename.replace("/", "-").replace("\\", "-"))
            output_path = self.get_unique_path(requested_path)

            self.report_status("正在啟動 Excel...")
            with xw.App(visible=False) as app:
                wb = None
                try:
                    wb = app.books.open(template_path)

                    # --- Sheet 1: KBB&KGB invoice ---
                    self.report_status("正在填寫發票資料...")
                    sht_inv = wb.sheets['KBB&KGB invoice']
                    sht_inv.range('K1').value = invoice_no
                    sht_inv.range('K2').value = date_slash

                    start_row = 13
                    default_rows = 3  
                    target_rows = len(df)
                    diff = target_rows - default_rows
                    
                    if diff > 0:
                        sht_inv.range(f'{start_row + default_rows}:{start_row + default_rows + diff - 1}').insert('down')
                        sht_inv.range(f'{start_row}:{start_row}').copy()
                        sht_inv.range(f'{start_row + 1}:{start_row + target_rows - 1}').paste(paste='formats')
                    elif diff < 0:
                        sht_inv.range(f'{start_row + target_rows}:{start_row + default_rows - 1}').delete()

                    data_to_write = []
                    for i, row in df.iterrows():
                        returns_cell = str(row.get('預期退回', 'KBB')) if "KBB" in return_val and "Mail in" not in return_val else "KBB"
                        
                        # [修正重點]：只要是 KBB 相關模式（一般 KBB 或鋰電池 KBB），RMA# 皆抓「退回訂單」
                        if return_val in ["KBB", "KBB Battery"]:
                            rma_value = str(row.get('退回訂單', ''))
                        else:
                            rma_value = str(row.get('維修', ''))
                        
                        data_to_write.append([
                            i + 1, 
                            str(row.get('零件', '')), 
                            rma_value, 
                            str(row.get('零件說明', '')), 
                            None, None, None, 1, 
                            returns_cell, 
                            self.unit_price, 
                            self.unit_price, 
                            None
                        ])
                    
                    if data_to_write: sht_inv.range(f'A{start_row}').value = data_to_write

                    footer_total_row, footer_qty_row = 16 + diff, 18 + diff
                    sht_inv.range(f'J{footer_total_row}').value = "Total:"
                    sht_inv.range(f'K{footer_total_row}').formula = f"=SUM(K13:K{12 + target_rows})"
                    sht_inv.range(f'K{footer_qty_row}').value = target_rows
                    # --- Sheet 3: ePacking List ---
                    self.report_status("正在填寫 ePacking List...")
                    try:
                        sht_pack = wb.sheets['ePacking List']
                        sht_pack.range('A2:AD200').value = None
                        csv_cols = df.columns.tolist()
                        final_headers = csv_cols[1:] if csv_cols and "no" in str(csv_cols[0]).lower() else csv_cols
                        final_data = df.iloc[:, 1:].fillna('').values.tolist() if csv_cols and "no" in str(csv_cols[0]).lower() else df.fillna('').values.tolist()
                        sht_pack.range('B1').value = final_headers
                        sht_pack.range('B2').value = final_data
                        sht_pack.range('A2').value = [[i + 1] for i in range(len(df))]
                    except Exception as e:
                        raise RuntimeError(f"寫入「ePacking List」工作表失敗：{e}") from e

                    # --- Sheet: 條碼 ---
                    if return_val == "KBB Battery":
                        self.report_status("正在填寫條碼資料...")
                        try:
                            sht_barcode = wb.sheets['條碼']
                            row_count = len(df)
                            if row_count > 1:
                                sht_barcode.range(f'5:{5 + row_count - 2}').insert('down')
                                sht_barcode.range('4:4').copy()
                                sht_barcode.range(f'5:{5 + row_count - 2}').paste()
                            sht_barcode.range('A4').options(transpose=True).value = [i+1 for i in range(row_count)]
                            if '維修' in df.columns: sht_barcode.range('B4').options(transpose=True).value = df['維修'].astype(str).tolist()
                            if '退回訂單' in df.columns: sht_barcode.range('D4').options(transpose=True).value = df['退回訂單'].astype(str).tolist()
                            if '零件' in df.columns: sht_barcode.range('E4').options(transpose=True).value = df['零件'].astype(str).tolist()
                            if '零件說明' in df.columns: sht_barcode.range('F4').options(transpose=True).value = df['零件說明'].astype(str).tolist()
                        except Exception as e:
                            raise RuntimeError(f"寫入「條碼」工作表失敗：{e}") from e

                    self.report_status("正在儲存 Excel...")
                    wb.save(output_path)
                finally:
                    if wb is not None:
                        wb.close()

            dhl_generated = False
            warnings = []
            if return_val in ["Mail in", "KBB"]:
                self.report_status("正在生成 DHL CSV...")
                dhl_success, dhl_result, unknown_countries = self.generate_dhl_csv(df, downloads_path, invoice_no)
                if not dhl_success:
                    raise RuntimeError(f"Excel 已儲存，但 DHL CSV 生成失敗：{dhl_result}")
                dhl_generated = True
                if unknown_countries:
                    countries = "、".join(unknown_countries)
                    warnings.append(
                        f"無法辨識以下國家：{countries}\n"
                        "DHL CSV 已照常輸出，國家代碼暫用 CN，"
                        "原始國家已寫入 J 欄備註。"
                    )

            result = f"{output_path}\n(+ DHL CSV)" if dhl_generated else output_path
            self.report_result(True, result, warnings)
        except Exception as e:
            self.report_result(False, str(e))

    def finish_generation(self, success, result_msg, warnings=None):
        self.progress.stop()
        self.progress.pack_forget()
        self.gen_btn.config(state="normal")
        if success:
            lines = result_msg.split('\n')
            msg_text = f"檔案已生成：\n{os.path.basename(lines[0])}" + ("\n(已產生 DHL 上傳檔)" if len(lines) > 1 else "")
            self.status_label.config(text="✅ 生成成功！", foreground="#008000")
            if warnings:
                messagebox.showwarning("國家資料提醒", "\n\n".join(warnings))
            if messagebox.askyesno("成功", f"{msg_text}\n\n是否立即打開 Excel？"): self.open_file(lines[0])
        else:
            self.status_label.config(text="❌ 發生錯誤", foreground="#FF3B30")
            messagebox.showerror("錯誤", f"發生錯誤：\n{result_msg}")

    def open_file(self, file_path):
        try:
            if platform.system() == "Darwin": subprocess.run(["open", file_path], check=True)
            elif platform.system() == "Windows": os.startfile(file_path)
            else: subprocess.run(["xdg-open", file_path], check=True)
        except: pass

if __name__ == "__main__":
    root = tk.Tk()
    app = ReturnBotV1_2(root)
    root.mainloop()
