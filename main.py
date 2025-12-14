import sys
import os
import re
import shutil
import xlrd
import openpyxl
from datetime import datetime

from PySide6.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, 
                               QHBoxLayout, QPushButton, QLabel, QProgressBar, 
                               QTableWidget, QTableWidgetItem, QFileDialog, 
                               QMessageBox, QHeaderView, QAbstractItemView)
from PySide6.QtCore import Qt, QThread, Signal, QSettings, QUrl
from PySide6.QtGui import QColor, QDesktopServices, QAction

# ==========================================
# 1. HELPER FUNCTIONS
# ==========================================

def sanitize_filename(name):
    """Membersihkan karakter ilegal untuk nama file"""
    if not name: return "Unknown"
    clean = re.sub(r'[\\/*?:"<>|]', "", str(name)).strip()
    return clean if clean else "Unknown"

def clean_currency(value):
    if value in [None, ""]: return 0
    if isinstance(value, (int, float)): return value
    str_val = str(value).replace("Rp", "").replace(".", "").replace(",", ".").strip()
    try: return float(str_val)
    except: return 0

def extract_year_from_date(date_str):
    try:
        if not date_str: return str(datetime.now().year)
        parts = date_str.split('-')
        if len(parts) == 3:
            yy = parts[2]
            return f"20{yy}" if len(yy) == 2 else yy
    except:
        pass
    return str(datetime.now().year)

# --- PARSER KHUSUS .XLS (XLRD) ---
def parse_xls_classic(filepath):
    try:
        wb = xlrd.open_workbook(filepath, formatting_info=False)
        sheet = wb.sheet_by_index(0)
        
        def get_xls_val(addr):
            match = re.match(r"([A-Z]+)([0-9]+)", addr.upper())
            if not match: return None
            col_str, row_str = match.groups()
            row = int(row_str) - 1
            col = 0
            for char in col_str:
                col = col * 26 + (ord(char) - ord('A') + 1)
            col -= 1
            try: return sheet.cell_value(row, col)
            except: return None

        def get_xls_date_pack(addr):
            val = get_xls_val(addr)
            if isinstance(val, float):
                try:
                    dt_tuple = xlrd.xldate_as_tuple(val, wb.datemode)
                    dt_obj = datetime(*dt_tuple)
                    return dt_obj.strftime("%d-%b-%y"), dt_obj
                except: pass
            return str(val), datetime.min

        date_str, date_obj = get_xls_date_pack("B3")
        sub_total = 0; penalty = 0; warranty = 0; total_cost = 0; cm_booked = 0; cr_booked = 0
        limit = min(sheet.nrows, 150)
        
        for r in range(9, limit):
            try:
                raw = sheet.cell_value(r, 0)
                txt = str(raw).upper() if raw else ""
            except: continue
            
            if "SUB TOTAL" in txt and sub_total == 0: sub_total = sheet.cell_value(r, 4)
            elif "PENALTY" in txt and penalty == 0: penalty = sheet.cell_value(r, 4)
            elif "WARRANTTY" in txt and warranty == 0: warranty = sheet.cell_value(r, 4)
            elif "TOTAL COST" in txt and total_cost == 0: total_cost = sheet.cell_value(r, 4)
            elif "CM BOOKED" in txt and cm_booked == 0: cm_booked = sheet.cell_value(r, 4)
            elif "CR BOOKED" in txt and cr_booked == 0: cr_booked = sheet.cell_value(r, 4)

        return {
            "status": "OK",
            "_sort_date": date_obj,
            "Project No": get_xls_val("K4"),
            "Cust Name": get_xls_val("K3"),
            "Proj Date": date_str,
            "Kurs": clean_currency(get_xls_val("B4")),
            "Project Value": clean_currency(get_xls_val("B5")),
            "Sub Total": clean_currency(sub_total),
            "Penalty": clean_currency(penalty),
            "Warranty": clean_currency(warranty),
            "Total Cost": clean_currency(total_cost),
            "CM Booked": clean_currency(cm_booked),
            "CR Booked": clean_currency(cr_booked)
        }
    except Exception as e:
        return {"status": "ERROR", "msg": str(e), "_sort_date": datetime.min}

# --- PARSER KHUSUS .XLSX (OPENPYXL) ---
def parse_xlsx_modern(filepath):
    try:
        wb = openpyxl.load_workbook(filepath, data_only=True)
        sheet = wb.active
        
        def get_xlsx_val(addr):
            return sheet[addr].value

        def get_xlsx_date_pack(addr):
            val = sheet[addr].value
            if isinstance(val, datetime):
                return val.strftime("%d-%b-%y"), val
            return str(val) if val else "", datetime.min

        date_str, date_obj = get_xlsx_date_pack("B3")
        sub_total = 0; penalty = 0; warranty = 0; total_cost = 0; cm_booked = 0; cr_booked = 0
        limit = min(sheet.max_row, 150)
        
        for r in range(10, limit + 1):
            try:
                cell_val = sheet.cell(row=r, column=1).value
                txt = str(cell_val).upper() if cell_val else ""
            except: continue
            
            val_col = 5
            if "SUB TOTAL" in txt and sub_total == 0: sub_total = sheet.cell(row=r, column=val_col).value
            elif "PENALTY" in txt and penalty == 0: penalty = sheet.cell(row=r, column=val_col).value
            elif "WARRANTY" in txt and warranty == 0: warranty = sheet.cell(row=r, column=val_col).value
            elif "TOTAL COST" in txt and total_cost == 0: total_cost = sheet.cell(row=r, column=val_col).value
            elif "CM BOOKED" in txt and cm_booked == 0: cm_booked = sheet.cell(row=r, column=val_col).value
            elif "CR BOOKED" in txt and cr_booked == 0: cr_booked = sheet.cell(row=r, column=val_col).value

        return {
            "status": "OK",
            "_sort_date": date_obj,
            "Project No": get_xlsx_val("K4"),
            "Cust Name": get_xlsx_val("K3"),
            "Proj Date": date_str,
            "Kurs": clean_currency(get_xlsx_val("B4")),
            "Project Value": clean_currency(get_xlsx_val("B5")),
            "Sub Total": clean_currency(sub_total),
            "Penalty": clean_currency(penalty),
            "Warranty": clean_currency(warranty),
            "Total Cost": clean_currency(total_cost),
            "CM Booked": clean_currency(cm_booked),
            "CR Booked": clean_currency(cr_booked)
        }

    except Exception as e:
        return {"status": "ERROR", "msg": str(e), "_sort_date": datetime.min}

def extract_dispatcher(filepath):
    ext = os.path.splitext(filepath)[1].lower()
    data = {}
    
    if ext == ".xls":
        data = parse_xls_classic(filepath)
    elif ext == ".xlsx":
        data = parse_xlsx_modern(filepath)
    else:
        return {"status": "SKIP", "msg": "Format tidak didukung", "_sort_date": datetime.min}
    
    if data["status"] == "OK":
        if not data.get("Project No"):
            data["status"] = "PARSING ERROR"
            data["msg"] = "Project No Kosong"
    
    return data

# ==========================================
# 2. WORKER THREAD (PREVIEW SCANNER + SORT + DEDUP)
# ==========================================

class PreviewWorker(QThread):
    progress = Signal(int)
    finished = Signal(list)
    
    def __init__(self, folder_path):
        super().__init__()
        self.folder_path = folder_path
        
    def run(self):
        try:
            all_files = [f for f in os.listdir(self.folder_path) if f.lower().endswith(('.xls', '.xlsx'))]
            all_files = [f for f in all_files if not f.startswith("~$")]
        except:
            all_files = []
        
        total = len(all_files)
        results = []
        
        # 1. PARSE SEMUA FILE
        for i, filename in enumerate(all_files):
            path = os.path.join(self.folder_path, filename)
            data = extract_dispatcher(path)
            data["filename"] = filename
            data["path"] = path
            results.append(data)
            if total > 0: self.progress.emit(int((i+1)/total * 100))

        # 2. LOGIKA DUPLIKAT
        id_counts = {}
        for item in results:
            if item["status"] == "OK":
                pid = str(item.get("Project No", "")).strip()
                if pid:
                    id_counts[pid] = id_counts.get(pid, 0) + 1
        
        for item in results:
            if item["status"] == "OK":
                pid = str(item.get("Project No", "")).strip()
                if pid and id_counts.get(pid, 0) > 1:
                    item["status"] = "DUPLIKAT" 
        
        # 3. SORTING BY DATE
        results.sort(key=lambda x: x.get("_sort_date", datetime.min))
        self.finished.emit(results)

# ==========================================
# 3. GENERATOR THREAD
# ==========================================

class GeneratorWorker(QThread):
    log_msg = Signal(str)
    finished = Signal(str)
    
    def __init__(self, data_list, output_folder):
        super().__init__()
        self.data_list = data_list
        self.output_folder = output_folder
        
    def run(self):
        self.log_msg.emit("🚀 Memulai proses generate...")
        
        valid_data = [d for d in self.data_list if d["status"] in ["OK", "DUPLIKAT"]]
        copied_count = 0
        created_paths = set()
        
        # 1. COPY & RENAME FILES
        for item in valid_data:
            try:
                old_path = item["path"]
                ext = os.path.splitext(old_path)[1]
                p_id = sanitize_filename(item['Project No'])
                cust = sanitize_filename(item['Cust Name'])
                year = extract_year_from_date(item.get('Proj Date'))
                
                base_name = f"PCM {p_id} {year} {cust}"
                new_name = f"{base_name}{ext}"
                new_path = os.path.join(self.output_folder, new_name)
                
                counter = 1
                while new_path in created_paths or os.path.exists(new_path) and new_path not in created_paths:
                    if new_path not in created_paths: break
                    new_name = f"{base_name} ({counter}){ext}"
                    new_path = os.path.join(self.output_folder, new_name)
                    counter += 1

                created_paths.add(new_path)
                shutil.copy2(old_path, new_path)
                copied_count += 1
                
            except Exception as e:
                self.log_msg.emit(f"❌ Gagal copy {item['filename']}: {e}")

        self.log_msg.emit(f"✅ Berhasil menyalin {copied_count} file.")

        # 2. GENERATE SUMMARY EXCEL
        try:
            self.log_msg.emit("📊 Membuat file summary...")
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "PCM SUMMARY"
            
            headers = ["No", "Project No", "Proj Date", "Cust Name", "Project Value", 
                       "Kurs", "BARANG&JASA", "Penalty", "Warranty", "Total Cost", 
                       "CM Booked", "CR Booked", "CM IDR", "CM %", "Status File"]
            
            # Styles
            header_font = openpyxl.styles.Font(bold=True, color="FFFFFF")
            header_fill = openpyxl.styles.PatternFill("solid", fgColor="2196F3")
            
            # --- COLOR MARKING UNTUK OUTPUT (KUNING) ---
            duplicate_fill = openpyxl.styles.PatternFill("solid", fgColor="FFFF00") # Kuning
            
            border = openpyxl.styles.Border(left=openpyxl.styles.Side(style='thin'), 
                                            right=openpyxl.styles.Side(style='thin'), 
                                            top=openpyxl.styles.Side(style='thin'), 
                                            bottom=openpyxl.styles.Side(style='thin'))

            ws.append(headers)
            for cell in ws[1]:
                cell.font = header_font
                cell.fill = header_fill
                cell.alignment = openpyxl.styles.Alignment(horizontal='center')
            
            for idx, item in enumerate(valid_data, 1):
                row = idx + 1
                f_cm_idr = f"=E{row}-J{row}"
                f_cm_pct = f"=IF(E{row}=0, 0, M{row}/E{row})"
                
                file_status = "COPIED"
                if item["status"] == "DUPLIKAT":
                    file_status = "DUPLICATE INPUT"

                row_data = [
                    idx, item["Project No"], item["Proj Date"], item["Cust Name"], item["Project Value"],
                    "IDR", item["Sub Total"], item["Penalty"], item["Warranty"], item["Total Cost"],
                    item["CM Booked"], item["CR Booked"], f_cm_idr, f_cm_pct, file_status
                ]
                ws.append(row_data)
                
                for c, val in enumerate(row_data, 1):
                    cell = ws.cell(row=row, column=c)
                    cell.border = border
                    if c in [5, 7, 8, 9, 10, 11, 12, 13]: cell.number_format = '#,##0'
                    if c == 14: cell.number_format = '0.00%'
                    
                    # --- APPLY WARNA KUNING JIKA DUPLIKAT ---
                    if item["status"] == "DUPLIKAT":
                        cell.fill = duplicate_fill

            current_year = datetime.now().year
            summary_name = f"PCM SUMMARY {current_year}.xlsx"
            summary_path = os.path.join(self.output_folder, summary_name)
            
            wb.save(summary_path)
            self.finished.emit(summary_path) # Return path summary

        except Exception as e:
            self.finished.emit(f"ERROR: {e}")

# ==========================================
# 4. MAIN WINDOW
# ==========================================

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("PCM Generator v1.0.0")
        self.resize(1000, 650)
        
        # Registry Settings (Untuk ingat folder terakhir)
        self.settings = QSettings("FahmiSoft", "PCMGenerator")
        
        central = QWidget()
        self.setCentralWidget(central)
        layout = QVBoxLayout(central)
        
        # --- INPUT OUTPUT SELECTION ---
        grp_io = QWidget()
        layout_io = QVBoxLayout(grp_io)
        
        h1 = QHBoxLayout()
        self.lbl_input = QLabel("Input Folder: (Belum Dipilih)")
        self.lbl_input.setStyleSheet("background: #e3f2fd; padding: 5px; border-radius: 4px;")
        btn_input = QPushButton("1. Pilih Folder INPUT")
        btn_input.clicked.connect(self.select_input)
        h1.addWidget(self.lbl_input); h1.addWidget(btn_input)
        
        h2 = QHBoxLayout()
        self.lbl_output = QLabel("Output Folder: (Belum Dipilih)")
        self.lbl_output.setStyleSheet("background: #e8f5e9; padding: 5px; border-radius: 4px;")
        btn_output = QPushButton("2. Pilih Folder OUTPUT")
        btn_output.clicked.connect(self.select_output)
        h2.addWidget(self.lbl_output); h2.addWidget(btn_output)
        
        layout_io.addLayout(h1); layout_io.addLayout(h2)
        layout.addWidget(grp_io)
        
        # --- TABLE ---
        self.table = QTableWidget()
        cols = ["Nama File Asli", "Project ID", "Customer", "Date", "Nilai Project", "Status"]
        self.table.setColumnCount(len(cols))
        self.table.setHorizontalHeaderLabels(cols)
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        self.table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        layout.addWidget(self.table)
        
        # --- ACTION ---
        h3 = QHBoxLayout()
        self.progress = QProgressBar()
        self.btn_gen = QPushButton("3. GENERATE FILES")
        self.btn_gen.setStyleSheet("QPushButton { background-color: #2196F3; color: white; font-weight: bold; padding: 10px; } QPushButton:disabled { background-color: #ccc; }")
        self.btn_gen.setEnabled(False)
        self.btn_gen.clicked.connect(self.start_generation)
        h3.addWidget(self.progress); h3.addWidget(self.btn_gen)
        layout.addLayout(h3)
        
        # --- STATUS BAR (NEW) ---
        self.setup_statusbar()
        
        self.input_dir = ""; self.output_dir = ""; self.data_cache = []
        self.scan_worker = None; self.gen_worker = None

        # Load Last Settings
        self.load_settings()

    def setup_statusbar(self):
        status_bar = self.statusBar()
        
        # Kiri: Versi
        lbl_version = QLabel("  PCM Summary Generator v. 1.0.0  ")
        status_bar.addWidget(lbl_version)
        
        # Kanan: About
        btn_about = QPushButton("About Application")
        btn_about.setFlat(True)
        btn_about.setStyleSheet("font-weight: bold; color: #555;")
        btn_about.clicked.connect(self.show_about_dialog)
        status_bar.addPermanentWidget(btn_about)

    def load_settings(self):
        # Restore last input/output folder
        last_in = self.settings.value("last_input_dir")
        last_out = self.settings.value("last_output_dir")
        
        if last_in and os.path.exists(last_in):
            self.input_dir = last_in
            self.lbl_input.setText(last_in)
            self.run_preview_scan() # Auto scan jika ada history
            
        if last_out and os.path.exists(last_out):
            self.output_dir = last_out
            self.lbl_output.setText(last_out)
            self.check_ready()

    def select_input(self):
        start_dir = self.input_dir if self.input_dir else ""
        path = QFileDialog.getExistingDirectory(self, "Pilih Input Folder", start_dir)
        if path:
            self.input_dir = path
            self.lbl_input.setText(path)
            self.settings.setValue("last_input_dir", path) # Save setting
            self.run_preview_scan()

    def select_output(self):
        start_dir = self.output_dir if self.output_dir else ""
        path = QFileDialog.getExistingDirectory(self, "Pilih Output Folder", start_dir)
        if path:
            self.output_dir = path
            self.lbl_output.setText(path)
            self.settings.setValue("last_output_dir", path) # Save setting
            self.check_ready()

    def run_preview_scan(self):
        self.table.setRowCount(0)
        self.btn_gen.setEnabled(False)
        self.scan_worker = PreviewWorker(self.input_dir)
        self.scan_worker.progress.connect(self.progress.setValue)
        self.scan_worker.finished.connect(self.on_preview_done)
        self.scan_worker.start()
        
    def on_preview_done(self, results):
        self.data_cache = results
        self.table.setRowCount(len(results))
        for r, item in enumerate(results):
            status = item["status"]
            color = QColor(Qt.black)
            bg_color = QColor(Qt.white)
            
            if status == "DUPLIKAT":
                bg_color = QColor("#FFEB3B"); status = "DUPLIKAT (Diproses)"
            elif status != "OK":
                bg_color = QColor("#FFCDD2"); color = QColor(Qt.red)
            
            def make_item(text):
                it = QTableWidgetItem(str(text))
                it.setForeground(color); it.setBackground(bg_color)
                return it

            self.table.setItem(r, 0, make_item(item["filename"]))
            self.table.setItem(r, 1, make_item(item.get("Project No", "-")))
            self.table.setItem(r, 2, make_item(item.get("Cust Name", "-")))
            self.table.setItem(r, 3, make_item(item.get("Proj Date", "-")))
            val = item.get("Project Value", 0)
            fmt_val = f"{val:,.0f}" if isinstance(val, (int, float)) else str(val)
            self.table.setItem(r, 4, make_item(fmt_val))
            self.table.setItem(r, 5, make_item(status))
        
        self.check_ready()

    def check_ready(self):
        valid_count = sum(1 for d in self.data_cache if d["status"] in ["OK", "DUPLIKAT"])
        is_ready = bool(self.input_dir and self.output_dir and valid_count > 0)
        self.btn_gen.setEnabled(is_ready)
        if is_ready: self.btn_gen.setText(f"3. GENERATE ({valid_count} File Valid)")
        else: self.btn_gen.setText("3. GENERATE (Menunggu Input/Output)")

    def start_generation(self):
        if not self.output_dir: return

        try:
            files_in_output = [f for f in os.listdir(self.output_dir) if not f.startswith('.')]
            if files_in_output:
                msg_box = QMessageBox(self)
                msg_box.setIcon(QMessageBox.Warning)
                msg_box.setWindowTitle("Folder Output Tidak Kosong")
                msg_box.setText(f"Folder Output berisi {len(files_in_output)} file/folder.")
                msg_box.setInformativeText("File lama mungkin akan tertimpa.\nLanjutkan?")
                btn_overwrite = msg_box.addButton("Lanjutkan (Overwrite)", QMessageBox.AcceptRole)
                btn_cancel = msg_box.addButton("Batalkan", QMessageBox.RejectRole)
                msg_box.exec()
                if msg_box.clickedButton() == btn_cancel: return
        except Exception as e:
            QMessageBox.critical(self, "Error", str(e)); return
        
        self.gen_worker = GeneratorWorker(self.data_cache, self.output_dir)
        self.gen_worker.log_msg.connect(lambda s: self.progress.setFormat(s))
        self.gen_worker.finished.connect(self.on_generation_finished)
        self.progress.setValue(0); self.progress.setRange(0, 0)
        self.btn_gen.setEnabled(False)
        self.gen_worker.start()

    def on_generation_finished(self, result_msg):
        self.progress.setRange(0, 100); self.progress.setValue(100); self.progress.setFormat("Selesai")
        self.btn_gen.setEnabled(True); self.btn_gen.setText("GENERATE SELESAI")
        
        # Cek apakah result_msg berisi path (sukses) atau error
        if "ERROR:" in result_msg:
            QMessageBox.critical(self, "Gagal", result_msg)
        else:
            # Sukses
            reply = QMessageBox.question(self, "Sukses", 
                                         f"Generate selesai!\nFile Summary:\n{result_msg}\n\n"
                                         "Buka folder output di Windows Explorer?",
                                         QMessageBox.Yes | QMessageBox.No)
            
            if reply == QMessageBox.Yes:
                # Buka Folder Output
                QDesktopServices.openUrl(QUrl.fromLocalFile(self.output_dir))

    def show_about_dialog(self):
        info = (
            "<h3>PCM Summary Generator</h3>"
            "<p>Version: 1.0.0</p>"
            "<p>Aplikasi untuk rekapitulasi otomatis Project Cost Management.</p>"
            "<hr>"
            "<p><b>Developer:</b> Fahmi Fauzi Rahman</p>"
            "<p><b>Contact:</b> 0853-1740-4760</p>"
            "<hr>"
            "<p><b>Credits / Libraries Used:</b></p>"
            "<ul>"
            "<li>Python 3</li>"
            "<li>PySide6 (Qt for Python)</li>"
            "<li>openpyxl (Excel Modern Support)</li>"
            "<li>xlrd (Excel Classic Support)</li>"
            "</ul>"
        )
        QMessageBox.about(self, "About Application", info)

if __name__ == "__main__":
    app = QApplication(sys.argv)
    app.setStyle("Fusion")
    w = MainWindow()
    w.show()
    sys.exit(app.exec())