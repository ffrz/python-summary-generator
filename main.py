import sys
import os
import re
import xlrd
from datetime import datetime
import time

# Library Watchdog
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler

from PySide6.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, 
                               QHBoxLayout, QPushButton, QLabel, QProgressBar, 
                               QTableWidget, QTableWidgetItem, QFileDialog, QMessageBox, QHeaderView, QCheckBox)
from PySide6.QtCore import Qt, QThread, Signal, QTimer

# ==========================================
# BAGIAN 1: HELPER & PARSER
# ==========================================

def addr_to_idx(address):
    match = re.match(r"([A-Z]+)([0-9]+)", address.upper())
    if not match: raise ValueError(f"Bad Addr: {address}")
    col_str, row_str = match.groups()
    row_idx = int(row_str) - 1
    col_idx = 0
    for char in col_str:
        col_idx = col_idx * 26 + (ord(char) - ord('A') + 1)
    return row_idx, col_idx - 1

def get_val(sheet, address):
    try:
        r, c = addr_to_idx(address)
        return sheet.cell_value(r, c)
    except: return None 

def clean_currency(value):
    if value in [None, ""]: return 0
    if isinstance(value, (int, float)): return value
    str_val = str(value).replace("Rp", "").replace(".", "").replace(",", ".").strip()
    try: return float(str_val)
    except: return 0

def format_excel_date(value, workbook):
    if value is None or value == "": return ""
    try:
        if isinstance(value, float):
            date_tuple = xlrd.xldate_as_tuple(value, workbook.datemode)
            return datetime(*date_tuple).strftime("%d-%b-%y")
        return str(value)
    except: return str(value)

def extract_one_file(filepath):
    try:
        wb = xlrd.open_workbook(filepath, formatting_info=False) 
        sheet = wb.sheet_by_index(0) 

        exchange_rate = get_val(sheet, "B4")
        date_updated = get_val(sheet, "B3")
        proj_value = get_val(sheet, "B5") 
        cust_name  = get_val(sheet, "K3") 
        project_no = get_val(sheet, "K4") 
        formatted_date = format_excel_date(date_updated, wb)

        sub_total = 0
        warranty = 0
        cm_booked = 0
        cr_booked = 0
        penalty = 0
        total_cost = 0
        
        limit = min(sheet.nrows, 150)
        for r in range(9, limit): 
            try:
                raw_val = sheet.cell_value(r, 0)
                cell_text = str(raw_val).upper() if raw_val else ""
            except: continue
            
            if sub_total == 0 and "SUB TOTAL" in cell_text: sub_total = sheet.cell_value(r, 4) 
            elif penalty == 0 and "PENALTY" in cell_text: penalty = sheet.cell_value(r, 4)
            elif warranty == 0 and "WARRANTTY" in cell_text: warranty = sheet.cell_value(r, 4)
            elif total_cost == 0 and "TOTAL COST" in cell_text: total_cost = sheet.cell_value(r, 4)
            elif cm_booked == 0 and "CM BOOKED" in cell_text: cm_booked = sheet.cell_value(r, 4)
            elif cr_booked == 0 and "CR BOOKED" in cell_text: cr_booked = sheet.cell_value(r, 4)

        return {
            "status": "OK", "Project No": project_no, "Cust Name": cust_name, "Proj Date": formatted_date,
            "Kurs": clean_currency(exchange_rate), "Project Value": clean_currency(proj_value),
            "Sub Total": clean_currency(sub_total), "Penalty": clean_currency(penalty),
            "Warranty": clean_currency(warranty), "Total Cost": clean_currency(total_cost),
            "CM Booked": clean_currency(cm_booked), "CR Booked": clean_currency(cr_booked)
        }
    except Exception as e:
        return {"status": "ERROR", "msg": str(e), "file": os.path.basename(filepath)}

# ==========================================
# BAGIAN 2: WORKER & WATCHER
# ==========================================

class ExcelWorker(QThread):
    progress_updated = Signal(int)
    status_updated = Signal(str)
    finished_data = Signal(list)

    def __init__(self, folder_path):
        super().__init__()
        self.folder_path = folder_path

    def run(self):
        self.status_updated.emit("⚡ Auto-Scanning...")
        all_files = [f for f in os.listdir(self.folder_path) if f.lower().endswith('.xls')]
        total_files = len(all_files)
        
        if total_files == 0:
            self.status_updated.emit("Folder kosong (tidak ada .xls)")
            self.finished_data.emit([])
            return

        summary_data = []
        for i, filename in enumerate(all_files):
            if self.isInterruptionRequested(): break
            
            full_path = os.path.join(self.folder_path, filename)
            data = extract_one_file(full_path)
            if data["status"] == "OK": summary_data.append(data)
            
            progress = int((i + 1) / total_files * 100)
            self.progress_updated.emit(progress)

        self.finished_data.emit(summary_data)

class FolderChangeHandler(FileSystemEventHandler):
    def __init__(self, signal_emitter):
        self.signal_emitter = signal_emitter

    def on_any_event(self, event):
        if event.is_directory: return
        if not event.src_path.lower().endswith('.xls'): return
        self.signal_emitter.emit()

class WatcherThread(QThread):
    folder_changed = Signal()

    def __init__(self, folder_path):
        super().__init__()
        self.folder_path = folder_path
        self.observer = Observer()

    def run(self):
        event_handler = FolderChangeHandler(self.folder_changed)
        self.observer.schedule(event_handler, self.folder_path, recursive=False)
        self.observer.start()
        try:
            while not self.isInterruptionRequested():
                self.msleep(500)
        finally:
            self.observer.stop()
            self.observer.join()

    def stop(self):
        self.requestInterruption()

# ==========================================
# BAGIAN 3: MAIN WINDOW
# ==========================================

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("PCM Generator 2025 - Auto Sync 🔄")
        self.resize(1100, 700)

        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        layout = QVBoxLayout(central_widget)

        # --- HEADER ---
        header_layout = QHBoxLayout()
        self.lbl_path = QLabel("Pilih folder input...")
        self.lbl_path.setStyleSheet("border: 1px solid #ccc; padding: 5px; background: #f0f0f0;")
        self.btn_browse = QPushButton("📁 1. Pilih Folder Input")
        header_layout.addWidget(self.lbl_path)
        header_layout.addWidget(self.btn_browse)
        layout.addLayout(header_layout)

        # --- TARGET FILE SELECTION ---
        target_layout = QHBoxLayout()
        self.lbl_target = QLabel("File Summary Target: Belum dipilih")
        self.lbl_target.setStyleSheet("border: 1px solid #ccc; padding: 5px; background: #fffde7;")
        self.btn_target = QPushButton("📄 2. Pilih File Summary")
        target_layout.addWidget(self.lbl_target)
        target_layout.addWidget(self.btn_target)
        layout.addLayout(target_layout)

        # --- CONTROLS ---
        ctrl_layout = QHBoxLayout()
        
        self.chk_auto_scan = QCheckBox("Auto-Scan Input")
        self.chk_auto_scan.setChecked(True)
        self.chk_auto_scan.setStyleSheet("color: #2196F3; font-weight: bold;")

        self.chk_auto_save = QCheckBox("Auto-Save to Excel")
        self.chk_auto_save.setChecked(False) 
        self.chk_auto_save.setStyleSheet("color: #4CAF50; font-weight: bold;")
        self.chk_auto_save.setEnabled(False) 

        self.btn_force_scan = QPushButton("🔄 Force Refresh")
        self.btn_manual_save = QPushButton("💾 Manual Save")
        self.btn_manual_save.setEnabled(False)

        ctrl_layout.addWidget(self.chk_auto_scan)
        ctrl_layout.addWidget(self.chk_auto_save)
        ctrl_layout.addWidget(self.btn_force_scan)
        ctrl_layout.addWidget(self.btn_manual_save)
        layout.addLayout(ctrl_layout)

        # --- STATUS & PROGRESS ---
        self.lbl_status = QLabel("Idle")
        self.lbl_status.setStyleSheet("font-size: 14px; padding: 5px;")
        self.progress_bar = QProgressBar()
        layout.addWidget(self.lbl_status)
        layout.addWidget(self.progress_bar)

        # --- TABLE ---
        self.table = QTableWidget()
        self.columns = ["Project No", "Proj Date", "Cust Name", "Project Value", "Kurs", 
                   "BARANG&JASA", "Penalty", "Warranty", "Cost (estd)", 
                   "CM Booked", "CR Booked", "CM %", "Ket"]
        self.table.setColumnCount(len(self.columns))
        self.table.setHorizontalHeaderLabels(self.columns)
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.ResizeToContents)
        layout.addWidget(self.table)

        # --- SYSTEM VARS ---
        self.selected_folder = ""
        self.target_file = "" 
        self.all_data = [] 
        self.worker = None
        self.watcher_thread = None
        
        self.debounce_timer = QTimer()
        self.debounce_timer.setSingleShot(True)
        self.debounce_timer.setInterval(1500) 

        # --- SIGNALS ---
        self.btn_browse.clicked.connect(self.select_folder)
        self.btn_target.clicked.connect(self.select_target_file)
        self.btn_force_scan.clicked.connect(self.trigger_scan)
        self.btn_manual_save.clicked.connect(lambda: self.save_to_legacy_xls(silent=False))
        self.debounce_timer.timeout.connect(self.trigger_scan)

    def select_folder(self):
        folder = QFileDialog.getExistingDirectory(self, "Pilih Folder Input")
        if folder:
            self.selected_folder = folder
            self.lbl_path.setText(folder)
            self.trigger_scan()
            self.start_watcher()

    def select_target_file(self):
        file_path, _ = QFileDialog.getOpenFileName(self, "Pilih File Summary Master", "", "Excel Files (*.xls)")
        if file_path:
            self.target_file = file_path
            self.lbl_target.setText(f"Target: {os.path.basename(file_path)}")
            self.chk_auto_save.setEnabled(True)
            self.chk_auto_save.setChecked(True) 
            self.btn_manual_save.setEnabled(True)

    def start_watcher(self):
        if self.watcher_thread and self.watcher_thread.isRunning():
            self.watcher_thread.stop()
            self.watcher_thread.wait()

        self.watcher_thread = WatcherThread(self.selected_folder)
        self.watcher_thread.folder_changed.connect(self.on_folder_changed)
        self.watcher_thread.start()
        self.lbl_status.setText("👁️ Monitoring Active")

    def on_folder_changed(self):
        if self.chk_auto_scan.isChecked():
            self.lbl_status.setText("⏳ Detected change... Waiting...")
            self.lbl_status.setStyleSheet("color: orange; font-weight: bold;")
            self.debounce_timer.start() 

    def trigger_scan(self):
        if not self.selected_folder: return
        if self.worker and self.worker.isRunning(): return 

        self.table.setRowCount(0)
        self.progress_bar.setValue(0)
        self.lbl_status.setStyleSheet("color: black;")
        
        self.worker = ExcelWorker(self.selected_folder)
        self.worker.progress_updated.connect(self.progress_bar.setValue)
        self.worker.status_updated.connect(self.lbl_status.setText)
        self.worker.finished_data.connect(self.populate_table)
        self.worker.finished.connect(self.on_scan_finished)
        self.worker.start()

    def populate_table(self, data_list):
        self.all_data = data_list 
        self.table.setRowCount(len(data_list))
        
        for row, item in enumerate(data_list):
            proj_val = item.get("Project Value", 0)
            cm_val = item.get("CM Booked", 0)
            cm_percent = 0
            if proj_val > 0: cm_percent = (cm_val / proj_val) * 100
            item["CM %"] = f"{cm_percent:.2f}%"

            def fmt(val): return "{:,.0f}".format(val).replace(",", ".") if isinstance(val, (int, float)) else val

            self.table.setItem(row, 0, QTableWidgetItem(str(item.get("Project No"))))
            self.table.setItem(row, 1, QTableWidgetItem(str(item.get("Proj Date"))))
            self.table.setItem(row, 2, QTableWidgetItem(str(item.get("Cust Name"))))
            self.table.setItem(row, 3, QTableWidgetItem(fmt(item.get("Project Value"))))
            self.table.setItem(row, 4, QTableWidgetItem(fmt(item.get("Kurs"))))
            self.table.setItem(row, 5, QTableWidgetItem(fmt(item.get("Sub Total"))))
            self.table.setItem(row, 6, QTableWidgetItem(fmt(item.get("Penalty"))))
            self.table.setItem(row, 7, QTableWidgetItem(fmt(item.get("Warranty"))))
            self.table.setItem(row, 8, QTableWidgetItem(fmt(item.get("Total Cost"))))
            self.table.setItem(row, 9, QTableWidgetItem(fmt(item.get("CM Booked"))))
            self.table.setItem(row, 10, QTableWidgetItem(fmt(item.get("CR Booked"))))
            self.table.setItem(row, 11, QTableWidgetItem(item["CM %"]))
            self.table.setItem(row, 12, QTableWidgetItem(""))

    def on_scan_finished(self):
        if self.chk_auto_save.isChecked() and self.target_file:
            self.save_to_legacy_xls(silent=True) 
        else:
            ts = datetime.now().strftime("%H:%M:%S")
            self.lbl_status.setText(f"Scan selesai {ts}. Menunggu save manual.")

    def get_hardcoded_mapping(self):
        return {
            "No": 2, "Project No": 3, "Proj Date": 5, "Cust Name": 6,
            "Project Value": 8, "Kurs": 9, "Sub Total": 11, "Penalty": 12,
            "Warranty": 13, "Total Cost": 15, "CM Booked": 16, "CR Booked": 17,
            "CM %": 19
        }

    def save_to_legacy_xls(self, silent=False):
        if not self.target_file:
            if not silent: QMessageBox.warning(self, "Warning", "Pilih file target dulu!")
            return

        import xlwt
        from xlutils.copy import copy
        
        try:
            # --- 1. DETEKSI LOCKING SAAT MEMBUKA (READ) ---
            rb = xlrd.open_workbook(self.target_file, formatting_info=True)
            
        except PermissionError:
            # HANDLER 1: File terkunci saat mau dibaca
            msg = "❌ GAGAL BACA: File Excel sedang dibuka! Tutup file lalu coba lagi."
            self.lbl_status.setText(msg)
            self.lbl_status.setStyleSheet("color: red; font-weight: bold;")
            if not silent: QMessageBox.critical(self, "File Terkunci", msg)
            return
        except Exception as e:
            if not silent: QMessageBox.critical(self, "Error", str(e))
            return
        
        target_sheet_name = "PCM"
        sheet_idx = 0
        for idx, name in enumerate(rb.sheet_names()):
            if name.strip().upper() == target_sheet_name:
                sheet_idx = idx
                break
        r_sheet = rb.sheet_by_index(sheet_idx)
        wb = copy(rb)
        w_sheet = wb.get_sheet(sheet_idx)

        # STYLES
        borders = xlwt.Borders()
        borders.left = xlwt.Borders.THIN
        borders.right = xlwt.Borders.THIN
        borders.top = xlwt.Borders.THIN
        borders.bottom = xlwt.Borders.THIN
        
        style_center = xlwt.XFStyle(); style_center.borders = borders; style_center.alignment = xlwt.Alignment()
        style_center.alignment.horz = xlwt.Alignment.HORZ_CENTER; style_center.alignment.vert = xlwt.Alignment.VERT_CENTER
        
        style_date = xlwt.XFStyle(); style_date.borders = borders; style_date.alignment = xlwt.Alignment()
        style_date.alignment.horz = xlwt.Alignment.HORZ_CENTER; style_date.num_format_str = 'D-MMM-YY'
        
        style_text = xlwt.XFStyle(); style_text.borders = borders; style_text.alignment = xlwt.Alignment()
        style_text.alignment.horz = xlwt.Alignment.HORZ_LEFT; style_text.alignment.vert = xlwt.Alignment.VERT_CENTER
        
        style_num = xlwt.XFStyle(); style_num.borders = borders; style_num.alignment = xlwt.Alignment()
        style_num.alignment.horz = xlwt.Alignment.HORZ_RIGHT; style_num.num_format_str = '#,##0'
        
        style_pct = xlwt.XFStyle(); style_pct.borders = borders; style_pct.alignment = xlwt.Alignment()
        style_pct.alignment.horz = xlwt.Alignment.HORZ_RIGHT; style_pct.num_format_str = '0.00%'

        START_ROW_INDEX = 4 
        col_indices = self.get_hardcoded_mapping()
        proj_col_idx = col_indices["Project No"] 

        existing_projects = set()
        write_row = START_ROW_INDEX
        while write_row < r_sheet.nrows:
            try:
                p_val = r_sheet.cell_value(write_row, proj_col_idx)
                if not str(p_val).strip(): break
                existing_projects.add(str(p_val).strip())
                write_row += 1
            except IndexError: break
        
        added_count = 0
        for item in self.all_data:
            p_no = str(item.get("Project No")).strip()
            if p_no in existing_projects: continue

            no_col = col_indices["No"]
            existing_no = None
            try: existing_no = r_sheet.cell_value(write_row, no_col)
            except: pass
            
            if not (isinstance(existing_no, (int, float)) and existing_no > 0):
                calc_no = (write_row - START_ROW_INDEX) + 1
                w_sheet.write(write_row, no_col, calc_no, style_center)

            excel_row = write_row + 1 
            for key, col_idx in col_indices.items():
                if key == "No": continue
                val = item.get(key)
                use_style = style_text

                if key == "CM %":
                    formula_str = f"IF(I{excel_row}=0,0, S{excel_row}/I{excel_row})"
                    w_sheet.write(write_row, col_idx, xlwt.Formula(formula_str), style_pct)
                    continue

                if key in ["Project Value", "Kurs", "Sub Total", "Penalty", "Warranty", 
                           "Total Cost", "CM Booked", "CR Booked"]:
                    use_style = style_num
                    try: val = float(val)
                    except: val = 0
                elif key == "Proj Date": use_style = style_date

                if val is not None: w_sheet.write(write_row, col_idx, val, use_style)
            
            w_sheet.write(write_row, 18, xlwt.Formula(f"I{excel_row}-P{excel_row}"), style_num)
            w_sheet.write(write_row, 7, "IDR", style_center)
            write_row += 1
            added_count += 1

        try:
            # --- 2. DETEKSI LOCKING SAAT MENYIMPAN (WRITE) ---
            wb.save(self.target_file)
            
            ts = datetime.now().strftime("%H:%M:%S")
            msg = f"✅ Auto-Saved to Excel at {ts} (+{added_count} Data)"
            self.lbl_status.setText(msg)
            self.lbl_status.setStyleSheet("color: green; font-weight: bold;")
            
            if not silent: 
                QMessageBox.information(self, "Update Sukses", f"Data Baru: {added_count}\nBerhasil update ke file .xls")
                
        except PermissionError:
             # HANDLER 2: File terkunci saat mau disimpan
            msg = "❌ GAGAL SAVE: File Excel sedang dibuka! Tutup file lalu coba lagi."
            self.lbl_status.setText(msg)
            self.lbl_status.setStyleSheet("color: red; font-weight: bold;")
            if not silent: QMessageBox.critical(self, "Gagal", msg)
        except Exception as e:
            self.lbl_status.setText(f"Error Save: {e}")

if __name__ == "__main__":
    app = QApplication(sys.argv)
    app.setStyle("Fusion")
    window = MainWindow()
    window.show()
    sys.exit(app.exec())