import os
import sys 
from PySide6.QtWidgets import (QMainWindow, QWidget, QVBoxLayout, 
                               QHBoxLayout, QPushButton, QLabel, QProgressBar, 
                               QTableWidget, QTableWidgetItem, QFileDialog, 
                               QMessageBox, QHeaderView, QAbstractItemView,
                               QDialog, QTextEdit)
from PySide6.QtCore import Qt, QSettings, QUrl, QTimer
from PySide6.QtGui import QColor, QDesktopServices, QFont

from workers import WatcherThread, PreviewWorker, GeneratorWorker

# --- KELAS DIALOG BANTUAN ---
class HelpDialog(QDialog):
    def __init__(self, content, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Panduan Pengguna - PCM Summary Generator")
        self.resize(700, 600) # Ukuran sedikit diperbesar agar lega
        
        layout = QVBoxLayout(self)
        
        # Area Teks Read-Only
        self.text_area = QTextEdit()
        self.text_area.setPlainText(content)
        self.text_area.setReadOnly(True)
        self.text_area.setFont(QFont("Consolas", 10)) # Font Monospace
        
        layout.addWidget(self.text_area)
        
        # Tombol Tutup
        btn_close = QPushButton("Tutup")
        btn_close.clicked.connect(self.close)
        layout.addWidget(btn_close)

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle(" PCM Summary Generator v1.0.0")
        self.resize(1000, 650)
        self.settings = QSettings("FahmiSoft", "PCMGenerator")
        
        # --- VARIABEL UNTUK JENDELA HELP (Agar tidak blocking) ---
        self.help_window = None 

        central = QWidget()
        self.setCentralWidget(central)
        layout = QVBoxLayout(central)
        
        # --- IO ---
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
        self.table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.table.setSelectionMode(QAbstractItemView.SingleSelection)
        self.table.setEditTriggers(QAbstractItemView.NoEditTriggers)

        # --- HEADER SETUP ---
        header = self.table.horizontalHeader()
        header.setSectionResizeMode(QHeaderView.ResizeMode.Interactive)
        header.setStretchLastSection(True)

        self.table.setColumnWidth(0, 300) 
        self.table.setColumnWidth(1, 100) # Project ID
        self.table.setColumnWidth(4, 150) # Nilai Project
        
        # --- FITUR SORTING ---
        self.table.setSortingEnabled(True)
        
        # --- FITUR DOUBLE CLICK ---
        self.table.cellDoubleClicked.connect(self.on_table_double_click)
        
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
        
        self.setup_statusbar()
        self.input_dir = ""; self.output_dir = ""; self.data_cache = []
        self.scan_worker = None; self.gen_worker = None
        self.watcher_thread = None
        
        self.debounce_timer = QTimer()
        self.debounce_timer.setSingleShot(True)
        self.debounce_timer.setInterval(1500)
        self.debounce_timer.timeout.connect(self.run_preview_scan)

        self.load_settings()

    def setup_statusbar(self):
        status_bar = self.statusBar()
        lbl_version = QLabel("PCM Summary Generator v 1.0.0")
        status_bar.addWidget(lbl_version)
        
        # --- TOMBOL BANTUAN ---
        btn_help = QPushButton("Bantuan / Help")
        btn_help.setFlat(True)
        btn_help.setStyleSheet("font-weight: bold; color: #2196F3;") 
        btn_help.clicked.connect(self.open_help_dialog) 
        status_bar.addPermanentWidget(btn_help)

        # --- TOMBOL ABOUT ---
        btn_about = QPushButton("About")
        btn_about.setFlat(True)
        btn_about.setStyleSheet("font-weight: bold; color: #555;")
        btn_about.clicked.connect(self.show_about_dialog)
        status_bar.addPermanentWidget(btn_about)

    def open_help_dialog(self):
        """
        Membuka jendela bantuan secara Non-Blocking (Modeless).
        User tetap bisa klik Main Window saat jendela ini terbuka.
        """
        filename = "USER_MANUAL.txt"
        
        # Logika mencari path file text
        if getattr(sys, 'frozen', False):
            base_path = os.path.dirname(sys.executable)
        else:
            base_path = os.path.dirname(os.path.abspath(__file__))
            
        help_path = os.path.join(base_path, filename)
        
        content = ""
        if os.path.exists(help_path):
            try:
                with open(help_path, 'r', encoding='utf-8') as f:
                    content = f.read()
            except Exception as e:
                content = f"Error membaca file:\n{e}"
        else:
            content = (f"File '{filename}' tidak ditemukan.\n\n"
                       f"Lokasi pencarian: {help_path}\n\n"
                       "Pastikan file manual ada di folder aplikasi.")

        # --- LOGIKA JENDELA NON-BLOCKING ---
        # 1. Jika jendela belum pernah dibuat, atau sudah dihancurkan -> Buat baru
        if self.help_window is None:
            self.help_window = HelpDialog(content, self)
            # Opsional: Jika user menutup jendela, kita set variable ke None lagi (opsional)
            # Tapi di sini kita biarkan objectnya hidup agar posisi/ukuran tersimpan selama aplikasi jalan
        
        # 2. Tampilkan Jendela (Non-Blocking)
        self.help_window.show()
        
        # 3. Bawa ke depan (agar tidak tertutup main window jika sudah terbuka sebelumnya)
        self.help_window.raise_()
        self.help_window.activateWindow()

    def load_settings(self):
        last_in = self.settings.value("last_input_dir")
        last_out = self.settings.value("last_output_dir")
        if last_in and os.path.exists(last_in):
            self.input_dir = last_in
            self.lbl_input.setText(last_in)
            self.start_watcher()
            self.run_preview_scan()
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
            self.settings.setValue("last_input_dir", path)
            self.start_watcher()
            self.run_preview_scan()

    def select_output(self):
        start_dir = self.output_dir if self.output_dir else ""
        path = QFileDialog.getExistingDirectory(self, "Pilih Output Folder", start_dir)
        if path:
            self.output_dir = path
            self.lbl_output.setText(path)
            self.settings.setValue("last_output_dir", path)
            self.check_ready()

    def start_watcher(self):
        if self.watcher_thread: self.watcher_thread.stop(); self.watcher_thread.wait()
        self.watcher_thread = WatcherThread(self.input_dir)
        self.watcher_thread.folder_changed.connect(self.on_folder_change_detected)
        self.watcher_thread.start()

    def on_folder_change_detected(self):
        self.statusBar().showMessage("🔍 Mendeteksi perubahan file... Menunggu...", 2000)
        self.debounce_timer.start()

    def run_preview_scan(self):
        self.table.setSortingEnabled(False)
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
            if isinstance(val, (int, float)):
                # Langkah 1: Format standar (1,250,000)
                # Langkah 2: Replace koma menjadi titik (1.250.000)
                fmt_val = f"{val:,.0f}".replace(",", ".")
                
                # OPSI: Jika ingin tambah 'Rp', uncomment baris bawah ini:
                # fmt_val = f"Rp {fmt_val}" 
            else:
                fmt_val = str(val)
            project_val_item = make_item(fmt_val)
            project_val_item.setTextAlignment(Qt.AlignRight | Qt.AlignVCenter)
            self.table.setItem(r, 4, project_val_item)

            self.table.setItem(r, 5, make_item(status))
        
        self.table.setSortingEnabled(True)
        self.check_ready()
        self.statusBar().showMessage(f"Scan selesai. Total {len(results)} file.", 3000)

    # --- FITUR DOUBLE CLICK OPEN EXCEL ---
    def on_table_double_click(self, row, col):
        if row < 0 or row >= len(self.data_cache): return
        
        filename_in_table = self.table.item(row, 0).text()
        selected_file = None
        for item in self.data_cache:
            if item["filename"] == filename_in_table:
                selected_file = item
                break
        
        if not selected_file: return

        reply = QMessageBox.question(self, "Edit File Input", 
                                     f"Apakah Anda ingin memodifikasi file ini?\n\n{selected_file['filename']}",
                                     QMessageBox.Yes | QMessageBox.No)
        
        if reply == QMessageBox.Yes:
            file_path = selected_file["path"]
            if os.path.exists(file_path):
                success = QDesktopServices.openUrl(QUrl.fromLocalFile(file_path))
                if not success:
                    folder = os.path.dirname(file_path)
                    QDesktopServices.openUrl(QUrl.fromLocalFile(folder))
            else:
                QMessageBox.warning(self, "Error", "File tidak ditemukan!")

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
        self.btn_gen.setEnabled(True); self.btn_gen.setText("GENERATE ULANG")
        
        if "ERROR:" in result_msg:
            QMessageBox.critical(self, "Gagal", result_msg)
        else:
            reply = QMessageBox.question(self, "Sukses", 
                                         f"Generate selesai!\nFile Summary:\n{result_msg}\n\n"
                                         "Buka folder output di Windows Explorer?",
                                         QMessageBox.Yes | QMessageBox.No)
            if reply == QMessageBox.Yes:
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
        QMessageBox.about(self, "About", info)