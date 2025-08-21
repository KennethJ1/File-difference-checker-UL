from PySide6.QtWidgets import (QApplication, QMainWindow, QFileDialog, QTextEdit, QPushButton,
                                QVBoxLayout, QWidget, QLabel, QComboBox, QHBoxLayout, QProgressBar,
                                QListWidget, QListWidgetItem, QMessageBox, QLineEdit, QDoubleSpinBox, QColorDialog)
from PySide6.QtCore import Qt, QThread, Signal
from PySide6.QtGui import QColor
import sys
import os
import typing

# allow importing src/core modules when running this script directly
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))
from core import runner


class CompareWorker(QThread):
    progress = Signal(int)
    finished = Signal(object)
    error = Signal(str)

    def __init__(self, files: typing.List[str], options: dict, file_type: typing.Optional[str] = None):
        super().__init__()
        self.files = files
        self.options = options or {}
        self.file_type = file_type

    def _emit_progress(self, p: int):
        try:
            self.progress.emit(int(p))
        except Exception:
            pass

    def run(self):
        try:
            result = runner.run_compare(self.files, options=self.options, file_type=self.file_type, progress_cb=self._emit_progress, return_meta=True)
            self.finished.emit(result)
        except Exception as e:
            self.error.emit(str(e))


class FileListWidget(QListWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setAcceptDrops(True)

    def dragEnterEvent(self, event):
        md = event.mimeData()
        if md.hasUrls():
            event.acceptProposedAction()
        else:
            event.ignore()

    def dropEvent(self, event):
        md = event.mimeData()
        added = []
        if md.hasUrls():
            for url in md.urls():
                path = url.toLocalFile()
                if path and os.path.isfile(path):
                    # prevent directories and duplicates
                    added.append(path)
        # emit a custom signal by calling parent's handler if present
        parent = self.parent()
        if hasattr(parent, 'handle_dropped_files'):
            parent.handle_dropped_files(added)
        else:
            # fallback: add items directly
            for p in added:
                if not any(self.item(i).text() == p for i in range(self.count())):
                    self.addItem(QListWidgetItem(p))


class AppWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Difference Checker")
        self.setGeometry(100, 100, 900, 600)
        self.selected_files: typing.List[str] = []
        self.output_path: typing.Optional[str] = None
        self.worker: typing.Optional[CompareWorker] = None

        # create logs directory and file
        project_root = os.path.abspath(os.path.join(os.path.dirname(__file__), '..', '..'))
        logs_dir = os.path.join(project_root, 'logs')
        os.makedirs(logs_dir, exist_ok=True)
        import datetime
        self.log_file = os.path.join(logs_dir, f"diffcheck_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.log")
        # write header
        with open(self.log_file, 'a', encoding='utf-8') as f:
            f.write(f"Difference Checker log started: {datetime.datetime.now().isoformat()}\n")

        self.initUI()

    def write_log(self, msg: str):
        """Append to GUI log widget and save to disk."""
        import datetime
        ts = datetime.datetime.now().isoformat()
        line = f"[{ts}] {msg}"
        try:
            self.log.append(line)
        except Exception:
            pass
        try:
            with open(self.log_file, 'a', encoding='utf-8') as f:
                f.write(line + '\n')
        except Exception:
            pass

    def initUI(self):
        main_layout = QVBoxLayout()

        label = QLabel("Select the files you want to compare (2 files):")
        main_layout.addWidget(label)

        # file list widget (supports drag & drop)
        self.file_list_widget = FileListWidget(self)
        main_layout.addWidget(self.file_list_widget)

        file_buttons_layout = QHBoxLayout()
        self.add_button = QPushButton("Add Files")
        self.add_button.clicked.connect(self.add_files)
        file_buttons_layout.addWidget(self.add_button)

        self.clear_button = QPushButton("Clear")
        self.clear_button.clicked.connect(self.clear_files)
        file_buttons_layout.addWidget(self.clear_button)

        main_layout.addLayout(file_buttons_layout)

        # file type selection (Auto / Excel / PDF)
        type_layout = QHBoxLayout()
        type_layout.addWidget(QLabel("File type (auto-detect):"))
        self.file_type = QComboBox()
        self.file_type.addItems(["Auto", "Excel", "PDF"])  # future: Word, etc.
        type_layout.addWidget(self.file_type)

        # output path selector
        self.output_btn = QPushButton("Choose Output Path")
        self.output_btn.clicked.connect(self.choose_output)
        type_layout.addWidget(self.output_btn)

        # display selected output path next to button
        self.output_path_field = QLineEdit()
        self.output_path_field.setReadOnly(True)
        self.output_path_field.setPlaceholderText("No output path selected — will auto-generate if omitted")
        type_layout.addWidget(self.output_path_field)

        main_layout.addLayout(type_layout)

        # compare button and progress
        action_layout = QHBoxLayout()
        self.compare_button = QPushButton("Compare")
        self.compare_button.clicked.connect(self.start_compare)
        action_layout.addWidget(self.compare_button)

        self.progress_bar = QProgressBar()
        self.progress_bar.setRange(0, 100)
        self.progress_bar.setValue(0)
        action_layout.addWidget(self.progress_bar)

        main_layout.addLayout(action_layout)

        # log area
        self.log = QTextEdit()
        self.log.setReadOnly(True)
        main_layout.addWidget(self.log)

        # engine-specific options area
        opts_layout = QHBoxLayout()
        # Excel key header
        opts_layout.addWidget(QLabel("Excel key header:"))
        self.excel_key_header = QLineEdit()
        self.excel_key_header.setPlaceholderText("S.no")
        self.excel_key_header.setText("S.no")
        opts_layout.addWidget(self.excel_key_header)

        # PDF threshold
        opts_layout.addWidget(QLabel("PDF visual threshold (0-1):"))
        self.pdf_threshold = QDoubleSpinBox()
        self.pdf_threshold.setRange(0.0, 1.0)
        self.pdf_threshold.setSingleStep(0.05)
        self.pdf_threshold.setValue(0.6)
        self.pdf_threshold.setToolTip("Controls how much visual difference is needed to highlight a change. Lower values = more sensitive.")
        opts_layout.addWidget(self.pdf_threshold)

        # PDF colors - use color picker buttons
        opts_layout.addWidget(QLabel("PDF text color File1:"))
        self.pdf_color1_btn = QPushButton()
        self.pdf_color1_btn.setFixedWidth(50)
        self._pdf_color1 = QColor('#FF0000')
        self._apply_color_to_button(self.pdf_color1_btn, self._pdf_color1)
        self.pdf_color1_btn.clicked.connect(self.choose_color1)
        opts_layout.addWidget(self.pdf_color1_btn)

        opts_layout.addWidget(QLabel("PDF text color File2:"))
        self.pdf_color2_btn = QPushButton()
        self.pdf_color2_btn.setFixedWidth(50)
        self._pdf_color2 = QColor('#00FF00')
        self._apply_color_to_button(self.pdf_color2_btn, self._pdf_color2)
        self.pdf_color2_btn.clicked.connect(self.choose_color2)
        opts_layout.addWidget(self.pdf_color2_btn)

        main_layout.addLayout(opts_layout)

        container = QWidget()
        container.setLayout(main_layout)
        self.setCentralWidget(container)

    def _apply_color_to_button(self, btn: QPushButton, qcolor: QColor):
        ss = f"background-color: {qcolor.name()}; border: 1px solid #555;"
        btn.setStyleSheet(ss)

    def choose_color1(self):
        c = QColorDialog.getColor(self._pdf_color1, self, "Choose color for File 1")
        if c.isValid():
            self._pdf_color1 = c
            self._apply_color_to_button(self.pdf_color1_btn, c)

    def choose_color2(self):
        c = QColorDialog.getColor(self._pdf_color2, self, "Choose color for File 2")
        if c.isValid():
            self._pdf_color2 = c
            self._apply_color_to_button(self.pdf_color2_btn, c)

    def handle_dropped_files(self, paths: typing.List[str]):
        # called by FileListWidget when files are dropped
        for p in paths:
            if p not in self.selected_files:
                self.selected_files.append(p)
                self.file_list_widget.addItem(QListWidgetItem(p))
        # limit to two files
        while len(self.selected_files) > 2:
            self.selected_files.pop()
            self.file_list_widget.takeItem(self.file_list_widget.count()-1)

    def add_files(self):
        options = QFileDialog.Options()
        files, _ = QFileDialog.getOpenFileNames(self, "Select files to compare", "", "All Files (*);;Excel (*.xlsx *.xls);;PDF (*.pdf)", options=options)
        if not files:
            return
        for f in files:
            if f not in self.selected_files:
                self.selected_files.append(f)
                self.file_list_widget.addItem(QListWidgetItem(f))
        # limit to first two files
        while len(self.selected_files) > 2:
            self.selected_files.pop()
            self.file_list_widget.takeItem(self.file_list_widget.count()-1)

    def clear_files(self):
        self.selected_files = []
        self.file_list_widget.clear()

    def choose_output(self):
        # suggest an output filename based on selected files
        suggested = "comparison_output"
        if self.selected_files:
            base = os.path.splitext(os.path.basename(self.selected_files[0]))[0]
            other = os.path.splitext(os.path.basename(self.selected_files[-1]))[0]
            suggested = f"{base}_vs_{other}"
        # determine suggested extension based on selected type
        cur_type = self.file_type.currentText()
        if cur_type == 'PDF':
            suggested += '.pdf'
            filter_str = "PDF (*.pdf);;All Files (*)"
        else:
            suggested += '.xlsx'
            filter_str = "Excel (*.xlsx);;All Files (*)"

        fpath, _ = QFileDialog.getSaveFileName(self, "Select output file", suggested, filter_str)
        if fpath:
            self.output_path = fpath
            self.output_path_field.setText(fpath)
            self.write_log(f"Output path set: {fpath}")

    def start_compare(self):
        if len(self.selected_files) < 2:
            QMessageBox.warning(self, "Need 2 files", "Please select two files to compare.")
            return

        # determine file type override
        ftype = None
        ft_sel = self.file_type.currentText()
        if ft_sel == "Excel":
            ftype = "excel"
        elif ft_sel == "PDF":
            ftype = "pdf"

        # build options mapping - harmonize keys across engines
        opts = {}
        # if user selected an output path in UI, use it; otherwise leave to engines to auto-generate
        if self.output_path:
            opts["output_path"] = self.output_path

        # engine specific options
        opts["key_header"] = self.excel_key_header.text().strip() or "S.no"
        # PDF
        opts["threshold"] = float(self.pdf_threshold.value())

        # get colors from pickers and convert to (r,g,b,a)
        def _qcolor_to_rgba(qc: QColor, alpha: int = 30):
            if not qc or not qc.isValid():
                return None
            return (qc.red(), qc.green(), qc.blue(), alpha)

        c1 = _qcolor_to_rgba(self._pdf_color1, alpha=30)
        c2 = _qcolor_to_rgba(self._pdf_color2, alpha=30)
        if c1:
            opts["text_color1"] = c1
        if c2:
            opts["text_color2"] = c2

        # Harmonize common names: progress_cb/output_path/return_meta are already used
        opts["return_meta"] = True

        # disable UI while running
        self.compare_button.setEnabled(False)
        self.add_button.setEnabled(False)
        self.clear_button.setEnabled(False)
        self.progress_bar.setValue(0)
        self.write_log("Starting comparison...")

        # start worker thread
        self.worker = CompareWorker(self.selected_files[:2], options=opts, file_type=ftype)
        self.worker.progress.connect(self.on_progress)
        self.worker.finished.connect(self.on_finished)
        self.worker.error.connect(self.on_error)
        self.worker.start()

    def on_progress(self, p: int):
        self.progress_bar.setValue(p)

    def on_finished(self, result):
        # result is (output_path, meta)
        self.compare_button.setEnabled(True)
        self.add_button.setEnabled(True)
        self.clear_button.setEnabled(True)
        try:
            output, meta = result
            self.write_log(f"Comparison finished. Output: {output}")
            self.write_log(str(meta))
            # try to open output if obvious
            if isinstance(output, str) and os.path.exists(output):
                self.write_log("Opening output file...")
                try:
                    if sys.platform.startswith('win'):
                        os.startfile(output)
                    else:
                        import subprocess
                        subprocess.Popen(['xdg-open', output])
                except Exception as e:
                    self.write_log(f"Failed to open output: {e}")
        except Exception:
            # some engines may return just a path
            self.write_log(f"Comparison finished: {result}")
        self.progress_bar.setValue(100)

    def on_error(self, msg: str):
        self.compare_button.setEnabled(True)
        self.add_button.setEnabled(True)
        self.clear_button.setEnabled(True)
        QMessageBox.critical(self, "Error during comparison", msg)
        self.write_log(f"Error: {msg}")
        self.progress_bar.setValue(0)


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = AppWindow()
    window.show()
    sys.exit(app.exec())

