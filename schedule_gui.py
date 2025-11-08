# -*- coding: utf-8 -*-
import sys, os
from PyQt6.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QLabel, QSpinBox, QRadioButton,
    QPushButton, QMessageBox, QButtonGroup, QHBoxLayout, QFrame,
    QDialog, QTableWidget, QTableWidgetItem, QHeaderView, QDialogButtonBox,
    QProgressDialog
)
from PyQt6.QtGui import QPixmap, QFont
from PyQt6.QtCore import Qt, QThread, pyqtSignal
from datetime import datetime
from schedule_logic import create_schedule

# --- UNC パス対応 ---
DEFAULT_IMPORT = r"\\PC011\Users\yasumoku\Desktop\タカラ関係\工程表"
DEFAULT_OUTPUT = r"\\PC009\share01\日程表"

# --- 難読化された期限チェック ---
def __hidden_expire_check__():
    import math
    bd = [50, 48, 50, 53, 48, 56, 48, 54]  # "20250806"
    yy = int("".join([chr(c) for c in bd[0:4]]))
    mm = int("".join([chr(c) for c in bd[4:6]]))
    dd = int("".join([chr(c) for c in bd[6:8]]))
    h = 13
    mi = 21
    base = datetime(yy, mm, dd, h, mi)
    expire_min = int("FFFFF", 16)
    now = datetime.now()
    check_val = (now - base).total_seconds()/60
    if check_val > expire_min:
        app = QApplication([])
        QMessageBox.critical(None, "使用不可", "このアプリは使用できません。\n管理担当者に確認してください。")
        sys.exit(0)

__hidden_expire_check__()

# --- PyQt6 高DPI対応 ---
from PyQt6 import QtCore
try:
    QtCore.QCoreApplication.setAttribute(QtCore.Qt.ApplicationAttribute.AA_EnableHighDpiScaling, True)
    QtCore.QCoreApplication.setAttribute(QtCore.Qt.ApplicationAttribute.AA_UseHighDpiPixmaps, True)
except AttributeError:
    pass

# --- UNC パス対応ディレクトリチェック ---
try:
    os.listdir(DEFAULT_IMPORT)
except Exception as e:
    app = QApplication(sys.argv)
    QMessageBox.critical(None, "ディレクトリ無効",
        f"参照先ディレクトリが存在しないかアクセスできません:\n{DEFAULT_IMPORT}\n{e}")
    sys.exit(0)

# --- Worker ---
class ScheduleWorker(QThread):
    finished = pyqtSignal(str, int, int, int, int)
    error = pyqtSignal(str)

    def __init__(self, year, month, day, filter_type, import_file, output_path):
        super().__init__()
        self.year = year
        self.month = month
        self.day = day
        self.filter_type = filter_type
        self.import_file = import_file
        self.output_path = output_path

    def run(self):
        try:
            save_file, gifu_new, shiga_new, gifu_old, shiga_old = create_schedule(
                self.year, self.month, self.day,
                self.filter_type,
                self.import_file,
                self.output_path
            )
            self.finished.emit(save_file, gifu_new, shiga_new, gifu_old, shiga_old)
        except Exception as e:
            self.error.emit(str(e))

# --- ファイル選択ダイアログ ---
class FileSelectDialog(QDialog):
    def __init__(self, files):
        super().__init__()
        self.setWindowTitle("対象データを選択")
        self.setMinimumWidth(500)
        layout = QVBoxLayout(self)

        self.table = QTableWidget(len(files), 2)
        self.table.setHorizontalHeaderLabels(["ファイル名", "更新日"])
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        self.table.setSelectionBehavior(QTableWidget.SelectionBehavior.SelectRows)
        self.table.setSelectionMode(QTableWidget.SelectionMode.SingleSelection)

        for i, (name, mtime) in enumerate(files):
            self.table.setItem(i, 0, QTableWidgetItem(name))
            self.table.setItem(i, 1, QTableWidgetItem(mtime))

        layout.addWidget(self.table)

        btns = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel)
        btns.accepted.connect(self.accept)
        btns.rejected.connect(self.reject)
        layout.addWidget(btns)

    def get_selected_file(self):
        row = self.table.currentRow()
        if row < 0:
            return None
        return self.table.item(row, 0).text()

# --- GUI ---
class CuteScheduleApp(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("日程表作成")
        self.setGeometry(100, 100, 700, 650)
        self.setStyleSheet("background-color:#fafafa; color:#333333;")
        layout = QVBoxLayout()
        layout.setSpacing(15)

        # 見出し
        title = QLabel("📋 日程表作成アプリ")
        title.setFont(QFont("Arial", 26, QFont.Weight.Bold))
        title.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(title)

        # 上部画像
        image_label = QLabel()
        if getattr(sys, 'frozen', False):
            script_dir = sys._MEIPASS
        else:
            script_dir = os.path.dirname(os.path.abspath(__file__))
        image_path = os.path.join(script_dir, "05.png")
        if os.path.exists(image_path):
            pixmap = QPixmap(image_path)
            pixmap = pixmap.scaled(600, pixmap.height(), Qt.AspectRatioMode.KeepAspectRatio)
            image_label.setPixmap(pixmap)
            image_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
            layout.addWidget(image_label)

        today = datetime.now()

        # 日付フレーム
        date_frame = QFrame()
        date_frame.setStyleSheet("background-color: #f0f0f0; border-radius: 10px; padding:10px;")
        date_layout = QHBoxLayout()
        date_layout.setSpacing(5)

        lbl_year = QLabel("年:")
        lbl_year.setFont(QFont("Arial", 14))
        lbl_year.setStyleSheet("color:#555555;")
        date_layout.addWidget(lbl_year)
        self.year_input = QSpinBox()
        self.year_input.setRange(2000, 2100)
        self.year_input.setValue(today.year)
        self.year_input.setFont(QFont("Arial", 14))
        date_layout.addWidget(self.year_input)

        lbl_month = QLabel("月:")
        lbl_month.setFont(QFont("Arial", 14))
        lbl_month.setStyleSheet("color:#555555;")
        date_layout.addWidget(lbl_month)
        self.month_input = QSpinBox()
        self.month_input.setRange(1, 12)
        self.month_input.setValue(today.month)
        self.month_input.setFont(QFont("Arial", 14))
        date_layout.addWidget(self.month_input)

        lbl_day = QLabel("日:")
        lbl_day.setFont(QFont("Arial", 14))
        lbl_day.setStyleSheet("color:#555555;")
        date_layout.addWidget(lbl_day)
        self.day_input = QSpinBox()
        self.day_input.setRange(1, 31)
        self.day_input.setValue(today.day)
        self.day_input.setFont(QFont("Arial", 14))
        date_layout.addWidget(self.day_input)

        arrow_style = """
        QSpinBox::up-button, QSpinBox::down-button { width: 25px; height: 25px; }
        QSpinBox::up-arrow, QSpinBox::down-arrow { image: none; }
        """
        for sb in [self.year_input, self.month_input, self.day_input]:
            sb.setStyleSheet(arrow_style)

        date_frame.setLayout(date_layout)
        layout.addWidget(date_frame)

        # フィルター
        filter_frame = QFrame()
        filter_frame.setStyleSheet("background-color: #e0f7fa; border-radius: 10px; padding:5px;")
        filter_layout = QHBoxLayout()
        filter_layout.setSpacing(10)
        lbl_filter = QLabel("フィルター:")
        lbl_filter.setFont(QFont("Arial", 14))
        lbl_filter.setStyleSheet("color:#333333;")
        filter_layout.addWidget(lbl_filter)
        self.rb_all = QRadioButton("全件")
        self.rb_all.setFont(QFont("Arial", 14))
        self.rb_all.setChecked(True)  # ← 全件をデフォルト
        self.rb_dollar = QRadioButton("新図面のみ")
        self.rb_dollar.setFont(QFont("Arial", 14))
        filter_layout.addWidget(self.rb_all)
        filter_layout.addWidget(self.rb_dollar)
        self.filter_group = QButtonGroup()
        self.filter_group.addButton(self.rb_all)
        self.filter_group.addButton(self.rb_dollar)
        filter_frame.setLayout(filter_layout)
        layout.addWidget(filter_frame)

        # 参照先・保存先
        import_frame = QFrame()
        import_frame.setStyleSheet("background-color: #fff3e0; border-radius: 10px; padding:5px;")
        import_layout = QVBoxLayout()
        self.import_label = QLabel(f"参照先: {DEFAULT_IMPORT}")
        self.import_label.setFont(QFont("Arial", 12))
        import_layout.addWidget(self.import_label)
        import_frame.setLayout(import_layout)
        layout.addWidget(import_frame)

        output_frame = QFrame()
        output_frame.setStyleSheet("background-color: #fff3e0; border-radius: 10px; padding:5px;")
        output_layout = QVBoxLayout()
        self.output_label = QLabel(f"保存先: {DEFAULT_OUTPUT}")
        self.output_label.setFont(QFont("Arial", 12))
        output_layout.addWidget(self.output_label)
        output_frame.setLayout(output_layout)
        layout.addWidget(output_frame)

        # 実行ボタン
        btn_run = QPushButton("実行")
        btn_run.setFont(QFont("Arial", 16, QFont.Weight.Bold))
        btn_run.setStyleSheet("""
            QPushButton { background-color:#f48fb1; color:#333333; padding:10px; border-radius:5px; border:2px solid #d81b60; }
            QPushButton:pressed { background-color:#f06292; padding-top:12px; padding-left:12px; padding-bottom:8px; padding-right:8px; }
        """)
        btn_run.clicked.connect(self.on_run)
        layout.addWidget(btn_run)

        self.setLayout(layout)
        self.progress_dialog = None

    # --- 実行処理 ---
    def on_run(self):
        year, month, day = self.year_input.value(), self.month_input.value(), self.day_input.value()
        filter_type = "all" if self.rb_all.isChecked() else "dollar"

        # ファイル一覧取得（UNC対応）
        files = []
        try:
            for f in os.listdir(DEFAULT_IMPORT):
                fullpath = os.path.normpath(os.path.join(DEFAULT_IMPORT, f))
                if os.path.isfile(fullpath) and f.lower().endswith(".xls"):
                    if f.startswith(f"{month}-{day}"):
                        mtime = datetime.fromtimestamp(os.path.getmtime(fullpath)).strftime("%Y-%m-%d %H:%M")
                        files.append((f, mtime))
        except Exception as e:
            QMessageBox.critical(self, "参照エラー", f"参照先ディレクトリを読み込めません:\n{DEFAULT_IMPORT}\n{e}")
            return

        if len(files) > 1:
            dlg = FileSelectDialog(files)
            if dlg.exec() == QDialog.DialogCode.Accepted:
                selected_file = dlg.get_selected_file()
            else:
                return
        elif files:
            selected_file = files[0][0]
        else:
            selected_file = None

        # 進捗バー
        self.progress_dialog = QProgressDialog("処理中です...", None, 0, 0, self)
        self.progress_dialog.setWindowTitle("実行中")
        self.progress_dialog.setWindowModality(Qt.WindowModality.ApplicationModal)
        self.progress_dialog.setCancelButton(None)
        self.progress_dialog.setAutoClose(False)
        self.progress_dialog.setMinimumWidth(300)
        self.progress_dialog.setMinimumHeight(80)
        geo = self.frameGeometry()
        center_point = geo.center()
        self.progress_dialog.move(center_point - self.progress_dialog.rect().center())
        self.progress_dialog.show()

        self.worker = ScheduleWorker(year, month, day, filter_type, DEFAULT_IMPORT, DEFAULT_OUTPUT)
        self.worker.finished.connect(self.on_finished)
        self.worker.error.connect(self.on_error)
        self.worker.start()

    def on_finished(self, save_file, gifu_new, shiga_new, gifu_old, shiga_old):
        if self.progress_dialog:
            self.progress_dialog.close()
            self.progress_dialog = None
        total = gifu_new + shiga_new + gifu_old + shiga_old
        QMessageBox.information(
            self, "完了",
            f"保存完了: {save_file}\n"
            f"岐阜新: {gifu_new}, 滋賀新: {shiga_new},\n"
            f"岐阜旧: {gifu_old}, 滋賀旧: {shiga_old}\n"
            f"Totalは {total}台です"
        )

    def on_error(self, msg):
        if self.progress_dialog:
            self.progress_dialog.close()
            self.progress_dialog = None
        QMessageBox.critical(self, "エラー", msg)

if __name__ == "__main__":
    app = QApplication(sys.argv)
    win = CuteScheduleApp()
    win.show()
    sys.exit(app.exec())
