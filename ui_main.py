from __future__ import annotations

import sys
from datetime import datetime
from pathlib import Path

from PyQt6.QtGui import QIcon
from PyQt6.QtWidgets import (
    QApplication,
    QDialog,
    QFileDialog,
    QGridLayout,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QMainWindow,
    QMessageBox,
    QPushButton,
    QSpinBox,
    QTextEdit,
    QVBoxLayout,
    QWidget,
    QComboBox,
    QProgressBar,
    QCheckBox,
)

from deepseek_client import DeepSeekClient, DeepSeekConfig
from settings import AppSettings, load_settings, save_settings
from worker import BatchWorker


class SettingsDialog(QDialog):
    def __init__(self, parent: QWidget, settings: AppSettings) -> None:
        super().__init__(parent)
        self.setWindowTitle("API 设置")
        self.resize(520, 240)
        self._settings = settings

        layout = QVBoxLayout(self)
        form = QGridLayout()
        layout.addLayout(form)

        self.base_url_edit = QLineEdit(settings.base_url)
        self.api_key_edit = QLineEdit(settings.api_key)
        self.api_key_edit.setEchoMode(QLineEdit.EchoMode.Password)
        self.model_edit = QLineEdit(settings.model)

        form.addWidget(QLabel("Base URL"), 0, 0)
        form.addWidget(self.base_url_edit, 0, 1, 1, 2)
        form.addWidget(QLabel("API Key"), 1, 0)
        form.addWidget(self.api_key_edit, 1, 1, 1, 2)
        form.addWidget(QLabel("Model"), 2, 0)
        form.addWidget(self.model_edit, 2, 1, 1, 2)

        buttons = QHBoxLayout()
        layout.addLayout(buttons)

        self.test_btn = QPushButton("测试连接")
        self.save_btn = QPushButton("保存")
        self.cancel_btn = QPushButton("取消")

        self.test_btn.clicked.connect(self._test_connection)
        self.save_btn.clicked.connect(self.accept)
        self.cancel_btn.clicked.connect(self.reject)

        buttons.addWidget(self.test_btn)
        buttons.addStretch(1)
        buttons.addWidget(self.save_btn)
        buttons.addWidget(self.cancel_btn)

    def apply_to_settings(self, settings: AppSettings) -> None:
        settings.base_url = self.base_url_edit.text().strip()
        settings.api_key = self.api_key_edit.text().strip()
        settings.model = self.model_edit.text().strip()

    def _test_connection(self) -> None:
        base_url = self.base_url_edit.text().strip()
        api_key = self.api_key_edit.text().strip()
        model = self.model_edit.text().strip()
        if not base_url or not api_key or not model:
            QMessageBox.warning(self, "提示", "请填写完整的 Base URL / API Key / Model。")
            return
        client = DeepSeekClient(
            DeepSeekConfig(
                base_url=base_url,
                api_key=api_key,
                model=model,
                timeout_seconds=15,
                max_retries=1,
            )
        )
        try:
            client.test_connection()
            QMessageBox.information(self, "成功", "连接成功。")
        except Exception as exc:
            QMessageBox.critical(self, "失败", f"连接失败：{exc}")


class MainWindow(QMainWindow):
    def __init__(self) -> None:
        super().__init__()
        self.setWindowTitle("DocAutoReviewer")
        self.resize(900, 600)
        self._set_window_icon()

        self.settings = load_settings()
        self.worker: BatchWorker | None = None

        root = QWidget(self)
        self.setCentralWidget(root)
        layout = QVBoxLayout(root)

        form = QGridLayout()
        layout.addLayout(form)

        self.input_dir_edit = QLineEdit()
        browse_btn = QPushButton("选择文件夹")
        browse_btn.clicked.connect(self._pick_folder)

        form.addWidget(QLabel("输入文件夹"), 0, 0)
        form.addWidget(self.input_dir_edit, 0, 1)
        form.addWidget(browse_btn, 0, 2)

        self.start_marker_edit = QLineEdit(self.settings.start_marker)
        self.end_marker_edit = QLineEdit(self.settings.end_marker)
        self.comment_label_edit = QLineEdit(self.settings.comment_label)
        self.overwrite_check = QCheckBox("覆盖输出文件")
        self.overwrite_check.setChecked(self.settings.overwrite_output)

        form.addWidget(QLabel("开始标记"), 1, 0)
        form.addWidget(self.start_marker_edit, 1, 1, 1, 2)
        form.addWidget(QLabel("结束标记"), 2, 0)
        form.addWidget(self.end_marker_edit, 2, 1, 1, 2)
        form.addWidget(QLabel("评语标记"), 3, 0)
        form.addWidget(self.comment_label_edit, 3, 1, 1, 1)
        form.addWidget(self.overwrite_check, 3, 2)

        self.base_url_edit = QLineEdit(self.settings.base_url)
        self.api_key_edit = QLineEdit(self.settings.api_key)
        self.api_key_edit.setEchoMode(QLineEdit.EchoMode.Password)
        self.model_edit = QLineEdit(self.settings.model)

        self.base_url_edit.setReadOnly(True)
        self.api_key_edit.setReadOnly(True)
        self.model_edit.setReadOnly(True)

        settings_btn = QPushButton("设置")
        settings_btn.clicked.connect(self._open_settings)

        form.addWidget(QLabel("Base URL"), 4, 0)
        form.addWidget(self.base_url_edit, 4, 1)
        form.addWidget(settings_btn, 4, 2)
        form.addWidget(QLabel("API Key"), 5, 0)
        form.addWidget(self.api_key_edit, 5, 1, 1, 2)
        form.addWidget(QLabel("Model"), 6, 0)
        form.addWidget(self.model_edit, 6, 1, 1, 2)

        self.style_combo = QComboBox()
        self.style_combo.addItems(["标准", "学术", "幽默", "极简"])
        if self.settings.style in ["标准", "学术", "幽默", "极简"]:
            self.style_combo.setCurrentText(self.settings.style)

        self.words_spin = QSpinBox()
        self.words_spin.setRange(50, 500)
        self.words_spin.setValue(self.settings.expected_words)

        form.addWidget(QLabel("风格"), 7, 0)
        form.addWidget(self.style_combo, 7, 1)
        form.addWidget(QLabel("预计字数"), 7, 2)
        form.addWidget(self.words_spin, 7, 3)

        buttons = QHBoxLayout()
        layout.addLayout(buttons)

        self.start_btn = QPushButton("Start")
        self.pause_btn = QPushButton("Pause")
        self.resume_btn = QPushButton("Resume")
        self.stop_btn = QPushButton("Stop")
        self.clear_log_btn = QPushButton("清空日志")

        self.pause_btn.setEnabled(False)
        self.resume_btn.setEnabled(False)
        self.stop_btn.setEnabled(False)

        self.start_btn.clicked.connect(self._start)
        self.pause_btn.clicked.connect(self._pause)
        self.resume_btn.clicked.connect(self._resume)
        self.stop_btn.clicked.connect(self._stop)
        self.clear_log_btn.clicked.connect(self._clear_log)

        buttons.addWidget(self.start_btn)
        buttons.addWidget(self.pause_btn)
        buttons.addWidget(self.resume_btn)
        buttons.addWidget(self.stop_btn)
        buttons.addWidget(self.clear_log_btn)

        self.progress = QProgressBar()
        self.progress.setValue(0)
        self.current_file_label = QLabel("当前文件：-")

        layout.addWidget(self.progress)
        layout.addWidget(self.current_file_label)

        self.log = QTextEdit()
        self.log.setReadOnly(True)
        layout.addWidget(self.log, 1)

    def _pick_folder(self) -> None:
        folder = QFileDialog.getExistingDirectory(self, "选择输入文件夹")
        if folder:
            self.input_dir_edit.setText(folder)

    def _set_window_icon(self) -> None:
        base_dir = Path(getattr(sys, "_MEIPASS", Path(__file__).parent))
        icon_path = base_dir / "report.ico"
        if icon_path.exists():
            self.setWindowIcon(QIcon(str(icon_path)))

    def _start(self) -> None:
        input_dir = Path(self.input_dir_edit.text().strip())
        if not input_dir.exists():
            QMessageBox.warning(self, "提示", "请输入有效的输入文件夹。")
            return
        if not self.api_key_edit.text().strip():
            QMessageBox.warning(self, "提示", "请输入 API Key。")
            return

        self._save_settings()
        self._append_log("开始处理。")

        self.worker = BatchWorker(
            input_dir=input_dir,
            base_url=self.base_url_edit.text().strip(),
            api_key=self.api_key_edit.text().strip(),
            model=self.model_edit.text().strip(),
            style=self.style_combo.currentText(),
            expected_words=int(self.words_spin.value()),
            start_marker=self.start_marker_edit.text().strip(),
            end_marker=self.end_marker_edit.text().strip(),
            comment_label=self.comment_label_edit.text().strip(),
            overwrite_output=self.overwrite_check.isChecked(),
        )
        self.worker.progress_changed.connect(self.progress.setValue)
        self.worker.current_file_changed.connect(
            lambda name: self.current_file_label.setText(f"当前文件：{name}")
        )
        self.worker.log_message.connect(self._append_log)
        self.worker.attention_message.connect(self._show_attention)
        self.worker.finished_all.connect(self._finished)

        self._set_running_state(True)
        self.worker.start()

    def _pause(self) -> None:
        if self.worker:
            self.worker.pause()
            self._append_log("已暂停（将在当前文件完成后生效）。")
            self.pause_btn.setEnabled(False)
            self.resume_btn.setEnabled(True)

    def _resume(self) -> None:
        if self.worker:
            self.worker.resume()
            self._append_log("已继续。")
            self.pause_btn.setEnabled(True)
            self.resume_btn.setEnabled(False)

    def _stop(self) -> None:
        if self.worker:
            self.worker.stop()
            self._append_log("已请求停止（将在当前文件完成后生效）。")
            self.stop_btn.setEnabled(False)

    def _finished(self) -> None:
        self._append_log("处理完成。")
        self._set_running_state(False)
        self.current_file_label.setText("当前文件：-")

    def _set_running_state(self, running: bool) -> None:
        self.start_btn.setEnabled(not running)
        self.pause_btn.setEnabled(running)
        self.resume_btn.setEnabled(False)
        self.stop_btn.setEnabled(running)

    def _append_log(self, message: str) -> None:
        ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        self.log.append(f"[{ts}] {message}")

    def _clear_log(self) -> None:
        self.log.clear()

    def _show_attention(self, message: str) -> None:
        QMessageBox.warning(self, "提示", message)

    def _open_settings(self) -> None:
        dialog = SettingsDialog(self, self.settings)
        if dialog.exec() == QDialog.DialogCode.Accepted:
            dialog.apply_to_settings(self.settings)
            save_settings(self.settings)
            self._refresh_settings_fields()
            QMessageBox.information(self, "已保存", "设置已保存。")

    def _refresh_settings_fields(self) -> None:
        self.base_url_edit.setText(self.settings.base_url)
        self.api_key_edit.setText(self.settings.api_key)
        self.model_edit.setText(self.settings.model)
        self.start_marker_edit.setText(self.settings.start_marker)
        self.end_marker_edit.setText(self.settings.end_marker)
        self.comment_label_edit.setText(self.settings.comment_label)
        self.overwrite_check.setChecked(self.settings.overwrite_output)

    def _save_settings(self) -> None:
        settings = AppSettings(
            base_url=self.base_url_edit.text().strip(),
            api_key=self.api_key_edit.text().strip(),
            model=self.model_edit.text().strip(),
            style=self.style_combo.currentText(),
            expected_words=int(self.words_spin.value()),
            start_marker=self.start_marker_edit.text().strip(),
            end_marker=self.end_marker_edit.text().strip(),
            comment_label=self.comment_label_edit.text().strip(),
            overwrite_output=self.overwrite_check.isChecked(),
        )
        save_settings(settings)
        self.settings = settings


def main() -> None:
    app = QApplication(sys.argv)
    base_dir = Path(getattr(sys, "_MEIPASS", Path(__file__).parent))
    icon_path = base_dir / "report.ico"
    if icon_path.exists():
        app.setWindowIcon(QIcon(str(icon_path)))
    window = MainWindow()
    window.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
