from __future__ import annotations

import sys
from datetime import datetime
from pathlib import Path

from PyQt6.QtGui import QIcon, QColor, QTextCharFormat
from PyQt6.QtWidgets import (
    QApplication,
    QDialog,
    QFileDialog,
    QGridLayout,
    QGroupBox,
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


STYLESHEET = """
QMainWindow {
    background-color: #f5f5f5;
}

QGroupBox {
    font-weight: bold;
    border: 1px solid #c0c0c0;
    border-radius: 6px;
    margin-top: 12px;
    padding-top: 10px;
    background-color: #ffffff;
}

QGroupBox::title {
    subcontrol-origin: margin;
    left: 12px;
    padding: 0 6px;
    color: #2c3e50;
}

QPushButton {
    background-color: #3498db;
    color: white;
    border: none;
    border-radius: 4px;
    padding: 8px 16px;
    font-size: 13px;
    min-width: 70px;
}

QPushButton:hover {
    background-color: #2980b9;
}

QPushButton:pressed {
    background-color: #1c6ea4;
}

QPushButton:disabled {
    background-color: #bdc3c7;
    color: #7f8c8d;
}

QPushButton#startBtn {
    background-color: #27ae60;
}
QPushButton#startBtn:hover {
    background-color: #219a52;
}

QPushButton#stopBtn {
    background-color: #e74c3c;
}
QPushButton#stopBtn:hover {
    background-color: #c0392b;
}

QPushButton#pauseBtn, QPushButton#resumeBtn {
    background-color: #f39c12;
}
QPushButton#pauseBtn:hover, QPushButton#resumeBtn:hover {
    background-color: #d68910;
}

QLineEdit, QComboBox, QSpinBox {
    border: 1px solid #bdc3c7;
    border-radius: 4px;
    padding: 6px 10px;
    background-color: #ffffff;
    font-size: 13px;
}

QLineEdit:focus, QComboBox:focus, QSpinBox:focus {
    border-color: #3498db;
}

QLineEdit:read-only {
    background-color: #ecf0f1;
    color: #7f8c8d;
}

QTextEdit {
    border: 1px solid #bdc3c7;
    border-radius: 4px;
    background-color: #ffffff;
    font-family: Consolas, Monaco, monospace;
    font-size: 12px;
}

QProgressBar {
    border: 1px solid #bdc3c7;
    border-radius: 4px;
    text-align: center;
    height: 22px;
}

QProgressBar::chunk {
    background-color: #27ae60;
    border-radius: 3px;
}

QLabel {
    color: #2c3e50;
    font-size: 13px;
}

QCheckBox {
    color: #2c3e50;
    font-size: 13px;
}

QCheckBox::indicator {
    width: 16px;
    height: 16px;
}
"""

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
        self.resize(800, 520)
        self._set_window_icon()
        self.setStyleSheet(STYLESHEET)

        self.settings = load_settings()
        self.worker: BatchWorker | None = None

        root = QWidget(self)
        self.setCentralWidget(root)
        main_layout = QVBoxLayout(root)
        main_layout.setSpacing(10)
        main_layout.setContentsMargins(12, 12, 12, 12)

        # === 顶部：输入文件夹（横向一行）===
        input_row = QHBoxLayout()
        input_row.setSpacing(8)
        input_label = QLabel("输入文件夹")
        input_label.setFixedWidth(70)
        self.input_dir_edit = QLineEdit()
        self.input_dir_edit.setPlaceholderText("选择包含 .docx 文件的文件夹...")
        browse_btn = QPushButton("浏览")
        browse_btn.setFixedWidth(60)
        browse_btn.clicked.connect(self._pick_folder)
        input_row.addWidget(input_label)
        input_row.addWidget(self.input_dir_edit, 1)
        input_row.addWidget(browse_btn)
        main_layout.addLayout(input_row)

        # === 中部：左右分栏 ===
        content_layout = QHBoxLayout()
        content_layout.setSpacing(12)

        # --- 左侧：配置面板 ---
        left_panel = QVBoxLayout()
        left_panel.setSpacing(10)

        # 标记与生成配置（合并）
        config_group = QGroupBox("标记与生成配置")
        config_layout = QGridLayout(config_group)
        config_layout.setSpacing(8)
        config_layout.setContentsMargins(10, 16, 10, 10)

        self.start_marker_edit = QLineEdit(self.settings.start_marker)
        self.start_marker_edit.setPlaceholderText("实验心得与反思")
        self.end_marker_edit = QLineEdit(self.settings.end_marker)
        self.end_marker_edit.setPlaceholderText("教师评语与评分")
        self.comment_label_edit = QLineEdit(self.settings.comment_label)
        self.comment_label_edit.setPlaceholderText("教师总评：")

        self.style_combo = QComboBox()
        self.style_combo.addItems(["标准", "学术", "幽默", "极简"])
        if self.settings.style in ["标准", "学术", "幽默", "极简"]:
            self.style_combo.setCurrentText(self.settings.style)

        self.words_spin = QSpinBox()
        self.words_spin.setRange(50, 500)
        self.words_spin.setValue(self.settings.expected_words)
        self.words_spin.setSuffix(" 字")

        self.overwrite_check = QCheckBox("覆盖已有输出")
        self.overwrite_check.setChecked(self.settings.overwrite_output)

        config_layout.addWidget(QLabel("开始标记"), 0, 0)
        config_layout.addWidget(self.start_marker_edit, 0, 1, 1, 3)
        config_layout.addWidget(QLabel("结束标记"), 1, 0)
        config_layout.addWidget(self.end_marker_edit, 1, 1, 1, 3)
        config_layout.addWidget(QLabel("评语标记"), 2, 0)
        config_layout.addWidget(self.comment_label_edit, 2, 1, 1, 3)
        config_layout.addWidget(QLabel("风格"), 3, 0)
        config_layout.addWidget(self.style_combo, 3, 1)
        config_layout.addWidget(QLabel("字数"), 3, 2)
        config_layout.addWidget(self.words_spin, 3, 3)
        config_layout.addWidget(self.overwrite_check, 4, 0, 1, 4)
        config_layout.setColumnStretch(1, 1)
        config_layout.setColumnStretch(3, 1)

        left_panel.addWidget(config_group)

        # API 配置（精简显示）
        api_group = QGroupBox("API 配置")
        api_layout = QVBoxLayout(api_group)
        api_layout.setSpacing(8)
        api_layout.setContentsMargins(10, 16, 10, 10)

        # 隐藏字段（用于保存和读取）
        self.base_url_edit = QLineEdit(self.settings.base_url)
        self.api_key_edit = QLineEdit(self.settings.api_key)
        self.model_edit = QLineEdit(self.settings.model)
        self.base_url_edit.hide()
        self.api_key_edit.hide()
        self.model_edit.hide()

        # 精简显示
        self.api_status_label = QLabel()
        self._update_api_status_label()
        settings_btn = QPushButton("修改 API 设置")
        settings_btn.clicked.connect(self._open_settings)

        api_layout.addWidget(self.api_status_label)
        api_layout.addWidget(settings_btn)

        left_panel.addWidget(api_group)
        left_panel.addStretch(1)

        # --- 右侧：日志面板 ---
        log_group = QGroupBox("运行日志")
        log_layout = QVBoxLayout(log_group)
        log_layout.setSpacing(6)
        log_layout.setContentsMargins(10, 16, 10, 10)

        self.log = QTextEdit()
        self.log.setReadOnly(True)
        log_layout.addWidget(self.log)

        clear_log_row = QHBoxLayout()
        clear_log_row.addStretch(1)
        self.clear_log_btn = QPushButton("清空")
        self.clear_log_btn.setFixedWidth(60)
        self.clear_log_btn.clicked.connect(self._clear_log)
        clear_log_row.addWidget(self.clear_log_btn)
        log_layout.addLayout(clear_log_row)

        # 左右放入 content_layout
        content_layout.addLayout(left_panel, 2)
        content_layout.addWidget(log_group, 3)

        main_layout.addLayout(content_layout, 1)

        # === 底部：进度条 + 当前文件 ===
        progress_row = QHBoxLayout()
        progress_row.setSpacing(12)
        self.progress = QProgressBar()
        self.progress.setValue(0)
        self.progress.setFixedHeight(20)
        self.current_file_label = QLabel("当前文件：-")
        self.current_file_label.setStyleSheet("color: #7f8c8d; font-style: italic;")
        self.current_file_label.setFixedWidth(200)
        progress_row.addWidget(self.progress, 1)
        progress_row.addWidget(self.current_file_label)
        main_layout.addLayout(progress_row)

        # === 底部：操作按钮 ===
        buttons = QHBoxLayout()
        buttons.setSpacing(10)

        self.start_btn = QPushButton("开始")
        self.start_btn.setObjectName("startBtn")
        self.pause_btn = QPushButton("暂停")
        self.pause_btn.setObjectName("pauseBtn")
        self.resume_btn = QPushButton("继续")
        self.resume_btn.setObjectName("resumeBtn")
        self.stop_btn = QPushButton("停止")
        self.stop_btn.setObjectName("stopBtn")

        self.pause_btn.setEnabled(False)
        self.resume_btn.setEnabled(False)
        self.stop_btn.setEnabled(False)

        self.start_btn.clicked.connect(self._start)
        self.pause_btn.clicked.connect(self._pause)
        self.resume_btn.clicked.connect(self._resume)
        self.stop_btn.clicked.connect(self._stop)

        buttons.addWidget(self.start_btn)
        buttons.addWidget(self.pause_btn)
        buttons.addWidget(self.resume_btn)
        buttons.addWidget(self.stop_btn)
        buttons.addStretch(1)

        main_layout.addLayout(buttons)

    def _update_api_status_label(self) -> None:
        model = self.settings.model or "未配置"
        has_key = "已配置" if self.settings.api_key else "未配置"
        self.api_status_label.setText(f"模型: {model}\nAPI Key: {has_key}")

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
        self.worker.log_message.connect(self._on_worker_log)
        self.worker.attention_message.connect(self._show_attention)
        self.worker.finished_all.connect(self._finished)

        self._set_running_state(True)
        self.worker.start()

    def _pause(self) -> None:
        if self.worker:
            self.worker.pause()
            self._append_warning("已暂停（将在当前文件完成后生效）。")
            self.pause_btn.setEnabled(False)
            self.resume_btn.setEnabled(True)

    def _resume(self) -> None:
        if self.worker:
            self.worker.resume()
            self._append_success("已继续。")
            self.pause_btn.setEnabled(True)
            self.resume_btn.setEnabled(False)

    def _stop(self) -> None:
        if self.worker:
            self.worker.stop()
            self._append_warning("已请求停止（将在当前文件完成后生效）。")
            self.stop_btn.setEnabled(False)

    def _finished(self) -> None:
        self._append_success("处理完成。")
        self._set_running_state(False)
        self.current_file_label.setText("当前文件：-")

    def _set_running_state(self, running: bool) -> None:
        self.start_btn.setEnabled(not running)
        self.pause_btn.setEnabled(running)
        self.resume_btn.setEnabled(False)
        self.stop_btn.setEnabled(running)

    def _append_log(self, message: str, level: str = "info") -> None:
        ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        color_map = {
            "info": "#2c3e50",
            "success": "#27ae60",
            "warning": "#f39c12",
            "error": "#e74c3c",
        }
        color = color_map.get(level, "#2c3e50")
        self.log.append(f'<span style="color:{color}">[{ts}] {message}</span>')

    def _append_success(self, message: str) -> None:
        self._append_log(message, "success")

    def _append_warning(self, message: str) -> None:
        self._append_log(message, "warning")

    def _append_error(self, message: str) -> None:
        self._append_log(message, "error")

    def _on_worker_log(self, message: str) -> None:
        """根据消息内容自动判断日志级别"""
        msg_lower = message.lower()
        if any(kw in message for kw in ["已保存", "完成", "成功", "跳过"]):
            if "失败" in message or "错误" in message:
                self._append_error(message)
            elif "跳过" in message or "未找到" in message:
                self._append_warning(message)
            else:
                self._append_success(message)
        elif any(kw in message for kw in ["失败", "错误", "error", "failed"]):
            self._append_error(message)
        elif any(kw in message for kw in ["警告", "warning", "未找到", "跳过"]):
            self._append_warning(message)
        else:
            self._append_log(message)

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
        self._update_api_status_label()

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
