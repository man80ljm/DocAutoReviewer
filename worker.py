from __future__ import annotations

import json
from pathlib import Path
from typing import Dict, List

from PyQt6.QtCore import QThread, pyqtSignal

from deepseek_client import DeepSeekClient, DeepSeekConfig
from docx_io import extract_reflection_text, insert_comment_and_save


class BatchWorker(QThread):
    progress_changed = pyqtSignal(int)
    current_file_changed = pyqtSignal(str)
    log_message = pyqtSignal(str)
    attention_message = pyqtSignal(str)
    finished_all = pyqtSignal()

    def __init__(
        self,
        input_dir: Path,
        base_url: str,
        api_key: str,
        model: str,
        style: str,
        expected_words: int,
        start_marker: str,
        end_marker: str,
        comment_label: str,
        overwrite_output: bool,
        parent=None,
    ) -> None:
        super().__init__(parent)
        self.input_dir = input_dir
        self.style = style
        self.expected_words = expected_words
        self.start_marker = start_marker
        self.end_marker = end_marker
        self.comment_label = comment_label
        self.overwrite_output = overwrite_output
        self._paused = False
        self._stop_requested = False
        self._client = DeepSeekClient(
            DeepSeekConfig(
                base_url=base_url,
                api_key=api_key,
                model=model,
                timeout_seconds=30,
                max_retries=2,
            )
        )

    def pause(self) -> None:
        self._paused = True

    def resume(self) -> None:
        self._paused = False

    def stop(self) -> None:
        self._stop_requested = True

    def run(self) -> None:
        output_dir = self.input_dir / "output"
        progress_path = output_dir / "progress.json"
        progress = self._load_progress(progress_path)

        files = self._collect_files(self.input_dir, output_dir)
        total = len(files)
        if total == 0:
            self.log_message.emit("未找到可处理的 .docx 文件。")
            self.finished_all.emit()
            return

        processed = 0
        for doc_path in files:
            if self._stop_requested:
                break
            while self._paused:
                self.msleep(200)
                if self._stop_requested:
                    break
            if self._stop_requested:
                break

            rel_key = str(doc_path.relative_to(self.input_dir))
            if progress.get(rel_key, {}).get("status") == "completed":
                processed += 1
                self._emit_progress(processed, total, doc_path.name, "已完成，跳过。")
                continue

            self.current_file_changed.emit(doc_path.name)
            try:
                reflection = extract_reflection_text(
                    doc_path,
                    start_marker=self.start_marker,
                    end_marker=self.end_marker,
                )
                if not reflection:
                    self._log_and_mark(
                        progress,
                        progress_path,
                        rel_key,
                        "skipped_reflection_missing",
                        "未找到反思内容，跳过。",
                    )
                    processed += 1
                    self._emit_progress(processed, total, doc_path.name, "未找到反思内容。")
                    continue

                comment = self._client.generate_comment(
                    reflection_text=reflection,
                    style=self.style,
                    expected_words=self.expected_words,
                )
                saved_path = insert_comment_and_save(
                    doc_path,
                    comment,
                    output_dir,
                    comment_label=self.comment_label,
                    overwrite_output=self.overwrite_output,
                )
                if not saved_path:
                    self._log_and_mark(
                        progress,
                        progress_path,
                        rel_key,
                        "skipped_insert_failed",
                        "未找到教师总评插入位置，跳过。",
                    )
                else:
                    self._log_and_mark(
                        progress,
                        progress_path,
                        rel_key,
                        "completed",
                        f"已保存：{saved_path.name}",
                    )
            except Exception as exc:
                if isinstance(exc, PermissionError):
                    self.attention_message.emit(f"{doc_path.name} 无法写入，请关闭文档后重试。")
                self._log_and_mark(
                    progress,
                    progress_path,
                    rel_key,
                    "error",
                    f"处理失败：{exc}",
                )

            processed += 1
            self._emit_progress(processed, total, doc_path.name, "完成。")

        self.finished_all.emit()

    def _emit_progress(self, processed: int, total: int, filename: str, note: str) -> None:
        percent = int((processed / total) * 100)
        self.progress_changed.emit(percent)
        self.log_message.emit(f"{filename} - {note}")

    def _log_and_mark(
        self,
        progress: Dict[str, Dict[str, str]],
        progress_path: Path,
        key: str,
        status: str,
        message: str,
    ) -> None:
        progress[key] = {"status": status}
        self._save_progress(progress_path, progress)
        self.log_message.emit(message)

    @staticmethod
    def _collect_files(input_dir: Path, output_dir: Path) -> List[Path]:
        files: List[Path] = []
        for path in input_dir.rglob("*.docx"):
            if path.name.startswith("~$"):
                continue
            try:
                path.relative_to(output_dir)
                continue
            except ValueError:
                pass
            files.append(path)
        return sorted(files)

    @staticmethod
    def _load_progress(progress_path: Path) -> Dict[str, Dict[str, str]]:
        if not progress_path.exists():
            return {}
        try:
            return json.loads(progress_path.read_text(encoding="utf-8"))
        except Exception:
            return {}

    @staticmethod
    def _save_progress(progress_path: Path, progress: Dict[str, Dict[str, str]]) -> None:
        progress_path.parent.mkdir(parents=True, exist_ok=True)
        progress_path.write_text(json.dumps(progress, ensure_ascii=False, indent=2), encoding="utf-8")
