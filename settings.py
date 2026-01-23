from __future__ import annotations

import json
from dataclasses import dataclass, asdict
from pathlib import Path
from typing import Any


SETTINGS_PATH = Path(__file__).parent / "settings.json"


@dataclass
class AppSettings:
    base_url: str = "https://api.deepseek.com"
    api_key: str = ""
    model: str = "deepseek-chat"
    style: str = "标准"
    expected_words: int = 150
    start_marker: str = ""
    end_marker: str = ""
    comment_label: str = ""
    overwrite_output: bool = True


def load_settings() -> AppSettings:
    if not SETTINGS_PATH.exists():
        return AppSettings()
    try:
        raw = json.loads(SETTINGS_PATH.read_text(encoding="utf-8"))
        return AppSettings(
            base_url=str(raw.get("base_url", AppSettings.base_url)),
            api_key=str(raw.get("api_key", "")),
            model=str(raw.get("model", AppSettings.model)),
            style=str(raw.get("style", AppSettings.style)),
            expected_words=int(raw.get("expected_words", AppSettings.expected_words)),
            start_marker=str(raw.get("start_marker", "")),
            end_marker=str(raw.get("end_marker", "")),
            comment_label=str(raw.get("comment_label", "")),
            overwrite_output=bool(raw.get("overwrite_output", True)),
        )
    except Exception:
        # If settings are corrupted, fall back to defaults.
        return AppSettings()


def save_settings(settings: AppSettings) -> None:
    data: dict[str, Any] = asdict(settings)
    SETTINGS_PATH.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
