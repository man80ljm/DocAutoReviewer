from __future__ import annotations

import time
from dataclasses import dataclass
from typing import Optional

import requests


@dataclass
class DeepSeekConfig:
    base_url: str
    api_key: str
    model: str
    timeout_seconds: int = 30
    max_retries: int = 2


class DeepSeekClient:
    def __init__(self, config: DeepSeekConfig) -> None:
        self._config = config

    def generate_comment(
        self,
        reflection_text: str,
        style: str,
        expected_words: int,
    ) -> str:
        prompt = self._build_prompt(reflection_text, style, expected_words)
        url = f"{self._config.base_url.rstrip('/')}/v1/chat/completions"
        headers = {
            "Authorization": f"Bearer {self._config.api_key}",
            "Content-Type": "application/json",
        }
        payload = {
            "model": self._config.model,
            "messages": [
                {
                    "role": "system",
                    "content": "你是认真负责的老师，给出简洁、有建设性的中文教师总评。",
                },
                {"role": "user", "content": prompt},
            ],
            "temperature": 0.7,
        }

        last_error: Optional[Exception] = None
        for attempt in range(self._config.max_retries + 1):
            try:
                resp = requests.post(
                    url,
                    headers=headers,
                    json=payload,
                    timeout=self._config.timeout_seconds,
                )
                resp.raise_for_status()
                data = resp.json()
                return self._extract_text(data)
            except Exception as exc:
                last_error = exc
                if attempt < self._config.max_retries:
                    time.sleep(1.0)
                    continue
                break
        raise RuntimeError(f"DeepSeek request failed: {last_error}") from last_error

    def test_connection(self) -> None:
        url = f"{self._config.base_url.rstrip('/')}/v1/models"
        headers = {
            "Authorization": f"Bearer {self._config.api_key}",
            "Content-Type": "application/json",
        }
        last_error: Optional[Exception] = None
        for attempt in range(self._config.max_retries + 1):
            try:
                resp = requests.get(url, headers=headers, timeout=self._config.timeout_seconds)
                resp.raise_for_status()
                return
            except Exception as exc:
                last_error = exc
                if attempt < self._config.max_retries:
                    time.sleep(1.0)
                    continue
                break
        raise RuntimeError(f"DeepSeek connection test failed: {last_error}") from last_error

    @staticmethod
    def _extract_text(data: dict) -> str:
        choices = data.get("choices", [])
        if not choices:
            raise RuntimeError("DeepSeek response missing choices.")
        message = choices[0].get("message", {})
        content = message.get("content", "")
        if not isinstance(content, str) or not content.strip():
            raise RuntimeError("DeepSeek response content empty.")
        return content.strip()

    @staticmethod
    def _build_prompt(reflection_text: str, style: str, expected_words: int) -> str:
        style_hint = {
            "标准": "语气正式、肯定优点并指出改进点。",
            "学术": "语气偏学术、客观、强调实验规范与逻辑。",
            "幽默": "语气轻松幽默但必须尊重鼓励，不讽刺不攻击。",
            "极简": "用尽量短的句子表达核心评价。",
        }.get(style, "语气正式、肯定优点并指出改进点。")
        return (
            "请根据下面学生反思生成一段中文“教师总评”，仅一段话，纯文本，无Markdown。"
            f"字数约{expected_words}字，允许上下浮动。"
            f"{style_hint}\n\n"
            "学生反思：\n"
            f"{reflection_text}"
        )
