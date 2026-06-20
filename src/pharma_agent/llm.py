from __future__ import annotations

import json
import os
import urllib.error
import urllib.request
from pathlib import Path


class GeminiClient:
    def __init__(self, model: str = "gemini-2.0-flash") -> None:
        self.api_key = self._resolve_api_key()
        self.model = model

    @property
    def enabled(self) -> bool:
        return bool(self.api_key)

    def generate_json(self, prompt: str) -> dict:
        if not self.enabled:
            raise RuntimeError("GEMINI_API_KEY is not configured.")

        url = (
            f"https://generativelanguage.googleapis.com/v1beta/models/"
            f"{self.model}:generateContent?key={self.api_key}"
        )
        payload = {
            "contents": [{"parts": [{"text": prompt}]}],
            "generationConfig": {
                "temperature": 0.4,
                "responseMimeType": "application/json",
            },
        }
        request = urllib.request.Request(
            url,
            data=json.dumps(payload).encode("utf-8"),
            headers={"Content-Type": "application/json"},
            method="POST",
        )
        try:
            with urllib.request.urlopen(request, timeout=60) as response:
                body = json.loads(response.read().decode("utf-8"))
        except urllib.error.HTTPError as exc:
            detail = exc.read().decode("utf-8", errors="ignore")
            raise RuntimeError(f"Gemini request failed with HTTP {exc.code}: {detail}") from exc
        except urllib.error.URLError as exc:
            raise RuntimeError(f"Gemini request failed: {exc.reason}") from exc

        try:
            raw_text = body["candidates"][0]["content"]["parts"][0]["text"]
            return json.loads(raw_text)
        except (KeyError, IndexError, json.JSONDecodeError) as exc:
            raise RuntimeError(f"Gemini returned an unexpected response: {body}") from exc

    def _resolve_api_key(self) -> str:
        direct = os.getenv("GEMINI_API_KEY", "").strip()
        if direct:
            return direct
        env_path = Path.cwd() / ".env"
        if not env_path.exists():
            return ""
        for raw_line in env_path.read_text(encoding="utf-8").splitlines():
            line = raw_line.strip()
            if not line or line.startswith("#") or "=" not in line:
                continue
            key, value = line.split("=", 1)
            if key.strip() == "GEMINI_API_KEY":
                return value.strip().strip('"').strip("'")
        return ""
