from __future__ import annotations

import json
import os
import urllib.error
import urllib.request
from pathlib import Path


class LLMClient:
    """Minimal provider wrapper used by the planner.

    The rest of the codebase should not need to care whether Groq or Gemini is
    active. This class resolves the configured provider once and exposes a single
    JSON-generation method to the planner.
    """

    def __init__(self) -> None:
        self.config = self._resolve_config()

    @property
    def enabled(self) -> bool:
        return bool(self.config["api_key"] and self.config["provider"])

    @property
    def provider(self) -> str:
        return self.config["provider"]

    @property
    def model(self) -> str:
        return self.config["model"]

    def generate_json(self, prompt: str) -> dict:
        """Route the prompt to the configured provider and require JSON output."""
        if not self.enabled:
            raise RuntimeError("No supported LLM API key is configured.")
        if self.provider == "groq":
            return self._generate_json_with_groq(prompt)
        if self.provider == "gemini":
            return self._generate_json_with_gemini(prompt)
        raise RuntimeError(f"Unsupported LLM provider: {self.provider}")

    def _generate_json_with_groq(self, prompt: str) -> dict:
        payload = {
            "model": self.model,
            "temperature": 0.4,
            "response_format": {"type": "json_object"},
            "messages": [
                {
                    "role": "system",
                    "content": "You are a presentation-planning assistant. Return valid JSON only.",
                },
                {"role": "user", "content": prompt},
            ],
        }
        request = urllib.request.Request(
            "https://api.groq.com/openai/v1/chat/completions",
            data=json.dumps(payload).encode("utf-8"),
            headers={
                "Content-Type": "application/json",
                "Accept": "application/json",
                "User-Agent": "pharma-presentation-agent/1.0",
                "Authorization": f"Bearer {self.config['api_key']}",
            },
            method="POST",
        )
        body = self._execute_request(request, "Groq")
        try:
            return json.loads(body["choices"][0]["message"]["content"])
        except (KeyError, IndexError, json.JSONDecodeError) as exc:
            raise RuntimeError(f"Groq returned an unexpected response: {body}") from exc

    def _generate_json_with_gemini(self, prompt: str) -> dict:
        url = f"https://generativelanguage.googleapis.com/v1beta/models/{self.model}:generateContent?key={self.config['api_key']}"
        payload = {
            "contents": [{"parts": [{"text": prompt}]}],
            "generationConfig": {"temperature": 0.4, "responseMimeType": "application/json"},
        }
        request = urllib.request.Request(
            url,
            data=json.dumps(payload).encode("utf-8"),
            headers={
                "Content-Type": "application/json",
                "Accept": "application/json",
                "User-Agent": "pharma-presentation-agent/1.0",
            },
            method="POST",
        )
        body = self._execute_request(request, "Gemini")
        try:
            return json.loads(body["candidates"][0]["content"]["parts"][0]["text"])
        except (KeyError, IndexError, json.JSONDecodeError) as exc:
            raise RuntimeError(f"Gemini returned an unexpected response: {body}") from exc

    def _execute_request(self, request: urllib.request.Request, provider_label: str) -> dict:
        try:
            with urllib.request.urlopen(request, timeout=60) as response:
                return json.loads(response.read().decode("utf-8"))
        except urllib.error.HTTPError as exc:
            detail = exc.read().decode("utf-8", errors="ignore")
            raise RuntimeError(f"{provider_label} request failed with HTTP {exc.code}: {detail}") from exc
        except urllib.error.URLError as exc:
            raise RuntimeError(f"{provider_label} request failed: {exc.reason}") from exc

    def _resolve_config(self) -> dict[str, str]:
        env_values = self._load_env_values()
        provider_hint = (os.getenv("LLM_PROVIDER") or env_values.get("LLM_PROVIDER") or "").strip().lower()
        groq_key = (os.getenv("GROQ_API_KEY") or env_values.get("GROQ_API_KEY") or "").strip()
        gemini_key = (os.getenv("GEMINI_API_KEY") or env_values.get("GEMINI_API_KEY") or "").strip()

        # Provider hint wins when both keys exist, otherwise prefer Groq then Gemini.
        if provider_hint == "groq" and groq_key:
            return {"provider": "groq", "api_key": groq_key, "model": (os.getenv("GROQ_MODEL") or env_values.get("GROQ_MODEL") or "llama-3.3-70b-versatile").strip()}
        if provider_hint == "gemini" and gemini_key:
            return {"provider": "gemini", "api_key": gemini_key, "model": (os.getenv("GEMINI_MODEL") or env_values.get("GEMINI_MODEL") or "gemini-2.0-flash").strip()}
        if groq_key:
            return {"provider": "groq", "api_key": groq_key, "model": (os.getenv("GROQ_MODEL") or env_values.get("GROQ_MODEL") or "llama-3.3-70b-versatile").strip()}
        if gemini_key:
            return {"provider": "gemini", "api_key": gemini_key, "model": (os.getenv("GEMINI_MODEL") or env_values.get("GEMINI_MODEL") or "gemini-2.0-flash").strip()}
        return {"provider": "", "api_key": "", "model": ""}

    @staticmethod
    def _load_env_values() -> dict[str, str]:
        env_path = Path.cwd() / ".env"
        if not env_path.exists():
            return {}
        values: dict[str, str] = {}
        for raw_line in env_path.read_text(encoding="utf-8-sig").splitlines():
            line = raw_line.strip().lstrip("\ufeff")
            if not line or line.startswith("#") or "=" not in line:
                continue
            key, value = line.split("=", 1)
            values[key.strip()] = value.strip().strip('"').strip("'")
        return values


GeminiClient = LLMClient
