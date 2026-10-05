"""Shared Gemini configuration and lightweight token reporting."""

from __future__ import annotations

import os
from pathlib import Path
from typing import Any

from google import genai
from google.genai import types


ROOT = Path(__file__).resolve().parent
DEFAULT_MODEL = "gemini-3.5-flash-lite"


def load_env(path: Path = ROOT / ".env") -> None:
    """Load simple KEY=VALUE pairs without another dependency."""
    if not path.exists():
        return
    for raw_line in path.read_text(encoding="utf-8").splitlines():
        line = raw_line.strip()
        if not line or line.startswith("#") or "=" not in line:
            continue
        key, value = line.split("=", 1)
        os.environ.setdefault(key.strip(), value.strip().strip('"').strip("'"))


def require_api_key() -> None:
    load_env()
    if not (os.getenv("GEMINI_API_KEY") or os.getenv("GOOGLE_API_KEY")):
        raise SystemExit("Missing GEMINI_API_KEY. Copy .env.example to .env and add your key.")


def client() -> genai.Client:
    require_api_key()
    return genai.Client()


def model_name() -> str:
    load_env()
    return os.getenv("GEMINI_MODEL", DEFAULT_MODEL)


def token_limit(name: str, default: int) -> int:
    load_env()
    try:
        return max(64, int(os.getenv(name, str(default))))
    except ValueError:
        return default


def config(*, max_tokens: int, system: str, tools: list[Any] | None = None) -> types.GenerateContentConfig:
    """Use deterministic, concise generation settings for live demos."""
    return types.GenerateContentConfig(
        system_instruction=system,
        temperature=0.2,
        max_output_tokens=max_tokens,
        tools=tools,
    )


def usage_line(response: Any) -> str:
    usage = getattr(response, "usage_metadata", None)
    if not usage:
        return "Token usage: unavailable"
    prompt = getattr(usage, "prompt_token_count", 0) or 0
    output = getattr(usage, "candidates_token_count", 0) or 0
    total = getattr(usage, "total_token_count", prompt + output) or prompt + output
    return f"Tokens — input: {prompt}, output: {output}, total: {total}"


def generate(prompt: str, *, system: str, max_tokens: int, tools: list[Any] | None = None):
    return client().models.generate_content(
        model=model_name(),
        contents=prompt,
        config=config(max_tokens=max_tokens, system=system, tools=tools),
    )
