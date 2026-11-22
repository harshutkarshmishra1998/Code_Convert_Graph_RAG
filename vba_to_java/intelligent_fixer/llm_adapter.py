"""Optional ChatGroq adapter for the intelligent fixer's JSON prompts."""
from __future__ import annotations

import importlib
import json
import os
import sys
from pathlib import Path
from typing import Any


class ChatGroqJSONAdapter:
    def __init__(self, client: Any, model: str):
        self.client = client
        self.model = model

    @classmethod
    def from_project(cls, project_root: str | Path) -> "ChatGroqJSONAdapter":
        root = Path(project_root).expanduser().resolve()
        if str(root) not in sys.path:
            sys.path.insert(0, str(root))
        try:
            api_keys = importlib.import_module("api_keys")
            from langchain_groq import ChatGroq

            model = os.getenv("LLM_MODEL") or api_keys._loaded["LLM"]
            return cls(ChatGroq(model=model, temperature=0), str(model))
        except Exception as exc:
            raise RuntimeError(
                "Could not initialize the intelligent-fixer ChatGroq client. "
                "Check GROQ_API and LLM in the project environment."
            ) from exc

    def __call__(self, prompt: dict[str, Any]) -> dict[str, Any]:
        response = self.client.invoke([
            (
                "system",
                "Return exactly one valid JSON object matching the requested schema. "
                "Do not wrap it in Markdown fences or add commentary.",
            ),
            ("human", json.dumps(prompt, ensure_ascii=False, default=str)),
        ])
        content = getattr(response, "content", None)
        if isinstance(content, list):
            content = "\n".join(
                part.get("text", "") if isinstance(part, dict) else str(part)
                for part in content
            )
        if not isinstance(content, str):
            raise TypeError(
                f"Expected text content from ChatGroq; got {type(content).__name__}"
            )

        text = content.strip()
        if text.startswith("```"):
            lines = text.splitlines()
            if lines and lines[-1].strip() == "```":
                text = "\n".join(lines[1:-1]).strip()
        result = json.loads(text)
        if not isinstance(result, dict):
            raise ValueError("ChatGroq response must be a JSON object")
        return result
