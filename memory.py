"""
Persistent memory for Jarvis.

Stores conversation history (last N turns) and long-term facts about Sir
in a single JSON file. Atomic writes prevent corruption on a mid-write crash.
"""

from __future__ import annotations

import datetime as dt
import json
import os
from pathlib import Path

HISTORY_LIMIT = 50  # turns kept in rolling history


class Memory:
    def __init__(self, path: Path) -> None:
        self.path = Path(path)
        self.data: dict = self._load()

    # -- file I/O ------------------------------------------------------------

    def _load(self) -> dict:
        if not self.path.exists():
            return {"history": [], "facts": [], "last_updated": None}
        try:
            raw = self.path.read_text(encoding="utf-8")
            data = json.loads(raw)
            data.setdefault("history", [])
            data.setdefault("facts", [])
            data.setdefault("last_updated", None)
            return data
        except (json.JSONDecodeError, OSError):
            # corrupt or unreadable — start fresh, but keep a backup
            try:
                backup = self.path.with_suffix(".bak")
                self.path.rename(backup)
            except OSError:
                pass
            return {"history": [], "facts": [], "last_updated": None}

    def save(self) -> None:
        self.data["last_updated"] = dt.datetime.now().isoformat(timespec="seconds")
        self.path.parent.mkdir(parents=True, exist_ok=True)
        tmp = self.path.with_suffix(self.path.suffix + ".tmp")
        tmp.write_text(
            json.dumps(self.data, ensure_ascii=False, indent=2),
            encoding="utf-8",
        )
        os.replace(tmp, self.path)

    # -- history -------------------------------------------------------------

    def get_history(self) -> list[dict]:
        return list(self.data.get("history", []))

    def set_history(self, history: list[dict]) -> None:
        # keep only last HISTORY_LIMIT turns to bound the file + context window
        self.data["history"] = list(history[-HISTORY_LIMIT:])

    # -- facts ---------------------------------------------------------------

    def get_facts(self) -> list[dict]:
        return list(self.data.get("facts", []))

    def add_fact(self, category: str, text: str) -> None:
        category = (category or "general").strip().lower()
        text = (text or "").strip()
        if not text:
            return
        # avoid exact duplicates
        for f in self.data.get("facts", []):
            if f.get("text", "").lower() == text.lower():
                return
        self.data.setdefault("facts", []).append({
            "category": category,
            "text": text,
            "added": dt.datetime.now().isoformat(timespec="seconds"),
        })
        self.save()

    def facts_by_category(self, category: str | None = None) -> list[dict]:
        facts = self.get_facts()
        if not category:
            return facts
        category = category.strip().lower()
        return [f for f in facts if f.get("category") == category]

    # -- summary helpers -----------------------------------------------------

    def last_topic(self) -> str | None:
        """Return a short snippet of the last user message, for briefings."""
        for entry in reversed(self.get_history()):
            if entry.get("role") == "user":
                content = entry.get("content")
                if isinstance(content, str) and content.strip():
                    snippet = content.strip()
                    return snippet[:80] + ("..." if len(snippet) > 80 else "")
        return None
