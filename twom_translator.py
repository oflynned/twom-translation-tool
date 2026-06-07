"""
TWoM Translation Tool

Workflow:
  1. Import .lang -> creates a working CSV
  2. Edit translations in this tool (or open CSV in Excel / share with team)
  3. Pack to .lang -> produces binary file for Storyteller
  4. Open in Storyteller -> Pack -> test in game
"""

from __future__ import annotations

import json
import re
import threading
import urllib.error
import urllib.parse
import urllib.request
from collections import Counter, defaultdict
from pathlib import Path
from typing import Optional

import tkinter as tk
from tkinter import filedialog, messagebox, ttk

from lang_parser import LangRecord, build_lang, load_csv, load_lang, save_csv

CFG = Path.home() / ".twom.json"
FILTER_OPTIONS = [
    ("all", "All"),
    ("untranslated", "Untranslated"),
    ("draft", "Draft"),
    ("reviewed", "Reviewed"),
    ("warnings", "Warnings"),
]
FILTER_LABELS = {key: label for key, label in FILTER_OPTIONS}
FILTER_LABEL_TO_KEY = {label: key for key, label in FILTER_OPTIONS}
ROW_STATES = {"untranslated", "draft", "reviewed"}
PLACEHOLDER_RE = re.compile(r"\{[^{}]+\}")
CARET_PLACEHOLDER_RE = re.compile(r"\^[^^\r\n]+\^")
SIMILAR_SPLIT_RE = re.compile(r"[/_.\-\s]+")
DEFAULT_PLACEHOLDER_TEMPLATES = ["{ms|}", "{fs|}"]
AI_MAX_OUTPUT_TOKENS = 2048

# Provider metadata drives the settings UI and request dispatch in one place.
AI_PROVIDER_SPECS = {
    "anthropic": {
        "label": "Anthropic",
        "key_label": "Anthropic API key",
        "help": "Get a key at console.anthropic.com",
        "default_model": "claude-sonnet-4-20250514",
    },
    "openai": {
        "label": "OpenAI",
        "key_label": "OpenAI API key",
        "help": "Get a key at platform.openai.com/api-keys",
        "default_model": "gpt-4o-mini",
    },
    "gemini": {
        "label": "Google Gemini",
        "key_label": "Gemini API key",
        "help": "Get a key at aistudio.google.com/apikey",
        "default_model": "gemini-2.5-flash",
    },
    "openrouter": {
        "label": "OpenRouter",
        "key_label": "OpenRouter API key",
        "help": "Get a key at openrouter.ai/keys",
        "default_model": "openai/gpt-4o-mini",
    },
}
AI_PROVIDER_LABELS = {key: spec["label"] for key, spec in AI_PROVIDER_SPECS.items()}
AI_PROVIDER_LABEL_TO_KEY = {label: key for key, label in AI_PROVIDER_LABELS.items()}


def normalize_provider(value: Optional[str]) -> str:
    return value if value in AI_PROVIDER_SPECS else "anthropic"


def load_cfg() -> dict:
    try:
        return json.loads(CFG.read_text(encoding="utf-8")) if CFG.exists() else {}
    except Exception:
        return {}


def save_cfg(cfg: dict) -> None:
    try:
        CFG.write_text(json.dumps(cfg, indent=2), encoding="utf-8")
    except Exception:
        pass


def migrate_cfg(cfg: dict) -> dict:
    if not isinstance(cfg, dict):
        cfg = {}

    provider = normalize_provider(cfg.get("ai_provider"))
    if "ai_provider" not in cfg and cfg.get("api_key"):
        provider = "anthropic"

    api_keys = cfg.get("api_keys")
    if not isinstance(api_keys, dict):
        api_keys = {}
    legacy_key = cfg.get("api_key")
    if isinstance(legacy_key, str) and legacy_key.strip() and not api_keys.get("anthropic"):
        api_keys["anthropic"] = legacy_key.strip()

    ai_models = cfg.get("ai_models")
    if not isinstance(ai_models, dict):
        ai_models = {}
    legacy_model = cfg.get("ai_model")
    if isinstance(legacy_model, str) and legacy_model.strip() and not ai_models.get(provider):
        ai_models[provider] = legacy_model.strip()

    cleaned_keys = {}
    for key, value in api_keys.items():
        if key in AI_PROVIDER_SPECS and isinstance(value, str) and value.strip():
            cleaned_keys[key] = value.strip()

    cleaned_models = {}
    for key, spec in AI_PROVIDER_SPECS.items():
        value = ai_models.get(key, "")
        cleaned_models[key] = value.strip() if isinstance(value, str) and value.strip() else spec["default_model"]

    cfg["ai_provider"] = provider
    cfg["api_keys"] = cleaned_keys
    cfg["ai_models"] = cleaned_models
    return cfg


def meta_path_for(csv_path: Optional[Path]) -> Optional[Path]:
    return Path(f"{csv_path}.twom-meta.json") if csv_path else None


def load_meta(csv_path: Optional[Path]) -> dict:
    meta_path = meta_path_for(csv_path)
    if not meta_path or not meta_path.exists():
        return {}
    try:
        return json.loads(meta_path.read_text(encoding="utf-8"))
    except Exception:
        return {}


def save_meta(csv_path: Optional[Path], data: dict) -> None:
    meta_path = meta_path_for(csv_path)
    if not meta_path:
        return
    meta_path.write_text(json.dumps(data, indent=2), encoding="utf-8")


def pack_meta_path_for(lang_path: Path) -> Path:
    return Path(f"{lang_path}.twom-pack.json")


def build_translation_prompt(text: str, src: str, tgt: str, item_key: str) -> str:
    return (
        f"Translate this {src} game text to {tgt} for 'This War of Mine' "
        f"(somber survival game). The localization key may clarify context or UI location. "
        f"Keep {{0}}, %s, \\n, and ^CharacterName^ intact. "
        f"Return only the translated text.\n\n"
        f"Localization key:\n{item_key}\n\n"
        f"Source text:\n{text}"
    )


def build_feedback_prompt(source: str, translation: str, src: str, tgt: str, item_key: str) -> str:
    return (
        f"Review this {tgt} translation of a {src} game localization string for 'This War of Mine'. "
        "Use the localization key as extra context if it helps disambiguate meaning. "
        "Focus on accuracy, tone, naturalness, grammar, and placeholder safety. "
        "Do not rewrite the whole translation unless needed. "
        "Give concise feedback with specific improvement ideas.\n\n"
        f"Localization key:\n{item_key}\n\n"
        f"Source:\n{source}\n\n"
        f"Translation:\n{translation}"
    )


def normalize_curly_placeholder(token: str) -> str:
    inner = token[1:-1]
    if "|" not in inner:
        return token
    selector, _value = inner.split("|", 1)
    selector = selector.strip()
    return f"{{{selector}|*}}" if selector else token


def curly_placeholder_counter(text: str) -> Counter:
    return Counter(normalize_curly_placeholder(token) for token in PLACEHOLDER_RE.findall(text))


def placeholder_insert_token(token: str) -> str:
    inner = token[1:-1] if token.startswith("{") and token.endswith("}") else ""
    if "|" not in inner:
        return token
    selector, _value = inner.split("|", 1)
    selector = selector.strip()
    return f"{{{selector}|}}" if selector else token


def _extract_error_message(payload: object) -> Optional[str]:
    if isinstance(payload, dict):
        error = payload.get("error")
        if isinstance(error, dict):
            for key in ("message", "details", "status"):
                value = error.get(key)
                if isinstance(value, str) and value.strip():
                    return value.strip()
        for key in ("message", "detail"):
            value = payload.get(key)
            if isinstance(value, str) and value.strip():
                return value.strip()
    return None


def provider_finish_reason(provider: str, payload: dict) -> Optional[str]:
    if provider == "anthropic":
        reason = payload.get("stop_reason")
        return reason.strip() if isinstance(reason, str) and reason.strip() else None
    if provider in {"openai", "openrouter"}:
        choices = payload.get("choices") or []
        if choices and isinstance(choices[0], dict):
            reason = choices[0].get("finish_reason")
            return reason.strip() if isinstance(reason, str) and reason.strip() else None
        return None
    if provider == "gemini":
        candidates = payload.get("candidates") or []
        if candidates and isinstance(candidates[0], dict):
            reason = candidates[0].get("finishReason")
            return reason.strip() if isinstance(reason, str) and reason.strip() else None
    return None


def ensure_not_truncated(provider: str, payload: dict) -> None:
    reason = provider_finish_reason(provider, payload)
    truncated_reasons = {"max_tokens", "length", "MAX_TOKENS"}
    if reason in truncated_reasons:
        raise RuntimeError(f"{AI_PROVIDER_SPECS[provider]['label']} truncated the response ({reason}).")


def post_json(url: str, body: dict, headers: dict[str, str]) -> dict:
    data = json.dumps(body).encode("utf-8")
    req = urllib.request.Request(url, data=data, method="POST")
    req.add_header("Content-Type", "application/json")
    for key, value in headers.items():
        req.add_header(key, value)
    try:
        with urllib.request.urlopen(req, timeout=30) as response:
            payload = json.loads(response.read())
    except urllib.error.HTTPError as exc:
        raw = exc.read().decode("utf-8", errors="replace")
        try:
            payload = json.loads(raw)
        except Exception:
            payload = raw
        message = _extract_error_message(payload) if isinstance(payload, dict) else raw.strip()
        raise RuntimeError(f"HTTP {exc.code}: {message or exc.reason}") from exc
    except urllib.error.URLError as exc:
        raise RuntimeError(str(exc.reason or exc)) from exc
    if not isinstance(payload, dict):
        raise RuntimeError("Provider returned an unexpected response.")
    return payload


def _extract_anthropic_text(payload: dict) -> str:
    parts = payload.get("content") or []
    text_parts = [part.get("text", "") for part in parts if isinstance(part, dict) and part.get("type") == "text"]
    result = "".join(text_parts).strip()
    if not result:
        raise RuntimeError("Anthropic returned no text.")
    return result


def _extract_openai_text(payload: dict) -> str:
    choices = payload.get("choices") or []
    if not choices:
        raise RuntimeError("Provider returned no choices.")
    message = choices[0].get("message", {})
    content = message.get("content", "")
    if isinstance(content, str):
        result = content.strip()
        if result:
            return result
    if isinstance(content, list):
        parts = []
        for item in content:
            if not isinstance(item, dict):
                continue
            text = item.get("text")
            if isinstance(text, str) and text.strip():
                parts.append(text)
            elif isinstance(text, dict):
                value = text.get("value")
                if isinstance(value, str) and value.strip():
                    parts.append(value)
        result = "".join(parts).strip()
        if result:
            return result
    raise RuntimeError("Provider returned no text.")


def _extract_gemini_text(payload: dict) -> str:
    candidates = payload.get("candidates") or []
    if not candidates:
        raise RuntimeError("Gemini returned no candidates.")
    content = candidates[0].get("content", {})
    parts = content.get("parts") or []
    text_parts = [part.get("text", "") for part in parts if isinstance(part, dict) and isinstance(part.get("text"), str)]
    result = "".join(text_parts).strip()
    if not result:
        raise RuntimeError("Gemini returned no text.")
    return result


def ai_complete(provider: str, model: str, key: str, system_prompt: str, user_prompt: str) -> str:
    messages = [{"role": "user", "content": user_prompt}]

    if provider == "anthropic":
        payload = post_json(
            "https://api.anthropic.com/v1/messages",
            {
                "model": model,
                "max_tokens": AI_MAX_OUTPUT_TOKENS,
                "system": system_prompt,
                "messages": messages,
            },
            {
                "x-api-key": key,
                "anthropic-version": "2023-06-01",
            },
        )
        ensure_not_truncated(provider, payload)
        return _extract_anthropic_text(payload)

    if provider == "openai":
        payload = post_json(
            "https://api.openai.com/v1/chat/completions",
            {
                "model": model,
                "max_completion_tokens": AI_MAX_OUTPUT_TOKENS,
                "messages": [
                    {
                        "role": "system",
                        "content": system_prompt,
                    },
                    {"role": "user", "content": user_prompt},
                ],
            },
            {"Authorization": f"Bearer {key}"},
        )
        ensure_not_truncated(provider, payload)
        return _extract_openai_text(payload)

    if provider == "openrouter":
        payload = post_json(
            "https://openrouter.ai/api/v1/chat/completions",
            {
                "model": model,
                "max_tokens": AI_MAX_OUTPUT_TOKENS,
                "messages": [
                    {
                        "role": "system",
                        "content": system_prompt,
                    },
                    {"role": "user", "content": user_prompt},
                ],
            },
            {"Authorization": f"Bearer {key}"},
        )
        ensure_not_truncated(provider, payload)
        return _extract_openai_text(payload)

    if provider == "gemini":
        payload = post_json(
            f"https://generativelanguage.googleapis.com/v1beta/models/{urllib.parse.quote(model, safe='')}:generateContent",
            {
                "systemInstruction": {"parts": [{"text": system_prompt}]},
                "contents": [{"role": "user", "parts": [{"text": user_prompt}]}],
                "generationConfig": {"maxOutputTokens": AI_MAX_OUTPUT_TOKENS},
            },
            {"x-goog-api-key": key},
        )
        ensure_not_truncated(provider, payload)
        return _extract_gemini_text(payload)

    raise RuntimeError(f"Unsupported AI provider: {provider}")


def ai_translate(text: str, src: str, tgt: str, item_key: str, provider: str, model: str, key: str) -> str:
    return ai_complete(
        provider,
        model,
        key,
        "You translate game localization strings. Return only the translated text.",
        build_translation_prompt(text, src, tgt, item_key),
    )


def ai_feedback(source: str, translation: str, src: str, tgt: str, item_key: str, provider: str, model: str, key: str) -> str:
    return ai_complete(
        provider,
        model,
        key,
        "You review game localization translations. Be constructive, concise, and practical.",
        build_feedback_prompt(source, translation, src, tgt, item_key),
    )


BG = "#FAFAF9"
BG2 = "#F2F1EF"
BORDER = "#E2E1DE"
TEXT = "#1C1C1A"
TEXT2 = "#6B6B68"
ACCENT = "#1D6F5A"
ACCENTL = "#EAF4F1"
ACCENTT = "#FFFFFF"
DANGER = "#C0392B"
DONE = "#27AE60"
WARN_BG = "#FCE8E6"

FONT = ("Segoe UI", 10)
FONTB = ("Segoe UI", 10, "bold")
FONTS = ("Segoe UI", 9)
FONTM = ("Consolas", 9)


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("TWoM Translation Tool")
        self.geometry("1320x860")
        self.minsize(980, 640)
        self.configure(bg=BG)

        self.cfg = migrate_cfg(load_cfg())
        self.ai_provider = tk.StringVar(value=self.cfg.get("ai_provider", "anthropic"))
        self._ai_provider_keys = {
            key: self.cfg.get("api_keys", {}).get(key, "")
            for key in AI_PROVIDER_SPECS
        }
        self._ai_provider_models = {
            key: self.cfg.get("ai_models", {}).get(key, AI_PROVIDER_SPECS[key]["default_model"])
            for key in AI_PROVIDER_SPECS
        }
        self._active_ai_provider = normalize_provider(self.ai_provider.get())
        self.active_api_key = tk.StringVar(value=self._ai_provider_keys.get(self._active_ai_provider, ""))
        self.ai_model = tk.StringVar(
            value=self._ai_provider_models.get(
                self._active_ai_provider,
                AI_PROVIDER_SPECS[self._active_ai_provider]["default_model"],
            )
        )
        self.tgt_lang = tk.StringVar(value=self.cfg.get("tgt_lang", "Irish"))
        self._ai_provider_name_var = tk.StringVar()
        self._ai_key_label_var = tk.StringVar()
        self._ai_key_hint_var = tk.StringVar()
        self._ai_model_hint_var = tk.StringVar()

        self.csv_path: Optional[Path] = None
        self.records: list[LangRecord] = []
        self.translations: dict[str, str] = {}
        self.review_state: dict[str, str] = {}

        self.changed = False
        self.ai_busy = False

        self._record_by_key: dict[str, LangRecord] = {}
        self._source_to_keys: dict[str, list[str]] = {}
        self._group_to_keys: dict[str, list[str]] = {}
        self._grp_ids: dict[str, Optional[str]] = {}
        self._group_iids: dict[Optional[str], str] = {}
        self._displayed_keys: list[str] = []

        self._cur_key: Optional[str] = None
        self._cur_grp: Optional[str] = None
        self._ai_suggestion_text = ""
        self._ai_output_kind = "suggestion"
        self._source_lang_name: Optional[str] = None
        self._last_pack_path: Optional[Path] = None

        self._search_q = tk.StringVar()
        self._filter_mode = tk.StringVar(value="all")
        self._status = tk.StringVar(value="Open a .lang file or existing CSV to begin.")
        self._progress = tk.StringVar(value="")
        self._state_var = tk.StringVar(value="State: untranslated")
        self._stats_var = tk.StringVar(value="Length: 0 chars")
        self._warning_var = tk.StringVar(value="No format warnings.")
        self._source_hint_var = tk.StringVar(value="No matching translations yet.")
        self._similar_hint_var = tk.StringVar(value="No similar key suggestions yet.")
        self._placeholder_label_var = tk.StringVar(value="Insert placeholder")
        self._ai_panel_label_var = tk.StringVar(value="AI suggestion")

        self._suspend_refresh = False

        self._build_styles()
        self._build_ui()
        self._build_menu()
        self._bind_shortcuts()
        self.protocol("WM_DELETE_WINDOW", self._on_close)

        self._search_q.trace_add("write", lambda *_: self._on_filter_inputs_changed())
        self._filter_mode.trace_add("write", lambda *_: self._on_filter_inputs_changed())
        self.ai_provider.trace_add("write", lambda *_: self._on_ai_provider_changed())
        self._refresh_ai_provider_vars()
        self.after(0, self._open_last_csv_on_startup)

    def _refresh_ai_provider_vars(self):
        provider = normalize_provider(self.ai_provider.get())
        spec = AI_PROVIDER_SPECS[provider]
        self._ai_provider_name_var.set(spec["label"])
        self._ai_key_label_var.set(spec["key_label"])
        self._ai_key_hint_var.set(spec["help"])
        self._ai_model_hint_var.set(f"Suggested model: {spec['default_model']}")

    def _sync_ai_preferences(self, provider: Optional[str] = None):
        provider = normalize_provider(provider or self.ai_provider.get())
        self._ai_provider_keys[provider] = self.active_api_key.get().strip()
        self._ai_provider_models[provider] = self.ai_model.get().strip() or AI_PROVIDER_SPECS[provider]["default_model"]
        self._active_ai_provider = provider

    def _on_ai_provider_changed(self):
        new_provider = normalize_provider(self.ai_provider.get())
        old_provider = self._active_ai_provider
        self._sync_ai_preferences(old_provider)
        self._active_ai_provider = new_provider
        self.active_api_key.set(self._ai_provider_keys.get(new_provider, ""))
        self.ai_model.set(
            self._ai_provider_models.get(new_provider, AI_PROVIDER_SPECS[new_provider]["default_model"])
            or AI_PROVIDER_SPECS[new_provider]["default_model"]
        )
        self._refresh_ai_provider_vars()

    def _save_settings_cfg(self):
        provider = normalize_provider(self.ai_provider.get())
        self._sync_ai_preferences(provider)
        selected_model = self._ai_provider_models.get(provider, AI_PROVIDER_SPECS[provider]["default_model"])
        payload = {
            "tgt_lang": self.tgt_lang.get(),
            "ai_provider": provider,
            "api_keys": {key: value for key, value in self._ai_provider_keys.items() if value},
            "ai_models": {
                key: self._ai_provider_models.get(key, spec["default_model"]) or spec["default_model"]
                for key, spec in AI_PROVIDER_SPECS.items()
            },
            "last_csv_path": str(self.csv_path) if self.csv_path else "",
            "api_key": self._ai_provider_keys.get("anthropic", ""),
            "ai_model": selected_model,
        }
        self.cfg.update(payload)
        save_cfg(self.cfg)

    def _get_ai_settings(self) -> Optional[tuple[str, str, str]]:
        provider = normalize_provider(self.ai_provider.get())
        self._sync_ai_preferences(provider)
        key = self._ai_provider_keys.get(provider, "")
        if not key:
            spec = AI_PROVIDER_SPECS[provider]
            self._status.set(f"Add a {spec['label']} API key in Settings to use AI translation.")
            self._open_settings()
            return None
        model = self._ai_provider_models.get(provider, AI_PROVIDER_SPECS[provider]["default_model"])
        return provider, model, key

    def _remember_current_csv(self):
        self.cfg["last_csv_path"] = str(self.csv_path) if self.csv_path else ""
        save_cfg(self.cfg)

    def _open_csv_path(self, path: Path, show_errors: bool = True) -> bool:
        try:
            records, translations = load_csv(path)
        except Exception as exc:
            if show_errors:
                messagebox.showerror("Open error", str(exc))
            return False
        self.csv_path = path
        self._source_lang_name = None
        self._last_pack_path = None
        self._load_records(records, translations, path.name)
        self._remember_current_csv()
        return True

    def _open_last_csv_on_startup(self):
        raw_path = self.cfg.get("last_csv_path", "")
        if not isinstance(raw_path, str) or not raw_path.strip():
            return
        path = Path(raw_path)
        if not path.exists():
            self.cfg["last_csv_path"] = ""
            save_cfg(self.cfg)
            self._status.set("Last opened CSV was not found. Open a .lang file or existing CSV to begin.")
            return
        if not self._open_csv_path(path, show_errors=False):
            self._status.set("Could not reopen the last CSV automatically. Open a .lang file or existing CSV to begin.")

    def _build_styles(self):
        style = ttk.Style(self)
        style.theme_use("clam")

        style.configure(".", background=BG, foreground=TEXT, font=FONT, borderwidth=0)
        style.configure("TFrame", background=BG)
        style.configure("TLabel", background=BG, foreground=TEXT, font=FONT)
        style.configure("Dim.TLabel", background=BG, foreground=TEXT2, font=FONTS)
        style.configure("Side.TLabel", background=BG2, foreground=TEXT2, font=FONTS)

        style.configure(
            "TEntry",
            fieldbackground="#FFFFFF",
            foreground=TEXT,
            insertcolor=ACCENT,
            borderwidth=1,
            relief="flat",
            padding=(6, 5),
        )

        style.configure(
            "Primary.TButton",
            background=ACCENT,
            foreground=ACCENTT,
            font=FONTB,
            borderwidth=0,
            relief="flat",
            padding=(14, 7),
        )
        style.map(
            "Primary.TButton",
            background=[("active", "#165947"), ("pressed", "#0F3D31")],
        )

        style.configure(
            "Ghost.TButton",
            background=BG,
            foreground=TEXT,
            font=FONT,
            borderwidth=1,
            relief="flat",
            padding=(10, 5),
        )
        style.map(
            "Ghost.TButton",
            background=[("active", BG2)],
            foreground=[("active", ACCENT)],
        )

        style.configure(
            "Strings.Treeview",
            background="#FFFFFF",
            foreground=TEXT,
            fieldbackground="#FFFFFF",
            rowheight=28,
            font=FONT,
            borderwidth=0,
        )
        style.configure(
            "Strings.Treeview.Heading",
            background=BG2,
            foreground=TEXT2,
            font=FONTS,
            relief="flat",
            borderwidth=0,
            padding=(8, 6),
        )
        style.map(
            "Strings.Treeview",
            background=[("selected", ACCENTL)],
            foreground=[("selected", ACCENT)],
        )

        style.configure(
            "Groups.Treeview",
            background=BG2,
            foreground=TEXT,
            fieldbackground=BG2,
            rowheight=26,
            font=FONTS,
            borderwidth=0,
        )
        style.configure(
            "Groups.Treeview.Heading",
            background=BG2,
            foreground=TEXT2,
            font=FONTS,
            relief="flat",
        )
        style.map(
            "Groups.Treeview",
            background=[("selected", ACCENTL)],
            foreground=[("selected", ACCENT)],
        )

        style.configure(
            "TScrollbar",
            background=BG2,
            troughcolor=BG,
            arrowcolor=TEXT2,
            borderwidth=0,
            width=8,
        )
        style.configure(
            "TProgressbar",
            background=ACCENT,
            troughcolor=BG2,
            borderwidth=0,
            thickness=3,
        )
        style.configure("TCombobox", fieldbackground="#FFFFFF", padding=(6, 4))

    def _build_ui(self):
        topbar = tk.Frame(self, bg=BG)
        topbar.pack(fill="x")

        left_bar = tk.Frame(topbar, bg=BG, padx=16, pady=10)
        left_bar.pack(side="left")

        tk.Label(
            left_bar,
            text="TWoM Translation",
            bg=BG,
            fg=TEXT,
            font=("Segoe UI", 12, "bold"),
        ).pack(side="left", padx=(0, 20))

        ttk.Button(left_bar, text="Import .lang", command=self._import_lang, style="Ghost.TButton").pack(side="left", padx=2)
        ttk.Button(left_bar, text="Open CSV", command=self._open_csv, style="Ghost.TButton").pack(side="left", padx=2)
        ttk.Button(left_bar, text="Save CSV", command=self._save_csv_action, style="Ghost.TButton").pack(side="left", padx=2)

        tk.Frame(left_bar, bg=BORDER, width=1).pack(side="left", fill="y", padx=10, pady=4)
        ttk.Button(left_bar, text="Pack to .lang", command=self._pack_lang, style="Primary.TButton").pack(side="left", padx=2)

        right_bar = tk.Frame(topbar, bg=BG, padx=16, pady=10)
        right_bar.pack(side="right")

        tk.Label(right_bar, text="Target language", bg=BG, fg=TEXT2, font=FONTS).pack(side="left", padx=(0, 6))
        ttk.Entry(right_bar, textvariable=self.tgt_lang, width=14).pack(side="left", padx=(0, 16))
        self._prog_bar = ttk.Progressbar(right_bar, orient="horizontal", mode="determinate", length=110)
        self._prog_bar.pack(side="left", padx=(0, 6))
        tk.Label(right_bar, textvariable=self._progress, bg=BG, fg=TEXT2, font=FONTS).pack(side="left")

        tk.Frame(self, bg=BORDER, height=1).pack(fill="x")

        searchbar = tk.Frame(self, bg=BG2, padx=16, pady=8)
        searchbar.pack(fill="x")

        tk.Label(searchbar, text="Search", bg=BG2, fg=TEXT2, font=FONTS).pack(side="left", padx=(0, 6))
        self._search_entry = ttk.Entry(searchbar, textvariable=self._search_q, width=28)
        self._search_entry.pack(side="left", padx=(0, 14))

        tk.Label(searchbar, text="Queue", bg=BG2, fg=TEXT2, font=FONTS).pack(side="left", padx=(0, 6))
        self._filter_combo = ttk.Combobox(
            searchbar,
            state="readonly",
            width=14,
            values=[label for _, label in FILTER_OPTIONS],
        )
        self._filter_combo.set(FILTER_LABELS[self._filter_mode.get()])
        self._filter_combo.pack(side="left", padx=(0, 14))
        self._filter_combo.bind("<<ComboboxSelected>>", self._on_filter_combo_changed)

        ttk.Button(searchbar, text="Previous", command=lambda: self._navigate_visible(-1), style="Ghost.TButton").pack(side="left", padx=2)
        ttk.Button(searchbar, text="Next in group", command=lambda: self._navigate_visible(1), style="Ghost.TButton").pack(side="left", padx=2)
        ttk.Button(searchbar, text="Next untranslated", command=self._goto_next_untranslated, style="Ghost.TButton").pack(side="left", padx=2)
        ttk.Button(searchbar, text="Next warning", command=self._goto_next_warning, style="Ghost.TButton").pack(side="left", padx=2)

        self._file_label = tk.Label(searchbar, text="No file", bg=BG2, fg=TEXT2, font=FONTS)
        self._file_label.pack(side="right")

        tk.Frame(self, bg=BORDER, height=1).pack(fill="x")

        content = tk.Frame(self, bg=BG)
        content.pack(fill="both", expand=True)

        sidebar = tk.Frame(content, bg=BG2, width=250)
        sidebar.pack(side="left", fill="y")
        sidebar.pack_propagate(False)

        tk.Label(sidebar, text="GROUPS", bg=BG2, fg=TEXT2, font=("Segoe UI", 8), pady=10, padx=12).pack(anchor="w")

        self.grp_tv = ttk.Treeview(sidebar, show="tree", selectmode="browse", style="Groups.Treeview")
        grp_sb = ttk.Scrollbar(sidebar, orient="vertical", command=self.grp_tv.yview)
        self.grp_tv.configure(yscrollcommand=grp_sb.set)
        grp_sb.pack(side="right", fill="y")
        self.grp_tv.pack(fill="both", expand=True)
        self.grp_tv.bind("<<TreeviewSelect>>", self._on_group_select)

        tk.Frame(content, bg=BORDER, width=1).pack(side="left", fill="y")

        right_side = tk.Frame(content, bg=BG)
        right_side.pack(side="left", fill="both", expand=True)

        table_frame = tk.Frame(right_side, bg=BG)
        table_frame.pack(fill="both", expand=True)

        self.table = ttk.Treeview(
            table_frame,
            columns=("source", "translation", "status"),
            show="headings",
            selectmode="browse",
            style="Strings.Treeview",
        )
        self.table.heading("source", text="Source")
        self.table.heading("translation", text="Translation")
        self.table.heading("status", text="State")
        self.table.column("source", width=420, minwidth=200, stretch=True)
        self.table.column("translation", width=420, minwidth=200, stretch=True)
        self.table.column("status", width=70, minwidth=70, stretch=False)

        self.table.tag_configure("untranslated", foreground=TEXT2)
        self.table.tag_configure("draft", foreground=ACCENT)
        self.table.tag_configure("reviewed", foreground=DONE)
        self.table.tag_configure("warning", foreground=DANGER)

        t_scroll_y = ttk.Scrollbar(table_frame, orient="vertical", command=self.table.yview)
        t_scroll_x = ttk.Scrollbar(table_frame, orient="horizontal", command=self.table.xview)
        self.table.configure(yscrollcommand=t_scroll_y.set, xscrollcommand=t_scroll_x.set)
        t_scroll_y.pack(side="right", fill="y")
        t_scroll_x.pack(side="bottom", fill="x")
        self.table.pack(fill="both", expand=True)
        self.table.bind("<<TreeviewSelect>>", self._on_row_select)
        self.table.bind("<Return>", lambda _e: self._focus_edit())
        self.table.bind("<Double-1>", lambda _e: self._focus_edit())

        tk.Frame(right_side, bg=BORDER, height=1).pack(fill="x")
        self._edit_panel = tk.Frame(right_side, bg=BG, padx=16, pady=12)
        self._edit_panel.pack(fill="x")
        self._build_edit_panel()

        tk.Frame(self, bg=BORDER, height=1).pack(fill="x")
        sbar = tk.Frame(self, bg=BG, padx=16, pady=5)
        sbar.pack(fill="x")
        tk.Label(sbar, textvariable=self._status, bg=BG, fg=TEXT2, font=FONTS, anchor="w").pack(side="left")
        self._ai_status = tk.Label(sbar, text="", bg=BG, fg=ACCENT, font=FONTS)
        self._ai_status.pack(side="right")

    def _build_edit_panel(self):
        key_row = tk.Frame(self._edit_panel, bg=BG)
        key_row.pack(fill="x", pady=(0, 8))
        self._key_lbl = tk.Label(key_row, text="Select a string to edit", bg=BG, fg=TEXT2, font=FONTM)
        self._key_lbl.pack(side="left")
        tk.Label(key_row, textvariable=self._state_var, bg=BG, fg=TEXT2, font=FONTS).pack(side="right")

        cols = tk.Frame(self._edit_panel, bg=BG)
        cols.pack(fill="x")
        cols.columnconfigure(0, weight=1)
        cols.columnconfigure(2, weight=1)

        src_frame = tk.Frame(cols, bg=BG)
        src_frame.grid(row=0, column=0, sticky="nsew", padx=(0, 8))
        tk.Label(src_frame, text="Source", bg=BG, fg=TEXT2, font=FONTS).pack(anchor="w")
        self._src_text = tk.Text(
            src_frame,
            height=5,
            font=FONT,
            bg=BG2,
            fg=TEXT,
            relief="flat",
            wrap="word",
            state="disabled",
            padx=8,
            pady=6,
            cursor="arrow",
        )
        self._src_text.pack(fill="x")

        mid_frame = tk.Frame(cols, bg=BG)
        mid_frame.grid(row=0, column=1, sticky="ns", padx=6)
        tk.Label(mid_frame, text="", bg=BG).pack(expand=True, fill="y")
        ttk.Button(
            mid_frame,
            text="Copy source ->",
            command=self._copy_source_to_translation,
            style="Ghost.TButton",
        ).pack()
        tk.Label(mid_frame, text="", bg=BG).pack(expand=True, fill="y")

        tgt_frame = tk.Frame(cols, bg=BG)
        tgt_frame.grid(row=0, column=2, sticky="nsew")
        tk.Label(tgt_frame, text="Translation", bg=BG, fg=TEXT2, font=FONTS).pack(anchor="w")
        self._tgt_text = tk.Text(
            tgt_frame,
            height=5,
            font=FONT,
            bg="#FFFFFF",
            fg=TEXT,
            relief="flat",
            wrap="word",
            padx=8,
            pady=6,
            insertbackground=ACCENT,
            highlightthickness=1,
            highlightbackground=BORDER,
            highlightcolor=ACCENT,
        )
        self._tgt_text.pack(fill="x")
        self._tgt_text.bind("<<Modified>>", self._on_tgt_modified)
        self._tgt_text.bind("<Control-Return>", lambda _e: self._save_edit())
        self._tgt_text.bind("<Control-Shift-Return>", lambda _e: self._save_edit(mark_reviewed=True))
        self._tgt_text.bind("<Escape>", lambda _e: self.table.focus_set())

        info_row = tk.Frame(self._edit_panel, bg=BG)
        info_row.pack(fill="x", pady=(8, 0))
        tk.Label(info_row, textvariable=self._stats_var, bg=BG, fg=TEXT2, font=FONTS).pack(side="left")
        tk.Label(info_row, textvariable=self._warning_var, bg=BG, fg=DANGER, font=FONTS).pack(side="right")

        action_row = tk.Frame(self._edit_panel, bg=BG)
        action_row.pack(fill="x", pady=(8, 0))
        ttk.Button(action_row, text="Save", command=self._save_edit, style="Ghost.TButton").pack(side="left", padx=(0, 6))
        ttk.Button(action_row, text="Mark reviewed", command=lambda: self._save_edit(mark_reviewed=True), style="Ghost.TButton").pack(side="left", padx=(0, 6))
        ttk.Button(action_row, text="Clear translation", command=self._clear_translation_editor, style="Ghost.TButton").pack(side="left", padx=(0, 6))
        self._placeholder_btn = ttk.Menubutton(
            action_row,
            textvariable=self._placeholder_label_var,
            style="Ghost.TButton",
            direction="below",
        )
        self._placeholder_menu = tk.Menu(self._placeholder_btn, tearoff=0, bg=BG, fg=TEXT)
        self._placeholder_btn["menu"] = self._placeholder_menu
        self._placeholder_btn.pack(side="left", padx=(0, 6))
        ttk.Button(action_row, text="AI feedback", command=self._ai_feedback_action, style="Ghost.TButton").pack(side="right")
        ttk.Button(action_row, text="AI suggest", command=self._ai_single, style="Ghost.TButton").pack(side="right", padx=(0, 6))

        hint_frame = tk.Frame(self._edit_panel, bg=BG)
        hint_frame.pack(fill="x", pady=(10, 0))
        hint_frame.columnconfigure(1, weight=1)

        tk.Label(hint_frame, text="Existing match", bg=BG, fg=TEXT2, font=FONTS).grid(row=0, column=0, sticky="nw", padx=(0, 8))
        self._source_hint_label = tk.Label(
            hint_frame,
            textvariable=self._source_hint_var,
            bg=BG,
            fg=TEXT,
            font=FONTS,
            justify="left",
            wraplength=530,
        )
        self._source_hint_label.grid(row=0, column=1, sticky="ew")
        self._apply_source_hint_btn = ttk.Button(hint_frame, text="Apply match", command=self._apply_source_hint, style="Ghost.TButton")
        self._apply_source_hint_btn.grid(row=0, column=2, sticky="e", padx=(8, 0))

        tk.Label(hint_frame, text="Similar key", bg=BG, fg=TEXT2, font=FONTS).grid(row=1, column=0, sticky="nw", padx=(0, 8), pady=(6, 0))
        self._similar_hint_label = tk.Label(
            hint_frame,
            textvariable=self._similar_hint_var,
            bg=BG,
            fg=TEXT,
            font=FONTS,
            justify="left",
            wraplength=530,
        )
        self._similar_hint_label.grid(row=1, column=1, sticky="ew", pady=(6, 0))
        self._apply_similar_hint_btn = ttk.Button(hint_frame, text="Apply similar", command=self._apply_similar_hint, style="Ghost.TButton")
        self._apply_similar_hint_btn.grid(row=1, column=2, sticky="e", padx=(8, 0), pady=(6, 0))

        ai_frame = tk.Frame(self._edit_panel, bg=BG)
        ai_frame.pack(fill="x", pady=(12, 0))
        top = tk.Frame(ai_frame, bg=BG)
        top.pack(fill="x")
        tk.Label(top, textvariable=self._ai_panel_label_var, bg=BG, fg=TEXT2, font=FONTS).pack(side="left")
        self._ai_accept_btn = ttk.Button(top, text="Accept", command=self._accept_ai_suggestion, style="Ghost.TButton")
        self._ai_accept_btn.pack(side="right", padx=(6, 0))
        self._ai_insert_btn = ttk.Button(top, text="Insert", command=self._insert_ai_suggestion, style="Ghost.TButton")
        self._ai_insert_btn.pack(side="right", padx=(6, 0))
        self._ai_replace_btn = ttk.Button(top, text="Replace", command=self._replace_with_ai_suggestion, style="Ghost.TButton")
        self._ai_replace_btn.pack(side="right")

        suggestion_wrap = tk.Frame(ai_frame, bg=BG2)
        suggestion_wrap.pack(fill="x", pady=(6, 0))
        ai_scroll_y = ttk.Scrollbar(suggestion_wrap, orient="vertical")
        ai_scroll_y.pack(side="right", fill="y")

        self._ai_suggestion_box = tk.Text(
            suggestion_wrap,
            height=6,
            font=FONT,
            bg=BG2,
            fg=TEXT,
            relief="flat",
            wrap="word",
            state="disabled",
            padx=8,
            pady=6,
            yscrollcommand=ai_scroll_y.set,
        )
        self._ai_suggestion_box.pack(side="left", fill="x", expand=True)
        ai_scroll_y.config(command=self._ai_suggestion_box.yview)

        self._source_hint_value = ""
        self._similar_hint_value = ""
        self._current_placeholders: list[str] = []
        self._set_ai_buttons_enabled(False)
        self._set_hint_buttons_enabled(False, False)
        self._refresh_placeholder_menu([])

    def _bind_shortcuts(self):
        self.bind_all("<Control-i>", lambda _e: self._import_lang())
        self.bind_all("<Control-o>", lambda _e: self._open_csv())
        self.bind_all("<Control-s>", lambda _e: self._save_csv_action())
        self.bind_all("<Control-p>", lambda _e: self._pack_lang())
        self.bind_all("<Control-f>", lambda _e: self._focus_search())
        self.bind_all("<Control-e>", lambda _e: self._focus_edit())
        self.bind_all("<Control-Return>", lambda _e: self._save_edit())
        self.bind_all("<Control-Shift-Return>", lambda _e: self._save_edit(mark_reviewed=True))
        self.bind_all("<F6>", lambda _e: self._focus_search())
        self.bind_all("<Control-Down>", lambda _e: self._navigate_visible(1))
        self.bind_all("<Control-Up>", lambda _e: self._navigate_visible(-1))
        self.bind_all("<Alt-Down>", lambda _e: self._navigate_visible(1))
        self.bind_all("<Alt-Up>", lambda _e: self._navigate_visible(-1))
        self.bind_all("<Control-g>", lambda _e: self._goto_next_untranslated())
        self.bind_all("<Control-G>", lambda _e: self._goto_next_untranslated())
        self.bind_all("<Control-Shift-W>", lambda _e: self._goto_next_warning())
        self.bind_all("<Alt-n>", lambda _e: self._goto_next_untranslated())
        self.bind_all("<Alt-N>", lambda _e: self._goto_next_untranslated())
        self.bind_all("<Alt-w>", lambda _e: self._goto_next_warning())
        self.bind_all("<Alt-W>", lambda _e: self._goto_next_warning())
        self.bind_all("<Control-1>", lambda _e: self._set_filter_mode("all"))
        self.bind_all("<Control-2>", lambda _e: self._set_filter_mode("untranslated"))
        self.bind_all("<Control-3>", lambda _e: self._set_filter_mode("draft"))
        self.bind_all("<Control-4>", lambda _e: self._set_filter_mode("reviewed"))
        self.bind_all("<Control-5>", lambda _e: self._set_filter_mode("warnings"))

    def _build_menu(self):
        menu = tk.Menu(self, tearoff=0, bg=BG, fg=TEXT, relief="flat")
        self.config(menu=menu)
        file_menu = tk.Menu(menu, tearoff=0, bg=BG, fg=TEXT)
        menu.add_cascade(label="File", menu=file_menu)
        file_menu.add_command(label="Import .lang...", command=self._import_lang, accelerator="Ctrl+I")
        file_menu.add_command(label="Open CSV...", command=self._open_csv, accelerator="Ctrl+O")
        file_menu.add_command(label="Save CSV", command=self._save_csv_action, accelerator="Ctrl+S")
        file_menu.add_separator()
        file_menu.add_command(label="Pack to .lang...", command=self._pack_lang, accelerator="Ctrl+P")
        file_menu.add_separator()
        file_menu.add_command(label="Settings...", command=self._open_settings)

        nav_menu = tk.Menu(menu, tearoff=0, bg=BG, fg=TEXT)
        menu.add_cascade(label="Navigate", menu=nav_menu)
        nav_menu.add_command(label="Focus Search", command=self._focus_search, accelerator="Ctrl+F / F6")
        nav_menu.add_command(label="Focus Translation", command=self._focus_edit, accelerator="Ctrl+E")
        nav_menu.add_separator()
        nav_menu.add_command(label="Previous Row", command=lambda: self._navigate_visible(-1), accelerator="Ctrl+Up")
        nav_menu.add_command(label="Next Row", command=lambda: self._navigate_visible(1), accelerator="Ctrl+Down")
        nav_menu.add_command(label="Next Untranslated", command=self._goto_next_untranslated, accelerator="Ctrl+G")
        nav_menu.add_command(label="Next Warning", command=self._goto_next_warning, accelerator="Ctrl+Shift+W")
        nav_menu.add_separator()
        nav_menu.add_command(label="Queue: All", command=lambda: self._set_filter_mode("all"), accelerator="Ctrl+1")
        nav_menu.add_command(label="Queue: Untranslated", command=lambda: self._set_filter_mode("untranslated"), accelerator="Ctrl+2")
        nav_menu.add_command(label="Queue: Draft", command=lambda: self._set_filter_mode("draft"), accelerator="Ctrl+3")
        nav_menu.add_command(label="Queue: Reviewed", command=lambda: self._set_filter_mode("reviewed"), accelerator="Ctrl+4")
        nav_menu.add_command(label="Queue: Warnings", command=lambda: self._set_filter_mode("warnings"), accelerator="Ctrl+5")

    def _group_for_key(self, key: str) -> str:
        parts = key.split("/")
        return parts[0] if len(parts) == 1 else "/".join(parts[:2])

    def _build_indexes(self):
        self._record_by_key = {record.key: record for record in self.records}
        source_to_keys: defaultdict[str, list[str]] = defaultdict(list)
        group_to_keys: defaultdict[str, list[str]] = defaultdict(list)
        for record in self.records:
            source_to_keys[record.value].append(record.key)
            group_to_keys[self._group_for_key(record.key)].append(record.key)
        self._source_to_keys = dict(source_to_keys)
        self._group_to_keys = dict(group_to_keys)

    def _keys_match_records(self, left: list[LangRecord], right: list[LangRecord]) -> bool:
        return len(left) == len(right) and all(a.key == b.key for a, b in zip(left, right))

    def _build_recovered_translations(
        self,
        source_records: list[LangRecord],
        packed_records: list[LangRecord],
    ) -> dict[str, str]:
        translations: dict[str, str] = {}
        for source_record, packed_record in zip(source_records, packed_records):
            if packed_record.value.strip() and packed_record.value != source_record.value:
                translations[source_record.key] = packed_record.value
        return translations

    def _load_pack_recovery(self, lang_path: Path, packed_records: list[LangRecord]) -> Optional[tuple[list[LangRecord], dict[str, str], str]]:
        meta_path = pack_meta_path_for(lang_path)
        if not meta_path.exists():
            return None
        try:
            payload = json.loads(meta_path.read_text(encoding="utf-8"))
            source_rows = payload.get("records", [])
            translation_rows = payload.get("translations", {})
            source_records = [
                LangRecord(key=row["key"], value=row["source"])
                for row in source_rows
                if isinstance(row, dict) and "key" in row and "source" in row
            ]
            if not self._keys_match_records(source_records, packed_records):
                return None
            translations = {
                key: value
                for key, value in translation_rows.items()
                if isinstance(value, str) and value.strip()
            }
            return source_records, translations, f"Recovered source context from {meta_path.name}"
        except Exception:
            return None

    def _load_csv_recovery(self, lang_path: Path, packed_records: list[LangRecord]) -> Optional[tuple[list[LangRecord], dict[str, str], str]]:
        for csv_path in sorted(lang_path.parent.glob("*.csv")):
            try:
                source_records, _translations = load_csv(csv_path)
            except Exception:
                continue
            if not self._keys_match_records(source_records, packed_records):
                continue
            translations = self._build_recovered_translations(source_records, packed_records)
            return source_records, translations, f"Recovered source context from {csv_path.name}"
        return None

    def _recover_lang_import(self, lang_path: Path, packed_records: list[LangRecord]) -> tuple[list[LangRecord], dict[str, str], Optional[str]]:
        recovered = self._load_pack_recovery(lang_path, packed_records)
        if recovered:
            return recovered
        recovered = self._load_csv_recovery(lang_path, packed_records)
        if recovered:
            return recovered
        return packed_records, {}, None

    def _load_sidecar(self):
        meta = load_meta(self.csv_path)
        review_state = meta.get("review_state", {})
        ui = meta.get("ui", {})
        valid_keys = {record.key for record in self.records}

        clean_state: dict[str, str] = {}
        if isinstance(review_state, dict):
            for key, state in review_state.items():
                if key in valid_keys and state in {"draft", "reviewed"} and self.translations.get(key, "").strip():
                    clean_state[key] = state

        for key, text in self.translations.items():
            if text.strip():
                clean_state.setdefault(key, "draft")

        self.review_state = clean_state
        self._loaded_ui = ui if isinstance(ui, dict) else {}

    def _save_sidecar(self):
        if not self.csv_path:
            return
        payload = {
            "review_state": {
                key: state
                for key, state in self.review_state.items()
                if state in {"draft", "reviewed"} and self.translations.get(key, "").strip()
            },
            "ui": {
                "current_key": self._cur_key,
                "group": self._cur_grp,
                "filter": self._filter_mode.get(),
                "search": self._search_q.get(),
                "source_lang_name": self._source_lang_name,
                "last_pack_path": str(self._last_pack_path) if self._last_pack_path else None,
            },
        }
        try:
            save_meta(self.csv_path, payload)
        except Exception as exc:
            self._status.set(f"Metadata save error: {exc}")

    def _entry_state(self, key: str) -> str:
        text = self.translations.get(key, "")
        if not text.strip():
            return "untranslated"
        return "reviewed" if self.review_state.get(key) == "reviewed" else "draft"

    def _row_status_text(self, key: str) -> str:
        if self._row_has_warning(key):
            return "Warn"
        state = self._entry_state(key)
        if state == "reviewed":
            return "Rev"
        if state == "draft":
            return "Draft"
        return "Todo"

    def _row_tag(self, key: str) -> str:
        return "warning" if self._row_has_warning(key) else self._entry_state(key)

    def _matching_base_keys(self) -> list[str]:
        query = self._search_q.get().lower().strip()
        keys: list[str] = []
        for record in self.records:
            key = record.key
            if self._cur_grp is not None and self._group_for_key(key) != self._cur_grp:
                continue
            translation = self.translations.get(key, "")
            if query and query not in key.lower() and query not in record.value.lower() and query not in translation.lower():
                continue
            keys.append(key)
        return keys

    def _filter_match(self, key: str) -> bool:
        mode = self._filter_mode.get()
        state = self._entry_state(key)
        if mode == "all":
            return True
        if mode == "untranslated":
            return state == "untranslated"
        if mode == "draft":
            return state == "draft"
        if mode == "reviewed":
            return state == "reviewed"
        if mode == "warnings":
            return self._row_has_warning(key)
        return True

    def _translated_count(self, keys: list[str]) -> int:
        return sum(1 for key in keys if self.translations.get(key, "").strip())

    def _build_groups(self):
        current = self._cur_grp
        self.grp_tv.delete(*self.grp_tv.get_children())
        self._grp_ids = {}
        self._group_iids = {}

        all_keys = [record.key for record in self.records]
        all_id = self.grp_tv.insert("", "end", text=f"All  ({self._translated_count(all_keys)}/{len(all_keys)})")
        self._grp_ids[all_id] = None
        self._group_iids[None] = all_id

        for group in sorted(self._group_to_keys):
            keys = self._group_to_keys[group]
            iid = self.grp_tv.insert("", "end", text=f"{group}  ({self._translated_count(keys)}/{len(keys)})")
            self._grp_ids[iid] = group
            self._group_iids[group] = iid

        self._select_group(current, refresh=False)

    def _load_records(self, records: list[LangRecord], translations: dict[str, str], path_label: str):
        self.records = records
        self.translations = translations
        self.changed = False
        self._cur_key = None
        self._cur_grp = None
        self._ai_suggestion_text = ""

        self._build_indexes()
        self._load_sidecar()
        self._build_groups()

        loaded_ui = getattr(self, "_loaded_ui", {})
        group = loaded_ui.get("group") if isinstance(loaded_ui, dict) else None
        filter_mode = loaded_ui.get("filter") if isinstance(loaded_ui, dict) else "all"
        search = loaded_ui.get("search") if isinstance(loaded_ui, dict) else ""
        current_key = loaded_ui.get("current_key") if isinstance(loaded_ui, dict) else None
        source_lang_name = loaded_ui.get("source_lang_name") if isinstance(loaded_ui, dict) else self._source_lang_name
        last_pack_path = loaded_ui.get("last_pack_path") if isinstance(loaded_ui, dict) else (
            str(self._last_pack_path) if self._last_pack_path else None
        )

        self._suspend_refresh = True
        self._filter_mode.set(filter_mode if filter_mode in FILTER_LABELS else "all")
        self._filter_combo.set(FILTER_LABELS[self._filter_mode.get()])
        self._search_q.set(search if isinstance(search, str) else "")
        self._suspend_refresh = False

        self._source_lang_name = source_lang_name if isinstance(source_lang_name, str) and source_lang_name.strip() else None
        self._last_pack_path = Path(last_pack_path) if isinstance(last_pack_path, str) and last_pack_path.strip() else None

        self._select_group(group if group in self._group_to_keys else None, refresh=False)
        self._refresh_table(select_candidates=[current_key] if current_key else None)
        self._update_progress()
        self._file_label.config(text=path_label)
        self._status.set(f"Loaded {len(records):,} strings  |  {len(translations):,} translated")
        self._save_sidecar()

    def _select_group(self, group: Optional[str], refresh: bool = True):
        self._cur_grp = group
        iid = self._group_iids.get(group) or self._group_iids.get(None)
        if iid:
            self.grp_tv.selection_set(iid)
            self.grp_tv.see(iid)
        if refresh:
            self._refresh_table(select_candidates=[self._cur_key] if self._cur_key else None)

    def _on_group_select(self, _event=None):
        selection = self.grp_tv.selection()
        if not selection:
            return
        self._cur_grp = self._grp_ids.get(selection[0])
        self._refresh_table(select_candidates=[self._cur_key] if self._cur_key else None)

    def _on_filter_combo_changed(self, _event=None):
        self._filter_mode.set(FILTER_LABEL_TO_KEY.get(self._filter_combo.get(), "all"))

    def _on_filter_inputs_changed(self):
        if self._suspend_refresh:
            return
        self._refresh_table(select_candidates=[self._cur_key] if self._cur_key else None)

    def _set_filter_mode(self, mode: str):
        if mode not in FILTER_LABELS:
            return
        self._filter_combo.set(FILTER_LABELS[mode])
        self._filter_mode.set(mode)

    def _refresh_table(self, select_candidates: Optional[list[Optional[str]]] = None, focus_editor: bool = False):
        self._sync_current_editor_if_needed()
        visible_keys = [key for key in self._matching_base_keys() if self._filter_match(key)]
        self.table.delete(*self.table.get_children())
        self._displayed_keys = visible_keys

        for key in visible_keys:
            record = self._record_by_key[key]
            self.table.insert(
                "",
                "end",
                iid=key,
                values=(record.value, self.translations.get(key, ""), self._row_status_text(key)),
                tags=(self._row_tag(key),),
            )

        target = None
        if select_candidates:
            for candidate in select_candidates:
                if candidate and candidate in visible_keys:
                    target = candidate
                    break
        if not target and self._cur_key in visible_keys:
            target = self._cur_key
        if not target and visible_keys:
            target = visible_keys[0]

        if target:
            self._select_table_key(target, focus_editor=focus_editor)
        else:
            self._clear_editor()

    def _select_table_key(self, key: str, focus_editor: bool = False):
        if key not in self.table.get_children():
            return
        self.table.selection_set(key)
        self.table.focus(key)
        self.table.see(key)
        self._on_row_select()
        if focus_editor:
            self._focus_edit()

    def _clear_editor(self):
        self._cur_key = None
        self._key_lbl.config(text="Select a string to edit")
        self._state_var.set("State: untranslated")
        self._set_source_text("")
        self._refresh_placeholder_menu([])
        self._set_editor_text("")
        self._warning_var.set("No format warnings.")
        self._stats_var.set("Length: 0 chars")
        self._source_hint_var.set("No matching translations yet.")
        self._similar_hint_var.set("No similar key suggestions yet.")
        self._source_hint_value = ""
        self._similar_hint_value = ""
        self._set_hint_buttons_enabled(False, False)
        self._set_ai_suggestion("")

    def _on_row_select(self, _event=None):
        selection = self.table.selection()
        if not selection:
            return
        key = selection[0]
        previous_key = self._cur_key
        if previous_key and previous_key != key:
            self._sync_current_editor_if_needed()
        record = self._record_by_key.get(key)
        if not record:
            return

        self._cur_key = key
        self._key_lbl.config(text=key)
        self._state_var.set(f"State: {self._entry_state(key)}")
        self._set_source_text(record.value)
        self._refresh_placeholder_menu(self._extract_placeholders(record.value))
        self._set_editor_text(self.translations.get(key, ""))
        self._refresh_editor_insights()
        self._refresh_translation_hints()
        self._set_ai_suggestion("")
        self._save_sidecar()

    def _set_source_text(self, text: str):
        self._src_text.config(state="normal")
        self._src_text.delete("1.0", "end")
        self._src_text.insert("1.0", text)
        self._src_text.config(state="disabled")

    def _set_editor_text(self, text: str):
        self._tgt_text.delete("1.0", "end")
        self._tgt_text.insert("1.0", text)
        self._tgt_text.edit_modified(False)

    def _current_editor_text(self) -> str:
        return self._tgt_text.get("1.0", "end-1c")

    def _focus_edit(self):
        self._tgt_text.focus_set()
        self._tgt_text.mark_set("insert", "end")

    def _focus_search(self):
        self._search_entry.focus_set()
        self._search_entry.selection_range(0, "end")

    def _on_tgt_modified(self, _event=None):
        if not self._tgt_text.edit_modified():
            return
        self._tgt_text.edit_modified(False)
        self._refresh_editor_insights()
        self._refresh_translation_hints()

    def _refresh_editor_insights(self):
        if not self._cur_key:
            self._warning_var.set("No format warnings.")
            self._stats_var.set("Length: 0 chars")
            return
        source = self._record_by_key[self._cur_key].value
        translation = self._current_editor_text()
        warnings = self._collect_warnings(source, translation, self._cur_key, include_empty=True)
        self._warning_var.set(" | ".join(warnings) if warnings else "No format warnings.")
        delta = len(translation) - len(source)
        sign = "+" if delta >= 0 else ""
        self._stats_var.set(
            f"Length: {len(translation)} chars | Source: {len(source)} | Delta: {sign}{delta}"
        )

    def _collect_warnings(self, source: str, translation: str, key: str, include_empty: bool) -> list[str]:
        warnings: list[str] = []
        translated = translation.strip()
        if include_empty and source.strip() and not translated:
            warnings.append("Translation is empty.")
            return warnings
        if not translated:
            return warnings

        if curly_placeholder_counter(source) != curly_placeholder_counter(translation):
            warnings.append("Curly placeholders do not match.")
        if Counter(CARET_PLACEHOLDER_RE.findall(source)) != Counter(CARET_PLACEHOLDER_RE.findall(translation)):
            warnings.append("Caret placeholders do not match.")
        if source.count("%s") != translation.count("%s"):
            warnings.append("%s token count changed.")
        if source.count("\\n") != translation.count("\\n") or source.count("\n") != translation.count("\n"):
            warnings.append("Line break tokens do not match.")

        src_lead = bool(source[:1].isspace())
        tgt_lead = bool(translation[:1].isspace())
        src_trail = bool(source[-1:].isspace())
        tgt_trail = bool(translation[-1:].isspace())
        if src_lead != tgt_lead or src_trail != tgt_trail:
            warnings.append("Leading or trailing whitespace differs.")

        peer_translations = {
            self.translations.get(other_key, "").strip()
            for other_key in self._source_to_keys.get(source, [])
            if other_key != key and self.translations.get(other_key, "").strip()
        }
        if peer_translations and translation.strip() not in peer_translations:
            warnings.append("This source text is translated differently elsewhere.")

        return warnings

    def _row_has_warning(self, key: str) -> bool:
        record = self._record_by_key.get(key)
        if not record:
            return False
        translation = self.translations.get(key, "")
        return bool(self._collect_warnings(record.value, translation, key, include_empty=False))

    def _refresh_translation_hints(self):
        if not self._cur_key:
            self._source_hint_var.set("No matching translations yet.")
            self._similar_hint_var.set("No similar key suggestions yet.")
            self._source_hint_value = ""
            self._similar_hint_value = ""
            self._set_hint_buttons_enabled(False, False)
            return

        current_text = self._current_editor_text().strip()
        same_source = self._best_same_source_match(self._cur_key, current_text)
        if same_source:
            self._source_hint_value = same_source["translation"]
            self._source_hint_var.set(f"{same_source['translation']}  [{same_source['key']}]")
        else:
            self._source_hint_value = ""
            self._source_hint_var.set("No matching translations yet.")

        similar = self._best_similar_key_match(self._cur_key, current_text)
        if similar:
            self._similar_hint_value = similar["translation"]
            self._similar_hint_var.set(f"{similar['translation']}  [{similar['key']}]")
        else:
            self._similar_hint_value = ""
            self._similar_hint_var.set("No similar key suggestions yet.")

        self._set_hint_buttons_enabled(bool(self._source_hint_value), bool(self._similar_hint_value))

    def _best_same_source_match(self, key: str, current_text: str) -> Optional[dict]:
        source = self._record_by_key[key].value
        candidates = []
        for other_key in self._source_to_keys.get(source, []):
            if other_key == key:
                continue
            translation = self.translations.get(other_key, "").strip()
            if not translation or translation == current_text:
                continue
            candidates.append(
                (
                    0 if self.review_state.get(other_key) == "reviewed" else 1,
                    other_key,
                    translation,
                )
            )
        if not candidates:
            return None
        candidates.sort()
        _, other_key, translation = candidates[0]
        return {"key": other_key, "translation": translation}

    def _best_similar_key_match(self, key: str, current_text: str) -> Optional[dict]:
        group = self._group_for_key(key)
        base_tokens = {token for token in SIMILAR_SPLIT_RE.split(key.lower()) if token}
        best_score: Optional[tuple[int, int, int]] = None
        best_match: Optional[dict] = None
        for other_key in self._group_to_keys.get(group, []):
            if other_key == key:
                continue
            translation = self.translations.get(other_key, "").strip()
            if not translation or translation == current_text:
                continue
            other_tokens = {token for token in SIMILAR_SPLIT_RE.split(other_key.lower()) if token}
            overlap = len(base_tokens & other_tokens)
            if overlap <= 0:
                continue
            score = (overlap, 1 if self.review_state.get(other_key) == "reviewed" else 0, -len(other_key))
            if best_score is None or score > best_score:
                best_score = score
                best_match = {"key": other_key, "translation": translation}
        return best_match

    def _set_hint_buttons_enabled(self, has_source: bool, has_similar: bool):
        self._apply_source_hint_btn.state(["!disabled"] if has_source else ["disabled"])
        self._apply_similar_hint_btn.state(["!disabled"] if has_similar else ["disabled"])

    def _apply_source_hint(self):
        if not self._source_hint_value:
            return
        self._set_editor_text(self._source_hint_value)
        self._refresh_editor_insights()
        self._focus_edit()

    def _apply_similar_hint(self):
        if not self._similar_hint_value:
            return
        self._set_editor_text(self._similar_hint_value)
        self._refresh_editor_insights()
        self._focus_edit()

    def _copy_source_to_translation(self):
        if not self._cur_key:
            return
        self._set_editor_text(self._record_by_key[self._cur_key].value)
        self._refresh_editor_insights()
        self._focus_edit()

    def _extract_placeholders(self, text: str) -> list[str]:
        placeholders: list[str] = []
        seen: set[str] = set()

        for token in PLACEHOLDER_RE.findall(text):
            insert_token = placeholder_insert_token(token)
            if insert_token not in seen:
                placeholders.append(insert_token)
                seen.add(insert_token)

        for token in CARET_PLACEHOLDER_RE.findall(text):
            if token not in seen:
                placeholders.append(token)
                seen.add(token)

        if "%s" in text and "%s" not in seen:
            placeholders.append("%s")
            seen.add("%s")

        if "\\n" in text and "\\n" not in seen:
            placeholders.append("\\n")
            seen.add("\\n")

        return placeholders

    def _available_placeholders(self) -> list[str]:
        placeholders: list[str] = []
        seen: set[str] = set()

        for token in DEFAULT_PLACEHOLDER_TEMPLATES:
            if token not in seen:
                placeholders.append(token)
                seen.add(token)

        for record in self.records:
            for token in self._extract_placeholders(record.value):
                if token not in seen:
                    placeholders.append(token)
                    seen.add(token)

        for text in self.translations.values():
            for token in self._extract_placeholders(text):
                if token not in seen:
                    placeholders.append(token)
                    seen.add(token)

        return placeholders

    def _refresh_placeholder_menu(self, placeholders: list[str]):
        self._current_placeholders = placeholders
        available_placeholders = self._available_placeholders()
        extra_placeholders = [token for token in available_placeholders if token not in placeholders]
        self._placeholder_menu.delete(0, "end")

        if not placeholders and not available_placeholders:
            self._placeholder_label_var.set("No placeholders")
            self._placeholder_menu.add_command(label="No placeholders available", state="disabled")
            self._placeholder_btn.state(["disabled"])
            return

        self._placeholder_label_var.set("Insert placeholder")
        if placeholders:
            self._placeholder_menu.add_command(label="Source placeholders", state="disabled")
        if len(placeholders) > 1:
            self._placeholder_menu.add_command(
                label="Insert all source placeholders",
                command=self._insert_all_placeholders,
            )
            self._placeholder_menu.add_separator()
        for token in placeholders:
            self._placeholder_menu.add_command(
                label=token,
                command=lambda value=token: self._insert_placeholder(value),
            )
        if extra_placeholders:
            if placeholders:
                self._placeholder_menu.add_separator()
            self._placeholder_menu.add_command(label="All available placeholders", state="disabled")
            for token in extra_placeholders:
                self._placeholder_menu.add_command(
                    label=token,
                    command=lambda value=token: self._insert_placeholder(value),
                )
        self._placeholder_btn.state(["!disabled"])

    def _insert_placeholder(self, token: str):
        self._tgt_text.insert("insert", token)
        self._tgt_text.edit_modified(False)
        self._refresh_editor_insights()
        self._focus_edit()

    def _insert_all_placeholders(self):
        if not self._current_placeholders:
            return
        text = " ".join(self._current_placeholders)
        self._tgt_text.insert("insert", text)
        self._tgt_text.edit_modified(False)
        self._refresh_editor_insights()
        self._focus_edit()

    def _clear_translation_editor(self):
        self._set_editor_text("")
        self._refresh_editor_insights()
        self._focus_edit()

    def _sync_current_editor_if_needed(self):
        if not self._cur_key:
            return
        editor_text = self._current_editor_text()
        if editor_text != self.translations.get(self._cur_key, ""):
            self._apply_translation_to_model(self._cur_key, editor_text, mark_reviewed=False)

    def _apply_translation_to_model(self, key: str, text: str, mark_reviewed: bool):
        previous_text = self.translations.get(key, "")
        previous_state = self.review_state.get(key)

        if text.strip():
            self.translations[key] = text
            self.review_state[key] = "reviewed" if mark_reviewed else "draft"
        else:
            self.translations.pop(key, None)
            self.review_state.pop(key, None)

        if previous_text != self.translations.get(key, "") or previous_state != self.review_state.get(key):
            self.changed = True

    def _save_edit(self, mark_reviewed: bool = False):
        if not self._cur_key:
            return
        old_visible = list(self._displayed_keys)
        current = self._cur_key
        index = old_visible.index(current) if current in old_visible else -1

        self._apply_translation_to_model(current, self._current_editor_text(), mark_reviewed=mark_reviewed)
        self._build_groups()
        self._update_progress()

        candidates: list[Optional[str]] = []
        if index >= 0:
            candidates.extend(old_visible[index + 1 :])
            candidates.extend(old_visible[:index])
            candidates.append(current)
        else:
            candidates.append(current)

        self._refresh_table(select_candidates=candidates, focus_editor=True)
        self._status.set(f"Saved {current} as {'reviewed' if mark_reviewed and self.translations.get(current, '').strip() else self._entry_state(current)}.")
        self._save_sidecar()

    def _navigate_visible(self, delta: int):
        if not self._displayed_keys:
            return
        if self._cur_key not in self._displayed_keys:
            self._select_table_key(self._displayed_keys[0], focus_editor=True)
            return
        index = self._displayed_keys.index(self._cur_key)
        next_index = (index + delta) % len(self._displayed_keys)
        self._select_table_key(self._displayed_keys[next_index], focus_editor=True)

    def _goto_next_matching(self, predicate):
        keys = self._matching_base_keys()
        if not keys:
            return
        start = keys.index(self._cur_key) if self._cur_key in keys else -1
        ordered = keys[start + 1 :] + keys[: start + 1] if start >= 0 else keys
        for key in ordered:
            if predicate(key):
                if key not in self._displayed_keys:
                    self._filter_mode.set("all")
                    self._filter_combo.set(FILTER_LABELS["all"])
                    self._refresh_table(select_candidates=[key], focus_editor=True)
                else:
                    self._select_table_key(key, focus_editor=True)
                return
        messagebox.showinfo("Queue", "No matching strings found in the current search/group scope.")

    def _goto_next_untranslated(self):
        self._goto_next_matching(lambda key: self._entry_state(key) == "untranslated")

    def _goto_next_warning(self):
        self._goto_next_matching(self._row_has_warning)

    def _update_progress(self):
        total = len(self.records)
        done = sum(1 for value in self.translations.values() if value.strip())
        if not total:
            self._progress.set("")
            self._prog_bar["value"] = 0
            return
        pct = int(done / total * 100)
        self._progress.set(f"{done:,} / {total:,}  ({pct}%)")
        self._prog_bar["value"] = pct

    def _import_lang(self):
        path = filedialog.askopenfilename(
            title="Import .lang file (source language)",
            filetypes=[("Lang files", "*.lang *.lngp"), ("All files", "*.*")],
        )
        if not path:
            return
        self._status.set(f"Importing {Path(path).name}...")
        self.update_idletasks()
        try:
            packed_records = load_lang(Path(path))
        except Exception as exc:
            messagebox.showerror("Import error", str(exc))
            return
        records, translations, recovery_note = self._recover_lang_import(Path(path), packed_records)

        default = Path(path).with_suffix(".csv")
        csv_out = filedialog.asksaveasfilename(
            title="Save working CSV",
            initialfile=default.name,
            initialdir=str(default.parent),
            defaultextension=".csv",
            filetypes=[("CSV files", "*.csv"), ("All files", "*.*")],
        )
        if not csv_out:
            return

        save_csv(Path(csv_out), records, translations)
        self.csv_path = Path(csv_out)
        self.review_state = {}
        self._source_lang_name = Path(path).name
        self._last_pack_path = None
        self._load_records(records, translations, Path(csv_out).name)
        self._remember_current_csv()
        status = f"Imported {len(records):,} strings from {Path(path).name}  |  Working CSV saved as {Path(csv_out).name}"
        if recovery_note:
            status += f"  |  {recovery_note}"
        self._status.set(status)

    def _open_csv(self):
        path = filedialog.askopenfilename(
            title="Open working CSV",
            filetypes=[("CSV files", "*.csv"), ("All files", "*.*")],
        )
        if not path:
            return
        self._open_csv_path(Path(path))

    def _save_csv_action(self):
        if not self.records:
            messagebox.showinfo("Nothing to save", "Import a .lang file first.")
            return
        if self.csv_path:
            self._do_save_csv(self.csv_path)
        else:
            path = filedialog.asksaveasfilename(
                title="Save CSV",
                defaultextension=".csv",
                filetypes=[("CSV files", "*.csv"), ("All files", "*.*")],
            )
            if path:
                self.csv_path = Path(path)
                self._do_save_csv(self.csv_path)

    def _do_save_csv(self, path: Path):
        self._sync_current_editor_if_needed()
        try:
            save_csv(path, self.records, self.translations)
            self.csv_path = path
            self._remember_current_csv()
            self.changed = False
            self._update_progress()
            self._build_groups()
            self._refresh_table(select_candidates=[self._cur_key] if self._cur_key else None)
            self._save_sidecar()
            self._status.set(f"Saved {path.name}  |  {len(self.translations):,} translations written")
        except Exception as exc:
            messagebox.showerror("Save error", str(exc))

    def _pack_lang(self):
        if not self.records:
            messagebox.showinfo("Nothing to pack", "Open a CSV or import a .lang file first.")
            return
        self._sync_current_editor_if_needed()
        if self.changed and self.csv_path:
            self._do_save_csv(self.csv_path)

        lang = self.tgt_lang.get().strip()
        if not lang:
            messagebox.showwarning("No language", "Enter a target language name.")
            return

        if self._last_pack_path:
            default_name = self._last_pack_path.name
            init_dir = str(self._last_pack_path.parent)
        elif self._source_lang_name:
            default_name = self._source_lang_name
            init_dir = str(self.csv_path.parent) if self.csv_path else "."
        else:
            default_name = f"{lang.lower()}.lang"
            init_dir = str(self.csv_path.parent) if self.csv_path else "."
        out = filedialog.asksaveasfilename(
            title="Pack to .lang binary",
            initialfile=default_name,
            initialdir=init_dir,
            defaultextension=".lang",
            filetypes=[("Lang files", "*.lang"), ("All files", "*.*")],
        )
        if not out:
            return

        try:
            data = build_lang(self.records, self.translations)
            Path(out).write_bytes(data)
            self._last_pack_path = Path(out)
            recovery_warning = self._write_pack_recovery(Path(out))
            self._save_sidecar()
            warnings = sum(1 for key in self._record_by_key if self._row_has_warning(key))
            done = sum(1 for value in self.translations.values() if value.strip())
            total = len(self.records)
            self._status.set(
                f"Packed {Path(out).name}  |  {done:,}/{total:,} strings  |  {warnings:,} warnings"
            )
            messagebox.showinfo(
                "Packed",
                (
                    f"Saved {Path(out).name}\n\n"
                    f"Warnings flagged in app: {warnings:,}\n\n"
                    "Next steps:\n"
                    "1. Copy the file into the game's Localizations folder\n"
                    "2. Open Storyteller, select your mod, and click Pack\n"
                    "3. Launch the game and select your language"
                ),
            )
            if recovery_warning:
                self._status.set(f"{self._status.get()}  |  {recovery_warning}")
        except Exception as exc:
            messagebox.showerror("Pack error", str(exc))

    def _write_pack_recovery(self, lang_path: Path):
        payload = {
            "records": [{"key": record.key, "source": record.value} for record in self.records],
            "translations": {
                key: value
                for key, value in self.translations.items()
                if isinstance(value, str) and value.strip()
            },
        }
        try:
            pack_meta_path_for(lang_path).write_text(json.dumps(payload, indent=2), encoding="utf-8")
            return None
        except Exception as exc:
            return f"Could not write recovery file: {exc}"

    def _set_ai_suggestion(self, text: str, kind: str = "suggestion"):
        self._ai_suggestion_text = text
        self._ai_output_kind = kind
        self._ai_panel_label_var.set("AI feedback" if kind == "feedback" else "AI suggestion")
        self._ai_suggestion_box.config(state="normal")
        self._ai_suggestion_box.delete("1.0", "end")
        if text:
            self._ai_suggestion_box.insert("1.0", text)
        self._ai_suggestion_box.config(state="disabled")
        self._set_ai_buttons_enabled(bool(text) and kind == "suggestion")

    def _set_ai_buttons_enabled(self, enabled: bool):
        state = ["!disabled"] if enabled else ["disabled"]
        self._ai_accept_btn.state(state)
        self._ai_replace_btn.state(state)
        self._ai_insert_btn.state(state)

    def _ai_single(self):
        if self.ai_busy:
            return
        ai_settings = self._get_ai_settings()
        if not ai_settings:
            return
        provider, model, key = ai_settings
        if not self._cur_key:
            return
        record = self._record_by_key.get(self._cur_key)
        if not record or not record.value.strip():
            return

        self.ai_busy = True
        provider_label = AI_PROVIDER_SPECS[provider]["label"]
        self._ai_status.config(text=f"{provider_label} translating...")
        src = self.csv_path.stem if self.csv_path else "English"
        tgt = self.tgt_lang.get().strip() or "target language"
        text = record.value

        def work():
            try:
                result = ai_translate(text, src, tgt, record.key, provider, model, key)
                self.after(0, lambda: self._on_ai_single_ready(result))
            except Exception as exc:
                self.after(0, lambda: self._status.set(f"AI error: {exc}"))
            finally:
                self.after(
                    0,
                    lambda: (
                        self._ai_status.config(text=""),
                        setattr(self, "ai_busy", False),
                    ),
                )

        threading.Thread(target=work, daemon=True).start()

    def _on_ai_single_ready(self, result: str):
        self._set_ai_suggestion(result, kind="suggestion")
        self._status.set("AI suggestion ready for review.")

    def _ai_feedback_action(self):
        if self.ai_busy:
            return
        ai_settings = self._get_ai_settings()
        if not ai_settings:
            return
        provider, model, key = ai_settings
        if not self._cur_key:
            return
        record = self._record_by_key.get(self._cur_key)
        if not record or not record.value.strip():
            return

        translation = self._current_editor_text().strip()
        if not translation:
            messagebox.showinfo("No translation", "Enter or draft a translation first, then ask for AI feedback.")
            return

        self.ai_busy = True
        provider_label = AI_PROVIDER_SPECS[provider]["label"]
        self._ai_status.config(text=f"{provider_label} reviewing...")
        src = self.csv_path.stem if self.csv_path else "English"
        tgt = self.tgt_lang.get().strip() or "target language"

        def work():
            try:
                result = ai_feedback(record.value, translation, src, tgt, record.key, provider, model, key)
                self.after(0, lambda: self._on_ai_feedback_ready(result))
            except Exception as exc:
                self.after(0, lambda: self._status.set(f"AI error: {exc}"))
            finally:
                self.after(
                    0,
                    lambda: (
                        self._ai_status.config(text=""),
                        setattr(self, "ai_busy", False),
                    ),
                )

        threading.Thread(target=work, daemon=True).start()

    def _on_ai_feedback_ready(self, result: str):
        self._set_ai_suggestion(result, kind="feedback")
        self._status.set("AI feedback ready for review.")

    def _replace_with_ai_suggestion(self):
        if not self._ai_suggestion_text:
            return
        self._set_editor_text(self._ai_suggestion_text)
        self._refresh_editor_insights()
        self._focus_edit()

    def _insert_ai_suggestion(self):
        if not self._ai_suggestion_text:
            return
        self._tgt_text.insert("insert", self._ai_suggestion_text)
        self._tgt_text.edit_modified(False)
        self._refresh_editor_insights()
        self._focus_edit()

    def _accept_ai_suggestion(self):
        if not self._ai_suggestion_text:
            return
        self._set_editor_text(self._ai_suggestion_text)
        self._refresh_editor_insights()
        self._save_edit(mark_reviewed=False)

    def _ai_batch(self):
        ai_settings = self._get_ai_settings()
        if not ai_settings:
            return
        provider, model, key = ai_settings
        untranslated = [
            record
            for record in self.records
            if record.value.strip() and not self.translations.get(record.key, "").strip()
        ]
        if not untranslated:
            messagebox.showinfo("Done", "All strings already translated.")
            return
        if not messagebox.askyesno(
            "Batch AI",
            (
                f"Auto-translate {len(untranslated):,} untranslated strings?\n"
                f"This will use {AI_PROVIDER_SPECS[provider]['label']} credits with model '{model}' "
                "and place the results in the Draft queue."
            ),
        ):
            return

        self.ai_busy = True
        provider_label = AI_PROVIDER_SPECS[provider]["label"]
        src = self.csv_path.stem if self.csv_path else "English"
        tgt = self.tgt_lang.get().strip() or "target language"

        def work():
            done = 0
            for record in untranslated:
                try:
                    result = ai_translate(record.value, src, tgt, record.key, provider, model, key)
                    self.translations[record.key] = result
                    self.review_state[record.key] = "draft"
                    done += 1
                    self.after(
                        0,
                        lambda value=done: self._ai_status.config(
                            text=f"{provider_label} {value:,}/{len(untranslated):,}"
                        ),
                    )
                except Exception as exc:
                    self.after(0, lambda err=exc: self._status.set(f"AI error: {err}"))
                    break
            self.after(0, lambda: self._ai_done(done, len(untranslated)))

        threading.Thread(target=work, daemon=True).start()

    def _ai_done(self, done: int, total: int):
        self.ai_busy = False
        if done:
            self.changed = True
            self._build_groups()
            self._update_progress()
            self._refresh_table(select_candidates=[self._cur_key] if self._cur_key else None)
        self._ai_status.config(text="")
        self._status.set(f"AI batch complete: {done:,}/{total:,} drafts created.")

    def _open_settings(self):
        window = tk.Toplevel(self)
        window.title("Settings")
        window.geometry("520x350")
        window.configure(bg=BG)
        window.resizable(False, False)
        window.grab_set()

        tk.Label(window, text="Settings", bg=BG, fg=TEXT, font=("Segoe UI", 13, "bold")).pack(
            pady=(20, 16), padx=24, anchor="w"
        )

        form = tk.Frame(window, bg=BG)
        form.pack(padx=24, fill="x")
        tk.Label(
            form,
            text="Each provider keeps its own saved API key and model.",
            bg=BG,
            fg=TEXT2,
            font=FONTS,
        ).pack(anchor="w", pady=(0, 10))
        tk.Label(form, text="Active AI provider", bg=BG, fg=TEXT2, font=FONTS).pack(anchor="w")
        provider_var = tk.StringVar(value=AI_PROVIDER_LABELS[normalize_provider(self.ai_provider.get())])
        provider_combo = ttk.Combobox(
            form,
            state="readonly",
            textvariable=provider_var,
            values=[label for label in AI_PROVIDER_LABELS.values()],
            width=24,
        )
        provider_combo.pack(anchor="w", pady=(0, 12))
        provider_combo.bind("<<ComboboxSelected>>", lambda *_: self.ai_provider.set(AI_PROVIDER_LABEL_TO_KEY[provider_var.get()]))

        tk.Label(form, text="Model id", bg=BG, fg=TEXT2, font=FONTS).pack(anchor="w")
        ttk.Entry(form, textvariable=self.ai_model, width=54).pack(fill="x")
        tk.Label(
            form,
            textvariable=self._ai_model_hint_var,
            bg=BG,
            fg=TEXT2,
            font=("Segoe UI", 8),
        ).pack(anchor="w", pady=(0, 10))

        tk.Label(form, textvariable=self._ai_key_label_var, bg=BG, fg=TEXT2, font=FONTS).pack(anchor="w")
        tk.Label(
            form,
            textvariable=self._ai_key_hint_var,
            bg=BG,
            fg=TEXT2,
            font=("Segoe UI", 8),
        ).pack(anchor="w", pady=(0, 4))

        entry = ttk.Entry(form, textvariable=self.active_api_key, width=54, show="*")
        entry.pack(fill="x")
        show_var = tk.BooleanVar()
        tk.Checkbutton(
            form,
            text="Show key",
            variable=show_var,
            command=lambda: entry.config(show="" if show_var.get() else "*"),
            bg=BG,
            fg=TEXT2,
            font=FONTS,
            activebackground=BG,
        ).pack(anchor="w", pady=4)

        buttons = tk.Frame(window, bg=BG)
        buttons.pack(pady=12, padx=24, anchor="w")

        def on_save():
            self._save_settings_cfg()
            window.destroy()

        ttk.Button(buttons, text="Save", command=on_save, style="Primary.TButton").pack(side="left", padx=(0, 8))
        ttk.Button(
            buttons,
            text="Batch translate all untranslated",
            command=lambda: [on_save(), self._ai_batch()],
            style="Ghost.TButton",
        ).pack(side="left")

    def _on_close(self):
        self._sync_current_editor_if_needed()
        if self.changed:
            reply = messagebox.askyesnocancel("Unsaved changes", "Save CSV before quitting?")
            if reply is None:
                return
            if reply and self.csv_path:
                self._do_save_csv(self.csv_path)
        self._save_settings_cfg()
        self._save_sidecar()
        self.destroy()


if __name__ == "__main__":
    app = App()
    app.mainloop()
