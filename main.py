"""
Jarvis — AI-powered Voice Assistant for Windows.

All voice input goes to a local Ollama model, which decides whether to
answer conversationally or execute a system action via tools. A background
monitor thread also pushes proactive observations (battery, disk, time of
day) into the same conversation flow, in JARVIS's voice.
"""

from __future__ import annotations

import argparse
import ctypes
import datetime as dt
import io
import json
import logging
import os
import platform
import queue
import re
import shutil
import subprocess
import sys
import threading
import time
import urllib.parse
import webbrowser
from ctypes import POINTER, cast
from difflib import SequenceMatcher
from pathlib import Path

if sys.stdout.encoding and sys.stdout.encoding.lower() != "utf-8":
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")
    sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding="utf-8", errors="replace")

import ollama
import pyttsx3
import speech_recognition as sr
from comtypes import CLSCTX_ALL
from dotenv import load_dotenv
from pycaw.pycaw import AudioUtilities, IAudioEndpointVolume

from memory import Memory
from monitor import Alert, Monitor
from personality import build_startup_briefing, build_system_prompt
from weather import get_weather

try:
    import psutil
except ImportError:
    psutil = None  # type: ignore

load_dotenv(Path(__file__).resolve().parent / ".env", override=True)

STARTUP_SHORTCUT_NAME = "JarvisAssistant.lnk"
NOTES_FILE = Path(__file__).resolve().parent / "data" / "notes.txt"


def _resolve_launch_command() -> list[str]:
    """Build a Windows-friendly command line for autostart registration."""

    script_path = Path(__file__).resolve()
    runner = Path(sys.executable)
    return [str(runner), str(script_path), "--autostart"]


def _startup_shortcut_path() -> Path:
    """Return the user's Startup folder shortcut path for Jarvis."""

    return (
        Path(os.environ["APPDATA"])
        / "Microsoft"
        / "Windows"
        / "Start Menu"
        / "Programs"
        / "Startup"
        / STARTUP_SHORTCUT_NAME
    )


def _ps_single_quote(value: str) -> str:
    """Escape a Python string for use inside a PowerShell single-quoted string."""

    return value.replace("'", "''")


def install_startup_task() -> None:
    """Create a Startup folder shortcut so Jarvis launches at user logon."""

    shortcut_path = _startup_shortcut_path()
    shortcut_path.parent.mkdir(parents=True, exist_ok=True)
    script_path = Path(__file__).resolve()
    launch_cmd = _resolve_launch_command()
    target_path = launch_cmd[0]

    powershell_script = (
        "$ws = New-Object -ComObject WScript.Shell; "
        f"$s = $ws.CreateShortcut('{_ps_single_quote(str(shortcut_path))}'); "
        f"$s.TargetPath = '{_ps_single_quote(target_path)}'; "
        f"$s.Arguments = '\"{_ps_single_quote(str(script_path))}\" --autostart'; "
        f"$s.WorkingDirectory = '{_ps_single_quote(str(script_path.parent))}'; "
        "$s.WindowStyle = 7; "
        "$s.Save()"
    )

    result = subprocess.run(
        ["powershell", "-NoProfile", "-Command", powershell_script],
        capture_output=True,
        text=True,
        check=False,
    )

    if result.returncode != 0:
        message = (result.stderr or result.stdout or "Unknown error").strip()
        raise RuntimeError(f"Could not create startup shortcut: {message}")


def remove_startup_task() -> None:
    """Remove the Startup folder shortcut created for Jarvis."""

    shortcut_path = _startup_shortcut_path()
    if shortcut_path.exists():
        shortcut_path.unlink()


def startup_task_exists() -> bool:
    """Check whether the Startup folder shortcut already exists."""

    return _startup_shortcut_path().exists()

# ---------------------------------------------------------------------------
# Tools definition for the local model (Ollama / OpenAI-style schema)
# ---------------------------------------------------------------------------

TOOLS = [
    {
        "type": "function",
        "function": {
            "name": "get_current_time",
            "description": "Get the current time and date",
            "parameters": {"type": "object", "properties": {}, "required": []},
        },
    },
    {
        "type": "function",
        "function": {
            "name": "open_application",
            "description": "Open/launch an application on Windows (e.g. Chrome, VS Code, Notepad, Calculator, Spotify, Discord, Word, Excel, Task Manager, Settings, etc.)",
            "parameters": {
                "type": "object",
                "properties": {"app_name": {"type": "string", "description": "Name of the application to open"}},
                "required": ["app_name"],
            },
        },
    },
    {
        "type": "function",
        "function": {
            "name": "close_application",
            "description": "Close/kill a running application",
            "parameters": {
                "type": "object",
                "properties": {"app_name": {"type": "string", "description": "Name of the application to close"}},
                "required": ["app_name"],
            },
        },
    },
    {
        "type": "function",
        "function": {
            "name": "google_search",
            "description": "Search Google for any query",
            "parameters": {
                "type": "object",
                "properties": {"query": {"type": "string", "description": "The search query"}},
                "required": ["query"],
            },
        },
    },
    {
        "type": "function",
        "function": {
            "name": "youtube_search",
            "description": "Search YouTube for videos",
            "parameters": {
                "type": "object",
                "properties": {"query": {"type": "string", "description": "The YouTube search query"}},
                "required": ["query"],
            },
        },
    },
    {
        "type": "function",
        "function": {
            "name": "open_website",
            "description": "Open a website URL or well-known site name in the browser",
            "parameters": {
                "type": "object",
                "properties": {"url": {"type": "string", "description": "URL or site name (e.g. youtube.com, github)"}},
                "required": ["url"],
            },
        },
    },
    {
        "type": "function",
        "function": {
            "name": "set_volume",
            "description": "Set system volume to a specific percentage, or adjust it up/down",
            "parameters": {
                "type": "object",
                "properties": {
                    "action": {"type": "string", "enum": ["set", "up", "down", "mute", "unmute"], "description": "Volume action"},
                    "level": {"type": "integer", "description": "Volume level 0-100 (only for 'set' action)"},
                },
                "required": ["action"],
            },
        },
    },
    {
        "type": "function",
        "function": {
            "name": "set_brightness",
            "description": "Set screen brightness to a specific percentage, or adjust it up/down",
            "parameters": {
                "type": "object",
                "properties": {
                    "action": {"type": "string", "enum": ["set", "up", "down"], "description": "Brightness action"},
                    "level": {"type": "integer", "description": "Brightness level 0-100 (only for 'set' action)"},
                },
                "required": ["action"],
            },
        },
    },
    {
        "type": "function",
        "function": {
            "name": "system_control",
            "description": "Control the Windows system: lock, restart, shutdown, sleep, log off, empty recycle bin",
            "parameters": {
                "type": "object",
                "properties": {
                    "action": {"type": "string", "enum": ["lock", "restart", "shutdown", "sleep", "logoff", "empty_recycle_bin"], "description": "System action to perform"},
                },
                "required": ["action"],
            },
        },
    },
    {
        "type": "function",
        "function": {
            "name": "open_folder",
            "description": "Open a folder in File Explorer (desktop, downloads, documents, pictures, videos, music, projects, or any path)",
            "parameters": {
                "type": "object",
                "properties": {"folder": {"type": "string", "description": "Folder name or full path"}},
                "required": ["folder"],
            },
        },
    },
    {
        "type": "function",
        "function": {
            "name": "take_screenshot",
            "description": "Take a screenshot of the screen",
            "parameters": {"type": "object", "properties": {}, "required": []},
        },
    },
    {
        "type": "function",
        "function": {
            "name": "media_control",
            "description": "Control media playback: play/pause, next track, previous track",
            "parameters": {
                "type": "object",
                "properties": {
                    "action": {"type": "string", "enum": ["play_pause", "next", "previous"], "description": "Media action"},
                },
                "required": ["action"],
            },
        },
    },
    {
        "type": "function",
        "function": {
            "name": "run_shell_command",
            "description": "Run a Windows shell/PowerShell command. Use for things like: checking wifi status, IP address, installed programs, system info, disk space, etc.",
            "parameters": {
                "type": "object",
                "properties": {"command": {"type": "string", "description": "The shell command to execute"}},
                "required": ["command"],
            },
        },
    },
    {
        "type": "function",
        "function": {
            "name": "get_weather",
            "description": "Get the current weather for a city. If no city is given, uses Sir's default city.",
            "parameters": {
                "type": "object",
                "properties": {"city": {"type": "string", "description": "Optional city name"}},
                "required": [],
            },
        },
    },
    {
        "type": "function",
        "function": {
            "name": "get_system_status",
            "description": "Get a one-line snapshot of the local machine: battery, CPU, RAM, disk free, uptime.",
            "parameters": {"type": "object", "properties": {}, "required": []},
        },
    },
    {
        "type": "function",
        "function": {
            "name": "remember_fact",
            "description": "Persist a long-term fact about Sir (preferences, projects, names, routines). Call this silently in addition to your spoken reply whenever you learn something worth keeping across sessions.",
            "parameters": {
                "type": "object",
                "properties": {
                    "category": {"type": "string", "description": "Short category like 'preference', 'project', 'routine'"},
                    "text": {"type": "string", "description": "The fact itself, in one short sentence"},
                },
                "required": ["category", "text"],
            },
        },
    },
    {
        "type": "function",
        "function": {
            "name": "recall_facts",
            "description": "Look up what you already know about Sir from long-term memory. Optional category filter.",
            "parameters": {
                "type": "object",
                "properties": {"category": {"type": "string", "description": "Optional category filter"}},
                "required": [],
            },
        },
    },
    {
        "type": "function",
        "function": {
            "name": "type_text",
            "description": "Type text into whatever window currently has keyboard focus, as if Sir typed it himself. Useful for filling forms, sending messages, writing notes.",
            "parameters": {
                "type": "object",
                "properties": {"text": {"type": "string", "description": "The text to type"}},
                "required": ["text"],
            },
        },
    },
    {
        "type": "function",
        "function": {
            "name": "press_keys",
            "description": "Press a keyboard shortcut or single key (e.g. 'ctrl+s', 'alt+f4', 'win+d', 'enter', 'esc').",
            "parameters": {
                "type": "object",
                "properties": {"keys": {"type": "string", "description": "Key or chord, e.g. 'ctrl+shift+t'"}},
                "required": ["keys"],
            },
        },
    },
    {
        "type": "function",
        "function": {
            "name": "file_operation",
            "description": "Create, read, append, or delete a text file on disk. Use this for note-taking, logging, or quick edits.",
            "parameters": {
                "type": "object",
                "properties": {
                    "action": {"type": "string", "enum": ["create", "read", "append", "delete", "exists"], "description": "What to do"},
                    "path": {"type": "string", "description": "Absolute or relative file path"},
                    "content": {"type": "string", "description": "Text content (for create/append)"},
                },
                "required": ["action", "path"],
            },
        },
    },
    {
        "type": "function",
        "function": {
            "name": "list_directory",
            "description": "List the contents of a folder on disk.",
            "parameters": {
                "type": "object",
                "properties": {"path": {"type": "string", "description": "Folder path"}},
                "required": ["path"],
            },
        },
    },
    {
        "type": "function",
        "function": {
            "name": "list_processes",
            "description": "List the top running processes by CPU or memory usage.",
            "parameters": {
                "type": "object",
                "properties": {
                    "sort_by": {"type": "string", "enum": ["cpu", "memory"], "description": "Sort key"},
                    "limit": {"type": "integer", "description": "Max processes to return (default 10)"},
                },
                "required": [],
            },
        },
    },
    {
        "type": "function",
        "function": {
            "name": "kill_process",
            "description": "Forcefully kill a running process by name or PID. Confirm with Sir for anything critical.",
            "parameters": {
                "type": "object",
                "properties": {"target": {"type": "string", "description": "Process name (e.g. 'chrome.exe') or numeric PID"}},
                "required": ["target"],
            },
        },
    },
    {
        "type": "function",
        "function": {
            "name": "window_control",
            "description": "Control the foreground or all windows: minimize all, restore all, show desktop, switch window, close active window.",
            "parameters": {
                "type": "object",
                "properties": {
                    "action": {"type": "string", "enum": ["minimize_all", "restore_all", "show_desktop", "switch_window", "close_active"], "description": "Window action"},
                },
                "required": ["action"],
            },
        },
    },
    {
        "type": "function",
        "function": {
            "name": "network_info",
            "description": "Get network information: current IP, WiFi SSID, connection status.",
            "parameters": {"type": "object", "properties": {}, "required": []},
        },
    },
    {
        "type": "function",
        "function": {
            "name": "toggle_wifi",
            "description": "Turn the WiFi adapter on or off.",
            "parameters": {
                "type": "object",
                "properties": {"action": {"type": "string", "enum": ["on", "off"], "description": "WiFi state"}},
                "required": ["action"],
            },
        },
    },
    {
        "type": "function",
        "function": {
            "name": "clipboard",
            "description": "Read or write the system clipboard.",
            "parameters": {
                "type": "object",
                "properties": {
                    "action": {"type": "string", "enum": ["read", "write"], "description": "Clipboard action"},
                    "text": {"type": "string", "description": "Text to write (for write action)"},
                },
                "required": ["action"],
            },
        },
    },
]

OLLAMA_MODEL = os.getenv("OLLAMA_MODEL", "llama3.2:3b").strip() or "llama3.2:3b"

# Tools whose result string is already a complete spoken answer.
# For these we skip the second LLM round-trip and speak the result directly,
# saving 2-4 seconds per command.
TERMINAL_TOOLS = {
    "get_current_time", "get_weather", "get_system_status", "recall_facts",
    "open_application", "close_application", "google_search", "youtube_search",
    "open_website", "open_folder", "take_screenshot", "media_control",
    "set_volume", "set_brightness", "system_control", "remember_fact",
    "type_text", "press_keys", "list_directory", "kill_process",
    "window_control", "network_info", "toggle_wifi", "clipboard",
}

# ---------------------------------------------------------------------------
# App & folder registries
# ---------------------------------------------------------------------------

APP_MAP: dict[str, tuple[tuple[str, ...], tuple[str, ...]]] = {
    # name -> (launch_targets, process_names)
    "chrome": (("chrome.exe", r"%ProgramFiles%\Google\Chrome\Application\chrome.exe"), ("chrome.exe",)),
    "firefox": (("firefox.exe",), ("firefox.exe",)),
    "edge": (("msedge.exe",), ("msedge.exe",)),
    "file explorer": (("explorer.exe",), ("explorer.exe",)),
    "explorer": (("explorer.exe",), ("explorer.exe",)),
    "vs code": (("code.cmd", "Code.exe", r"%LocalAppData%\Programs\Microsoft VS Code\Code.exe"), ("Code.exe",)),
    "vscode": (("code.cmd", "Code.exe", r"%LocalAppData%\Programs\Microsoft VS Code\Code.exe"), ("Code.exe",)),
    "notepad": (("notepad.exe",), ("notepad.exe",)),
    "calculator": (("calc.exe",), ("CalculatorApp.exe",)),
    "terminal": (("wt.exe", "powershell.exe"), ("WindowsTerminal.exe",)),
    "powershell": (("powershell.exe",), ("powershell.exe",)),
    "cmd": (("cmd.exe",), ("cmd.exe",)),
    "task manager": (("taskmgr.exe",), ("Taskmgr.exe",)),
    "settings": (("ms-settings:",), ()),
    "spotify": (("spotify.exe", r"%AppData%\Spotify\Spotify.exe"), ("Spotify.exe",)),
    "discord": (("discord.exe",), ("Discord.exe",)),
    "telegram": (("telegram.exe",), ("Telegram.exe",)),
    "word": (("winword.exe",), ("WINWORD.EXE",)),
    "excel": (("excel.exe",), ("EXCEL.EXE",)),
    "powerpoint": (("powerpnt.exe",), ("POWERPNT.EXE",)),
    "paint": (("mspaint.exe",), ("mspaint.exe",)),
    "snipping tool": (("snippingtool.exe",), ("SnippingTool.exe",)),
}

FOLDER_MAP: dict[str, Path] = {
    "desktop": Path.home() / "Desktop",
    "downloads": Path.home() / "Downloads",
    "documents": Path.home() / "Documents",
    "pictures": Path.home() / "Pictures",
    "videos": Path.home() / "Videos",
    "music": Path.home() / "Music",
    "projects": Path.home() / "Projects",
    "home": Path.home(),
}

SITE_MAP: dict[str, str] = {
    "youtube": "https://www.youtube.com", "google": "https://www.google.com",
    "github": "https://github.com", "facebook": "https://www.facebook.com",
    "twitter": "https://twitter.com", "reddit": "https://www.reddit.com",
    "whatsapp": "https://web.whatsapp.com", "gmail": "https://mail.google.com",
    "linkedin": "https://www.linkedin.com", "netflix": "https://www.netflix.com",
    "chatgpt": "https://chat.openai.com", "claude": "https://claude.ai",
    "amazon": "https://www.amazon.com", "stackoverflow": "https://stackoverflow.com",
}

# ---------------------------------------------------------------------------
# Tool execution
# ---------------------------------------------------------------------------

def exec_tool(name: str, inp: dict, vol_endpoint, memory: Memory | None = None) -> str:
    """Execute a tool call and return the result string."""

    if name == "get_current_time":
        now = dt.datetime.now()
        return now.strftime("It's %I:%M %p on %A, %B %d, %Y")

    if name == "open_application":
        app = inp["app_name"].lower().strip()
        entry = APP_MAP.get(app)
        if entry:
            _launch(entry[0])
            return f"Opened {app}"
        # try generic
        exe = shutil.which(app) or shutil.which(f"{app}.exe")
        if exe:
            subprocess.Popen([exe], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
            return f"Launched {app}"
        return f"Could not find application: {app}"

    if name == "close_application":
        app = inp["app_name"].lower().strip()
        entry = APP_MAP.get(app)
        procs = entry[1] if entry else (f"{app}.exe",)
        closed = False
        for p in procs:
            r = subprocess.run(["taskkill", "/IM", p, "/T", "/F"], capture_output=True, text=True, check=False)
            if r.returncode == 0:
                closed = True
        return f"Closed {app}" if closed else f"{app} is not running"

    if name == "google_search":
        q = inp["query"]
        webbrowser.open(f"https://www.google.com/search?q={urllib.parse.quote(q)}", new=2)
        return f"Searching Google for: {q}"

    if name == "youtube_search":
        q = inp["query"]
        webbrowser.open(f"https://www.youtube.com/results?search_query={urllib.parse.quote(q)}", new=2)
        return f"Searching YouTube for: {q}"

    if name == "open_website":
        url = inp["url"].strip().lower()
        if url in SITE_MAP:
            webbrowser.open(SITE_MAP[url], new=2)
            return f"Opened {url}"
        if not url.startswith(("http://", "https://")):
            url = f"https://{url}" if "." in url else f"https://www.{url}.com"
        webbrowser.open(url, new=2)
        return f"Opened {url}"

    if name == "set_volume":
        if vol_endpoint is None:
            return "Volume control is not available on this device"
        action = inp["action"]
        if action == "mute":
            vol_endpoint.SetMute(1, None)
            return "Muted"
        if action == "unmute":
            vol_endpoint.SetMute(0, None)
            return "Unmuted"
        if action == "set":
            lv = min(100, max(0, inp.get("level", 50)))
            vol_endpoint.SetMute(0, None)
            vol_endpoint.SetMasterVolumeLevelScalar(lv / 100.0, None)
            return f"Volume set to {lv}%"
        cur = vol_endpoint.GetMasterVolumeLevelScalar()
        delta = 0.10 if action == "up" else -0.10
        nv = min(1.0, max(0.0, cur + delta))
        vol_endpoint.SetMute(0, None)
        vol_endpoint.SetMasterVolumeLevelScalar(nv, None)
        return f"Volume {'raised' if action == 'up' else 'lowered'} to {int(nv * 100)}%"

    if name == "set_brightness":
        action = inp["action"]
        if action == "set":
            lv = min(100, max(0, inp.get("level", 50)))
            _brightness_set(lv)
            return f"Brightness set to {lv}%"
        delta = 10 if action == "up" else -10
        _brightness_change(delta)
        return f"Brightness {'increased' if action == 'up' else 'decreased'}"

    if name == "system_control":
        action = inp["action"]
        if action == "lock":
            ctypes.windll.user32.LockWorkStation()
            return "Computer locked"
        if action == "restart":
            subprocess.Popen(["shutdown", "/r", "/t", "3"])
            return "Restarting in 3 seconds"
        if action == "shutdown":
            subprocess.Popen(["shutdown", "/s", "/t", "3"])
            return "Shutting down in 3 seconds"
        if action == "sleep":
            ctypes.WinDLL("PowrProf.dll").SetSuspendState(False, False, False)
            return "Entering sleep mode"
        if action == "logoff":
            subprocess.Popen(["shutdown", "/l"])
            return "Logging off"
        if action == "empty_recycle_bin":
            ctypes.windll.shell32.SHEmptyRecycleBinW(None, None, 0x07)
            return "Recycle bin emptied"
        return f"Unknown action: {action}"

    if name == "open_folder":
        folder = inp["folder"].lower().strip()
        path = FOLDER_MAP.get(folder)
        if path and path.exists():
            os.startfile(str(path))
            return f"Opened {folder} folder"
        raw = Path(os.path.expandvars(inp["folder"])).expanduser()
        if raw.exists():
            os.startfile(str(raw))
            return f"Opened {raw}"
        return f"Folder not found: {folder}"

    if name == "take_screenshot":
        ts = dt.datetime.now().strftime("%Y%m%d_%H%M%S")
        path = Path.home() / "Pictures" / "Screenshots" / f"screenshot_{ts}.png"
        path.parent.mkdir(parents=True, exist_ok=True)
        subprocess.run(
            ["powershell", "-Command",
             f"Add-Type -AssemblyName System.Windows.Forms;"
             f"$b = [System.Drawing.Bitmap]::new([System.Windows.Forms.Screen]::PrimaryScreen.Bounds.Width,"
             f"[System.Windows.Forms.Screen]::PrimaryScreen.Bounds.Height);"
             f"$g = [System.Drawing.Graphics]::FromImage($b);"
             f"$g.CopyFromScreen(0, 0, 0, 0, $b.Size);"
             f"$b.Save('{path}')"],
            capture_output=True, check=False,
        )
        return f"Screenshot saved to {path}"

    if name == "media_control":
        keys = {"play_pause": 0xB3, "next": 0xB0, "previous": 0xB1}
        vk = keys.get(inp["action"])
        if vk:
            ctypes.windll.user32.keybd_event(vk, 0, 0, 0)
            ctypes.windll.user32.keybd_event(vk, 0, 2, 0)
            return f"Media: {inp['action']}"
        return "Unknown media action"

    if name == "run_shell_command":
        cmd = inp["command"]
        try:
            r = subprocess.run(cmd, shell=True, capture_output=True, text=True, timeout=15, check=False)
            output = (r.stdout + r.stderr).strip()
            return output[:2000] if output else "Command executed (no output)"
        except subprocess.TimeoutExpired:
            return "Command timed out after 15 seconds"
        except Exception as e:
            return f"Error: {e}"

    if name == "get_weather":
        return get_weather(inp.get("city"))

    if name == "get_system_status":
        if psutil is None:
            return "System status unavailable — psutil is not installed."
        try:
            parts = []
            cpu_name = (os.environ.get("PROCESSOR_IDENTIFIER") or platform.processor() or "").strip()
            if cpu_name:
                parts.append(cpu_name)
            physical = psutil.cpu_count(logical=False)
            logical = psutil.cpu_count(logical=True)
            if physical or logical:
                if physical and logical and physical != logical:
                    parts.append(f"{physical} cores, {logical} threads")
                elif logical:
                    parts.append(f"{logical} logical CPUs")
            b = psutil.sensors_battery()
            if b is not None:
                plug = "plugged in" if b.power_plugged else "unplugged"
                parts.append(f"battery {int(b.percent)}% ({plug})")
            cpu_pct = psutil.cpu_percent(interval=0.1)
            try:
                freq = psutil.cpu_freq()
            except Exception:
                freq = None
            if freq and freq.current:
                if freq.max and freq.max > 0:
                    parts.append(f"CPU {cpu_pct:.0f}% at {freq.current:.0f}/{freq.max:.0f} MHz")
                else:
                    parts.append(f"CPU {cpu_pct:.0f}% at {freq.current:.0f} MHz")
            else:
                parts.append(f"CPU {cpu_pct:.0f}%")
            parts.append(f"RAM {psutil.virtual_memory().percent:.0f}%")
            d = psutil.disk_usage("C:\\")
            parts.append(f"C: {round(d.free / (1024 ** 3), 1)} GB free")
            uptime_min = int((time.time() - psutil.boot_time()) / 60)
            hrs, mins = divmod(uptime_min, 60)
            parts.append(f"uptime {hrs}h {mins}m")
            return ", ".join(parts)
        except Exception as e:
            return f"Could not read system status: {e}"

    if name == "remember_fact":
        if memory is None:
            return "Memory unavailable."
        category = inp.get("category", "general")
        text = inp.get("text", "")
        if not text.strip():
            return "Nothing to remember."
        category_norm = str(category).strip().lower()
        text_norm = text.strip().lower()
        if category_norm in {"identity", "assistant", "system"} or any(
            phrase in text_norm
            for phrase in (
                "i am an artificial intelligence",
                "i am a language model",
                "designed to assist and communicate",
                "i am jarvis",
                "i am an ai",
            )
        ):
            return "I only store facts about Sir, not about myself."
        memory.add_fact(category, text)
        return "Noted."

    if name == "recall_facts":
        if memory is None:
            return "Memory unavailable."
        cat = inp.get("category")
        facts = memory.facts_by_category(cat) if cat else memory.get_facts()
        if not facts:
            return "I have no facts on record yet."
        return "; ".join(f"[{f['category']}] {f['text']}" for f in facts[-15:])

    if name == "type_text":
        text = inp.get("text", "")
        if not text:
            return "Nothing to type."
        try:
            _send_text(text)
            return f"Typed: {text[:80]}"
        except Exception as e:
            return f"Could not type: {e}"

    if name == "press_keys":
        keys = inp.get("keys", "").strip().lower()
        if not keys:
            return "No keys specified."
        try:
            _send_keys(keys)
            return f"Pressed {keys}"
        except Exception as e:
            return f"Could not press keys: {e}"

    if name == "file_operation":
        action = inp.get("action", "")
        path_str = inp.get("path", "")
        if not path_str:
            return "No path given."
        path = Path(os.path.expandvars(path_str)).expanduser()
        try:
            if action == "create":
                path.parent.mkdir(parents=True, exist_ok=True)
                path.write_text(inp.get("content", ""), encoding="utf-8")
                return f"Created {path}"
            if action == "read":
                if not path.exists():
                    return f"File not found: {path}"
                data = path.read_text(encoding="utf-8", errors="replace")
                return data[:1500] if data else "(empty file)"
            if action == "append":
                path.parent.mkdir(parents=True, exist_ok=True)
                with path.open("a", encoding="utf-8") as f:
                    f.write(inp.get("content", ""))
                return f"Appended to {path}"
            if action == "delete":
                if path.exists():
                    path.unlink()
                    return f"Deleted {path}"
                return f"File not found: {path}"
            if action == "exists":
                return f"{path} exists" if path.exists() else f"{path} does not exist"
            return f"Unknown file action: {action}"
        except Exception as e:
            return f"File error: {e}"

    if name == "list_directory":
        path_str = inp.get("path", "")
        path = Path(os.path.expandvars(path_str)).expanduser() if path_str else Path.cwd()
        try:
            if not path.exists():
                return f"Folder not found: {path}"
            entries = sorted(path.iterdir(), key=lambda p: (not p.is_dir(), p.name.lower()))[:40]
            parts = [f"{'[D]' if e.is_dir() else '   '} {e.name}" for e in entries]
            return f"{path}: " + ", ".join(parts) if parts else f"{path} is empty"
        except Exception as e:
            return f"Could not list folder: {e}"

    if name == "list_processes":
        if psutil is None:
            return "psutil is not installed."
        sort_by = inp.get("sort_by", "cpu")
        limit = int(inp.get("limit", 10))
        try:
            procs = []
            for p in psutil.process_iter(["pid", "name", "cpu_percent", "memory_percent"]):
                try:
                    procs.append(p.info)
                except Exception:
                    continue
            key = "cpu_percent" if sort_by == "cpu" else "memory_percent"
            procs.sort(key=lambda d: d.get(key) or 0, reverse=True)
            top = procs[:limit]
            return ", ".join(
                f"{d['name']}({d['pid']}) {d.get(key) or 0:.0f}%" for d in top
            )
        except Exception as e:
            return f"Could not list processes: {e}"

    if name == "kill_process":
        target = str(inp.get("target", "")).strip()
        if not target:
            return "No target specified."
        try:
            if target.isdigit():
                if psutil is None:
                    subprocess.run(["taskkill", "/PID", target, "/F"], capture_output=True, check=False)
                    return f"Killed PID {target}"
                psutil.Process(int(target)).kill()
                return f"Killed PID {target}"
            r = subprocess.run(["taskkill", "/IM", target, "/T", "/F"], capture_output=True, text=True, check=False)
            if r.returncode == 0:
                return f"Killed {target}"
            return f"{target} not running or could not be killed"
        except Exception as e:
            return f"Could not kill: {e}"

    if name == "window_control":
        action = inp.get("action", "")
        try:
            if action == "minimize_all":
                ctypes.windll.user32.keybd_event(0x5B, 0, 0, 0)  # win down
                ctypes.windll.user32.keybd_event(0x4D, 0, 0, 0)  # m down
                ctypes.windll.user32.keybd_event(0x4D, 0, 2, 0)
                ctypes.windll.user32.keybd_event(0x5B, 0, 2, 0)
                return "Minimized all windows"
            if action == "restore_all":
                ctypes.windll.user32.keybd_event(0x5B, 0, 0, 0)
                ctypes.windll.user32.keybd_event(0xA0, 0, 0, 0)  # shift
                ctypes.windll.user32.keybd_event(0x4D, 0, 0, 0)
                ctypes.windll.user32.keybd_event(0x4D, 0, 2, 0)
                ctypes.windll.user32.keybd_event(0xA0, 0, 2, 0)
                ctypes.windll.user32.keybd_event(0x5B, 0, 2, 0)
                return "Restored windows"
            if action == "show_desktop":
                ctypes.windll.user32.keybd_event(0x5B, 0, 0, 0)
                ctypes.windll.user32.keybd_event(0x44, 0, 0, 0)  # d
                ctypes.windll.user32.keybd_event(0x44, 0, 2, 0)
                ctypes.windll.user32.keybd_event(0x5B, 0, 2, 0)
                return "Showing desktop"
            if action == "switch_window":
                _send_keys("alt+tab")
                return "Switched window"
            if action == "close_active":
                _send_keys("alt+f4")
                return "Closed active window"
            return f"Unknown window action: {action}"
        except Exception as e:
            return f"Window control error: {e}"

    if name == "network_info":
        try:
            r = subprocess.run(
                ["powershell", "-Command",
                 "$ip=(Get-NetIPAddress -AddressFamily IPv4 | Where-Object {$_.PrefixOrigin -eq 'Dhcp'} | Select-Object -First 1 -ExpandProperty IPAddress);"
                 "$wifi=(netsh wlan show interfaces | Select-String 'SSID' | Where-Object {$_ -notmatch 'BSSID'} | Select-Object -First 1).ToString().Split(':')[1].Trim();"
                 "Write-Output \"IP $ip, WiFi $wifi\""],
                capture_output=True, text=True, timeout=8, check=False,
            )
            out = (r.stdout or "").strip()
            return out if out else "Could not retrieve network info."
        except Exception as e:
            return f"Network error: {e}"

    if name == "toggle_wifi":
        action = inp.get("action", "")
        cmd = "enabled" if action == "on" else "disabled"
        try:
            subprocess.run(
                ["netsh", "interface", "set", "interface", "Wi-Fi", cmd],
                capture_output=True, check=False,
            )
            return f"WiFi {action}"
        except Exception as e:
            return f"WiFi error: {e}"

    if name == "clipboard":
        action = inp.get("action", "")
        try:
            if action == "read":
                r = subprocess.run(
                    ["powershell", "-Command", "Get-Clipboard"],
                    capture_output=True, text=True, timeout=5, check=False,
                )
                txt = (r.stdout or "").strip()
                return txt[:1500] if txt else "(clipboard is empty)"
            if action == "write":
                text = inp.get("text", "")
                p = subprocess.Popen(
                    ["powershell", "-Command", "$input | Set-Clipboard"],
                    stdin=subprocess.PIPE, text=True,
                )
                p.communicate(text)
                return f"Copied to clipboard: {text[:80]}"
            return f"Unknown clipboard action: {action}"
        except Exception as e:
            return f"Clipboard error: {e}"

    return f"Unknown tool: {name}"


def _launch(targets: tuple[str, ...]) -> None:
    for t in targets:
        expanded = os.path.expandvars(t)
        if expanded.endswith(":"):
            os.startfile(expanded)
            return
        if Path(expanded).exists():
            subprocess.Popen([expanded], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
            return
        exe = shutil.which(expanded) or shutil.which(t)
        if exe:
            subprocess.Popen([exe], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
            return
        try:
            subprocess.Popen([t], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
            return
        except Exception:
            continue
    raise FileNotFoundError(f"Cannot launch any of: {targets}")


def _brightness_set(level: int) -> None:
    subprocess.run(["powershell", "-Command",
                     f"(Get-WmiObject -Namespace root/WMI -Class WmiMonitorBrightnessMethods).WmiSetBrightness(1,{level})"],
                    capture_output=True, check=False)


def _brightness_change(delta: int) -> None:
    subprocess.run(["powershell", "-Command",
                     f"$c=(Get-WmiObject -Namespace root/WMI -Class WmiMonitorBrightness).CurrentBrightness;"
                     f"$n=[Math]::Max(0,[Math]::Min(100,$c+{delta}));"
                     f"(Get-WmiObject -Namespace root/WMI -Class WmiMonitorBrightnessMethods).WmiSetBrightness(1,$n)"],
                    capture_output=True, check=False)


# --- Keyboard input helpers (for type_text and press_keys tools) ----------

_VK_MAP: dict[str, int] = {
    "ctrl": 0x11, "control": 0x11,
    "alt": 0x12,
    "shift": 0x10,
    "win": 0x5B, "windows": 0x5B, "super": 0x5B,
    "enter": 0x0D, "return": 0x0D,
    "esc": 0x1B, "escape": 0x1B,
    "tab": 0x09, "space": 0x20, "backspace": 0x08, "delete": 0x2E, "del": 0x2E,
    "up": 0x26, "down": 0x28, "left": 0x25, "right": 0x27,
    "home": 0x24, "end": 0x23, "pageup": 0x21, "pagedown": 0x22,
    "f1": 0x70, "f2": 0x71, "f3": 0x72, "f4": 0x73, "f5": 0x74, "f6": 0x75,
    "f7": 0x76, "f8": 0x77, "f9": 0x78, "f10": 0x79, "f11": 0x7A, "f12": 0x7B,
}

KEYEVENTF_KEYUP = 0x0002


def _vk_for(key: str) -> int | None:
    key = key.strip().lower()
    if key in _VK_MAP:
        return _VK_MAP[key]
    if len(key) == 1:
        c = key.upper()
        if c.isalnum():
            return ord(c)
    return None


def _send_keys(combo: str) -> None:
    """Press a key chord like 'ctrl+s', 'alt+f4', 'win+d'."""
    parts = [p.strip() for p in combo.replace(" ", "").split("+") if p.strip()]
    vks: list[int] = []
    for p in parts:
        vk = _vk_for(p)
        if vk is not None:
            vks.append(vk)
    if not vks:
        raise ValueError(f"Unknown key chord: {combo}")
    user32 = ctypes.windll.user32
    for vk in vks:
        user32.keybd_event(vk, 0, 0, 0)
    time.sleep(0.04)
    for vk in reversed(vks):
        user32.keybd_event(vk, 0, KEYEVENTF_KEYUP, 0)


def _send_text(text: str) -> None:
    """Type a Unicode string into the focused window using SendInput."""
    user32 = ctypes.windll.user32

    # SendInput structures
    PUL = ctypes.POINTER(ctypes.c_ulong)

    class KeyBdInput(ctypes.Structure):
        _fields_ = [("wVk", ctypes.c_ushort),
                    ("wScan", ctypes.c_ushort),
                    ("dwFlags", ctypes.c_ulong),
                    ("time", ctypes.c_ulong),
                    ("dwExtraInfo", PUL)]

    class HardwareInput(ctypes.Structure):
        _fields_ = [("uMsg", ctypes.c_ulong),
                    ("wParamL", ctypes.c_short),
                    ("wParamH", ctypes.c_ushort)]

    class MouseInput(ctypes.Structure):
        _fields_ = [("dx", ctypes.c_long), ("dy", ctypes.c_long),
                    ("mouseData", ctypes.c_ulong),
                    ("dwFlags", ctypes.c_ulong),
                    ("time", ctypes.c_ulong),
                    ("dwExtraInfo", PUL)]

    class _Input_I(ctypes.Union):
        _fields_ = [("ki", KeyBdInput),
                    ("mi", MouseInput),
                    ("hi", HardwareInput)]

    class Input(ctypes.Structure):
        _fields_ = [("type", ctypes.c_ulong), ("ii", _Input_I)]

    KEYEVENTF_UNICODE = 0x0004
    extra = ctypes.c_ulong(0)

    for ch in text:
        for flags in (KEYEVENTF_UNICODE, KEYEVENTF_UNICODE | KEYEVENTF_KEYUP):
            ki = KeyBdInput(0, ord(ch), flags, 0, ctypes.pointer(extra))
            inp = Input(1, _Input_I(ki=ki))
            user32.SendInput(1, ctypes.byref(inp), ctypes.sizeof(inp))
        time.sleep(0.005)


# ---------------------------------------------------------------------------
# Voice Assistant
# ---------------------------------------------------------------------------

class Jarvis:

    def __init__(self) -> None:
        self.logger = self._init_logger()
        self.rec = sr.Recognizer()
        self.rec.dynamic_energy_threshold = True
        # tighter pause threshold so we commit speech faster
        self.rec.pause_threshold = 0.4
        self.rec.non_speaking_duration = 0.3
        self.vol = self._init_volume()
        self.memory = Memory(Path(__file__).resolve().parent / "data" / "memory.json")
        self.history: list[dict] = self.memory.get_history()
        self.alert_queue: queue.Queue[Alert] = queue.Queue()
        self.monitor = Monitor(self.logger, self.alert_queue)
        self.running = True
        self.focus_mode = False
        self.tts_proc: subprocess.Popen | None = None
        self._start_tts_worker()
        self.logger.info(
            "MEMORY_LOAD | history=%d facts=%d",
            len(self.history),
            len(self.memory.get_facts()),
        )
        self.logger.info("TTS_MODE | persistent subprocess worker")

    def _init_logger(self) -> logging.Logger:
        logs_dir = Path(__file__).resolve().parent / "logs"
        logs_dir.mkdir(exist_ok=True)
        logger = logging.getLogger("jarvis")
        logger.setLevel(logging.INFO)
        logger.propagate = False
        if not logger.handlers:
            fh = logging.FileHandler(logs_dir / "jarvis.log", encoding="utf-8")
            fh.setFormatter(logging.Formatter("%(asctime)s | %(message)s"))
            logger.addHandler(fh)
        return logger

    # NOTE: Windows SAPI is happier when it lives in a dedicated subprocess.
    # We keep one long-lived worker whose main thread loops on stdin and speaks
    # each line synchronously, which avoids repeated startup cost.

    def _start_tts_worker(self) -> None:
        worker_code = (
            "import sys\n"
            "import comtypes.client\n"
            "from comtypes import CoInitialize, CoUninitialize\n"
            "CoInitialize()\n"
            "voice = comtypes.client.CreateObject('SAPI.SpVoice')\n"
            "voice.Rate = 0\n"
            "voice.Volume = 100\n"
            "selected_name = 'default'\n"
            "try:\n"
            "    voices = voice.GetVoices()\n"
            "    preferred = None\n"
            "    for i in range(voices.Count):\n"
            "        candidate = voices.Item(i)\n"
            "        desc = f'{candidate.GetDescription()} {candidate.Id}'.lower()\n"
            "        if any(token in desc for token in ('zira', 'hoda', 'arabic', 'naayf')):\n"
            "            preferred = candidate\n"
            "            selected_name = candidate.GetDescription()\n"
            "            break\n"
            "    if preferred is not None:\n"
            "        voice.Voice = preferred\n"
            "except Exception:\n"
            "    pass\n"
            "sys.stdout.write(f'READY|{selected_name}\\n')\n"
            "sys.stdout.flush()\n"
            "for line in sys.stdin:\n"
            "    text = line.rstrip('\\n').replace('\\\\n', ' ').strip()\n"
            "    if not text:\n"
            "        continue\n"
            "    if text == '__EXIT__':\n"
            "        break\n"
            "    try:\n"
            "        voice.Speak(text, 0)\n"
            "    except Exception as ex:\n"
            "        sys.stderr.write(f'TTS_ERR {ex}\\n')\n"
            "        sys.stderr.flush()\n"
            "CoUninitialize()\n"
        )
        try:
            self.tts_proc = subprocess.Popen(
                [sys.executable, "-u", "-c", worker_code],
                stdin=subprocess.PIPE,
                stdout=subprocess.PIPE,
                stderr=subprocess.DEVNULL,
                text=True,
                encoding="utf-8",
                bufsize=1,
            )
            # wait for the worker to print READY (means the speech worker is live)
            ready = self.tts_proc.stdout.readline() if self.tts_proc.stdout else ""
            if not ready.strip().startswith("READY"):
                self.logger.info("TTS_WORKER_NOT_READY | %r", ready)
            else:
                self.logger.info("TTS_WORKER_READY | %s", ready.strip())
        except Exception as e:
            self.logger.info("TTS_WORKER_START_ERROR | %s", e)
            self.tts_proc = None

    def _stop_tts_worker(self) -> None:
        if self.tts_proc and self.tts_proc.stdin:
            try:
                self.tts_proc.stdin.write("__EXIT__\n")
                self.tts_proc.stdin.flush()
                self.tts_proc.wait(timeout=3)
            except Exception:
                try:
                    self.tts_proc.kill()
                except Exception:
                    pass
        self.tts_proc = None

    def _init_volume(self):
        try:
            dev = AudioUtilities.GetSpeakers()
            if dev is None:
                return None
            iface = dev.Activate(IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
            return cast(iface, POINTER(IAudioEndpointVolume))
        except Exception:
            return None

    # -- I/O -----------------------------------------------------------------

    def status(self, icon: str, text: str) -> None:
        ts = dt.datetime.now().strftime("%H:%M:%S")
        print(f"  [{ts}]  {icon}  {text}", flush=True)

    def _save_note(self, note: str) -> str:
        """Append a quick note to disk with a timestamp."""

        note = note.strip()
        if not note:
            return "Nothing to save."
        NOTES_FILE.parent.mkdir(parents=True, exist_ok=True)
        ts = dt.datetime.now().strftime("%Y-%m-%d %H:%M")
        with NOTES_FILE.open("a", encoding="utf-8") as f:
            f.write(f"[{ts}] {note}\n")
        self.logger.info("NOTE_SAVED | %s", note)
        return "Note saved."

    def _read_notes(self, limit: int = 5) -> str:
        """Read the latest quick notes back to Sir."""

        if not NOTES_FILE.exists():
            return "You do not have any notes yet."
        try:
            lines = [line.strip() for line in NOTES_FILE.read_text(encoding="utf-8").splitlines() if line.strip()]
        except Exception as e:
            return f"Could not read notes: {e}"
        if not lines:
            return "You do not have any notes yet."
        latest = lines[-limit:]
        return "Latest notes: " + " | ".join(latest)

    def _set_focus_mode(self, enabled: bool) -> str:
        """Enable or disable focus mode for proactive alerts."""

        self.focus_mode = enabled
        state = "enabled" if enabled else "disabled"
        self.logger.info("FOCUS_MODE | %s", state)
        return "Focus mode enabled. Proactive alerts are paused." if enabled else "Focus mode disabled. Alerts are back on."

    @staticmethod
    def _command_help() -> str:
        """Return a compact spoken cheat sheet of the most useful commands."""

        return (
            "I can open apps like Chrome, File Explorer, and VS Code; search Google or YouTube; "
            "open websites and folders; lock the PC; control volume, screenshots, and media; "
            "tell time and system status; save quick notes; and toggle focus mode."
        )

    def _fast_reply(self, user_text: str) -> str | None:
        """Handle very common small-talk and device questions without Ollama."""

        low = re.sub(r"\s+", " ", user_text.lower()).strip()
        if not low:
            return None

        low = low.replace("jarvis", "").strip(" ,.!?;:")
        compact = re.sub(r"[^\w\s%:.-]", " ", low)
        compact = re.sub(r"\s+", " ", compact).strip()
        if not compact:
            return None

        if any(
            compact == greeting or compact.startswith(greeting + " ")
            for greeting in ("hi", "hello", "hey", "good morning", "good afternoon", "good evening")
        ):
            return "Hello, Sir. I am ready."

        if any(phrase in compact for phrase in ("how are you", "how are u", "how's it going", "whats up", "what's up")):
            return "I am functioning normally, Sir."

        if any(phrase in compact for phrase in ("who are you", "what are you", "what is your name", "what's your name")):
            return "I am Jarvis, your personal assistant."

        if any(phrase in compact for phrase in ("what can you do", "what do you do", "what are your abilities")):
            return self._command_help()

        if compact in {"help", "commands", "show commands", "list commands", "what can you do", "what do you do"}:
            return self._command_help()

        if compact in {
            "ok good", "okay good", "all good", "sounds good", "great",
            "nice", "cool", "understood", "got it", "thanks", "thank you", "thx",
            "ok", "okay", "alright", "fine",
        }:
            return "Understood, Sir."

        if any(phrase in compact for phrase in ("what time is it", "current time", "time now", "what's the time", "whats the time", "كام الساعة")):
            return exec_tool("get_current_time", {}, self.vol, self.memory)

        if any(phrase in compact for phrase in ("what date is it", "current date", "today's date", "whats the date", "what day is it")):
            return exec_tool("get_current_time", {}, self.vol, self.memory)

        status_triggers = ("cpu", "processor", "ram", "memory", "disk", "battery")
        if any(trigger in compact for trigger in status_triggers) and any(
            trigger in compact for trigger in ("speed", "usage", "status", "info", "how much", "current", "percent")
        ):
            return exec_tool("get_system_status", {}, self.vol, self.memory)

        if compact in {"system status", "system info", "pc status", "device status", "battery", "battery status", "battery level", "cpu", "processor", "ram", "memory"}:
            return exec_tool("get_system_status", {}, self.vol, self.memory)

        if compact in {"cpu info", "processor info", "hardware info", "device specs", "specs"}:
            return exec_tool("get_system_status", {}, self.vol, self.memory)

        return None

    def _fast_command(self, user_text: str) -> str | None:
        """Execute obvious commands locally before falling back to the model."""

        low = re.sub(r"\s+", " ", user_text.lower()).strip()
        if not low:
            return None

        low = low.replace("jarvis", "").strip(" ,.!?;:")
        compact = re.sub(r"[^\w\s%:./-]", " ", low)
        compact = re.sub(r"\s+", " ", compact).strip()
        if not compact:
            return None

        def after_prefix(prefixes: tuple[str, ...]) -> str:
            for prefix in prefixes:
                if compact.startswith(prefix):
                    return compact[len(prefix):].strip()
            return ""

        def app_name_for(target: str) -> str | None:
            aliases = [
                ("chrome", ("google chrome", "chrome", "browser")),
                ("file explorer", ("file explorer", "explorer", "files")),
                ("vs code", ("visual studio code", "vs code", "vscode", "code")),
                ("notepad", ("notepad", "text editor")),
                ("calculator", ("calculator", "calc")),
                ("terminal", ("windows terminal", "terminal", "wt")),
                ("powershell", ("powershell", "power shell", "ps")),
                ("cmd", ("command prompt", "cmd")),
                ("task manager", ("task manager",)),
                ("settings", ("settings", "windows settings")),
                ("spotify", ("spotify",)),
                ("discord", ("discord",)),
                ("telegram", ("telegram",)),
                ("word", ("microsoft word", "word")),
                ("excel", ("microsoft excel", "excel")),
                ("powerpoint", ("microsoft powerpoint", "powerpoint", "ppt")),
                ("paint", ("paint",)),
                ("snipping tool", ("snipping tool", "snip")),
            ]
            for canonical, patterns in aliases:
                if any(p in target for p in patterns):
                    return canonical
            return None

        if compact in {"focus mode", "do not disturb", "quiet mode", "pause alerts"}:
            return self._set_focus_mode(not self.focus_mode)
        if compact in {"focus mode on", "turn focus mode on", "do not disturb on", "quiet mode on", "alerts off", "pause alerts on", "mute alerts"}:
            return self._set_focus_mode(True)
        if compact in {"focus mode off", "turn focus mode off", "do not disturb off", "quiet mode off", "alerts on", "resume alerts", "resume"}:
            return self._set_focus_mode(False)

        if compact.startswith(("save note ", "note ", "remember note ", "quick note ")):
            note = after_prefix(("save note ", "note ", "remember note ", "quick note "))
            if note:
                return self._save_note(note)

        if compact in {"show notes", "read notes", "list notes", "my notes", "open notes"}:
            return self._read_notes()

        if compact in {"time", "what time is it", "current time", "time now", "what's the time", "whats the time"}:
            return exec_tool("get_current_time", {}, self.vol, self.memory)

        if compact in {"date", "what date is it", "current date", "today's date", "whats the date", "what day is it"}:
            return exec_tool("get_current_time", {}, self.vol, self.memory)

        if compact in {"system status", "system info", "pc status", "device status", "battery", "battery status", "battery level", "cpu", "processor", "ram", "memory"}:
            return exec_tool("get_system_status", {}, self.vol, self.memory)

        if compact in {"lock", "lock computer", "lock the computer", "lock pc", "lock device", "lock screen"} or compact.startswith("lock "):
            return exec_tool("system_control", {"action": "lock"}, self.vol, self.memory)

        if compact.startswith(("open ", "launch ", "start ")):
            target = after_prefix(("open ", "launch ", "start "))
            if not target:
                return None

            site_target = target.removeprefix("website ").removeprefix("site ").strip()
            folder_target = target.removeprefix("folder ").strip()
            app = app_name_for(target)
            if app:
                return exec_tool("open_application", {"app_name": app}, self.vol, self.memory)
            if folder_target in FOLDER_MAP:
                return exec_tool("open_folder", {"folder": folder_target}, self.vol, self.memory)
            if site_target in SITE_MAP or "." in site_target or site_target.startswith(("http://", "https://")):
                return exec_tool("open_website", {"url": site_target}, self.vol, self.memory)
            if Path(site_target).expanduser().exists():
                return exec_tool("open_folder", {"folder": site_target}, self.vol, self.memory)

        if compact.startswith(("close ", "quit ", "kill ")):
            target = after_prefix(("close ", "quit ", "kill "))
            if not target:
                return None
            app = app_name_for(target)
            if app:
                return exec_tool("close_application", {"app_name": app}, self.vol, self.memory)

        if compact.startswith(("search ", "google ", "look up ", "find ")):
            query = after_prefix(("search ", "google ", "look up ", "find "))
            if query:
                return exec_tool("google_search", {"query": query}, self.vol, self.memory)

        if "youtube" in compact and any(compact.startswith(prefix) for prefix in ("search ", "open ", "play ")):
            query = re.sub(r"^(search|open|play)\s+(on\s+)?youtube\s*", "", compact).strip()
            if query:
                return exec_tool("youtube_search", {"query": query}, self.vol, self.memory)
            return exec_tool("open_website", {"url": "youtube.com"}, self.vol, self.memory)

        if compact.startswith(("website ", "site ", "go to ", "open website ", "open site ")):
            target = after_prefix(("website ", "site ", "go to ", "open website ", "open site "))
            if target:
                return exec_tool("open_website", {"url": target}, self.vol, self.memory)

        if compact.startswith(("folder ", "open folder ")):
            target = after_prefix(("folder ", "open folder "))
            if target:
                return exec_tool("open_folder", {"folder": target}, self.vol, self.memory)

        if any(phrase in compact for phrase in ("weather", "forecast")):
            city = None
            match = re.search(r"\bin\s+(.+)$", compact)
            if match:
                city = match.group(1).strip()
            return exec_tool("get_weather", {"city": city} if city else {}, self.vol, self.memory)

        if any(phrase in compact for phrase in ("volume up", "raise volume", "turn volume up", "louder")):
            return exec_tool("set_volume", {"action": "up"}, self.vol, self.memory)
        if any(phrase in compact for phrase in ("volume down", "lower volume", "turn volume down", "quieter")):
            return exec_tool("set_volume", {"action": "down"}, self.vol, self.memory)
        if any(phrase in compact for phrase in ("mute", "silence")) and "unmute" not in compact:
            return exec_tool("set_volume", {"action": "mute"}, self.vol, self.memory)
        if "unmute" in compact:
            return exec_tool("set_volume", {"action": "unmute"}, self.vol, self.memory)

        if any(phrase in compact for phrase in ("screenshot", "screen shot", "capture screen", "take screenshot")):
            return exec_tool("take_screenshot", {}, self.vol, self.memory)

        if any(phrase in compact for phrase in ("play pause", "pause/play", "play or pause")):
            return exec_tool("media_control", {"action": "play_pause"}, self.vol, self.memory)
        if compact in {"next", "next track", "skip", "skip track"}:
            return exec_tool("media_control", {"action": "next"}, self.vol, self.memory)
        if compact in {"previous", "previous track", "back track", "go back"}:
            return exec_tool("media_control", {"action": "previous"}, self.vol, self.memory)

        return None

    # If the model leaks JSON like {"type":"say","text":"..."}, recover the
    # natural-language inside instead of speaking the literal braces.
    @staticmethod
    def _sanitize_speech(text: str) -> str:
        s = text.strip()
        if not s:
            return s
        # Detect a JSON-ish wrapper and pull the speakable field out
        if s.startswith("{") and s.endswith("}"):
            try:
                obj = json.loads(s)
                if isinstance(obj, dict):
                    for key in ("text", "message", "content", "say", "reply", "answer"):
                        v = obj.get(key)
                        if isinstance(v, str) and v.strip():
                            return v.strip()
                    t = obj.get("type")
                    if isinstance(t, str):
                        type_map = {
                            "greet": "Hello, Sir. I am ready.",
                            "hello": "Hello, Sir. I am ready.",
                            "confirm": "Certainly, Sir.",
                            "ack": "Certainly, Sir.",
                            "ready": "I am ready, Sir.",
                        }
                        mapped = type_map.get(t.strip().lower())
                        if mapped:
                            return mapped
                    return ""
            except json.JSONDecodeError:
                return ""
        # Strip stray markdown / code fences
        s = s.replace("```", "").replace("`", "")
        return s

    def say(self, text: str) -> None:
        if not text:
            return
        text = self._sanitize_speech(text)
        if not text:
            return
        self.status(">>", text)
        self.logger.info("SPEAK | %s", text)
        # send text to the persistent TTS worker via stdin (instant)
        if self.tts_proc is None or self.tts_proc.poll() is not None:
            # worker died — try to restart it once
            self.logger.info("TTS_WORKER_RESTART")
            self._start_tts_worker()
        if self.tts_proc and self.tts_proc.stdin:
            try:
                # collapse newlines so each utterance is exactly one line
                line = text.replace("\r", " ").replace("\n", " ").strip()
                self.tts_proc.stdin.write(line + "\n")
                self.tts_proc.stdin.flush()
            except (BrokenPipeError, OSError) as e:
                self.logger.info("TTS_WRITE_ERROR | %s", e)
                self._start_tts_worker()

    def listen(self, *, timeout: int | None = None, limit: int = 10,
               prompt: str | None = "Listening...") -> str | None:
        try:
            with sr.Microphone() as src:
                if prompt:
                    self.status("~", prompt)
                audio = self.rec.listen(src, timeout=timeout, phrase_time_limit=limit)
        except sr.WaitTimeoutError:
            return None
        except Exception as e:
            self.logger.info("MIC_ERROR | %s", e)
            time.sleep(1)
            return None

        try:
            text = self.rec.recognize_google(audio, language="en-US")
            self.status("<<", text)
            self.logger.info("HEARD | %s", text)
            return text.strip()
        except sr.UnknownValueError:
            return None
        except sr.RequestError:
            self.logger.info("RECOGNIZE_NET_ERROR")
            time.sleep(2)
            return None
        except Exception as e:
            # catch socket timeouts and other network issues gracefully
            self.logger.info("RECOGNIZE_ERROR | %s", e)
            time.sleep(1)
            return None

    # -- Ollama integration --------------------------------------------------

    def _chat(self, *, stream: bool = False):
        """Call the local Ollama model with current history + system prompt."""
        system_prompt = build_system_prompt(self.memory.get_facts())
        return ollama.chat(
            model=OLLAMA_MODEL,
            messages=[{"role": "system", "content": system_prompt}, *self.history],
            tools=TOOLS,
            options={
                "temperature": 0.3,
                "num_predict": 80,    # voice replies are short — cap aggressively
                "num_ctx": 2048,      # smaller context = faster prefill
                "top_k": 20,
                "top_p": 0.9,
                "num_thread": os.cpu_count() or 4,  # use all CPU cores
            },
            keep_alive="30m",
            stream=stream,
        )

    # split a buffer at the first sentence boundary; returns (sentence, rest)
    _SENT_RE = re.compile(r"^(.*?[\.\!\?])(\s|$)", re.DOTALL)

    @classmethod
    def _take_sentence(cls, buf: str) -> tuple[str, str]:
        m = cls._SENT_RE.match(buf)
        if not m:
            return "", buf
        end = m.end(1)
        # also consume the trailing whitespace if present
        if end < len(buf) and buf[end].isspace():
            end += 1
        return buf[:m.end(1)].strip(), buf[end:]

    def think(self, user_text: str, *, system_origin: bool = False) -> str:
        """Send user input to the local model, streaming the reply aloud.

        If ``system_origin`` is True, the message is treated as a proactive
        observation from the monitor rather than something Sir actually said.

        Returns "" if the reply was already spoken via streaming, otherwise
        returns text the caller should speak.
        """

        content = user_text
        fast_reply = None if system_origin else self._fast_reply(content)
        if fast_reply is not None:
            self.logger.info("FAST_REPLY | %s | %s", content, fast_reply)
            self.history.append({"role": "user", "content": content})
            if len(self.history) > 10:
                self.history = self.history[-10:]
            self.history.append({"role": "assistant", "content": fast_reply})
            return fast_reply

        fast_command = None if system_origin else self._fast_command(content)
        if fast_command is not None:
            self.logger.info("FAST_COMMAND | %s | %s", content, fast_command)
            self.history.append({"role": "user", "content": content})
            if len(self.history) > 10:
                self.history = self.history[-10:]
            self.history.append({"role": "assistant", "content": fast_command})
            return fast_command

        self.history.append({"role": "user", "content": content})

        # keep history tiny so prefill stays fast
        if len(self.history) > 10:
            self.history = self.history[-10:]

        try:
            return self._stream_reply()
        except ollama.ResponseError as e:
            self.logger.info("OLLAMA_ERROR | %s", e)
            err = str(e).lower()
            if "not found" in err or "pull" in err:
                return f"The model {OLLAMA_MODEL} is not installed. Run: ollama pull {OLLAMA_MODEL}"
            if "does not support tools" in err:
                return f"The model {OLLAMA_MODEL} does not support tool calling. Try llama 3.1 or qwen 2.5 instead."
            return "The local model returned an error. Check the log file for details."
        except (ConnectionError, ConnectionRefusedError) as e:
            self.logger.info("OLLAMA_ERROR | %s", e)
            return "My model is still coming online, Sir. Please try again in a moment."
        except Exception as e:
            self.logger.info("OLLAMA_ERROR | %s", e)
            return "My model is still starting up, Sir. Please try again in a moment."

    def _stream_reply(self) -> str:
        """Stream the model's response. Speak each completed sentence as it
        arrives. If tool calls appear, execute them and recurse. Returns "" if
        all speech was already emitted via streaming."""

        stream = self._chat(stream=True)

        full_text = ""
        sentence_buf = ""
        tool_calls: list = []
        spoke_anything = False

        for chunk in stream:
            msg = chunk.message

            tcs = getattr(msg, "tool_calls", None) or []
            if tcs:
                tool_calls.extend(tcs)

            piece = (getattr(msg, "content", None) or "")
            if not piece:
                continue
            full_text += piece
            sentence_buf += piece

            # Don't speak if we're inside what looks like a JSON blob — wait
            # for the full message so _sanitize_speech can recover it.
            stripped = full_text.lstrip()
            if stripped.startswith("{"):
                continue

            # Speak any completed sentences in the buffer
            while True:
                sent, sentence_buf = self._take_sentence(sentence_buf)
                if not sent:
                    break
                self.say(sent)
                spoke_anything = True

        # End of stream — record assistant turn
        assistant_entry: dict = {"role": "assistant", "content": full_text}
        if tool_calls:
            assistant_entry["tool_calls"] = [
                {
                    "function": {
                        "name": tc.function.name,
                        "arguments": tc.function.arguments,
                    }
                }
                for tc in tool_calls
            ]
        self.history.append(assistant_entry)

        # If the model wanted to call tools, do it now
        if tool_calls:
            return self._run_tool_calls(tool_calls)

        # No tools — speak anything still left in the buffer
        leftover = sentence_buf.strip()
        if leftover:
            # If we never spoke anything (e.g. it was all JSON), let caller
            # speak the sanitised full text instead.
            if not spoke_anything:
                return self._sanitize_speech(full_text)
            self.say(leftover)
        elif not spoke_anything:
            # Nothing got spoken at all (rare). Return sanitised text.
            return self._sanitize_speech(full_text)
        return ""

    def _run_tool_calls(self, tool_calls) -> str:
        """Execute a list of tool calls and continue the conversation."""
        tool_outputs = []
        all_terminal = True
        for tc in tool_calls:
            name = tc.function.name
            args = tc.function.arguments
            if isinstance(args, str):
                try:
                    args = json.loads(args)
                except json.JSONDecodeError:
                    args = {}

            self.status("*", f"Executing: {name}")
            self.logger.info("TOOL | %s | %s", name, json.dumps(args))
            try:
                result = exec_tool(name, args, self.vol, self.memory)
            except Exception as e:
                result = f"Tool error: {e}"
            self.logger.info("RESULT | %s | %s", name, result)
            tool_outputs.append(result)

            if name not in TERMINAL_TOOLS:
                all_terminal = False

            self.history.append({"role": "tool", "name": name, "content": result})

        # Fast path: every tool's result IS the answer — speak directly
        if all_terminal:
            spoken = " ".join(tool_outputs)
            self.history.append({"role": "assistant", "content": spoken})
            return spoken

        # Otherwise stream a follow-up summary from the model
        return self._stream_reply()

    # -- Proactive observations ---------------------------------------------

    def _handle_alert(self, alert: Alert) -> None:
        """Take a monitor Alert and let JARVIS announce it in his own voice."""
        self.logger.info("ALERT | %s | %s", alert.type, alert.message)
        self.status("!", f"Alert: {alert.type}")
        self.say(self._alert_speech(alert))

    def _drain_alerts(self) -> None:
        """Speak or suppress any queued proactive alerts."""

        while not self.alert_queue.empty():
            try:
                alert = self.alert_queue.get_nowait()
            except queue.Empty:
                break
            if self.focus_mode:
                self.logger.info("ALERT_SUPPRESSED | focus_mode | %s | %s", alert.type, alert.message)
                continue
            try:
                self._handle_alert(alert)
            except Exception as e:
                self.logger.info("ALERT_HANDLE_ERROR | %s", e)

    @staticmethod
    def _alert_speech(alert: Alert) -> str:
        """Return a short spoken line for proactive monitor alerts."""

        if alert.type == "morning":
            return "Good morning, Sir. I am ready."
        if alert.type == "lunch":
            return "Sir, it is around lunchtime."
        if alert.type == "late_night":
            return "Sir, it is late. You may want to wind down."
        if alert.type == "screen_time":
            return alert.message
        return alert.message

    # -- Wake word & main loop -----------------------------------------------

    # Phrases that obviously look like commands — accepted without a wake word
    HOT_STARTS = (
        "open ", "close ", "play ", "pause ", "stop ", "search ", "google ",
        "youtube ", "what ", "what's ", "whats ", "tell ", "show ", "set ",
        "turn ", "lock ", "shut ", "restart ", "sleep", "take ", "screenshot",
        "next ", "previous ", "mute", "unmute", "volume ", "brightness ",
    )

    # Common Google-Speech mishearings of "jarvis"
    WAKE_ALIASES = {
        "jarvis", "jervis", "jervas", "jarvas", "jervice", "jeeves", "jarvice",
        "jervice", "service", "harvest", "harvis", "javis", "jaris",
    }

    def _matches_wake(self, word: str) -> bool:
        if word in self.WAKE_ALIASES:
            return True
        # fuzzy fallback for other mishearings (e.g. "jaarvis", "djarvis")
        return SequenceMatcher(None, word, "jarvis").ratio() >= 0.7

    # Tiny noise utterances we should NOT treat as commands.
    NOISE_WORDS = {
        "uh", "um", "hmm", "mm", "ah", "oh", "ok", "okay", "yeah", "yep",
        "no", "nope", "huh", "eh", "hi", "hey",
    }

    def wait_for_wake(self) -> str | None:
        """Open-mic mode: any meaningful speech is treated as a command.

        Sir does not need to say "Jarvis" each time. The wake word is still
        accepted as a courtesy and will be stripped from the command if used.
        """
        # 5s timeout so the main loop can drain proactive alerts between waits
        heard = self.listen(timeout=5, limit=10, prompt="Listening...")
        if not heard:
            return None

        low = heard.lower().strip()
        words = low.split()

        # ignore tiny noise utterances ("uh", "ok", single short words)
        if len(words) == 1 and (words[0] in self.NOISE_WORDS or len(words[0]) < 3):
            return None

        # if the user did say the wake word, strip it and use what's after.
        # if there's nothing after, fall back to using the whole phrase as
        # context (still no "Yes?" interjection — Sir asked us not to ping back).
        wake_idx = -1
        for i, w in enumerate(words):
            if self._matches_wake(w):
                wake_idx = i
                break
        if wake_idx >= 0:
            after = " ".join(words[wake_idx + 1 :]).strip()
            return after if after else None

        # no wake word — accept the whole phrase as a command anyway
        return heard

    def _warmup_ollama(self) -> None:
        """Force Ollama to load the model into memory in a background thread."""
        delay = 2
        while self.running:
            try:
                ollama.chat(
                    model=OLLAMA_MODEL,
                    messages=[{"role": "user", "content": "ok"}],
                    options={"num_predict": 1, "temperature": 0.0},
                    keep_alive="30m",
                )
                self.logger.info("OLLAMA_WARMUP_DONE")
                return
            except Exception as e:
                self.logger.info("OLLAMA_WARMUP_ERROR | %s", e)
                time.sleep(delay)
                delay = min(delay * 2, 30)

    def run(self) -> None:
        if sys.platform != "win32":
            print("This assistant is for Windows only.")
            return

        # kick off model warm-up in parallel with mic calibration
        warmup_thread = threading.Thread(target=self._warmup_ollama, daemon=True)
        warmup_thread.start()

        # mic init
        try:
            with sr.Microphone() as src:
                self.status("INIT", "Calibrating microphone...")
                self.rec.adjust_for_ambient_noise(src, duration=0.6)
            self.status("OK", "Microphone ready.")
        except Exception as e:
            print(f"  Microphone error: {e}")
            self.say("Could not access the microphone.")
            return

        print()
        print("  ╔══════════════════════════════════════╗")
        print("  ║   JARVIS — AI Voice Assistant        ║")
        print("  ║   Just talk naturally, no wake word   ║")
        print("  ║   Ask questions or give commands      ║")
        print("  ║   Say 'stop' to exit                  ║")
        print("  ╚══════════════════════════════════════╝")
        print()

        # start the background system monitor
        try:
            self.monitor.start()
            self.logger.info("MONITOR_START")
        except Exception as e:
            self.logger.info("MONITOR_START_ERROR | %s", e)

        # speak the personalised JARVIS briefing instead of a static line
        try:
            briefing = build_startup_briefing(self.memory, self.monitor)
        except Exception as e:
            self.logger.info("BRIEFING_ERROR | %s", e)
            briefing = "At your service, Sir."
        self.say(briefing)

        while self.running:
            try:
                # drain any proactive alerts that fired while we were idle
                self._drain_alerts()

                user_input = self.wait_for_wake()
                if not user_input:
                    continue

                low = user_input.lower()
                if any(w in low for w in ("stop", "exit", "quit", "goodbye", "shut down assistant")):
                    self.say("Very good, Sir. Until next time.")
                    self.running = False
                    break

                # send to the local model
                reply = self.think(user_input)
                if reply:
                    self.say(reply)

                # persist conversation + facts after every turn
                try:
                    self.memory.set_history(self.history)
                    self.memory.save()
                except Exception as e:
                    self.logger.info("MEMORY_SAVE_ERROR | %s", e)
            except KeyboardInterrupt:
                raise
            except Exception as e:
                self.logger.info("LOOP_ERROR | %s", e)
                self.status("!", f"Recovered from error: {e}")
                time.sleep(1)

    def shutdown(self) -> None:
        """Flush memory + stop the monitor + kill TTS worker. Safe to call twice."""
        try:
            self.monitor.stop()
        except Exception as e:
            self.logger.info("MONITOR_STOP_ERROR | %s", e)
        try:
            self.memory.set_history(self.history)
            self.memory.save()
        except Exception as e:
            self.logger.info("MEMORY_SAVE_ERROR | %s", e)
        try:
            self._stop_tts_worker()
        except Exception as e:
            self.logger.info("TTS_STOP_ERROR | %s", e)


def parse_args(argv: list[str] | None = None) -> argparse.Namespace:
    """Parse CLI flags for startup management and background launch."""

    parser = argparse.ArgumentParser(description="Jarvis voice assistant for Windows")
    parser.add_argument(
        "--install-startup",
        action="store_true",
        help="Create a Startup folder shortcut so Jarvis starts automatically.",
    )
    parser.add_argument(
        "--remove-startup",
        action="store_true",
        help="Remove the Startup folder shortcut created for Jarvis.",
    )
    parser.add_argument(
        "--startup-status",
        action="store_true",
        help="Check whether the Startup folder shortcut already exists.",
    )
    parser.add_argument(
        "--gui",
        action="store_true",
        help="Open the Jarvis desktop dashboard.",
    )
    parser.add_argument(
        "--autostart",
        action="store_true",
        help=argparse.SUPPRESS,
    )
    return parser.parse_args(argv)


def main(argv: list[str] | None = None) -> int:
    args = parse_args(argv)

    if args.install_startup:
        try:
            install_startup_task()
            print(f"Auto-start enabled: {STARTUP_SHORTCUT_NAME}")
            return 0
        except Exception as exc:
            print(f"Failed to enable auto-start: {exc}", file=sys.stderr)
            return 1

    if args.remove_startup:
        try:
            remove_startup_task()
            print(f"Auto-start removed: {STARTUP_SHORTCUT_NAME}")
            return 0
        except Exception as exc:
            print(f"Failed to remove auto-start: {exc}", file=sys.stderr)
            return 1

    if args.startup_status:
        exists = startup_task_exists()
        print(f"Auto-start status: {'enabled' if exists else 'disabled'}")
        return 0

    if args.gui:
        from gui import main as gui_main

        gui_main()
        return 0

    jarvis = None
    try:
        jarvis = Jarvis()
        jarvis.run()
    except KeyboardInterrupt:
        if jarvis:
            jarvis.say("Interrupted, Sir. Until next time.")
    except Exception as e:
        print(f"  Fatal: {e}", file=sys.stderr)
        if jarvis:
            jarvis.logger.info("FATAL | %s", e)
        return 1
    finally:
        if jarvis:
            jarvis.shutdown()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
