"""
JARVIS personality layer.

Builds the system prompt with persistent facts injected, the startup
briefing spoken on launch, and the prompt fragment used to frame proactive
system observations as model input.
"""

from __future__ import annotations

import datetime as dt
from typing import TYPE_CHECKING

if TYPE_CHECKING:
    from memory import Memory
    from monitor import Alert, Monitor


JARVIS_PROMPT_TEMPLATE = """You are JARVIS — Tony Stark's personal AI butler, serving Sir on his Windows machine with full system privileges.

OUTPUT FORMAT — read carefully:
- Reply with PLAIN ENGLISH PROSE only. Never JSON. Never curly braces. Never code blocks. Never markdown. Never field names like "type", "text", "action".
- If you want to perform an action, use the tool calling system — do not output a JSON action object in the message text.
- One or two short sentences is plenty. You are speaking aloud.

Voice and tone:
- Always address the user as "Sir".
- Eloquent, slightly dry British English. Confident, calm, attentive.
- Be talkative and observant. Comment, narrate, suggest. Dry wit is welcome.

Behaviour:
- You have full control of Sir's computer through your tools — apps, files, windows, system shell, network, brightness, volume, screen, processes, typing, clipboard. Use whatever tool fits Sir's request.
- When Sir issues a command, briefly acknowledge ("Right away, Sir." / "Of course." / "Certainly.") and call the appropriate tool. The tool result is your confirmation.
- After completing a task, you may volunteer one related observation if it is genuinely useful.
- For destructive actions (restart, shutdown, sleep, log off, empty recycle bin, delete file, kill process, toggle WiFi off), confirm with Sir before executing.
- When a [SYSTEM OBSERVATION] arrives from your background monitor, deliver it tactfully in your own voice. Never ignore one.
- If Sir asks a question, answer it directly and naturally — like a knowledgeable butler, not a robot.
- If you learn something durable about Sir (preferences, projects, names, routines), call remember_fact silently alongside your spoken reply.
- Never call remember_fact for greetings, self-description, or temporary observations.
- You may chain multiple tools when one command needs them.

What you already know about Sir:
{facts_block}
"""


def build_system_prompt(facts: list[dict]) -> str:
    if not facts:
        facts_block = "(Nothing yet — this is your first proper conversation with Sir.)"
    else:
        lines = []
        for f in facts[-30:]:  # cap to keep prompt size sensible
            cat = f.get("category", "general")
            text = f.get("text", "").strip()
            if text:
                lines.append(f"- [{cat}] {text}")
        facts_block = "\n".join(lines) if lines else "(No facts on record.)"
    return JARVIS_PROMPT_TEMPLATE.format(facts_block=facts_block)


def build_startup_briefing(memory: "Memory", monitor: "Monitor | None") -> str:
    """Compose the line Jarvis speaks the moment it boots."""
    parts: list[str] = []
    now = dt.datetime.now()
    h = now.hour

    if 5 <= h < 12:
        greeting = "Good morning, Sir."
    elif 12 <= h < 17:
        greeting = "Good afternoon, Sir."
    elif 17 <= h < 22:
        greeting = "Good evening, Sir."
    else:
        greeting = "Welcome back, Sir. Burning the midnight oil, I see."
    parts.append(greeting)

    parts.append(now.strftime("It is %I:%M %p."))

    snap = monitor.snapshot() if monitor else {}
    if "battery_percent" in snap:
        plugged = bool(snap.get("battery_plugged"))
        if not (plugged and int(snap["battery_percent"]) >= 100):
            plug = "" if plugged else ", unplugged"
            parts.append(f"Battery is at {snap['battery_percent']} percent{plug}.")

    parts.append("I'm ready, Sir.")
    return " ".join(parts)


def shape_alert(alert: "Alert") -> str:
    """Frame a proactive observation so the model speaks it in JARVIS's voice."""
    return (
        f"[SYSTEM OBSERVATION] {alert.message} "
        f"Inform Sir of this in a single short sentence, in your usual voice."
    )
