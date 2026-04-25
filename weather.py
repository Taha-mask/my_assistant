"""
OpenWeatherMap client.

Returns a short, natural-language summary that the model can read aloud.
On any failure (no key, no internet, bad city) it returns a clear sentence
the model can relay verbatim — never raises.
"""

from __future__ import annotations

import os

try:
    import requests
except ImportError:  # pragma: no cover
    requests = None  # type: ignore

API_URL = "https://api.openweathermap.org/data/2.5/weather"
DEFAULT_FALLBACK_CITY = "Cairo"


def get_weather(city: str | None = None) -> str:
    if requests is None:
        return "Weather is unavailable — the requests library is not installed."

    api_key = os.getenv("OPENWEATHER_API_KEY", "").strip()
    if not api_key or api_key == "your_key_here":
        return (
            "Weather is unavailable — no OpenWeatherMap API key is configured. "
            "Sir may add one to the env file."
        )

    city = (city or os.getenv("DEFAULT_CITY") or DEFAULT_FALLBACK_CITY).strip()

    try:
        r = requests.get(
            API_URL,
            params={"q": city, "appid": api_key, "units": "metric"},
            timeout=6,
        )
    except requests.exceptions.ConnectionError:
        return "I cannot reach the weather service. The internet appears to be down."
    except requests.exceptions.Timeout:
        return "The weather service is taking too long to respond."
    except Exception as e:
        return f"Weather lookup failed: {e}"

    if r.status_code == 401:
        return "The weather API key is invalid."
    if r.status_code == 404:
        return f"I could not find weather data for {city}."
    if r.status_code != 200:
        return f"Weather service returned an error ({r.status_code})."

    try:
        data = r.json()
        temp = round(data["main"]["temp"])
        feels = round(data["main"]["feels_like"])
        humidity = data["main"]["humidity"]
        desc = data["weather"][0]["description"]
        name = data.get("name") or city
        return (
            f"{temp} degrees Celsius and {desc} in {name}, "
            f"feels like {feels}, humidity {humidity} percent."
        )
    except (KeyError, ValueError, TypeError) as e:
        return f"Weather data was malformed: {e}"
