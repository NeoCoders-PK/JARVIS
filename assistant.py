"""
assistant.py - Talks to Groq's OpenAI-compatible chat completions API,
with tool calling so JARVIS can propose actions (open an app, read/write
a file, etc.) for main.py to confirm with the user and execute.

Get a free API key at https://console.groq.com/keys (no credit card
needed). Free tier as of Sep 2026: 30 requests/min, 6,000 tokens/min,
14,400 requests/day - shared across all Groq models. That's roughly
10-30x Gemini's free daily request cap, and far more generous (and
clearly documented) than Mistral's free tier. No free tier from any
provider is truly unlimited, though - if you outgrow this too, Groq's
paid "Dev tier" (just adding a card, no minimum spend) gives ~10x these
limits again. Check https://console.groq.com/docs/rate-limits for
current numbers if you ever see a 429 again - free-tier limits shift
over time.
"""
import json
import time

import requests

API_URL = "https://api.groq.com/openai/v1/chat/completions"


def _to_tools(function_declarations):
    """Wraps our plain {name, description, parameters} declarations in
    the {"type": "function", "function": {...}} shape OpenAI-style APIs
    (including Groq) expect for the `tools` field."""
    if not function_declarations:
        return None
    return [{"type": "function", "function": decl} for decl in function_declarations]


def ask_groq(messages, api_key, model="openai/gpt-oss-120b", function_declarations=None):
    """
    `messages` is already in OpenAI chat format (see main.py's history):
      {"role": "system"|"user", "content": str}
      {"role": "assistant", "content": str|None, "tool_calls": [...] }
      {"role": "tool", "tool_call_id": str, "content": str}

    Returns one of:
      {"type": "text", "text": str}
      {"type": "function_calls", "calls": [{"id": str, "name": str, "args": dict}, ...]}
    Raises RuntimeError on HTTP/network errors.
    """
    payload = {"model": model, "messages": messages}
    tools = _to_tools(function_declarations)
    if tools:
        payload["tools"] = tools

    headers = {
        "Content-Type": "application/json",
        "Authorization": f"Bearer {api_key}",
    }

    # Brief retry on 429 - a single burst (e.g. a tool-call round-trip
    # firing two requests close together) can trip the free tier's
    # per-second cap even well under the daily limit. Retry-After is
    # normally a couple seconds; if this loop runs out you're genuinely
    # out of daily quota, not just momentarily bursting.
    last_err = None
    for attempt in range(3):
        resp = requests.post(API_URL, headers=headers, json=payload, timeout=45)
        if resp.status_code == 429:
            wait = float(resp.headers.get("Retry-After", 2 * (attempt + 1)))
            last_err = RuntimeError(f"Groq error 429: {resp.text[:300]}")
            time.sleep(min(wait, 10))
            continue
        if resp.status_code != 200:
            raise RuntimeError(f"Groq error {resp.status_code}: {resp.text[:300]}")
        data = resp.json()
        try:
            message = data["choices"][0]["message"]
        except (KeyError, IndexError):
            raise RuntimeError(f"Unexpected Groq response: {data}")

        tool_calls = message.get("tool_calls")
        if tool_calls:
            calls = []
            for tc in tool_calls:
                fn = tc["function"]
                try:
                    args = json.loads(fn.get("arguments") or "{}")
                except json.JSONDecodeError:
                    args = {}
                calls.append({"id": tc["id"], "name": fn["name"], "args": args})
            return {"type": "function_calls", "calls": calls}

        return {"type": "text", "text": (message.get("content") or "").strip()}

    raise last_err
