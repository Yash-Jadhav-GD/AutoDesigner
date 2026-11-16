"""
Claude mapper: sends a compact JSON slide model (text, roles, geometry, typography; no images)
to Claude, which picks an archetype and fills its schema. The result is schema-validated and
coverage-checked locally; one retry with the errors, then fall back to the rules mapper.

Responses are cached in .cache/understand/ keyed by the exact request, so re-runs and evals
make no API calls.
"""

import copy
import hashlib
import json
import os
from pathlib import Path
from typing import Any, Callable, Dict, List, Optional

from ..slide_model import SlideModel, coverage
from ..builders.registry import DESCRIPTIONS
from ..builders.schemas import ADAPTERS, json_schema, validate_data
from .rules import candidates, map_slide

MODEL = "claude-opus-5-5"
PROMPT_VERSION = "1"
CACHE_DIR = Path(__file__).resolve().parents[2] / ".cache" / "understand"
FALLBACK_BETA = "server-side-fallback-2026-07-01"


def credentials_available() -> bool:
    if os.environ.get("ANTHROPIC_API_KEY") or os.environ.get("ANTHROPIC_AUTH_TOKEN"):
        return True
    return (Path.home() / ".config" / "anthropic").exists()


# ------------------------------------------------------------------------------
# Prompt
# ------------------------------------------------------------------------------

def _strict(schema: Dict) -> Dict:
    """Structured outputs need every property required and no extra keys; drop defaults/titles."""
    s = copy.deepcopy(schema)

    def walk(n):
        if isinstance(n, dict):
            n.pop("default", None)
            n.pop("title", None)
            if n.get("type") == "object" and "properties" in n:
                n["required"] = list(n["properties"].keys())
                n["additionalProperties"] = False
            for v in n.values():
                walk(v)
        elif isinstance(n, list):
            for v in n:
                walk(v)
    walk(s)
    return s


def _envelope_schema() -> Dict:
    defs: Dict[str, Any] = {}
    variants = []
    for name in ADAPTERS:
        sch = _strict(json_schema(name))
        for k, v in sch.pop("$defs", {}).items():
            defs[k] = v
        variants.append(sch)
    return {
        "type": "object",
        "properties": {
            "archetype": {"type": "string", "enum": list(ADAPTERS)},
            "confidence": {"type": "number"},
            "reasons": {"type": "array", "items": {"type": "string"}},
            "data": {"anyOf": variants},
        },
        "required": ["archetype", "confidence", "reasons", "data"],
        "additionalProperties": False,
        "$defs": defs,
    }


def _system_prompt() -> str:
    catalogue = "\n".join(
        f"### {name}\n{DESCRIPTIONS[name]}\nData schema:\n{json.dumps(json_schema(name), ensure_ascii=False)}"
        for name in ADAPTERS
    )
    return f"""You map one PowerPoint draft slide onto the best-fitting layout archetype for a GEP-branded deck.

You receive the slide as JSON: every text element with its position and size in points (960x540 slide),
font size, bold flag, fill colour and text; tables as cell grids; filled shapes that are not text
(useful as visual containers, badges and legend swatches).

Rules:
- Copy text verbatim from the elements into the data. Never paraphrase, translate, shorten, merge words or invent text.
- Every word on the slide must appear somewhere in the data. If text does not fit a field naturally,
  choose a more general archetype (card_grid is always valid) rather than dropping it.
- Use geometry and typography to infer roles: larger/bold text is a heading; repeated aligned blocks are
  lists or grids; colour fills shared between shapes and a legend encode categories.
- Prefer a branded archetype only when the slide's structure genuinely matches its description.
- confidence is your probability (0-1) that the archetype and field assignment are what a designer would choose.

Archetypes:
{catalogue}
"""


def _slide_payload(m: SlideModel) -> Dict:
    els = []
    for e in m.elements:
        if e.kind in ("group", "connector"):
            continue
        if e.kind == "shape" and not e.fill:
            continue
        row = {"id": e.id, "kind": e.kind, "x": round(e.x), "y": round(e.y), "w": round(e.w), "h": round(e.h)}
        if e.kind == "table":
            row["table"] = e.table
        elif e.kind == "text":
            row.update({"size": e.size, "bold": e.bold, "text": e.text})
        if e.fill:
            row["fill"] = e.fill
        if e.placeholder:
            row["placeholder"] = e.placeholder
        els.append(row)
    return {"slide": m.index, "elements": els}


# ------------------------------------------------------------------------------
# Call
# ------------------------------------------------------------------------------

def _call(system: str, messages: List[Dict]) -> Dict:
    import anthropic

    client = anthropic.Anthropic()
    response = client.messages.create(
        model=MODEL,
        max_tokens=16000,
        system=[{"type": "text", "text": system, "cache_control": {"type": "ephemeral"}}],
        messages=messages,
        output_config={"effort": "medium", "format": {"type": "json_schema", "schema": _envelope_schema()}},
        # Server-side refusal fallback; this SDK version has no typed `fallbacks` parameter.
        extra_headers={"anthropic-beta": FALLBACK_BETA},
        extra_body={"fallbacks": "default"},
    )
    if response.stop_reason == "refusal":
        raise RuntimeError("Claude declined this slide")
    if response.stop_reason == "max_tokens":
        raise RuntimeError("Claude response truncated (max_tokens)")
    text = next(b.text for b in response.content if b.type == "text")
    return json.loads(text)


def _cache_path(key: Dict) -> Path:
    digest = hashlib.sha256(json.dumps(key, sort_keys=True, ensure_ascii=False).encode()).hexdigest()[:32]
    return CACHE_DIR / f"{digest}.json"


def map_slide_llm(m: SlideModel, log: Callable = print, use_cache: bool = True) -> Dict[str, Any]:
    rules_result = map_slide(m)
    payload = _slide_payload(m)
    payload["rules_suggestion"] = [{"archetype": c["archetype"], "score": c["score"]} for c in candidates(m)[:3]]
    system = _system_prompt()
    key = {"model": MODEL, "v": PROMPT_VERSION, "system": system, "payload": payload}
    path = _cache_path(key)

    if use_cache and path.exists():
        return json.loads(path.read_text(encoding="utf-8"))

    messages: List[Dict] = [{"role": "user", "content": json.dumps(payload, ensure_ascii=False)}]
    log(f"  [claude] sending slide {m.index} ({len(payload['elements'])} elements) to {MODEL}")
    try:
        for attempt in range(2):
            out = _call(system, messages)
            problems = []
            try:
                data = validate_data(out["archetype"], out["data"])
            except Exception as e:
                data, problems = None, [f"data does not match the {out['archetype']} schema: {e}"]
            if data is not None:
                missing = coverage(m, data)["missing"]
                if missing:
                    problems.append("these words from the slide are missing from data: " + " ".join(missing[:80]))
            if not problems:
                result = {"archetype": out["archetype"], "data": data, "confidence": out["confidence"],
                          "reasons": out.get("reasons", []), "source": "llm", "model": MODEL}
                CACHE_DIR.mkdir(parents=True, exist_ok=True)
                path.write_text(json.dumps(result, ensure_ascii=False, indent=1), encoding="utf-8")
                return result
            messages += [
                {"role": "assistant", "content": json.dumps(out, ensure_ascii=False)},
                {"role": "user", "content": "Fix these problems and answer again:\n- " + "\n- ".join(problems)},
            ]
        reason = "Claude result failed validation twice"
    except Exception as e:
        reason = f"Claude call failed: {e.__class__.__name__}: {str(e)[:200]}"
    log(f"  [claude] slide {m.index}: {reason}; using rules result")
    rules_result["reasons"] = rules_result["reasons"] + [reason]
    return rules_result
