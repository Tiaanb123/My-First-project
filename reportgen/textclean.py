"""Keep raw ticket text fit for a client-facing report.

CarlSource comments sometimes carry SSH steps, share-drive links and whole JSON config dumps.
`limit_text` cuts a section off at the first such line and caps its length;
`summarise_freeform` builds a short stand-in for tickets that don't use the
Reported/Findings/Remedy template.
"""
from __future__ import annotations

import re

MAX_BULLET_CHARS = 700

_TECH_LINE_RE = re.compile(
    r"""(?ix)
      [{}]                                         # JSON / config blocks
    | https?://|sharepoint|\.sharepoint\.
    | \b(?:ssh|sudo|docker|chmod|systemctl|scp)\b
    | \S{60,}                                      # long unbroken token
    | ^\s*(?:resolution|steps?|procedure|commands?)\s*:?\s*$
    | ^\s*step\s*\d
    """
)
_DATE_PREFIX_RE = re.compile(r"^\s*\d{1,2}/\d{1,2}\s*:?\s*")
_PLACEHOLDER_RE = re.compile(r"^\s*(?:xx|\d{1,2})\s*/\s*(?:xx|\d{1,2})\s*:?\s*$", re.IGNORECASE)
_SKIP_LINE_RE = re.compile(
    r"""(?ix)
      ^\s*(?:call\s*ref(?:erence)?|help)\s*:?[\d,.&+\s]*(?:and)?[\d,.&+\s]*$
    | ^\s*[A-Za-z ]{1,25}:\s*$                     # bare label such as "Background:"
    | ^.*\.(?:png|jpe?g|gif|bmp)\s*$               # pasted screenshot names
    """
)
_SENTENCE_SPLIT_RE = re.compile(r"(?<=[.!?])\s+")


def is_placeholder(text: str) -> bool:
    """True for empty text or a lone template date marker such as 'xx/09:'."""
    return not text.strip() or bool(_PLACEHOLDER_RE.match(text))


def _cap(text: str, wo: str, limit: int) -> str:
    if len(text) <= limit:
        return text
    cut = text[:limit]
    boundary = cut.rfind(". ")
    if boundary >= limit // 2:
        cut = cut[: boundary + 1]
    elif (boundary := cut.rfind(" ")) > 0:
        cut = cut[:boundary]
    return f"{cut.rstrip()} … (see WO {wo} for full detail)"


def limit_text(text: str, wo: str, limit: int = MAX_BULLET_CHARS) -> str:
    """Drop everything from the first technical line onward, then cap the length."""
    if not text.strip():
        return ""
    kept: list[str] = []
    technical = False
    for line in text.splitlines():
        if _TECH_LINE_RE.search(line):
            technical = True
            break
        if not _PLACEHOLDER_RE.match(line):  # stray template date markers ("xx/09:")
            kept.append(line)
    body = "\n".join(ln.strip() for ln in kept if ln.strip())
    body = _cap(body, wo, limit)
    if technical:
        note = f"(Further technical detail recorded in WO {wo}; omitted here for readability.)"
        body = f"{body} {note}".strip()
    return body


def summarise_freeform(freeform: str, wo: str, sentences: int = 3) -> str:
    """First few sentences of unstructured ticket text, technical tail removed."""
    units: list[str] = []
    for raw in freeform.splitlines():
        line = _DATE_PREFIX_RE.sub("", raw).replace("*", "").strip()
        if not line or _SKIP_LINE_RE.match(line) or _PLACEHOLDER_RE.match(line):
            continue
        if _TECH_LINE_RE.search(line):
            break
        for unit in _SENTENCE_SPLIT_RE.split(line):
            unit = unit.strip()
            if unit:
                units.append(unit if unit[-1] in ".!?" else unit + ".")
    return _cap(" ".join(units[:sentences]), wo, MAX_BULLET_CHARS)
