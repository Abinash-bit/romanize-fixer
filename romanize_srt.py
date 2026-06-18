"""
romanize_srt.py
---------------
Romanize (transliterate) a native-script subtitle .srt file using the Anthropic
Claude API, then put the romanized text back onto the original timecodes.

Pipeline (see app.py "SRT Romanizer" tab):
    native-script .srt
      -> parse into blocks, strip the timecodes + index lines (token saving)
      -> send the bare text lines to claude-opus-4-8, get romanized lines back
      -> drop the romanized lines back onto the original timecodes
      -> romanized .srt  (then fed through the normal dictionary fixer)

Only the subtitle TEXT is sent to the model — timecodes and block numbers stay
local, which is the whole point of stripping them (fewer tokens per request).

Public API
==========
    text, stats = romanize_srt_bytes(srt_bytes, "Hindi", progress_cb=...)
    docx_bytes  = srt_text_to_docx_bytes(text)   # hand to the existing fixer
    LANGUAGES                                     # dropdown choices
"""

from __future__ import annotations
import io
import os
import re
import json
import hashlib
import logging

log = logging.getLogger("romanize.srt")

# claude-opus-4-8: Anthropic's most capable Opus-tier model. Transliteration is
# mechanical and format-sensitive, so we run with adaptive thinking off (omit it)
# and low effort — fast, cheap, and less likely to "overthink" the line format.
MODEL = "claude-opus-4-8"
BATCH_SIZE = 40          # lines per API call — small enough to keep 1:1 alignment

# label shown in the dropdown -> script hint handed to the model
LANGUAGES = {
    "Hindi":    "Hindi, written in the Devanagari script",
    "Punjabi":  "Punjabi, written in the Gurmukhi script",
    "Marathi":  "Marathi, written in the Devanagari script",
    "Tamil":    "Tamil, written in the Tamil script",
    "Telugu":   "Telugu, written in the Telugu script",
    "Kannada":  "Kannada, written in the Kannada script",
    "Gujarati": "Gujarati, written in the Gujarati script",
}

INDEX_RE = re.compile(r'^\d+$')
NUMBERED_RE = re.compile(r'^\s*(\d+)\t(.*)$')          # strict: "12\t<text>"
LOOSE_RE = re.compile(r'^\s*(\d+)\s*[\t.):\-]\s?(.*)$')  # fallback: "12. <text>" etc.


def anthropic_available() -> bool:
    """True if the anthropic SDK is importable (used to show a friendly error)."""
    try:
        import anthropic  # noqa: F401
        return True
    except Exception:
        return False


# ──────────────────────────────────────────────────────────────
# On-disk cache  (same file + language  ->  no second API call)
# ──────────────────────────────────────────────────────────────
# Bumped when the prompt/parsing logic changes so old results aren't reused.
_CACHE_VERSION = "v2"
# Override with ROMANIZE_CACHE_DIR to point at persistent storage (e.g. a mounted
# volume). On Streamlit Community Cloud the default dir is ephemeral — it works
# while the app is awake but is wiped on every reboot/redeploy.
CACHE_DIR = os.environ.get("ROMANIZE_CACHE_DIR") or os.path.join(
    os.path.dirname(os.path.abspath(__file__)), ".romanize_cache")


def _cache_key(data: bytes, language: str) -> str:
    h = hashlib.sha256()
    h.update(_CACHE_VERSION.encode("utf-8"))
    h.update(b"\x00")
    h.update(language.encode("utf-8"))
    h.update(b"\x00")
    h.update(data)        # exact file bytes — any edit is a different key
    return h.hexdigest()


def _cache_load(key: str):
    try:
        with open(os.path.join(CACHE_DIR, key + ".json"), "r", encoding="utf-8") as f:
            return json.load(f)
    except (OSError, ValueError):
        return None


def _cache_store(key: str, romanized_text: str, stats: dict):
    try:
        os.makedirs(CACHE_DIR, exist_ok=True)
        path = os.path.join(CACHE_DIR, key + ".json")
        tmp = path + ".tmp"
        with open(tmp, "w", encoding="utf-8") as f:
            json.dump({"romanized_text": romanized_text, "stats": stats},
                      f, ensure_ascii=False)
        os.replace(tmp, path)   # atomic — never leaves a half-written cache file
    except OSError as e:
        log.warning("Could not write romanization cache: %s", e)


def clear_cache() -> int:
    """Delete all cached romanizations. Returns how many were removed."""
    n = 0
    try:
        for name in os.listdir(CACHE_DIR):
            if name.endswith(".json"):
                os.remove(os.path.join(CACHE_DIR, name))
                n += 1
    except OSError:
        pass
    return n


# ──────────────────────────────────────────────────────────────
# SRT parsing / serialisation
# ──────────────────────────────────────────────────────────────
def decode_srt(data: bytes) -> str:
    for enc in ("utf-8-sig", "utf-8"):
        try:
            return data.decode(enc)
        except UnicodeDecodeError:
            continue
    return data.decode("latin-1", errors="replace")


class _Block:
    __slots__ = ("index", "timecode", "lines")

    def __init__(self, index, timecode, lines):
        self.index = index        # original index line (str) or None
        self.timecode = timecode  # original "00:00:01,200 --> ..." (str) or None
        self.lines = lines         # list[str] of text lines (mutated in place)


def parse_srt(text: str):
    """Tolerant SRT parser. Splits on blank lines; finds the timecode line by the
    '-->' marker so missing/extra index lines don't throw it off."""
    text = text.replace("\r\n", "\n").replace("\r", "\n")
    chunks = re.split(r'\n[ \t]*\n', text.strip("\n"))
    blocks = []
    for chunk in chunks:
        if not chunk.strip():
            continue
        lines = chunk.split("\n")
        tc_i = next((i for i, l in enumerate(lines) if '-->' in l), None)
        if tc_i is None:
            # no timecode in this chunk — keep its lines as plain text
            blocks.append(_Block(None, None, lines))
            continue
        index = None
        for l in lines[:tc_i]:
            if INDEX_RE.match(l.strip()):
                index = l.strip()
        blocks.append(_Block(index, lines[tc_i].strip(), lines[tc_i + 1:]))
    return blocks


def serialize_srt(blocks) -> str:
    out = []
    for b in blocks:
        parts = []
        if b.index is not None:
            parts.append(b.index)
        if b.timecode is not None:
            parts.append(b.timecode)
        parts.extend(b.lines)
        out.append("\n".join(parts))
    return "\n\n".join(out).strip() + "\n"


def _collect(blocks):
    """Flatten every non-blank text line into a list, keeping (block, line) refs
    so the romanized output can be written straight back."""
    refs, texts = [], []
    for bi, b in enumerate(blocks):
        for lj, line in enumerate(b.lines):
            if line.strip():
                refs.append((bi, lj))
                texts.append(line)
    return refs, texts


# ──────────────────────────────────────────────────────────────
# Anthropic transliteration
# ──────────────────────────────────────────────────────────────
SYSTEM_TEMPLATE = (
    "You are an expert transliterator for Indian-language subtitles (Same Language "
    "Subtitling, used for reading practice).\n\n"
    "Every user message contains numbered subtitle lines in {language}. For each "
    "line, TRANSLITERATE it into the Latin/Roman alphabet — write how the words "
    "SOUND using ordinary English letters. Do NOT translate the meaning into "
    "English; only convert the script to Roman letters.\n\n"
    "Output rules (follow EXACTLY):\n"
    "- Return one output line for every input line, in the same order.\n"
    "- Each output line MUST be: the same number, then a single TAB character "
    "(\\t), then the romanized text. Example:  12\\tmera dil kho gaya\n"
    "- Use natural, widely-used phonetic romanization (the common way these "
    "words are written in Roman script).\n"
    "- Use normal sentence capitalization: capitalize the first letter of each "
    "sentence and proper nouns (names, places). Keep all other words lowercase. "
    "Do NOT lowercase everything, and do NOT capitalize every word.\n"
    "- Convert native-script digits to Western digits (e.g. १९४७ -> 1947).\n"
    "- Keep punctuation, musical symbols (like ♪), and any text that is already "
    "in Roman letters unchanged.\n"
    "- If a line cannot be transliterated, repeat it unchanged (still with its "
    "number and tab).\n"
    "- Never merge, split, drop, reorder, or add lines. Output ONLY the numbered "
    "lines — no commentary, headings, code fences, or blank lines."
)

# effort is an optimisation; drop it automatically on an SDK that doesn't know it
_EXTRA = {"output_config": {"effort": "low"}}


def _parse_numbered(output: str, n: int):
    """Map the model's reply back to n lines. Prefers the explicit 'N\\t...' form;
    falls back to positional mapping if the count lines up but numbering drifted."""
    result = [None] * n
    for line in output.split("\n"):
        m = NUMBERED_RE.match(line)
        if not m:
            continue
        idx = int(m.group(1))
        if 1 <= idx <= n:
            result[idx - 1] = m.group(2)

    if any(x is None for x in result):
        cand = []
        for line in output.split("\n"):
            if not line.strip():
                continue
            m = LOOSE_RE.match(line)
            cand.append(m.group(2) if m else line.strip())
        if len(cand) == n:
            for i in range(n):
                if result[i] is None:
                    result[i] = cand[i]
    return result


def _romanize_batch(client, batch, language_desc):
    """Romanize one batch of lines. Returns (lines, n_fell_back_to_original)."""
    global _EXTRA
    numbered = "\n".join(f"{i}\t{t}" for i, t in enumerate(batch, 1))
    system = [{
        "type": "text",
        "text": SYSTEM_TEMPLATE.format(language=language_desc),
        "cache_control": {"type": "ephemeral"},   # reused across batches in a run
    }]
    max_tokens = min(16000, max(1500, len(batch) * 200))
    kwargs = dict(
        model=MODEL,
        max_tokens=max_tokens,
        system=system,
        messages=[{"role": "user", "content": numbered}],
    )
    # stream (avoids HTTP timeouts on large batches); .get_final_message() gives
    # the assembled reply without hand-handling events.
    try:
        with client.messages.stream(**kwargs, **_EXTRA) as stream:
            msg = stream.get_final_message()
    except TypeError:
        # installed SDK predates output_config/effort — retry without it
        _EXTRA = {}
        with client.messages.stream(**kwargs) as stream:
            msg = stream.get_final_message()

    text = "".join(b.text for b in msg.content if getattr(b, "type", None) == "text")
    parsed = _parse_numbered(text, len(batch))

    out, n_missing = [], 0
    for orig, rom in zip(batch, parsed):
        if rom is None or not rom.strip():
            out.append(orig)      # safest fallback: keep the source line
            n_missing += 1
        else:
            out.append(rom)
    return out, n_missing


def romanize_srt_bytes(data: bytes, language: str, progress_cb=None, api_key=None,
                       use_cache: bool = True):
    """Romanize a native-script .srt.

    Returns (romanized_srt_text, stats) where stats = {lines, blocks, missing, cached}.
    `progress_cb(done, total)` is called after each batch (optional).

    If `use_cache` and an identical (file bytes + language) was romanized before,
    the stored result is returned WITHOUT contacting Anthropic.
    """
    key = _cache_key(data, language)
    if use_cache:
        hit = _cache_load(key)
        if hit is not None:
            stats = dict(hit.get("stats") or {})
            stats["cached"] = True
            log.info("Romanization cache HIT (%s…) — no API call", key[:8])
            return hit["romanized_text"], stats

    import anthropic  # imported lazily so the rest of the app works without it

    language_desc = LANGUAGES.get(language, language)
    blocks = parse_srt(decode_srt(data))
    refs, texts = _collect(blocks)
    total = len(texts)
    if total == 0:
        stats = {"lines": 0, "blocks": len(blocks), "missing": 0, "cached": False}
        text = serialize_srt(blocks)
        if use_cache:
            _cache_store(key, text, stats)
        return text, stats

    client = anthropic.Anthropic(api_key=api_key) if api_key else anthropic.Anthropic()

    romanized, missing = [], 0
    for start in range(0, total, BATCH_SIZE):
        batch = texts[start:start + BATCH_SIZE]
        out, n_missing = _romanize_batch(client, batch, language_desc)
        romanized.extend(out)
        missing += n_missing
        if progress_cb:
            progress_cb(min(start + len(batch), total), total)

    for (bi, lj), rom in zip(refs, romanized):
        blocks[bi].lines[lj] = rom

    text = serialize_srt(blocks)
    stats = {"lines": total, "blocks": len(blocks), "missing": missing, "cached": False}
    log.info("Romanized %d lines across %d blocks (%d kept original)",
             total, len(blocks), missing)
    if use_cache:
        _cache_store(key, text, stats)
    return text, stats


def srt_text_to_docx_bytes(srt_text: str) -> bytes:
    """Wrap an .srt (one line per paragraph) in a .docx so it can be fed to the
    existing dictionary fixer, which expects a 'Romanized subtitle docx'."""
    from docx import Document
    doc = Document()
    for line in srt_text.split("\n"):
        doc.add_paragraph(line)
    buf = io.BytesIO()
    doc.save(buf)
    buf.seek(0)
    return buf.getvalue()
