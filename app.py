import streamlit as st
import re
import copy
import io
import time
import logging
from docx import Document
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from collections import Counter

from suggest import SuggestEngine

# ─────────────────────────────────────────────
# Logging — shows in the terminal where you ran `streamlit run app.py`
# ─────────────────────────────────────────────
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s | %(levelname)-5s | %(name)s | %(message)s",
    datefmt="%H:%M:%S",
)
log = logging.getLogger("romanize.app")

# ─────────────────────────────────────────────
# Tier styling
# ─────────────────────────────────────────────
# exact   : confirmed dictionary correction        -> yellow highlight
# fix     : engine fix, higher confidence (review) -> green  highlight
# suggest : engine guess, lower confidence (review)-> cyan   highlight
TIER_HL = {"exact": "yellow", "fix": "green", "suggest": "cyan"}

FIX_CUTOFF = 0.75       # >= this -> "fix" tier (green). below -> "suggest"
SUGGEST_FLOOR = 0.60    # below this -> ignore entirely (the 0.50-0.60 band was
                        # ~50% accurate — a coin flip — so we don't show it)


# ─────────────────────────────────────────────
# Dictionary loading
# ─────────────────────────────────────────────
def _clean_note(s: str) -> str:
    """Strip human annotation notes like [accent] / (as per dialogue)."""
    s = re.sub(r"\[[^\]]*\]", "", s)
    s = re.sub(r"\([^)]*\)", "", s)
    return s.strip()


def load_dictionary(file_bytes: bytes):
    """Return (exact_replacements, pairs).

    exact_replacements: {variant_lower: planetread_word}  (verbatim, as before)
    pairs:              [(wrong, right), ...]  cleaned, for the learning engine
    """
    t0 = time.perf_counter()
    log.info("Loading dictionary (%d KB)…", len(file_bytes) // 1024)
    doc = Document(io.BytesIO(file_bytes))
    replacements = {}
    pairs = []
    for table in doc.tables:
        for row in table.rows:
            cells = [c.text.strip() for c in row.cells]
            if len(cells) < 3:
                continue
            ai_word, pr_word = cells[1].strip(), cells[2].strip()
            if not ai_word or not pr_word:
                continue
            if ai_word.lower() in ("ai version", "ai_version"):
                continue
            for variant in ai_word.split('/'):
                variant = variant.strip()
                if variant:
                    replacements[variant.lower()] = pr_word
            a, b = _clean_note(ai_word), _clean_note(pr_word)
            if a and b:
                pairs.append((a, b))
    log.info("Dictionary loaded: %d exact variants, %d clean pairs in %.2fs",
             len(replacements), len(pairs), time.perf_counter() - t0)
    return replacements, pairs


# ─────────────────────────────────────────────
# Run XML builders
# ─────────────────────────────────────────────
def _base_rPr(source_run_xml=None):
    if source_run_xml is not None:
        src = source_run_xml.find(qn('w:rPr'))
        if src is not None:
            return copy.deepcopy(src)
    return OxmlElement('w:rPr')


def _make_t(text: str):
    t = OxmlElement('w:t')
    t.text = text
    if text.startswith(' ') or text.endswith(' '):
        t.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
    return t


def _clean_rPr(rPr):
    for tag in ('w:color', 'w:strike', 'w:highlight'):
        for el in rPr.findall(qn(tag)):
            rPr.remove(el)
    return rPr


def build_red_run(text, source_run_xml=None):
    r = OxmlElement('w:r')
    rPr = _clean_rPr(_base_rPr(source_run_xml))
    color = OxmlElement('w:color')
    color.set(qn('w:val'), 'FF0000')
    rPr.append(color)
    r.append(rPr)
    r.append(_make_t(text))
    return r


def build_highlight_run(text, color, source_run_xml=None):
    r = OxmlElement('w:r')
    rPr = _clean_rPr(_base_rPr(source_run_xml))
    hl = OxmlElement('w:highlight')
    hl.set(qn('w:val'), color)
    rPr.append(hl)
    r.append(rPr)
    r.append(_make_t(text))
    return r


def build_normal_run(text, source_run_xml=None):
    r = OxmlElement('w:r')
    r.append(_base_rPr(source_run_xml))
    r.append(_make_t(text))
    return r


# ─────────────────────────────────────────────
# Per-paragraph replacement
# ─────────────────────────────────────────────
def build_exact_regex(replacements):
    """Compile ONE regex matching any dictionary key (longest first). Built once
    per run instead of 600 separate regexes per paragraph — the big speed win."""
    keys = sorted(replacements.keys(), key=len, reverse=True)
    if not keys:
        return None
    alt = "|".join(re.escape(k) for k in keys)
    return re.compile(r'(?<![A-Za-z])(' + alt + r')(?![A-Za-z])', re.IGNORECASE)


def replace_in_paragraph(para, exact_re, replacements, engine, enable_engine):
    """Return list of change dicts: {orig, repl, tier, conf, reason}."""
    full_text = "".join(run.text for run in para.runs)
    if not full_text.strip():
        return []

    # 1) exact dictionary spans — single regex pass, matches are non-overlapping
    spans = []  # (start, end, replacement, tier, conf, reason)
    if exact_re is not None:
        for m in exact_re.finditer(full_text):
            rep = replacements.get(m.group(1).lower())
            if rep is not None:
                spans.append((m.start(), m.end(), rep, "exact", 1.0, "dictionary"))

    # 2) engine suggestions on remaining words
    if enable_engine and engine is not None:
        for m in re.finditer(r'[A-Za-z]+', full_text):
            if any(s < m.end() and m.start() < e for s, e, *_ in spans):
                continue
            sug = engine.suggest(m.group(0))
            if sug is None or sug.confidence < SUGGEST_FLOOR:
                continue
            tier = "fix" if sug.confidence >= FIX_CUTOFF else "suggest"
            spans.append((m.start(), m.end(), sug.correction, tier, sug.confidence, sug.reason))

    if not spans:
        return []

    spans.sort(key=lambda x: x[0])
    changes = [
        {"orig": full_text[s:e], "repl": rep, "tier": tier, "conf": conf, "reason": reason}
        for s, e, rep, tier, conf, reason in spans
    ]

    # rebuild the paragraph
    first_run_xml = para.runs[0]._r if para.runs else None
    p_xml = para._p
    for r in p_xml.findall(qn('w:r')):
        p_xml.remove(r)

    pos = 0
    for (start, end, replacement, tier, conf, reason) in spans:
        if pos < start:
            p_xml.append(build_normal_run(full_text[pos:start], first_run_xml))
        p_xml.append(build_red_run(full_text[start:end], first_run_xml))
        p_xml.append(build_normal_run(' ', first_run_xml))
        p_xml.append(build_highlight_run(replacement, TIER_HL[tier], first_run_xml))
        pos = end
    if pos < len(full_text):
        p_xml.append(build_normal_run(full_text[pos:], first_run_xml))

    return changes


def process_document(dict_bytes, input_bytes, enable_engine, progress_cb=None):
    """progress_cb(done, total, stage_msg) is called as work proceeds (optional)."""
    t_start = time.perf_counter()

    def report(done, total, msg):
        if progress_cb:
            progress_cb(done, total, msg)

    report(0, 1, "Reading dictionary…")
    replacements, pairs = load_dictionary(dict_bytes)
    exact_re = build_exact_regex(replacements)

    report(0, 1, "Training pattern engine…")
    engine = SuggestEngine(pairs) if enable_engine else None

    report(0, 1, "Opening subtitle file…")
    doc = Document(io.BytesIO(input_bytes))

    # gather every paragraph up-front so we know the total for the progress bar
    paras = list(doc.paragraphs)
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                paras.extend(cell.paragraphs)
    total = len(paras)
    log.info("Scanning %d paragraphs (engine=%s)…", total, "on" if enable_engine else "off")

    all_changes = []
    t_scan = time.perf_counter()
    for i, para in enumerate(paras, 1):
        all_changes.extend(replace_in_paragraph(para, exact_re, replacements, engine, enable_engine))
        if i % 200 == 0 or i == total:
            log.info("  …%d/%d paragraphs, %d changes so far", i, total, len(all_changes))
            report(i, total, f"Scanning paragraphs… {i}/{total}")
    log.info("Scan done: %d changes in %.2fs", len(all_changes), time.perf_counter( ) - t_scan)

    if engine is not None:
        log.info("Engine cache: %d unique words looked up", len(engine._cache))

    report(total, total, "Saving fixed file…")
    out = io.BytesIO()
    doc.save(out)
    out.seek(0)
    log.info("process_document finished in %.2fs total", time.perf_counter() - t_start)
    return out, replacements, all_changes


def build_dictionary_export(changes):
    """Make a 3-column docx (same format) of engine-found pairs to paste back."""
    seen, rows = set(), []
    for ch in changes:
        if ch["tier"] == "exact":
            continue
        key = (ch["orig"].lower(), ch["repl"])
        if key in seen:
            continue
        seen.add(key)
        rows.append((ch["orig"], ch["repl"]))

    doc = Document()
    doc.add_heading("Engine-found corrections (review before merging)", level=1)
    table = doc.add_table(rows=1, cols=3)
    table.style = "Table Grid"
    hdr = table.rows[0].cells
    hdr[0].text, hdr[1].text, hdr[2].text = "", "AI Version", "Planet Read Version"
    for i, (a, b) in enumerate(rows, 1):
        cells = table.add_row().cells
        cells[0].text, cells[1].text, cells[2].text = str(i), a, b
    out = io.BytesIO()
    doc.save(out)
    out.seek(0)
    return out, len(rows)


# regex for an SRT timecode line: 00:00:01,200 --> 00:00:04,000
TIMECODE_RE = re.compile(
    r'\d{1,2}:\d{2}:\d{2}[,.]\d{3}\s*-->\s*\d{1,2}:\d{2}:\d{2}[,.]\d{3}'
)


def build_srt(input_bytes, replacements):
    """Build a clean .srt from the subtitle Word file.

    - applies ONLY the confirmed dictionary (yellow) corrections — wrong word is
      replaced by the correct word, with NO red/green/cyan markup at all
    - engine guesses are NOT applied (original word kept) — they're unconfirmed
    - timecodes / block numbers already present in the Word text are preserved

    Returns (srt_bytes, num_timecodes, num_corrections).
    """
    exact_re = build_exact_regex(replacements)
    n_fixes = 0

    def fix_line(text):
        nonlocal n_fixes
        if exact_re is None:
            return text

        def repl(m):
            nonlocal n_fixes
            rep = replacements.get(m.group(1).lower())
            if rep is None:
                return m.group(0)
            n_fixes += 1
            return rep
        return exact_re.sub(repl, text)

    doc = Document(io.BytesIO(input_bytes))
    lines = [fix_line(p.text) for p in doc.paragraphs]
    srt_text = "\n".join(lines).strip() + "\n"
    n_codes = len(TIMECODE_RE.findall(srt_text))
    log.info("SRT built: %d timecodes, %d dictionary corrections applied", n_codes, n_fixes)
    return srt_text.encode("utf-8"), n_codes, n_fixes


# ─────────────────────────────────────────────
# Streamlit UI
# ─────────────────────────────────────────────
st.set_page_config(page_title="PlanetRead – Romanized Docx Fixer", page_icon="📖", layout="wide")

st.markdown("""
    <h1 style='color:#E05C00;'>📖 PlanetRead – Romanized Docx Fixer</h1>
    <p style='color:#555; font-size:16px;'>
        Upload the <b>Universal Dictionary</b> and a <b>Romanized subtitle docx</b>.
    </p>
    <p style='font-size:14px;'>
      <span style='background:yellow;padding:2px 8px;border-radius:4px;'>yellow</span>
      = confirmed dictionary fix &nbsp;·&nbsp;
      <span style='background:#9bffb0;padding:2px 8px;border-radius:4px;'>green</span>
      = engine fix (please confirm) &nbsp;·&nbsp;
      <span style='background:#9bf6ff;padding:2px 8px;border-radius:4px;'>cyan</span>
      = engine guess (review carefully)
    </p>
    <hr>
""", unsafe_allow_html=True)

enable_engine = st.sidebar.checkbox("🧠 Enable AI pattern engine", value=True)
st.sidebar.caption(
    "Learns spelling-correction patterns from your dictionary and applies them to "
    "words **not** in the table. Offline, no internet. ~75–80% precise on novel words, "
    "so its fixes are flagged for your review, never silently applied."
)

col1, col2 = st.columns(2)
with col1:
    st.subheader("📚 Step 1 — Upload Dictionary")
    dict_file = st.file_uploader("Universal Dictionary (.docx)", type=["docx"], key="dict")
with col2:
    st.subheader("🎬 Step 2 — Upload Subtitle File")
    input_file = st.file_uploader("Romanized Subtitle (.docx)", type=["docx"], key="input")

st.markdown("<br>", unsafe_allow_html=True)

if dict_file and input_file:
    if st.button("🚀 Run Replacements", type="primary", use_container_width=True):
        progress = st.progress(0.0, text="Starting…")
        status = st.empty()
        t_ui = time.perf_counter()

        def on_progress(done, total, msg):
            frac = min(1.0, done / total) if total else 0.0
            progress.progress(frac, text=msg)
            status.caption(f"⏱️ {time.perf_counter() - t_ui:.1f}s — {msg}")

        try:
                dict_bytes = dict_file.read()
                input_bytes = input_file.read()
                output_bytes, replacements, all_changes = process_document(
                    dict_bytes, input_bytes, enable_engine,
                    progress_cb=on_progress,
                )
                progress.progress(1.0, text="Done!")

                by_tier = {"exact": [], "fix": [], "suggest": []}
                for ch in all_changes:
                    by_tier[ch["tier"]].append(ch)
                total = len(all_changes)

                st.success(f"✅ Done! **{total} change(s)** across the document.")

                m1, m2, m3, m4 = st.columns(4)
                m1.metric("Dictionary Pairs", len(replacements))
                m2.metric("✅ Dictionary Fixes", len(by_tier["exact"]))
                m3.metric("🟢 Engine Fixes", len(by_tier["fix"]))
                m4.metric("🔵 Engine Guesses", len(by_tier["suggest"]))

                def render_tier(title, items, color, note):
                    if not items:
                        return
                    st.markdown(f"### {title}")
                    if note:
                        st.caption(note)
                    counts = Counter((c["orig"], c["repl"]) for c in items)
                    confs = {(c["orig"], c["repl"]): c["conf"] for c in items}
                    rows = sorted(counts.items(), key=lambda x: -x[1])
                    h1, h2, h3, h4 = st.columns([3, 3, 1, 1])
                    h1.markdown("<b>Wrong (in file)</b>", unsafe_allow_html=True)
                    h2.markdown("<b>Correction</b>", unsafe_allow_html=True)
                    h3.markdown("<b>Conf.</b>", unsafe_allow_html=True)
                    h4.markdown("<b>Count</b>", unsafe_allow_html=True)
                    st.divider()
                    for (orig, corr), cnt in rows:
                        c1, c2, c3, c4 = st.columns([3, 3, 1, 1])
                        c1.markdown(f"<span style='color:red;font-weight:bold'>{orig}</span>", unsafe_allow_html=True)
                        c2.markdown(f"<span style='background:{color};padding:2px 10px;border-radius:4px;font-weight:bold'>{corr}</span>", unsafe_allow_html=True)
                        cf = confs[(orig, corr)]
                        c3.markdown("—" if cf >= 1.0 else f"{cf*100:.0f}%")
                        c4.markdown(f"**{cnt}**")

                render_tier("✅ Confirmed Dictionary Fixes", by_tier["exact"], "yellow",
                            "Exact matches from your dictionary.")
                render_tier("🟢 Engine Fixes — please confirm", by_tier["fix"], "#9bffb0",
                            "Learned patterns, higher confidence. Review and accept in Word.")
                render_tier("🔵 Engine Guesses — review carefully", by_tier["suggest"], "#9bf6ff",
                            "Lower confidence. The engine is unsure — check each before accepting.")

                st.markdown("<br>", unsafe_allow_html=True)
                base = input_file.name.rsplit(".docx", 1)[0]
                d1, d2, d3 = st.columns(3)
                with d1:
                    st.markdown("#### 📥 Fixed Word File")
                    out_name = base + "_fixed.docx"
                    st.download_button(
                        f"⬇️ {out_name}", data=output_bytes, file_name=out_name,
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                        use_container_width=True, type="primary",
                    )
                    st.caption("Review file with red/yellow/green/cyan markup.")
                with d2:
                    st.markdown("#### 🎞️ Clean SRT")
                    srt_bytes, n_codes, n_fixes = build_srt(input_bytes, replacements)
                    srt_name = base + ".srt"
                    st.download_button(
                        f"⬇️ {srt_name}", data=srt_bytes, file_name=srt_name,
                        mime="application/x-subrip",
                        use_container_width=True, type="primary",
                    )
                    if n_codes == 0:
                        st.warning("⚠️ No timecodes found in the file — the .srt may "
                                   "not have valid timings.")
                    else:
                        st.caption(f"Plain subtitles, {n_codes} timecodes kept. "
                                   f"Only confirmed dictionary fixes applied ({n_fixes}); "
                                   "no markup, no engine guesses.")
                with d3:
                    st.markdown("#### 🔁 Grow the Dictionary")
                    exp_bytes, n_new = build_dictionary_export(all_changes)
                    st.download_button(
                        f"⬇️ Engine-found pairs ({n_new})",
                        data=exp_bytes, file_name="engine_found_pairs.docx",
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                        use_container_width=True, disabled=(n_new == 0),
                    )
                    st.caption("Review, delete wrong rows, paste into the master dictionary.")

        except Exception as e:
                log.exception("Processing failed")
                st.error(f"Something went wrong: {e}")
                st.exception(e)

elif not dict_file and not input_file:
    st.info("👆 Upload both files above to get started.")
elif not dict_file:
    st.warning("Please upload the Dictionary file.")
else:
    st.warning("Please upload the Subtitle file.")

st.markdown("<br><hr><p style='text-align:center; color:#aaa; font-size:13px;'>PlanetRead · Romanized Shabdkosh Tool</p>", unsafe_allow_html=True)
