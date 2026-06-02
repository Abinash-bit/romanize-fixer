import streamlit as st
import re
import copy
import io
from docx import Document
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from collections import Counter

# ─────────────────────────────────────────────
# Core logic
# ─────────────────────────────────────────────

def load_dictionary(file_bytes: bytes) -> dict:
    doc = Document(io.BytesIO(file_bytes))
    replacements = {}
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
            replacements[ai_word.lower()] = pr_word
    return replacements


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


def build_red_run(text, source_run_xml=None):
    r = OxmlElement('w:r')
    rPr = _base_rPr(source_run_xml)
    for tag in ('w:color', 'w:strike', 'w:highlight'):
        for el in rPr.findall(qn(tag)):
            rPr.remove(el)
    color = OxmlElement('w:color')
    color.set(qn('w:val'), 'FF0000')
    rPr.append(color)
    r.append(rPr)
    r.append(_make_t(text))
    return r


def build_yellow_highlight_run(text, source_run_xml=None):
    r = OxmlElement('w:r')
    rPr = _base_rPr(source_run_xml)
    for tag in ('w:color', 'w:strike', 'w:highlight'):
        for el in rPr.findall(qn(tag)):
            rPr.remove(el)
    hl = OxmlElement('w:highlight')
    hl.set(qn('w:val'), 'yellow')
    rPr.append(hl)
    r.append(rPr)
    r.append(_make_t(text))
    return r


def build_normal_run(text, source_run_xml=None):
    r = OxmlElement('w:r')
    r.append(_base_rPr(source_run_xml))
    r.append(_make_t(text))
    return r


def replace_in_paragraph(para, replacements: dict):
    full_text = "".join(run.text for run in para.runs)
    if not full_text.strip():
        return 0, []

    sorted_keys = sorted(replacements.keys(), key=len, reverse=True)
    spans = []
    for key in sorted_keys:
        pattern = r'(?<![A-Za-z])' + re.escape(key) + r'(?![A-Za-z])'
        for m in re.finditer(pattern, full_text, flags=re.IGNORECASE):
            if not any(s < m.end() and m.start() < e for s, e, _ in spans):
                spans.append((m.start(), m.end(), replacements[key]))

    if not spans:
        return 0, []

    spans.sort(key=lambda x: x[0])
    changed = [(full_text[s:e], r) for s, e, r in spans]

    first_run_xml = para.runs[0]._r if para.runs else None
    p_xml = para._p
    for r in p_xml.findall(qn('w:r')):
        p_xml.remove(r)

    pos = 0
    for (start, end, replacement) in spans:
        if pos < start:
            p_xml.append(build_normal_run(full_text[pos:start], first_run_xml))
        p_xml.append(build_red_run(full_text[start:end], first_run_xml))
        p_xml.append(build_normal_run(' ', first_run_xml))
        p_xml.append(build_yellow_highlight_run(replacement, first_run_xml))
        pos = end

    if pos < len(full_text):
        p_xml.append(build_normal_run(full_text[pos:], first_run_xml))

    return len(spans), changed


def process_document(dict_bytes: bytes, input_bytes: bytes):
    replacements = load_dictionary(dict_bytes)
    doc = Document(io.BytesIO(input_bytes))
    total, all_changes = 0, []
    for para in doc.paragraphs:
        n, ch = replace_in_paragraph(para, replacements)
        total += n; all_changes.extend(ch)
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for para in cell.paragraphs:
                    n, ch = replace_in_paragraph(para, replacements)
                    total += n; all_changes.extend(ch)
    out = io.BytesIO()
    doc.save(out)
    out.seek(0)
    return out, total, replacements, all_changes


# ─────────────────────────────────────────────
# Streamlit UI
# ─────────────────────────────────────────────

st.set_page_config(
    page_title="PlanetRead – Romanized Docx Fixer",
    page_icon="📖",
    layout="wide"
)

st.markdown("""
    <h1 style='color:#E05C00;'>📖 PlanetRead – Romanized Docx Fixer</h1>
    <p style='color:#555; font-size:16px;'>
        Upload the <b>Universal Dictionary</b> and a <b>Romanized subtitle docx</b>.<br>
        Each correction will show:
        <span style='color:red; font-weight:bold;'>wrong word</span>
        &nbsp;
        <span style='background:yellow; padding:2px 6px;'>correct word</span>
    </p>
    <hr>
""", unsafe_allow_html=True)

col1, col2 = st.columns(2)
with col1:
    st.subheader("📚 Step 1 — Upload Dictionary")
    dict_file = st.file_uploader("Universal Dictionary (.docx)", type=["docx"], key="dict")
with col2:
    st.subheader("🎬 Step 2 — Upload Subtitle File")
    input_file = st.file_uploader("Romanized Subtitle (.docx)", type=["docx"], key="input")

st.markdown("<br>", unsafe_allow_html=True)

if dict_file and input_file:
    with st.expander("👁️ Preview Dictionary (first 20 pairs)", expanded=False):
        d = load_dictionary(dict_file.read())
        dict_file.seek(0)
        items = list(d.items())[:20]
        st.table({"AI Version": [k for k, v in items], "PlanetRead Version": [v for k, v in items]})
        st.caption(f"Total pairs loaded: **{len(d)}**")

    st.markdown("<br>", unsafe_allow_html=True)

    if st.button("🚀 Run Replacements", type="primary", use_container_width=True):
        with st.spinner("Processing document…"):
            try:
                output_bytes, total, replacements, all_changes = process_document(
                    dict_file.read(), input_file.read()
                )

                st.success(f"✅ Done! **{total} replacement(s)** made across the document.")

                m1, m2, m3 = st.columns(3)
                m1.metric("Dictionary Pairs", len(replacements))
                m2.metric("Replacements Made", total)
                m3.metric("Unique Words Changed", len(set(o for o, _ in all_changes)))

                if all_changes:
                    st.markdown("### 🔄 Changes Made")

                    change_counts = Counter((o, r) for o, r in all_changes)
                    rows = sorted(change_counts.items(), key=lambda x: -x[1])

                    # Header row
                    h1, h2, h3, h4 = st.columns([0.5, 3, 3, 1])
                    h1.markdown("<span style='color:#888; font-size:13px'>#</span>", unsafe_allow_html=True)
                    h2.markdown("<b>AI Version (Wrong)</b>", unsafe_allow_html=True)
                    h3.markdown("<b>PlanetRead Version (Fixed)</b>", unsafe_allow_html=True)
                    h4.markdown("<b>Count</b>", unsafe_allow_html=True)
                    st.divider()

                    for i, ((orig, corr), cnt) in enumerate(rows, 1):
                        c1, c2, c3, c4 = st.columns([0.5, 3, 3, 1])
                        c1.markdown(f"<span style='color:#aaa; font-size:13px'>{i}</span>", unsafe_allow_html=True)
                        c2.markdown(f"<span style='color:red; font-weight:bold; font-size:15px'>{orig}</span>", unsafe_allow_html=True)
                        c3.markdown(f"<span style='background:yellow; padding:2px 10px; border-radius:4px; font-weight:bold; font-size:15px'>{corr}</span>", unsafe_allow_html=True)
                        c4.markdown(f"**{cnt}**")

                st.markdown("<br>", unsafe_allow_html=True)
                st.markdown("### 📥 Download Fixed File")
                out_name = input_file.name.replace(".docx", "_fixed.docx")
                st.download_button(
                    label=f"⬇️ Download  {out_name}",
                    data=output_bytes,
                    file_name=out_name,
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    use_container_width=True,
                    type="primary"
                )

            except Exception as e:
                st.error(f"Something went wrong: {e}")
                st.exception(e)

elif not dict_file and not input_file:
    st.info("👆 Upload both files above to get started.")
elif not dict_file:
    st.warning("Please upload the Dictionary file.")
else:
    st.warning("Please upload the Subtitle file.")

st.markdown("<br><hr><p style='text-align:center; color:#aaa; font-size:13px;'>PlanetRead · Romanized Shabdkosh Tool</p>", unsafe_allow_html=True)
