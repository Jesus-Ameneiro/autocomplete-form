"""
Template Document Filler — Streamlit App (v2)
Two modes: Document View (live inline preview) and Form View.
Exports to DOCX and PDF.
"""

import re, io, html as html_mod, copy
import streamlit as st
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from lxml import etree

# ─── Constants ───────────────────────────────────────────────────────────────

PLACEHOLDER_RE = re.compile(r"\[([^\[\]]+)\]")
WD_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

MULTILINE_KEYWORDS = (
    "descripción", "description", "notas", "notes", "dirección", "address",
    "párrafo", "paragraph", "cuerpo", "body", "contenido", "content",
    "detalles", "details", "resumen", "summary", "comentarios", "comments",
)

# ─── Document Parsing Helpers ────────────────────────────────────────────────

def _is_list_paragraph(para):
    """Check if a paragraph is a list/bullet item via its XML."""
    ppr = para._element.find(f".//{{{WD_NS}}}pPr")
    if ppr is not None:
        return ppr.find(f".//{{{WD_NS}}}numPr") is not None
    return False


def _get_alignment(para):
    """Return CSS text-align for a paragraph."""
    a = para.alignment
    if a == WD_ALIGN_PARAGRAPH.RIGHT:
        return "right"
    elif a == WD_ALIGN_PARAGRAPH.CENTER:
        return "center"
    elif a == WD_ALIGN_PARAGRAPH.JUSTIFY:
        return "justify"
    return "left"


def _apply_run_fmt(html_fragment, run):
    """Wrap an HTML fragment with the run's bold / italic / size formatting."""
    if not run:
        return html_fragment
    is_bold = getattr(run, "bold", False)
    is_italic = getattr(run, "italic", False)
    font_size = None
    if run.font and run.font.size:
        font_size = round(run.font.size / 12700)  # EMU → pt
    if is_bold:
        html_fragment = f"<strong>{html_fragment}</strong>"
    if is_italic:
        html_fragment = f"<em>{html_fragment}</em>"
    if font_size and font_size != 12:
        html_fragment = f'<span style="font-size:{font_size}pt">{html_fragment}</span>'
    return html_fragment


def render_runs_html(runs, mode, values, lower_map):
    """
    Render a sequence of runs to HTML, correctly handling placeholders
    that span multiple runs (e.g. '[' in run 1, 'XX/XX' in run 2, ']' in run 3).
    """
    if not runs:
        return ""

    # 1. Build combined text and per-char → run index mapping
    full_text = ""
    char_run = []  # char index → run index
    for ri, run in enumerate(runs):
        for _ in run.text:
            char_run.append(ri)
        full_text += run.text

    if not full_text.strip():
        return ""

    # 2. Find placeholder spans in the combined text
    ph_spans = [(m.start(), m.end(), m.group(1).strip()) for m in PLACEHOLDER_RE.finditer(full_text)]

    # 3. Walk the text, emitting HTML for normal text and placeholders
    parts = []
    pos = 0

    def _emit_plain(start, end):
        """Emit plain (non-placeholder) text, respecting run boundaries."""
        if start >= end:
            return
        i = start
        while i < end:
            ri = char_run[i]
            # Find how far this run extends within [start, end)
            j = i + 1
            while j < end and char_run[j] == ri:
                j += 1
            fragment = html_mod.escape(full_text[i:j])
            parts.append(_apply_run_fmt(fragment, runs[ri]))
            i = j

    for ph_start, ph_end, ph_name in ph_spans:
        # Plain text before this placeholder
        _emit_plain(pos, ph_start)

        # Determine the replacement HTML for this placeholder
        key = ph_name.lower()
        display_name = lower_map.get(key, ph_name)
        val = values.get(display_name, "")
        # Use formatting from the run that contains the opening bracket
        fmt_run = runs[char_run[ph_start]]

        if mode == "blank":
            inner = html_mod.escape(full_text[ph_start:ph_end])
        elif val.strip():
            escaped_val = html_mod.escape(val)
            inner = (f'<span class="filled-value">{escaped_val}</span>'
                     if mode == "edit" else escaped_val)
        else:
            if mode == "edit":
                inner = f'<span class="empty-placeholder">{html_mod.escape(ph_name)}</span>'
            else:
                inner = html_mod.escape(full_text[ph_start:ph_end])

        parts.append(_apply_run_fmt(inner, fmt_run))
        pos = ph_end

    # Remaining text after the last placeholder
    _emit_plain(pos, len(full_text))

    return "".join(parts)


# ─── Parse Document into Structured Elements ────────────────────────────────

def parse_document(doc):
    """
    Walk the document body XML to get elements in order (paragraphs, tables,
    bullet items) since python-docx's doc.paragraphs skips over tables' ordering.
    """
    elements = []  # list of dicts
    body = doc.element.body

    # Map paragraph elements to python-docx Paragraph objects
    para_map = {}
    for p in doc.paragraphs:
        para_map[id(p._element)] = p

    # Map table elements
    table_map = {}
    for t in doc.tables:
        table_map[id(t._element)] = t

    current_list_items = []

    for child in body:
        tag = etree.QName(child.tag).localname
        if tag == "p":
            para = para_map.get(id(child))
            if para is None:
                continue
            if _is_list_paragraph(para):
                current_list_items.append(para)
            else:
                # Flush any accumulated list
                if current_list_items:
                    elements.append({"type": "list", "items": current_list_items})
                    current_list_items = []
                elements.append({"type": "paragraph", "para": para})
        elif tag == "tbl":
            if current_list_items:
                elements.append({"type": "list", "items": current_list_items})
                current_list_items = []
            tbl = table_map.get(id(child))
            if tbl:
                elements.append({"type": "table", "table": tbl})

    if current_list_items:
        elements.append({"type": "list", "items": current_list_items})

    return elements


# ─── Render Elements to HTML ─────────────────────────────────────────────────

def render_paragraph_html(para, mode, values, lower_map):
    """Render a single paragraph to HTML."""
    align = _get_alignment(para)
    runs = para.runs
    if not runs:
        return f'<p style="text-align:{align};margin:4px 0">&nbsp;</p>'

    inner = render_runs_html(runs, mode, values, lower_map)

    if not inner.strip():
        inner = "&nbsp;"

    return f'<p style="text-align:{align};margin:6px 0;line-height:1.5">{inner}</p>'


def render_table_html(table, mode, values, lower_map):
    """Render a table to HTML."""
    rows_html = []
    for ri, row in enumerate(table.rows):
        cells_html = []
        for cell in row.cells:
            cell_parts = []
            for para in cell.paragraphs:
                cell_parts.append(render_runs_html(para.runs, mode, values, lower_map))
            inner = " ".join(p for p in cell_parts if p)
            tag = "th" if ri == 0 else "td"
            cells_html.append(f"<{tag}>{inner or '&nbsp;'}</{tag}>")
        rows_html.append("<tr>" + "".join(cells_html) + "</tr>")
    return '<table class="doc-table">' + "".join(rows_html) + "</table>"


def render_list_html(items, mode, values, lower_map):
    """Render a list of bullet paragraphs."""
    lis = []
    for para in items:
        inner = render_runs_html(para.runs, mode, values, lower_map)
        lis.append(f"<li>{inner}</li>")
    return "<ul>" + "".join(lis) + "</ul>"


def render_document_html(doc, mode, values, lower_map):
    """
    Render the full document as styled HTML.
    mode: 'edit' | 'preview' | 'blank'
    """
    elements = parse_document(doc)
    parts = []
    for el in elements:
        if el["type"] == "paragraph":
            parts.append(render_paragraph_html(el["para"], mode, values, lower_map))
        elif el["type"] == "table":
            parts.append(render_table_html(el["table"], mode, values, lower_map))
        elif el["type"] == "list":
            parts.append(render_list_html(el["items"], mode, values, lower_map))

    body_html = "\n".join(parts)
    css = _get_display_css(mode)

    return f"""<!DOCTYPE html>
<html><head><meta charset="utf-8"><style>{css}</style></head>
<body><div class="doc-page">{body_html}</div></body></html>"""


def _get_display_css(mode):
    return """
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;600;700&display=swap');
    * { box-sizing: border-box; margin: 0; padding: 0; }
    body {
        font-family: 'Inter', 'Segoe UI', Arial, sans-serif;
        font-size: 11pt;
        color: #1a1a1a;
        background: transparent;
    }
    .doc-page {
        max-width: 720px;
        margin: 0 auto;
        padding: 32px 40px;
        background: white;
        border: 1px solid #d0d0d0;
        border-radius: 4px;
        box-shadow: 0 1px 4px rgba(0,0,0,0.06);
        min-height: 600px;
    }
    p { margin: 6px 0; line-height: 1.55; }
    ul { margin: 8px 0 8px 28px; }
    li { margin: 4px 0; line-height: 1.55; }
    .doc-table {
        width: 100%; border-collapse: collapse; margin: 12px 0; font-size: 10pt;
    }
    .doc-table th, .doc-table td {
        border: 1px solid #999; padding: 6px 10px; text-align: left;
    }
    .doc-table th {
        background: #f0f0f0; font-weight: 700;
    }
    .filled-value {
        background: #d4edda; color: #155724;
        padding: 1px 5px; border-radius: 3px;
        border-bottom: 2px solid #28a745;
        font-weight: 600;
    }
    .empty-placeholder {
        background: #fff3cd; color: #856404;
        padding: 1px 5px; border-radius: 3px;
        border: 1px dashed #ffc107;
        font-style: italic;
        font-size: 0.92em;
    }
    """



# ─── Placeholder Discovery ──────────────────────────────────────────────────

def iter_all_paragraphs(doc):
    yield from doc.paragraphs
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                yield from cell.paragraphs
                for nested in cell.tables:
                    for nr in nested.rows:
                        for nc in nr.cells:
                            yield from nc.paragraphs
    for section in doc.sections:
        for hf in (
            section.header, section.footer,
            section.first_page_header, section.first_page_footer,
            section.even_page_header, section.even_page_footer,
        ):
            if hf is not None and hf.is_linked_to_previous is False:
                yield from hf.paragraphs
                for table in hf.tables:
                    for row in table.rows:
                        for cell in row.cells:
                            yield from cell.paragraphs


def find_placeholders(doc):
    """Return (ordered_list, lower_map) where lower_map maps lowercase → display name."""
    seen_lower = {}
    order = []
    for para in iter_all_paragraphs(doc):
        full_text = "".join(run.text for run in para.runs)
        for match in PLACEHOLDER_RE.finditer(full_text):
            name = match.group(1).strip()
            key = name.lower()
            if key not in seen_lower:
                seen_lower[key] = name
                order.append(name)
    return order, seen_lower


# ─── DOCX Replacement (run-aware, formatting-preserving) ────────────────────

def _replace_in_paragraph(paragraph, repl_lower):
    runs = paragraph.runs
    if not runs:
        return
    char_origins = []
    for ri, run in enumerate(runs):
        for ci in range(len(run.text)):
            char_origins.append((ri, ci))
    full_text = "".join(r.text for r in runs)
    matches = []
    for m in PLACEHOLDER_RE.finditer(full_text):
        key = m.group(1).strip().lower()
        if key in repl_lower:
            matches.append((m.start(), m.end(), repl_lower[key]))
    if not matches:
        return
    for start, end, value in reversed(matches):
        first_ri, first_off = char_origins[start]
        last_ri, last_off = char_origins[end - 1]
        if first_ri == last_ri:
            r = runs[first_ri]
            r.text = r.text[:first_off] + value + r.text[last_off + 1:]
        else:
            runs[first_ri].text = runs[first_ri].text[:first_off] + value
            runs[last_ri].text = runs[last_ri].text[last_off + 1:]
            for mid in range(first_ri + 1, last_ri):
                runs[mid].text = ""


def apply_replacements(doc, replacements):
    repl_lower = {k.lower(): v for k, v in replacements.items()}
    for para in iter_all_paragraphs(doc):
        _replace_in_paragraph(para, repl_lower)


# ─── PDF Generation via fpdf2 (pure Python, no native deps) ─────────────────

def generate_pdf(doc, values, lower_map):
    """Generate PDF bytes from the document with filled values using fpdf2."""
    from fpdf import FPDF
    import os

    pdf = FPDF(orientation="P", unit="mm", format="Letter")
    pdf.set_auto_page_break(auto=True, margin=20)

    # Register a Unicode TTF font (DejaVu Sans ships with most Linux distros
    # and is available on Streamlit Cloud's Debian base image)
    font_dir = "/usr/share/fonts/truetype/dejavu"
    font_name = "DejaVu"
    if os.path.isfile(os.path.join(font_dir, "DejaVuSans.ttf")):
        pdf.add_font(font_name, "", os.path.join(font_dir, "DejaVuSans.ttf"))
        pdf.add_font(font_name, "B", os.path.join(font_dir, "DejaVuSans-Bold.ttf"))
        pdf.add_font(font_name, "I", os.path.join(font_dir, "DejaVuSans-Oblique.ttf"))
        pdf.add_font(font_name, "BI", os.path.join(font_dir, "DejaVuSans-BoldOblique.ttf"))
    else:
        # Fallback: Helvetica (ASCII only, accented chars will be lossy)
        font_name = "Helvetica"

    pdf.add_page()
    pdf.set_margins(18, 18, 18)
    pdf.set_font(font_name, size=11)

    def _resolve(text):
        """Replace [placeholder] tokens in a text string."""
        def _sub(m):
            key = m.group(1).strip().lower()
            display = lower_map.get(key, m.group(1).strip())
            val = values.get(display, "")
            return val if val.strip() else m.group(0)
        return PLACEHOLDER_RE.sub(_sub, text)

    elements = parse_document(doc)

    for el in elements:
        if el["type"] == "paragraph":
            para = el["para"]
            runs = para.runs
            full_text = "".join(r.text for r in runs)
            resolved = _resolve(full_text)

            if not resolved.strip():
                pdf.ln(4)
                continue

            # Determine alignment
            align = _get_alignment(para)
            align_map = {"left": "L", "right": "R", "center": "C", "justify": "J"}
            pdf_align = align_map.get(align, "L")

            # Determine font size from first run
            font_size = 11
            if runs and runs[0].font and runs[0].font.size:
                font_size = round(runs[0].font.size / 12700)

            # Check if entire paragraph is bold
            is_bold = all(r.bold for r in runs if r.text.strip()) if runs else False
            style = "B" if is_bold else ""

            pdf.set_font(font_name, style=style, size=font_size)
            pdf.multi_cell(0, font_size * 0.45, resolved, align=pdf_align)
            pdf.ln(1)

        elif el["type"] == "table":
            table = el["table"]
            num_cols = len(table.rows[0].cells) if table.rows else 0
            if num_cols == 0:
                continue
            col_width = (pdf.w - pdf.l_margin - pdf.r_margin) / num_cols

            for ri, row in enumerate(table.rows):
                is_header = ri == 0
                pdf.set_font(font_name, "B" if is_header else "", 9)
                row_height = 7
                for ci, cell in enumerate(row.cells):
                    cell_text = _resolve(cell.text)
                    x_before = pdf.get_x()
                    y_before = pdf.get_y()
                    pdf.rect(x_before, y_before, col_width, row_height)
                    if is_header:
                        pdf.set_fill_color(235, 235, 235)
                        pdf.rect(x_before, y_before, col_width, row_height, "F")
                    pdf.set_xy(x_before + 1, y_before + 1)
                    pdf.cell(col_width - 2, row_height - 2, cell_text, align="L")
                    pdf.set_xy(x_before + col_width, y_before)
                pdf.ln(row_height)
            pdf.ln(2)

        elif el["type"] == "list":
            for para in el["items"]:
                full_text = "".join(r.text for r in para.runs)
                resolved = _resolve(full_text)
                is_bold = any(r.bold for r in para.runs if r.text.strip())

                pdf.set_font(font_name, "", 11)
                pdf.cell(8, 5, "\u2022", align="R")  # bullet
                pdf.set_font(font_name, "B" if is_bold else "", 11)
                pdf.multi_cell(0, 5, " " + resolved, align="J")
                pdf.ln(1)

    buf = io.BytesIO()
    pdf.output(buf)
    buf.seek(0)
    return buf.getvalue()


# ─── Streamlit App ───────────────────────────────────────────────────────────

st.set_page_config(page_title="Template Filler", page_icon="📝", layout="wide")

st.markdown("""
<style>
    .block-container { max-width: 1200px; padding-top: 1.5rem; }
    div[data-testid="stFileUploader"] { margin-bottom: 0.5rem; }
    .placeholder-tag {
        display: inline-block; background: #e8f0fe; color: #1a56db;
        padding: 2px 8px; border-radius: 4px; font-family: monospace;
        font-size: 0.82em; margin: 2px;
    }
    .stTabs [data-baseweb="tab-list"] { gap: 4px; }
    .stTabs [data-baseweb="tab"] {
        padding: 8px 20px; font-weight: 600;
    }
    /* Compact inputs in sidebar */
    section[data-testid="stSidebar"] .stTextInput > div > div { font-size: 0.9em; }
</style>
""", unsafe_allow_html=True)

st.title("📝 Template Document Filler")
st.caption(
    "Upload a `.docx` template with `[placeholder]` fields. "
    "Fill in the values and export as DOCX or PDF."
)

uploaded_file = st.file_uploader("Upload your Word template", type=["docx"])

if uploaded_file is not None:
    file_bytes = uploaded_file.read()
    doc = Document(io.BytesIO(file_bytes))
    placeholders, lower_map = find_placeholders(doc)

    if not placeholders:
        st.warning("No `[placeholder]` fields found. Use square brackets, e.g. `[Name]`.")
        st.stop()

    # ── Initialize session state for values ──────────────────────────
    if "ph_values" not in st.session_state:
        st.session_state.ph_values = {ph: "" for ph in placeholders}
    else:
        # Ensure new placeholders from a different upload are covered
        for ph in placeholders:
            if ph not in st.session_state.ph_values:
                st.session_state.ph_values[ph] = ""

    # ── Tabs ─────────────────────────────────────────────────────────
    tab_doc, tab_form = st.tabs(["📄 Document View", "📋 Form View"])

    # ══════════════════════════════════════════════════════════════════
    # TAB 1 — Document View (live inline preview)
    # ══════════════════════════════════════════════════════════════════
    with tab_doc:
        left_col, right_col = st.columns([1, 2], gap="large")

        # ── Left: Input Fields ───────────────────────────────────────
        with left_col:
            st.markdown("#### Fields")
            st.caption(f"{len(placeholders)} placeholder(s) detected")

            for ph in placeholders:
                is_multiline = any(kw in ph.lower() for kw in MULTILINE_KEYWORDS)
                if is_multiline:
                    st.session_state.ph_values[ph] = st.text_area(
                        ph, value=st.session_state.ph_values[ph],
                        key=f"doc_{ph}", height=80,
                    )
                else:
                    st.session_state.ph_values[ph] = st.text_input(
                        ph, value=st.session_state.ph_values[ph],
                        key=f"doc_{ph}",
                    )

        # ── Right: Document Preview ──────────────────────────────────
        with right_col:
            preview_mode = st.toggle("Preview mode", value=False, key="preview_toggle",
                                     help="Toggle to see the clean final version")
            mode = "preview" if preview_mode else "edit"

            doc_html = render_document_html(
                doc, mode, st.session_state.ph_values, lower_map
            )

            st.components.v1.html(doc_html, height=900, scrolling=True)

        # ── Download Section ─────────────────────────────────────────
        st.markdown("---")
        st.markdown("#### Download")
        dl1, dl2, _ = st.columns([1, 1, 2])

        replacements = {k: v for k, v in st.session_state.ph_values.items() if v.strip()}

        with dl1:
            # DOCX
            fresh_doc = Document(io.BytesIO(file_bytes))
            apply_replacements(fresh_doc, replacements)
            buf = io.BytesIO()
            fresh_doc.save(buf)
            st.download_button(
                "⬇️ Download DOCX",
                data=buf.getvalue(),
                file_name=uploaded_file.name.replace(".docx", "_filled.docx"),
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                use_container_width=True,
            )

        with dl2:
            # PDF
            pdf_bytes = generate_pdf(doc, st.session_state.ph_values, lower_map)
            st.download_button(
                "⬇️ Download PDF",
                data=pdf_bytes,
                file_name=uploaded_file.name.replace(".docx", "_filled.pdf"),
                mime="application/pdf",
                use_container_width=True,
            )

    # ══════════════════════════════════════════════════════════════════
    # TAB 2 — Form View (classic form + generate)
    # ══════════════════════════════════════════════════════════════════
    with tab_form:
        st.markdown("#### Detected placeholders")
        tags_html = " ".join(
            f'<span class="placeholder-tag">[{ph}]</span>' for ph in placeholders
        )
        st.markdown(tags_html, unsafe_allow_html=True)
        st.markdown("")

        st.subheader("Fill in the fields")

        with st.form("template_form"):
            form_values = {}
            for ph in placeholders:
                is_multiline = any(kw in ph.lower() for kw in MULTILINE_KEYWORDS)
                default = st.session_state.ph_values.get(ph, "")
                if is_multiline:
                    form_values[ph] = st.text_area(
                        ph, value=default, key=f"form_{ph}", height=80
                    )
                else:
                    form_values[ph] = st.text_input(
                        ph, value=default, key=f"form_{ph}"
                    )

            submitted = st.form_submit_button(
                "Generate Document", type="primary", use_container_width=True
            )

        if submitted:
            # Sync values to session state so Document View picks them up
            for ph in placeholders:
                st.session_state.ph_values[ph] = form_values[ph]

            empty = [k for k, v in form_values.items() if not v.strip()]
            if empty:
                st.warning(
                    f"**{len(empty)}** field(s) left empty: "
                    + ", ".join(f"`[{f}]`" for f in empty)
                )

            replacements = {k: v for k, v in form_values.items() if v.strip()}

            # ── Preview ──────────────────────────────────────────────
            with st.expander("📖 Preview", expanded=True):
                preview_html = render_document_html(
                    doc, "preview", form_values, lower_map
                )
                st.components.v1.html(preview_html, height=800, scrolling=True)

            # ── Downloads ────────────────────────────────────────────
            st.markdown("#### Download")
            c1, c2, _ = st.columns([1, 1, 2])

            with c1:
                fresh_doc = Document(io.BytesIO(file_bytes))
                apply_replacements(fresh_doc, replacements)
                buf = io.BytesIO()
                fresh_doc.save(buf)
                st.download_button(
                    "⬇️ Download DOCX",
                    data=buf.getvalue(),
                    file_name=uploaded_file.name.replace(".docx", "_filled.docx"),
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    use_container_width=True,
                    key="form_dl_docx",
                )
            with c2:
                pdf_bytes = generate_pdf(doc, form_values, lower_map)
                st.download_button(
                    "⬇️ Download PDF",
                    data=pdf_bytes,
                    file_name=uploaded_file.name.replace(".docx", "_filled.pdf"),
                    mime="application/pdf",
                    use_container_width=True,
                    key="form_dl_pdf",
                )

else:
    st.markdown("---")
    st.markdown("**How it works**")
    st.markdown(
        "1. Create a Word document with placeholders like "
        "`[Client Name]`, `[Invoice Date]`, `[Total Amount]`.\n"
        "2. Upload it here.\n"
        "3. Use the **Document View** tab to fill fields while seeing the live preview, "
        "or the **Form View** tab for a traditional form.\n"
        "4. Toggle **Preview mode** to see the clean result.\n"
        "5. Download as **DOCX** or **PDF**."
    )
