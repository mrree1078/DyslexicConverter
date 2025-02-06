import streamlit as st
from docx import Document
import fitz  # PyMuPDF
import re
import io
from reportlab.lib.pagesizes import letter
from reportlab.lib.styles import ParagraphStyle
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer
from reportlab.lib.enums import TA_LEFT
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.lib.colors import HexColor

# Initialize session state
if 'settings' not in st.session_state:
    st.session_state.settings = {
        'text_size': 12,
        'line_spacing': 2.0,
        'char_spacing': 1,
        'bg_color': '#FFFFFF',
        'text_color': '#000000',
        'current_font': 'Arial',
    }

COLOR_FILTERS = {
    'Default': ('#000000', '#FFFFFF'),
    'Cream': ('#000000', '#FFFFEA'),
    'Soft Pink': ('#003366', '#E6F0FF'),
    'Dark Mode': ('#FFFFFF', '#1A1A1A')
}

FONT_OPTIONS = ['Arial']

def configure_fonts():
    """Register fonts with fallback handling"""
    try:
        pdfmetrics.registerFont(TTFont('Arial', 'arial.ttf'))
        pdfmetrics.registerFont(TTFont('Arial-Bold', 'arialbd.ttf'))
    except:
        st.warning("Some fonts not found - using defaults")

def process_word(word):
    """Process words with current settings - bolding first half"""
    clean_word = re.sub(r'^\W+|\W+$', '', word)
    if not clean_word:
        return (word, '')

    split = (len(clean_word) + 1) // 2
    bold = clean_word[:split]
    normal = clean_word[split:] + word[len(clean_word):]
    return (bold, normal)

def create_pdf(text, settings):
    """PDF generation with character spacing, no margins"""
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(
        buffer,
        pagesize=letter,
    )

    style = ParagraphStyle(
        'MainStyle',
        fontName=settings['current_font'],
        fontSize=settings['text_size'],
        leading=settings['text_size'] * settings['line_spacing'],
        textColor=HexColor(settings['text_color']),
        backColor=HexColor(settings['bg_color']),
        alignment=TA_LEFT,
        splitLongWords=False,
        spaceBefore=settings['text_size'] * 0.5,
        charSpace=settings['char_spacing'] # Character spacing setting is here
    )

    content = []
    paragraphs = text.split("\n\n")

    for para_text in paragraphs:
        processed_words = []
        words = re.findall(r'\S+|\s+', para_text)
        for word in words:
            if word.strip() == '':
                processed_words.append(word)
            else:
                bold, normal = process_word(word)
                processed_words.append(f'<font name="{settings["current_font"]}-Bold">{bold}</font>{normal}')

        paragraph_content = ''.join(processed_words)
        content.append(Paragraph(paragraph_content, style))
        content.append(Spacer(1, 6))

    def add_background(canvas, doc):
        canvas.saveState()
        canvas.setFillColor(HexColor(settings['bg_color']))
        canvas.rect(0, 0, doc.width, doc.height, fill=1)
        canvas.restoreState()

    doc.build(content, onFirstPage=add_background, onLaterPages=add_background)
    buffer.seek(0)
    return buffer.getvalue()


def read_file(uploaded_file):
    """Reads text from DOCX or PDF files."""
    file_type = uploaded_file.type
    if file_type == "application/vnd.openxmlformats-officedocument.wordprocessingml.document": # DOCX
        try:
            document = Document(uploaded_file)
            text = ""
            for paragraph in document.paragraphs:
                text += paragraph.text + "\n\n"
            return text
        except Exception as e:
            st.error(f"Error reading DOCX file: {e}")
            return None
    elif file_type == "application/pdf": # PDF
        try:
            pdf_document = fitz.open(stream=uploaded_file.read(), filetype="pdf")
            text = ""
            for page in pdf_document:
                text += page.get_text() + "\n\n"
            return text
        except Exception as e:
            st.error(f"Error reading PDF file: {e}")
            return None
    else:
        st.error("Unsupported file type. Please upload DOCX or PDF.")
        return None

def process_text(text):
    """Processes the entire text into bold and normal word parts for preview."""
    processed_text = []
    words = re.findall(r'\S+|\s+', text)
    for word in words:
        if word.strip() == '':
            processed_text.append(("", word))
        else:
            processed_text.append(process_word(word))
    return processed_text


# Streamlit UI
st.title("Dyslexic Document Converter")
configure_fonts()

# Sidebar controls
with st.sidebar:
    st.header("Document Settings")

    st.session_state.settings['current_font'] = st.selectbox(
        "Font Style",
        FONT_OPTIONS,
        index=FONT_OPTIONS.index(st.session_state.settings['current_font'])
    )

    color_preset = st.selectbox("Color Theme", list(COLOR_FILTERS.keys()))
    text_color, bg_color = COLOR_FILTERS[color_preset]
    st.session_state.settings.update({
        'text_color': text_color,
        'bg_color': bg_color
    })

    st.session_state.settings['text_size'] = st.slider("Text Size (pt)", 8, 24, 12)
    st.session_state.settings['line_spacing'] = st.slider("Line Spacing", 1.0, 3.0, 2.0, 0.25)
    st.session_state.settings['char_spacing'] = st.slider("Character Spacing (pt)", 0, 5, 1)

# Main interface
uploaded_file = st.file_uploader("Upload DOCX/PDF", type=['docx', 'pdf'])

if uploaded_file:
    text = read_file(uploaded_file)
    if text:
        preview_content = []
        processed = process_text(text)

        for bold, normal in processed:
            preview_content.append(
                f'<span style="font-family: {st.session_state.settings["current_font"]}; '
                f'font-size: {st.session_state.settings["text_size"]}pt; '
                f'line-height: {st.session_state.settings["line_spacing"]}; '
                f'letter-spacing: {st.session_state.settings["char_spacing"]}pt;'
                f'color: {st.session_state.settings["text_color"]};">'
                f'<b>{bold}</b>{normal}</span>'
            )

        st.markdown(
            f'<div style="background:{st.session_state.settings["bg_color"]};padding:20px;border-radius:10px;">'
            f'{"".join(preview_content)}</div>',
            unsafe_allow_html=True
        )

        # PDF is generated automatically whenever settings or text changes
        st.session_state.pdf_bytes = create_pdf(text, st.session_state.settings)

        if 'pdf_bytes' in st.session_state:
            st.sidebar.download_button(
                "Download PDF",
                st.session_state.pdf_bytes,
                "document.pdf",
                "application/pdf"
            )

# Instructions sidebar
with st.sidebar:
    st.markdown("---")
    st.markdown("**Instructions**")
    st.markdown("1. Upload document\n2. Adjust settings\n3. Preview updates automatically\n4. Download PDF")