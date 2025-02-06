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

# --- Initialize session state ---
if 'settings' not in st.session_state:
    st.session_state.settings = {
        'text_size': 12,
        'line_spacing': 2.0,
        'char_spacing': 1,
        'bg_color': '#FFFFFF',
        'text_color': '#000000',
        'current_font': 'Arial',
    }
if 'pdf_bytes' not in st.session_state:
    st.session_state.pdf_bytes = None
if 'error_message' not in st.session_state:
    st.session_state.error_message = None

COLOR_FILTERS = {
    'Default': ('#000000', '#FFFFFF'),
    'Cream': ('#000000', '#FFFFEA'),
    'Soft Blue': ('#003366', '#E6F0FF'),
    'Dark Mode': ('#FFFFFF', '#1A1A1A')
}

FONT_OPTIONS = ['Arial']  # Arial only as per user's request - can be expanded

# --- Font Configuration ---
def configure_fonts_app():
    try:
        pdfmetrics.registerFont(TTFont('Arial', 'arial.ttf'))
        pdfmetrics.registerFont(TTFont('Arial-Bold', 'arialbd.ttf'))
    except Exception as e:
        st.warning(f"Font configuration issue: {e}. Using default fonts.")

# --- Text Processing ---
def process_word_app(word):
    clean_word = re.sub(r'^\W+|\W+$', '', word)
    if not clean_word:
        return (word, '')
    split = (len(clean_word) + 1) // 2
    bold = clean_word[:split]
    normal = clean_word[split:] + word[len(clean_word):]
    return (bold, normal)

def process_text_for_preview(text):
    processed_text = []
    words = re.findall(r'\S+|\s+', text)
    for word in words:
        if word.strip() == '':
            processed_text.append(("", word))
        else:
            processed_text.append(process_word_app(word))
    return processed_text

# --- PDF Generation ---
def create_pdf_document(text, settings):
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
        charSpace=settings['char_spacing']
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
                bold, normal = process_word_app(word)
                processed_words.append(f'<font name="{settings["current_font"]}-Bold">{bold}</font>{normal}')

        paragraph_content = ''.join(processed_words)
        content.append(Paragraph(paragraph_content, style))
        content.append(Spacer(1, 6))

    def add_background(canvas, doc):
        canvas.saveState()
        canvas.setFillColor(HexColor(settings['bg_color']))
        canvas.rect(0, 0, doc.width, doc.height, fill=1)
        canvas.restoreState()

    try:
        doc.build(content, onFirstPage=add_background, onLaterPages=add_background)
    except Exception as e:
        raise RuntimeError(f"Failed to generate PDF: {e}")

    buffer.seek(0)
    return buffer.getvalue()

def read_document_file(uploaded_file):
    """Reads text from DOCX or PDF files."""
    text = None
    error = None
    file_type = uploaded_file.type
    try:
        if file_type == "application/vnd.openxmlformats-officedocument.wordprocessingml.document":  # DOCX
            doc = Document(uploaded_file)
            text = "\n\n".join([paragraph.text for paragraph in doc.paragraphs])
        elif file_type == "application/pdf":  # PDF
            pdf_doc = fitz.open(stream=uploaded_file.read(), filetype="pdf")
            text = "\n\n".join([page.get_text() for page in pdf_doc])
        else:
            error = "Unsupported file type. Please upload DOCX or PDF."
    except Exception as e:
        error = f"Error reading file: {e}"
    return text, error

# --- Streamlit UI ---
st.title("Accessibility Document Converter")
configure_fonts_app()

# --- Sidebar Controls ---
with st.sidebar:
    st.header("Document Settings")
    st.session_state.settings['current_font'] = st.selectbox("Font Style", FONT_OPTIONS, index=FONT_OPTIONS.index(st.session_state.settings['current_font']))
    color_preset = st.selectbox("Color Theme", list(COLOR_FILTERS.keys()))
    text_color, bg_color = COLOR_FILTERS[color_preset]
    st.session_state.settings.update({'text_color': text_color, 'bg_color': bg_color})
    st.session_state.settings['text_size'] = st.slider("Text Size (pt)", 8, 24, st.session_state.settings['text_size'])
    st.session_state.settings['line_spacing'] = st.slider("Line Spacing", 1.0, 3.0, st.session_state.settings['line_spacing'], 0.25)
    st.session_state.settings['char_spacing'] = st.slider("Character Spacing (pt)", 0, 5, st.session_state.settings['char_spacing'])

    st.download_button(
        label="Download PDF",
        data=st.session_state.pdf_bytes if st.session_state.pdf_bytes else b'',
        file_name="document.pdf",
        mime="application/pdf",
        disabled=st.session_state.pdf_bytes is None
    )

    st.markdown("---")
    st.markdown("**Instructions**")
    st.markdown("1. Upload document\n2. Adjust settings\n3. Preview updates automatically\n4. Download PDF")

# --- Main Interface ---
uploaded_file = st.file_uploader("Upload DOCX/PDF", type=['docx', 'pdf'])

if uploaded_file:
    st.session_state.error_message = None  # Clear previous error
    text, read_error = read_document_file(uploaded_file)
    
    if read_error:
        st.session_state.error_message = read_error
        st.session_state.pdf_bytes = None
    elif not text.strip():
        st.session_state.error_message = "The document contains no text."
        st.session_state.pdf_bytes = None
    else:
        # Generate preview
        preview_content = []
        processed_preview = process_text_for_preview(text)
        for bold, normal in processed_preview:
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
        
        # Generate PDF
        try:
            st.session_state.pdf_bytes = create_pdf_document(text, st.session_state.settings)
        except Exception as e:
            st.session_state.error_message = str(e)
            st.session_state.pdf_bytes = None

    # Display error message if any
    if st.session_state.error_message:
        st.error(st.session_state.error_message)

else:
    st.session_state.pdf_bytes = None
    if st.session_state.error_message:
        st.error(st.session_state.error_message)
