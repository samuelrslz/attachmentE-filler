import streamlit as st
from docx import Document
import io
from datetime import date
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter
from reportlab.lib.units import inch
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.lib.utils import ImageReader
from PIL import Image
import textwrap
import os

st.title("Word to Annotated PDF")

# Helper: render DOCX text into a PDF (ReportLab) and apply overlays
def create_pdf_from_docx(docx_file, overlays, page_size=letter, margin=50, font_name='Helvetica', font_size=12):
    # Read document text
    doc = Document(docx_file)
    paragraphs = [p.text for p in doc.paragraphs]

    buffer = io.BytesIO()
    c = canvas.Canvas(buffer, pagesize=page_size)
    page_width, page_height = page_size

    # Prepare text layout
    max_width = page_width - 2 * margin
    # We'll use a simple word-wrap using textwrap but measure with stringWidth for better fit
    lines = []
    for para in paragraphs:
        if not para:
            lines.append("")
            continue
        # conservative average char width
        wrapper = textwrap.TextWrapper(width=100)
        # break paragraph into candidate lines, then refine using stringWidth
        words = para.split()
        cur = ""
        for w in words:
            test = (cur + " " + w).strip()
            width = pdfmetrics.stringWidth(test, font_name, font_size)
            if width <= max_width:
                cur = test
            else:
                if cur:
                    lines.append(cur)
                cur = w
        if cur:
            lines.append(cur)
        lines.append("")

    # Draw text across pages
    y = page_height - margin
    leading = font_size * 1.2
    c.setFont(font_name, font_size)
    for line in lines:
        if y - leading < margin:
            c.showPage()
            c.setFont(font_name, font_size)
            y = page_height - margin
        c.drawString(margin, y - leading, line)
        y -= leading

    # At this point we have the base PDF content; to add overlays we need to (re)draw them on each page
    # ReportLab canvas draws in current page; but since we already advanced pages as needed while drawing text,
    # we need to reposition to the first page again. Easiest approach: regenerate pages but draw overlays as we go.
    # Simpler: we will create a new PDF where we re-render text and overlays together.
    buffer.seek(0)

    # Create final canvas and render again with overlays
    final = io.BytesIO()
    c2 = canvas.Canvas(final, pagesize=page_size)
    y = page_height - margin
    c2.setFont(font_name, font_size)
    for line in lines:
        if y - leading < margin:
            # draw overlays that belong to this page (we'll draw overlays on every page by default)
            # overlays are applied on each page in this implementation
            draw_overlays_on_canvas(c2, overlays, page_width, page_height)
            c2.showPage()
            c2.setFont(font_name, font_size)
            y = page_height - margin
        c2.drawString(margin, y - leading, line)
        y -= leading

    # draw overlays for the last page
    draw_overlays_on_canvas(c2, overlays, page_width, page_height)
    c2.save()
    final.seek(0)
    return final.getvalue()


def draw_overlays_on_canvas(c, overlays, page_width, page_height):
    """Draw overlays (images/text) onto current ReportLab canvas page.
    Coordinates supplied are interpreted as top-left origin (like typical screen pixels).
    """
    for ov in overlays:
        if ov['type'] == 'image':
            path = ov['path']
            if not os.path.isfile(path):
                continue
            img = Image.open(path)
            img_w, img_h = img.size
            # If width/height provided, use them (in points), else use image size (assume 72 dpi -> 1px=1pt approx)
            w = ov.get('w', img_w)
            h = ov.get('h', img_h)
            x_top = ov.get('x', 0)
            y_top = ov.get('y', 0)
            # Convert top-left to ReportLab bottom-left origin
            x = x_top
            y = page_height - y_top - h
            try:
                c.drawImage(ImageReader(img), x, y, width=w, height=h, mask='auto')
            except Exception:
                pass
        elif ov['type'] == 'text':
            text = ov.get('text', '')
            x_top = ov.get('x', 0)
            y_top = ov.get('y', 0)
            size = ov.get('size', 12)
            c.setFont(ov.get('font', 'Helvetica'), size)
            x = x_top
            y = page_height - y_top - size
            c.drawString(x, y, text)


st.sidebar.header('Overlay settings')
name = st.sidebar.text_input('Your name', 'Your Name')
date_override = st.sidebar.text_input('Date (leave blank for today)', '')
if not date_override:
    date_str = str(date.today())
else:
    date_str = date_override

# Default example overlay positions (top-left coordinates in points)
default_overlays = [
    {'type': 'image', 'path': 'checkmark.png', 'x': 100, 'y': 100, 'w': 20, 'h': 20},
    {'type': 'image', 'path': 'signature.png', 'x': 200, 'y': 250, 'w': 150, 'h': 60},
    {'type': 'text', 'text': name, 'x': 200, 'y': 200, 'size': 14},
    {'type': 'text', 'text': f'Date: {date_str}', 'x': 200, 'y': 230, 'size': 12},
]

# Upload Word document
uploaded_file = st.file_uploader('Choose a Word document', type=['docx'])

if uploaded_file:
    st.write('Converting Word to PDF...')
    try:
        pdf_bytes = create_pdf_from_docx(uploaded_file, default_overlays)
        st.success('Converted to PDF (with overlays).')

        st.download_button('Download Annotated PDF', data=pdf_bytes, file_name='annotated.pdf')
    except Exception as e:
        st.error(f'Error creating PDF: {e}')
