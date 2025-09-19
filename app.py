import streamlit as st
from docx import Document
import io
import fitz  # PyMuPDF
from datetime import date

st.title("Word to Annotated PDF")

# Upload Word document
uploaded_file = st.file_uploader("Choose a Word document", type=["docx"])

if uploaded_file:
    doc = Document(uploaded_file)

    # Combine all paragraphs into one string
    full_text = "\n\n".join([para.text for para in doc.paragraphs])

    # Create PDF and add text (auto page breaks)
    pdf_doc = fitz.open()
    page = pdf_doc.new_page()
    text_rect = fitz.Rect(50, 50, 550, 800)  # adjust page margins
    page.insert_textbox(text_rect, full_text, fontsize=12, align=0)

    st.write("Word converted to PDF!")

    if st.button("Annotate PDF and Download"):
        # Add annotations (images + text)
        for page in pdf_doc:
            page.insert_image(fitz.Rect(100, 100, 120, 120), filename="checkmark.png")
            page.insert_image(fitz.Rect(200, 250, 300, 300), filename="signature.png")
            page.insert_text((200, 200), "Your Name", fontsize=14)
            page.insert_text((200, 230), f"Date: {date.today()}", fontsize=12)

        pdf_stream = io.BytesIO()
        pdf_doc.save(pdf_stream)
        st.download_button(
            "Download Annotated PDF",
            data=pdf_stream.getvalue(),
            file_name="annotated.pdf"
        )
