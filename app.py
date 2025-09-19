import streamlit as st
import fitz  # PyMuPDF
from docx import Document
import io
from datetime import date
from PIL import Image

st.title("Word to Annotated PDF")

# Upload Word document
uploaded_file = st.file_uploader("Choose a Word document", type=["docx"])

if uploaded_file:
    doc = Document(uploaded_file)

    # Convert Word to PDF (simple: one page per paragraph)
    pdf_doc = fitz.open()  # new PDF
    for para in doc.paragraphs:
        page = pdf_doc.new_page()
        page.insert_text((50, 50), para.text, fontsize=12)

    st.write("Word converted to PDF!")

    if st.button("Annotate PDF and Download"):
        # Annotate each page
        for page in pdf_doc:
            # Add checkmark image
            page.insert_image(fitz.Rect(100, 100, 120, 120), filename="checkmark.png")
            # Add signature image
            page.insert_image(fitz.Rect(200, 250, 300, 300), filename="signature.png")
            # Add text
            page.insert_text((200, 200), "Your Name", fontsize=14)
            page.insert_text((200, 230), f"Date: {date.today()}", fontsize=12)

        # Save annotated PDF to memory
        pdf_stream = io.BytesIO()
        pdf_doc.save(pdf_stream)
        st.download_button(
            "Download Annotated PDF",
            data=pdf_stream.getvalue(),
            file_name="annotated.pdf"
        )
