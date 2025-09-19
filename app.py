import streamlit as st
from docx import Document
import io
import fitz  # PyMuPDF
from datetime import date

st.title("Word Filler and PDF Export with Annotations")

# Upload Word document
uploaded_file = st.file_uploader("Choose a Word document", type=["docx"])

if uploaded_file:
    # Load the Word document
    doc = Document(uploaded_file)

    st.write("Document loaded! Here are the paragraphs:")
    for i, para in enumerate(doc.paragraphs):
        st.write(f"{i+1}: {para.text}")

    # Replace placeholders
    placeholder = st.text_input("Text to replace:", "PLACEHOLDER")
    replacement = st.text_input("Replacement text:", "Hello World")

    if st.button("Fill Word and Export PDF"):
        # Replace text in Word
        for para in doc.paragraphs:
            if placeholder in para.text:
                para.text = para.text.replace(placeholder, replacement)

        # Save Word to memory
        word_stream = io.BytesIO()
        doc.save(word_stream)
        st.download_button("Download filled Word", data=word_stream.getvalue(), file_name="filled.docx")

        # Convert Word to simple PDF using PyMuPDF
        pdf_doc = fitz.open()  # new PDF
        for para in doc.paragraphs:
            page = pdf_doc.new_page()
            page.insert_text((50, 50), para.text, fontsize=12)

        # Annotate PDF with images and text at specific coordinates
        for page in pdf_doc:
            # Example positions (x, y in points)
            page.insert_image(fitz.Rect(100, 100, 120, 120), filename="checkmark.png")
            page.insert_text((200, 200), "Your Name", fontsize=14)
            page.insert_text((200, 230), f"Date: {date.today()}", fontsize=12)
            page.insert_image(fitz.Rect(200, 250, 300, 300), filename="signature.png")

        pdf_stream = io.BytesIO()
        pdf_doc.save(pdf_stream)
        st.download_button("Download annotated PDF", data=pdf_stream.getvalue(), file_name="filled_annotated.pdf")
