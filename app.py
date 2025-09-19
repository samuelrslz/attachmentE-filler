import streamlit as st
from docx import Document
import io
from fpdf import FPDF

st.title("Word Filler and PDF Export")

# Upload Word document
uploaded_file = st.file_uploader("Choose a Word document", type=["docx"])

if uploaded_file:
    # Load the Word document
    doc = Document(uploaded_file)

    st.write("Document loaded! Here are the paragraphs:")
    for i, para in enumerate(doc.paragraphs):
        st.write(f"{i+1}: {para.text}")

    # Example: Replace placeholder text
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

        # Convert Word to simple PDF
        pdf = FPDF()
        pdf.add_page()
        pdf.set_auto_page_break(auto=True, margin=15)
        pdf.set_font("Arial", size=12)

        for para in doc.paragraphs:
            pdf.multi_cell(0, 10, para.text)

        pdf_stream = io.BytesIO()
        pdf.output(pdf_stream)
        st.download_button("Download PDF", data=pdf_stream.getvalue(), file_name="filled.pdf")
