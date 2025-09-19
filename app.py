import streamlit as st
import fitz  # PyMuPDF
from datetime import date
import subprocess
import os

st.title("Word → PDF Annotator (Online)")

st.markdown("""
Upload a Word document (.docx). 
The app will convert it to PDF, place your annotations (name, date, checkmarks, signature), 
and let you download the final PDF.
""")

# --- File Upload ---
uploaded_docx = st.file_uploader("Upload Word Document", type=["docx"])

# --- Coordinates Input ---
st.sidebar.header("Coordinates (points)")
name_x = st.sidebar.number_input("Name X", value=100)
name_y = st.sidebar.number_input("Name Y", value=150)
date_x = st.sidebar.number_input("Date X", value=100)
date_y = st.sidebar.number_input("Date Y", value=180)
check_x = st.sidebar.number_input("Checkmark X", value=200)
check_y = st.sidebar.number_input("Checkmark Y", value=250)
check_w = st.sidebar.number_input("Checkmark Width", value=20)
check_h = st.sidebar.number_input("Checkmark Height", value=20)
sig_x = st.sidebar.number_input("Signature X", value=100)
sig_y = st.sidebar.number_input("Signature Y", value=400)
sig_w = st.sidebar.number_input("Signature Width", value=200)
sig_h = st.sidebar.number_input("Signature Height", value=60)

# --- Static images in repo ---
checkmark_path = "checkmark.png"
signature_path = "signature.png"

if st.button("Generate PDF") and uploaded_docx:
    # --- Save uploaded Word file ---
    input_docx = "input.docx"
    with open(input_docx, "wb") as f:
        f.write(uploaded_docx.read())

    # --- Convert Word to PDF using LibreOffice (Streamlit Cloud has it) ---
    st.info("Converting Word to PDF...")
    subprocess.run([
        "libreoffice", "--headless", "--convert-to", "pdf", input_docx, "--outdir", "."
    ])
    temp_pdf = input_docx.replace(".docx", ".pdf")

    # --- Annotate PDF ---
    st.info("Annotating PDF...")
    pdf = fitz.open(temp_pdf)
    page = pdf[0]  # first page

    # Add text
    page.insert_text((name_x, name_y), "Samuel Rios-Lazo", fontsize=12, color=(0,0,0))
    page.insert_text((date_x, date_y), str(date.today()), fontsize=12, color=(0,0,0))

    # Add checkmark image
    rect_check = fitz.Rect(check_x, check_y, check_x+check_w, check_y+check_h)
    with open(checkmark_path, "rb") as f:
        page.insert_image(rect_check, stream=f.read())

    # Add signature image
    rect_sig = fitz.Rect(sig_x, sig_y, sig_x+sig_w, sig_y+sig_h)
    with open(signature_path, "rb") as f:
        page.insert_image(rect_sig, stream=f.read())

    final_pdf = "final.pdf"
    pdf.save(final_pdf)
    pdf.close()

    # --- Download button ---
    st.success("✅ PDF generated!")
    with open(final_pdf, "rb") as f:
        st.download_button("Download Final PDF", f, file_name="final.pdf")
