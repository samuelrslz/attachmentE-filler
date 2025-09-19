import fitz  # PyMuPDF
from datetime import date
import os

# Optional: for Word → PDF conversion (works on Windows/Mac)
try:
    from docx2pdf import convert
    DOCX2PDF_AVAILABLE = True
except ImportError:
    DOCX2PDF_AVAILABLE = False


def convert_word_to_pdf(input_docx, output_pdf):
    """Convert Word to PDF using docx2pdf (Windows/Mac only)."""
    if not DOCX2PDF_AVAILABLE:
        raise ImportError("docx2pdf not installed or unsupported on this OS. Use LibreOffice instead.")
    convert(input_docx, output_pdf)


def annotate_pdf(input_pdf, output_pdf, annotations):
    """
    Annotate a PDF file.
    annotations: list of dicts with keys:
        - type: 'text' or 'image'
        - value: string (for text) or path to image (for image)
        - x, y: top-left coordinates
        - w, h: (only for images) target size
    """
    pdf = fitz.open(input_pdf)
    page = pdf[0]  # first page only, extend if needed

    for ann in annotations:
        if ann["type"] == "text":
            page.insert_text(
                (ann["x"], ann["y"]),
                ann["value"],
                fontsize=ann.get("fontsize", 12),
                color=(0, 0, 0)
            )
        elif ann["type"] == "image":
            rect = fitz.Rect(
                ann["x"], ann["y"],
                ann["x"] + ann["w"],
                ann["y"] + ann["h"]
            )
            with open(ann["value"], "rb") as f:
                page.insert_image(rect, stream=f.read())

    pdf.save(output_pdf)
    pdf.close()


if __name__ == "__main__":
    # --- INPUT FILES ---
    input_docx = "input.docx"
    temp_pdf = "converted.pdf"
    final_pdf = "final.pdf"

    checkmark_path = "checkmark.png"
    signature_path = "signature.png"

    # --- STEP 1: Convert Word to PDF ---
    if DOCX2PDF_AVAILABLE:
        convert_word_to_pdf(input_docx, temp_pdf)
    else:
        # fallback: if on Linux, use LibreOffice command line instead
        os.system(f'libreoffice --headless --convert-to pdf "{input_docx}" --outdir .')
        temp_pdf = input_docx.replace(".docx", ".pdf")

    # --- STEP 2: Define annotations ---
    annotations = [
        {"type": "text", "value": "Samuel Rios-Lazo", "x": 150, "y": 512, "fontsize": 12},
        {"type": "text", "value": str(date.today()), "x": 120, "y": 565, "fontsize": 12},
        {"type": "image", "value": checkmark_path, "x": 200, "y": 250, "w": 20, "h": 20},
        {"type": "image", "value": signature_path, "x": 100, "y": 517, "w": 100, "h": 25},
    ]

    # --- STEP 3: Annotate PDF ---
    annotate_pdf(temp_pdf, final_pdf, annotations)

    print(f"✅ Done! Saved as {final_pdf}")
