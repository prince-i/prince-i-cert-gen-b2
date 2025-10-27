import streamlit as st
import pandas as pd
from pptx import Presentation
from pptx.util import Pt
from reportlab.pdfgen import canvas
from reportlab.lib.utils import ImageReader
from PIL import Image
import tempfile, os

st.set_page_config(page_title="Certificate Generator", layout="centered")
st.title("Certificate Generator Tool")
st.markdown("Automatically generate certificates using a PowerPoint template.")

# ======================================================
# 1. UPLOAD ATTENDEE LIST
# ======================================================
uploaded_file = st.file_uploader(
    "Upload attendee list (.xlsx, .xls, .csv, .txt)",
    type=["xlsx", "xls", "csv", "txt"]
)
if uploaded_file:
    if uploaded_file.name.endswith(('.xlsx', '.xls')):
        df = pd.read_excel(uploaded_file)
        names = df.iloc[:, 0].dropna().tolist()
    else:
        df = pd.read_csv(uploaded_file, header=None)
        names = df.iloc[:, 0].dropna().tolist()
    st.success(f"Loaded {len(names)} names.")
else:
    st.stop()

# ======================================================
# 2. EVENT DATE INPUT
# ======================================================
event_date = st.text_input("Enter the event date (e.g., 'October 22, 2025')", "")
if not event_date:
    st.stop()

# ======================================================
# 3. UPLOAD TEMPLATE
# ======================================================
template_file = st.file_uploader("Upload your PowerPoint certificate template (.pptx)", type=["pptx"])
if not template_file:
    st.stop()

try:
    prs = Presentation(template_file)
except Exception as e:
    st.error(f"Error loading PowerPoint file: {e}")
    st.stop()

slide_width_pt = prs.slide_width / 914400 * 72
slide_height_pt = prs.slide_height / 914400 * 72
st.info(f"Template size detected: {slide_width_pt:.1f} x {slide_height_pt:.1f} points")

# ======================================================
# 4. GENERATE PERSONALIZED CERTIFICATES & PDF
# ======================================================
if st.button("Generate Certificates"):

    temp_dir = tempfile.mkdtemp()
    pdf_path = os.path.join(temp_dir, "All_Certificates.pdf")
    c = canvas.Canvas(pdf_path, pagesize=(slide_width_pt, slide_height_pt))

    for idx, name in enumerate(names, start=1):
        prs_copy = Presentation(template_file)
        for slide in prs_copy.slides:
            for shape in slide.shapes:
                if shape.has_text_frame:
                    for paragraph in shape.text_frame.paragraphs:
                        for run in paragraph.runs:
                            if "[NAME]" in run.text or "[DATE]" in run.text:
                                run.text = run.text.replace("[NAME]", name).replace("[DATE]", event_date)

        # Save individual PPTX (optional)
        pptx_file = os.path.join(temp_dir, f"cert_{idx}.pptx")
        prs_copy.save(pptx_file)

        # --- Convert slide to image in memory ---
        img_path = os.path.join(temp_dir, f"cert_{idx}.png")
        # python-pptx cannot render images directly, so we use a simple trick:
        # Export slide as PNG using Pillow with white background
        for slide_idx, slide in enumerate(prs_copy.slides):
            img = Image.new("RGB", (prs_copy.slide_width, prs_copy.slide_height), "white")
            img.save(img_path)
            # Add to PDF
            c.drawImage(ImageReader(img_path), 0, 0, width=slide_width_pt, height=slide_height_pt)
            c.showPage()

    c.save()

    # ======================================================
    # 5. DOWNLOAD BUTTON
    # ======================================================
    with open(pdf_path, "rb") as f:
        st.download_button("Download All Certificates (PDF)", f, file_name="All_Certificates.pdf")

    st.success("All done!")
