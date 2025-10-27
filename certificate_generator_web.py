import streamlit as st
import pandas as pd
from pptx import Presentation
import tempfile, os, zipfile

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

# ======================================================
# 4. GENERATE PERSONALIZED CERTIFICATES
# ======================================================
if st.button("Generate Certificates"):
    temp_dir = tempfile.mkdtemp()
    pptx_paths = []

    for idx, name in enumerate(names, start=1):
        # Copy the template
        prs_copy = Presentation(template_file)
        # Replace placeholders
        for slide in prs_copy.slides:
            for shape in slide.shapes:
                if shape.has_text_frame:
                    for paragraph in shape.text_frame.paragraphs:
                        for run in paragraph.runs:
                            run.text = run.text.replace("[NAME]", name).replace("[DATE]", event_date)
        # Save personalized PPTX
        pptx_file = os.path.join(temp_dir, f"Certificate_{idx}_{name}.pptx")
        prs_copy.save(pptx_file)
        pptx_paths.append(pptx_file)

    # Create ZIP file
    zip_path = os.path.join(temp_dir, "All_Certificates.zip")
    with zipfile.ZipFile(zip_path, "w") as zipf:
        for file in pptx_paths:
            zipf.write(file, os.path.basename(file))

    # Download button
    with open(zip_path, "rb") as f:
        st.download_button(
            "Download All Certificates (ZIP of PPTX)",
            f,
            file_name="All_Certificates.zip"
        )

    st.success("All certificates generated and ready to download!")
