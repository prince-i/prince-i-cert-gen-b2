import streamlit as st
from pptx import Presentation
from pptx.util import Inches, Pt
from io import BytesIO
from datetime import date

# --- Streamlit app ---
st.title("Certificate Generator")

# User inputs
name = st.text_input("Recipient Name")
course = st.text_input("Course Name", "Python Basics")
date_str = st.text_input("Date", date.today().strftime("%B %d, %Y"))

# Button to generate certificate
if st.button("Generate Certificate"):
    if not name.strip():
        st.warning("Please enter a name.")
    else:
        # --- Create PowerPoint certificate ---
        prs = Presentation()

        # Slide layout (blank)
        slide_layout = prs.slide_layouts[6]
        slide = prs.slides.add_slide(slide_layout)

        # Add title
        title_box = slide.shapes.add_textbox(Inches(1), Inches(1.5), Inches(8), Inches(1))
        title_tf = title_box.text_frame
        title_tf.text = "Certificate of Completion"
        title_tf.paragraphs[0].font.size = Pt(40)

        # Add recipient name
        name_box = slide.shapes.add_textbox(Inches(1), Inches(3), Inches(8), Inches(1))
        name_tf = name_box.text_frame
        name_tf.text = name
        name_tf.paragraphs[0].font.size = Pt(32)

        # Add course info
        course_box = slide.shapes.add_textbox(Inches(1), Inches(4), Inches(8), Inches(1))
        course_tf = course_box.text_frame
        course_tf.text = f"For completing the course: {course}"
        course_tf.paragraphs[0].font.size = Pt(24)

        # Add date
        date_box = slide.shapes.add_textbox(Inches(1), Inches(5), Inches(8), Inches(1))
        date_tf = date_box.text_frame
        date_tf.text = f"Date: {date_str}"
        date_tf.paragraphs[0].font.size = Pt(20)

        # Save to BytesIO
        output = BytesIO()
        prs.save(output)
        output.seek(0)

        # Download link
        st.download_button(
            label="Download Certificate (.pptx)",
            data=output,
            file_name=f"{name}_certificate.pptx",
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
        )

        st.success("Certificate generated successfully!")
