import streamlit as st
import pandas as pd
from pptx import Presentation
from pptx.util import Inches
from streamlit_cropper import st_cropper
from PIL import Image
import io
import docx


# --- Helper: Parse Consultant Schedule ---
def get_duty_senior(date_val, doc_file):
    doc = docx.Document(doc_file)
    table = doc.tables[0]
    data = [[cell.text for cell in row.cells] for row in table.rows]
    df = pd.DataFrame(data[1:], columns=data[0])

    # Filter by G.C date
    match = df[df['G.C'] == str(date_val)]
    if not match.empty:
        return f"{match['EOPD'].values[0]} / {match['Ward'].values[0]}"
    return "Not Found"


# --- App Interface ---
st.set_page_config(layout="wide")
st.title("OrthoCase PPT Generator")

# Sidebar: File Uploads
with st.sidebar:
    st.header("Setup")
    cons_file = st.file_uploader("Upload Consultant Schedule (DOCX)", type="docx")
    target_date = st.text_input("Enter Date (G.C)", "1")

# Form: Patient Data
with st.form("patient_data"):
    col1, col2 = st.columns(2)
    with col1:
        name = st.text_input("Patient Name")
        mrn = st.text_input("MRN")
    with col2:
        age_sex = st.text_input("Age / Sex")
        moi = st.text_input("MOI / Duration")

    # Image Section
    st.subheader("1. Pre-op / Injury Images")
    pre_op_file = st.file_uploader("Upload Pre-op/X-ray", type=['png', 'jpg', 'jpeg'], key="pre")
    if pre_op_file:
        img = Image.open(pre_op_file)
        pre_op_cropped = st_cropper(img, realtime_update=True, box_color='red')

    st.subheader("2. Post-op / Follow-up Images")
    post_op_file = st.file_uploader("Upload Post-op/CT", type=['png', 'jpg', 'jpeg'], key="post")
    if post_op_file:
        img2 = Image.open(post_op_file)
        post_op_cropped = st_cropper(img2, realtime_update=True, box_color='green')

    submitted = st.form_submit_button("Generate PowerPoint")

# --- PPT Generation Logic ---
if submitted:
    prs = Presentation()

    # Senior Info
    duty_team = get_duty_senior(target_date, cons_file) if cons_file else "Manual"

    # Slide 1: Pre-op
    slide1 = prs.slides.add_slide(prs.slide_layouts[6])
    slide1.shapes.add_textbox(Inches(0.5), Inches(0.5), Inches(8), Inches(1)).text = \
        f"Date: {target_date} | Team: {duty_team}\n{name} ({age_sex}) | MRN: {mrn}\nMOI: {moi}"

    if 'pre_op_cropped' in locals():
        # Save cropped image to memory
        buf = io.BytesIO()
        pre_op_cropped.save(buf, format="PNG")
        slide1.shapes.add_picture(buf, Inches(0.5), Inches(2), width=Inches(6))

    # Slide 2: Post-op (Only if uploaded)
    if 'post_op_cropped' in locals():
        slide2 = prs.slides.add_slide(prs.slide_layouts[6])
        slide2.shapes.add_textbox(Inches(0.5), Inches(0.5), Inches(8), Inches(1)).text = f"Post-op: {name}"
        buf2 = io.BytesIO()
        post_op_cropped.save(buf2, format="PNG")
        slide2.shapes.add_picture(buf2, Inches(0.5), Inches(2), width=Inches(6))

    # Save to buffer for download
    pptx_buf = io.BytesIO()
    prs.save(pptx_buf)

    st.download_button("Download Presentation", pptx_buf, "Case_Report.pptx")