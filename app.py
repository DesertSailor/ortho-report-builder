import streamlit as st
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from streamlit_cropper import st_cropper
from PIL import Image
import io

# --- Session State Initialization ---
if 'cases' not in st.session_state:
    st.session_state.cases = []
if 'case_counter' not in st.session_state:
    st.session_state.case_counter = 0

st.set_page_config(page_title="Ortho Morning Report", layout="wide")
st.title("🏥 Orthopedic Morning Report Builder")

# --- 1. Duty Team Setup (Top of the Page) ---
st.header("👥 Today's Duty Team")
with st.expander("Tap to enter Consultants & Residents on Duty", expanded=True):
    c1, c2 = st.columns(2)
    with c1:
        st.subheader("Consultants")
        c_ward = st.text_input("Consultant - Ward", placeholder="Dr. ...")
        c_emer = st.text_input("Consultant - Emergency", placeholder="Dr. ...")
        c_pelv = st.text_input("Consultant - Pelvic", placeholder="Dr. ...")
        c_sport = st.text_input("Consultant - Sport", placeholder="Dr. ...")
    with c2:
        st.subheader("Residents")
        res1 = st.text_input("Resident 1", placeholder="Dr. ...")
        res2 = st.text_input("Resident 2", placeholder="Dr. ...")
        res3 = st.text_input("Resident 3", placeholder="Dr. ...")
        res4 = st.text_input("Resident 4", placeholder="Dr. ...")
        res5 = st.text_input("Resident 5", placeholder="Dr. ...")
        res6 = st.text_input("Resident 6", placeholder="Dr. ...")

st.markdown("---")

# --- 2. Case Entry Section ---
st.header("📝 Add Patient Case")

# Use counter to fully reset file uploaders and inputs after clicking 'Add Case'
idx = st.session_state.case_counter

cc1, cc2 = st.columns(2)
with cc1:
    p_name = st.text_input("Patient Initials / Name", key=f"name_{idx}")
    p_mrn = st.text_input("MRN", key=f"mrn_{idx}")
    p_age = st.text_input("Age", key=f"age_{idx}")
with cc2:
    p_sex = st.selectbox("Sex", ["M", "F"], key=f"sex_{idx}")
    p_moi = st.text_input("MOI (e.g., RTA, Fall, FDA)", key=f"moi_{idx}")
    p_duration = st.text_input("Duration / Delay", key=f"dur_{idx}")

# The Operated Toggle
is_operated = st.radio("Was this case operated during the shift?", ["No", "Yes"], key=f"op_{idx}")

# Image Cropping Areas
cropped_pre = None
cropped_post = None

st.subheader("📸 Radiograph Processing")
img_col1, img_col2 = st.columns(2)

with img_col1:
    pre_file = st.file_uploader("Upload Injury / Pre-Op Image", type=["jpg", "jpeg", "png"], key=f"pre_file_{idx}")
    if pre_file:
        st.caption("Drag corners to crop your Pre-Op film:")
        raw_pre = Image.open(pre_file)
        cropped_pre = st_cropper(raw_pre, realtime_update=True, box_color='#FF0000', aspect_ratio=None, key=f"crop_pre_{idx}")

with img_col2:
    if is_operated == "Yes":
        post_file = st.file_uploader("Upload Post-Op / Fixation Image", type=["jpg", "jpeg", "png"], key=f"post_file_{idx}")
        if post_file:
            st.caption("Drag corners to crop your Post-Op film:")
            raw_post = Image.open(post_file)
            cropped_post = st_cropper(raw_post, realtime_update=True, box_color='#00FF00', aspect_ratio=None, key=f"crop_post_{idx}")

# Add Case Submission Button
if st.button("➕ Save This Case"):
    if p_name or p_mrn:
        new_case = {
            "name": p_name, "mrn": p_mrn, "age": p_age, "sex": p_sex, "moi": p_moi, "duration": p_duration,
            "operated": is_operated,
            "pre_img": cropped_pre,
            "post_img": cropped_post if is_operated == "Yes" else None
        }
        st.session_state.cases.append(new_case)
        st.session_state.case_counter += 1  # Increments the index to instantly clear fields
        st.rerun()
    else:
        st.error("Please provide at least a Patient Name or MRN before saving.")

st.markdown("---")

# --- 3. Queue & PPT Presentation Generation ---
if st.session_state.cases:
    st.header(f"📋 Current Presentation Queue ({len(st.session_state.cases)} Cases Added)")
    
    # List current patients added
    for i, c in enumerate(st.session_state.cases):
        status_tag = "🔴 Operated" if c['operated'] == "Yes" else "🟢 Conservative"
        st.write(f"**Case {i+1}:** {c['name']} ({c['age']}/{c['sex']}) — MRN: {c['mrn']} | {status_tag}")
        
    if st.button("🗑️ Clear Presentation Queue"):
        st.session_state.cases = []
        st.session_state.case_counter = 0
        st.rerun()

    st.subheader("🚀 Export Final Deck")
    if st.button("Generate Presentation"):
        prs = Presentation()
        prs.slide_width = Inches(13.333)
        prs.slide_height = Inches(7.5)
        blank_layout = prs.slide_layouts[6]
        
        # --- Slide 1: Cover/Duty Team Layout ---
        slide1 = prs.slides.add_slide(blank_layout)
        tb = slide1.shapes.add_textbox(Inches(0.7), Inches(0.8), Inches(11.9), Inches(6.0))
        tf = tb.text_frame
        tf.word_wrap = True
        
        p_title = tf.paragraphs[0]
        p_title.text = "Orthopedic Department Duty Report"
        p_title.font.size = Pt(36)
        p_title.font.bold = True
        
        p_team = tf.add_paragraph()
        p_team.text = f"\n🔹 **Consultants on Duty:**\n" \
                     f"   • Ward: {c_ward if c_ward else 'None'}\n" \
                     f"   • Emergency: {c_emer if c_emer else 'None'}\n" \
                     f"   • Pelvic: {c_pelv if c_pelv else 'None'}\n" \
                     f"   • Sport: {c_sport if c_sport else 'None'}\n\n" \
                     f"🔹 **Residents on Duty:**\n" \
                     f"   • {', '.join([r for r in [res1, res2, res3, res4, res5, res6] if r])}"
        p_team.font.size = Pt(18)

        # --- Dynamic Case-by-Case Slides ---
        for c in st.session_state.cases:
            # Slide A: Pre-Op / Presentation Slide
            slide_pre = prs.slides.add_slide(blank_layout)
            tb_meta = slide_pre.shapes.add_textbox(Inches(0.5), Inches(0.4), Inches(12.333), Inches(1.2))
            tf_meta = tb_meta.text_frame
            
            p_m = tf_meta.paragraphs[0]
            p_m.text = f"{c['name']} | {c['age']}/{c['sex']} | MRN: {c['mrn']} | MOI: {c['moi']} ({c['duration']})"
            p_m.font.size = Pt(24)
            p_m.font.bold = True
            
            p_sub = tf_meta.add_paragraph()
            p_sub.text = f"Pre-Operative Status [{ 'OPERATED CASE' if c['operated'] == 'Yes' else 'CONSERVATIVE MANAGEMENT' }]"
            p_sub.font.size = Pt(16)
            p_sub.font.color.rgb = RGBColor(220, 50, 50)
            
            if c['pre_img']:
                img_buf = io.BytesIO()
                c['pre_img'].save(img_buf, format="PNG")
                img_buf.seek(0)
                slide_pre.shapes.add_picture(img_buf, Inches(2.5), Inches(1.8), width=Inches(8.333))

            # Slide B: Post-Op Slide (Only if operated AND post-op picture exists)
            if c['operated'] == "Yes" and c['post_img']:
                slide_post = prs.slides.add_slide(blank_layout)
                tb_meta2 = slide_post.shapes.add_textbox(Inches(0.5), Inches(0.4), Inches(12.333), Inches(1.2))
                tf_meta2 = tb_meta2.text_frame
                
                p_m2 = tf_meta2.paragraphs[0]
                p_m2.text = f"{c['name']} | {c['age']}/{c['sex']} | MRN: {c['mrn']} (Post-Op Check)"
                p_m2.font.size = Pt(24)
                p_m2.font.bold = True
                
                p_sub2 = tf_meta2.add_paragraph()
                p_sub2.text = "Post-Operative / Internal Fixation Layout"
                p_sub2.font.size = Pt(16)
                p_sub2.font.color.rgb = RGBColor(50, 180, 50)
                
                img_buf2 = io.BytesIO()
                c['post_img'].save(img_buf2, format="PNG")
                img_buf2.seek(0)
                slide_post.shapes.add_picture(img_buf2, Inches(2.5), Inches(1.8), width=Inches(8.333))

        # Save Presentation Output
        final_ppt_buf = io.BytesIO()
        prs.save(final_ppt_buf)
        final_ppt_buf.seek(0)
        
        st.download_button(
            label="💾 Save PowerPoint File to Device",
            data=final_ppt_buf,
            file_name="Morning_Report_Presentation.pptx",
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
        )
