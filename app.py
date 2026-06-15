import streamlit as st
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from streamlit_cropper import st_cropper
from PIL import Image
import io
from datetime import date

# --- Formatting Helpers ---
def clean_dr_name(name):
    """Ensures name is properly capitalized and prefixed with Dr."""
    name = name.strip()
    if not name:
        return ""
    if name.lower().startswith("dr"):
        core = name[2:].strip().lstrip(".").strip()
        return f"Dr. {core}"
    return f"Dr. {name}"

# --- Session State Initialization ---
if 'cases' not in st.session_state:
    st.session_state.cases = []
if 'case_counter' not in st.session_state:
    st.session_state.case_counter = 0

st.set_page_config(page_title="Ortho Morning Report Builder", layout="wide")
st.title("🏥 Orthopedic Morning Report Builder")

# --- 1. Global Presentation Settings ---
st.header("📅 Presentation Settings")
# Calendar selector for the main report date
presentation_date = st.date_input("Date of Presentation (Calendar)", date.today())
presentation_date_str = presentation_date.strftime("%B %d, %Y")

st.markdown("---")

# --- 2. Duty Team Setup ---
st.header("👥 Today's Duty Team")
with st.expander("Review Consultants & Residents on Duty", expanded=True):
    c1, c2 = st.columns(2)
    with c1:
        st.subheader("Consultants")
        cons_1 = st.text_input("Consultant 1", placeholder="Name")
        cons_2 = st.text_input("Consultant 2", placeholder="Name")
        cons_3 = st.text_input("Consultant 3", placeholder="Name")
        cons_4 = st.text_input("Consultant 4", placeholder="Name")
        cons_5 = st.text_input("Consultant 5", value="Dr. Solomon")
        cons_6 = st.text_input("Consultant 6", value="Dr. Dawit")
    with c2:
        st.subheader("Residents")
        res1 = st.text_input("Resident 1", placeholder="Name")
        res2 = st.text_input("Resident 2", placeholder="Name")
        res3 = st.text_input("Resident 3", placeholder="Name")
        res4 = st.text_input("Resident 4", placeholder="Name")
        res5 = st.text_input("Resident 5", placeholder="Name")
        res6 = st.text_input("Resident 6", placeholder="Name")

st.markdown("---")

# --- 3. Case Entry Section ---
st.header("📝 Add Patient Case")
idx = st.session_state.case_counter

cc1, cc2 = st.columns(2)
with cc1:
    p_name = st.text_input("Patient Initials / Name", key=f"name_{idx}")
    
    # Enforce numeric data verification for MRN
    raw_mrn = st.text_input("MRN (Numbers Only)", key=f"raw_mrn_{idx}")
    p_mrn = "".join(filter(str.isdigit, raw_mrn))
    if raw_mrn and not raw_mrn.isdigit():
        st.warning("⚠️ Non-numeric characters stripped automatically.")
        
    p_age = st.text_input("Age", key=f"age_{idx}")
    p_sex = st.selectbox("Sex", ["M", "F"], key=f"sex_{idx}")
    
    # Individual Patient Date of Injury Calendar Selector
    p_doi = st.date_input("Date of Injury", date.today(), key=f"doi_{idx}")
    p_doi_str = p_doi.strftime("%b %d, %Y")

with cc2:
    # MOI Configuration Matrix
    moi_options = ["RTA", "Fall from Height", "FDA", "Sports Injury", "Direct Blow", "Others"]
    selected_moi = st.selectbox("Mechanism of Injury (MOI)", moi_options, key=f"moi_sel_{idx}")
    if selected_moi == "Others":
        p_moi = st.text_input("Specify Custom MOI", key=f"moi_txt_{idx}")
    else:
        p_moi = selected_moi

    # Duration Configuration Matrix
    dur_options = ["< 6 hours", "6-12 hours", "12-24 hours", "24-48 hours", "> 2 days", "Others"]
    selected_dur = st.selectbox("Duration / Presentation Delay", dur_options, key=f"dur_sel_{idx}")
    if selected_dur == "Others":
        p_duration = st.text_input("Specify Custom Duration", key=f"dur_txt_{idx}")
    else:
        p_duration = selected_dur

    is_operated = st.radio("Was this case operated during the shift?", ["No", "Yes"], key=f"op_{idx}")

# Processing Multiple Picture Upload Blocks
cropped_pre_list = []
cropped_post_list = []

st.subheader("📸 Radiograph Batch Processing")
img_col1, img_col2 = st.columns(2)

with img_col1:
    pre_files = st.file_uploader("Upload Injury / Pre-Op Images (Multiple Allowed)", type=["jpg", "jpeg", "png"], accept_multiple_files=True, key=f"pre_files_{idx}")
    if pre_files:
        for f_idx, f in enumerate(pre_files):
            st.write(f"🔍 **Crop Pre-Op Image [{f_idx + 1}]:**")
            raw_img = Image.open(f)
            c_img = st_cropper(raw_img, realtime_update=True, box_color='#FF0000', aspect_ratio=None, key=f"crop_pre_{idx}_{f_idx}")
            cropped_pre_list.append(c_img)

with img_col2:
    if is_operated == "Yes":
        post_files = st.file_uploader("Upload Post-Op / Fixation Images (Multiple Allowed)", type=["jpg", "jpeg", "png"], accept_multiple_files=True, key=f"post_files_{idx}")
        if post_files:
            for f_idx, f in enumerate(post_files):
                st.write(f"🔍 **Crop Post-Op Image [{f_idx + 1}]:**")
                raw_img = Image.open(f)
                c_img = st_cropper(raw_img, realtime_update=True, box_color='#00FF00', aspect_ratio=None, key=f"crop_post_{idx}_{f_idx}")
                cropped_post_list.append(c_img)

# Commit Entries to Session Cache
if st.button("➕ Save This Case"):
    if p_name or p_mrn:
        new_case = {
            "name": p_name, "mrn": p_mrn, "age": p_age, "sex": p_sex, "moi": p_moi, "duration": p_duration,
            "doi": p_doi_str,
            "operated": is_operated,
            "pre_imgs": cropped_pre_list.copy(),
            "post_imgs": cropped_post_list.copy() if is_operated == "Yes" else []
        }
        st.session_state.cases.append(new_case)
        st.session_state.case_counter += 1
        st.rerun()
    else:
        st.error("Please provide at least a Patient Identifier or MRN before saving.")

st.markdown("---")

# --- 4. Queue Display & File Processing Engine ---
if st.session_state.cases:
    st.header(f"📋 Presentation Queue ({len(st.session_state.cases)} Cases Added)")
    for i, c in enumerate(st.session_state.cases):
        tag = "🔴 Operated" if c['operated'] == "Yes" else "🟢 Conservative"
        st.write(f"**Case {i+1}:** {c['name']} ({c['age']}/{c['sex']}) — MRN: {c['mrn']} | DOI: {c['doi']} | {tag}")
        
    if st.button("🗑️ Reset Entire Queue"):
        st.session_state.cases = []
        st.session_state.case_counter = 0
        st.rerun()

    st.subheader("🚀 Export Slide Deck")
    if st.button("Compile PowerPoint Presentation"):
        prs = Presentation()
        prs.slide_width = Inches(13.333)
        prs.slide_height = Inches(7.5)
        blank_layout = prs.slide_layouts[6]
        
        # --- Slide 1: Cover Presentation ---
        slide1 = prs.slides.add_slide(blank_layout)
        tb = slide1.shapes.add_textbox(Inches(0.8), Inches(1.0), Inches(11.733), Inches(5.5))
        tf = tb.text_frame
        tf.word_wrap = True
        
        p_title = tf.paragraphs[0]
        p_title.text = f"Orthopedic Department Duty Report\nDate: {presentation_date_str}"
        p_title.font.size = Pt(36)
        p_title.font.bold = True
        
        # Format names dynamically with Dr. prefix verification
        seniors = [clean_dr_name(s) for s in [cons_1, cons_2, cons_3, cons_4, cons_5, cons_6] if s.strip()]
        residents = [clean_dr_name(r) for r in [res1, res2, res3, res4, res5, res6] if r.strip()]
        
        p_team = tf.add_paragraph()
        p_team.text = f"\n🔹 **Consultants on Duty:**\n{', '.join(seniors) if seniors else 'None Specified'}\n\n" \
                     f"🔹 **Residents on Duty:**\n{', '.join(residents) if residents else 'None Specified'}"
        p_team.font.size = Pt(18)

        # --- Case Document Generation ---
        for c in st.session_state.cases:
            # Slide A: Pre-Op Presentation
            slide_pre = prs.slides.add_slide(blank_layout)
            
            # TITLE BOX: Exclusively Patient Name, Age/Sex, and MRN
            title_box = slide_pre.shapes.add_textbox(Inches(0.5), Inches(0.4), Inches(12.333), Inches(0.8))
            tf_title = title_box.text_frame
            p_t = tf_title.paragraphs[0]
            p_t.text = f"{c['name']}  |  {c['age']}/{c['sex']}  |  MRN: {c['mrn']}"
            p_t.font.size = Pt(28)
            p_t.font.bold = True
            
            # DESCRIPTION BOX: Shifted down below title for diagnostic values
            desc_box = slide_pre.shapes.add_textbox(Inches(0.5), Inches(1.2), Inches(12.333), Inches(0.6))
            tf_desc = desc_box.text_frame
            p_d = tf_desc.paragraphs[0]
            p_d.text = f"MOI: {c['moi']}    |    Date of Injury: {c['doi']}    |    Presentation Delay: {c['duration']}    |    Status: Pre-Op"
            p_d.font.size = Pt(15)
            p_d.font.color.rgb = RGBColor(220, 50, 50)
            
            # Mount Multiple Pre-Op Images Side-by-Side
            if c['pre_imgs']:
                n_pre = len(c['pre_imgs'])
                img_w = min(5.8, 12.333 / max(1, n_pre))
                gap = 0.2
                left_start = 0.5 + (12.333 - (n_pre * img_w + (n_pre - 1) * gap)) / 2
                
                for i, img_obj in enumerate(c['pre_imgs']):
                    img_buf = io.BytesIO()
                    img_obj.save(img_buf, format="PNG")
                    img_buf.seek(0)
                    slide_pre.shapes.add_picture(img_buf, Inches(max(0.5, left_start + i * (img_w + gap))), Inches(2.0), width=Inches(img_w))

            # Slide B: Post-Op Presentation (If Active)
            if c['operated'] == "Yes" and c['post_imgs']:
                slide_post = prs.slides.add_slide(blank_layout)
                
                # TITLE BOX (Post-Op Check)
                title_box2 = slide_post.shapes.add_textbox(Inches(0.5), Inches(0.4), Inches(12.333), Inches(0.8))
                tf_title2 = title_box2.text_frame
                p_t2 = tf_title2.paragraphs[0]
                p_t2.text = f"{c['name']}  |  {c['age']}/{c['sex']}  |  MRN: {c['mrn']}"
                p_t2.font.size = Pt(28)
                p_t2.font.bold = True
                
                # DESCRIPTION BOX (Post-Op Check)
                desc_box2 = slide_post.shapes.add_textbox(Inches(0.5), Inches(1.2), Inches(12.333), Inches(0.6))
                tf_desc2 = desc_box2.text_frame
                p_d2 = tf_desc2.paragraphs[0]
                p_d2.text = f"MOI: {c['moi']}    |    Date of Injury: {c['doi']}    |    Status: Post-Operative Fixation Checks"
                p_d2.font.size = Pt(15)
                p_d2.font.color.rgb = RGBColor(50, 180, 50)
                
                # Mount Multiple Post-Op Images Side-by-Side
                n_post = len(c['post_imgs'])
                img_w2 = min(5.8, 12.333 / max(1, n_post))
                gap = 0.2
                left_start2 = 0.5 + (12.333 - (n_post * img_w2 + (n_post - 1) * gap)) / 2
                
                for i, img_obj in enumerate(c['post_imgs']):
                    img_buf = io.BytesIO()
                    img_obj.save(img_buf, format="PNG")
                    img_buf.seek(0)
                    slide_post.shapes.add_picture(img_buf, Inches(max(0.5, left_start2 + i * (img_w2 + gap))), Inches(2.0), width=Inches(img_w2))

        # Output Final File
        final_ppt_buf = io.BytesIO()
        prs.save(final_ppt_buf)
        final_ppt_buf.seek(0)
        
        st.download_button(
            label="💾 Save PowerPoint File to Device",
            data=final_ppt_buf,
            file_name=f"Morning_Report_{presentation_date.strftime('%Y_%m_%d')}.pptx",
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
        )
