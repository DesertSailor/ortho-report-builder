
import streamlit as st
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from streamlit_cropper import st_cropper
from PIL import Image
import io
import os
import pickle
from datetime import date

# ==========================================================
# CONFIGURATION
# ==========================================================
st.set_page_config(
    page_title="Orthopedic Morning Report Builder",
    layout="wide"
)

st.title("🏥 Orthopedic Morning Report Builder")

BACKUP_PREFIX = "morning_report_"

# ==========================================================
# HELPER FUNCTIONS
# ==========================================================

def clean_dr_name(name):
    """Format doctor names consistently."""
    name = name.strip()

    if not name:
        return ""

    if name.lower() == "none":
        return ""

    if name.lower().startswith("dr"):
        core = name[2:].strip().lstrip(".").strip()
        return f"Dr. {core.title()}"

    return f"Dr. {name.title()}"


def image_to_bytes(img):
    """Convert PIL image to bytes for storage."""
    buffer = io.BytesIO()
    img.save(buffer, format="PNG")
    return buffer.getvalue()


def bytes_to_image(img_bytes):
    """Convert bytes back to PIL image."""
    return Image.open(io.BytesIO(img_bytes))


# ==========================================================
# SESSION PRIVACY
# ==========================================================
st.sidebar.header("🔐 Session Isolation")

user_key = st.sidebar.text_input(
    "Initials / Shift ID",
    value="Resident_1",
    help="Cases are isolated using this identifier."
)

safe_key = "".join(c for c in user_key if c.isalnum()).lower()

if not safe_key:
    safe_key = "default"

BACKUP_FILE = f"{BACKUP_PREFIX}{safe_key}.pkl"


def save_backup():
    with open(BACKUP_FILE, "wb") as f:
        pickle.dump(st.session_state.cases, f)


# ==========================================================
# INITIALIZE SESSION
# ==========================================================
if "active_user" not in st.session_state:
    st.session_state.active_user = ""

if (
    st.session_state.active_user != safe_key
):

    st.session_state.active_user = safe_key

    if os.path.exists(BACKUP_FILE):
        try:
            with open(BACKUP_FILE, "rb") as f:
                st.session_state.cases = pickle.load(f)

        except Exception:
            st.session_state.cases = []

    else:
        st.session_state.cases = []

    st.session_state.case_counter = len(
        st.session_state.cases
    )


if st.session_state.cases:
    st.sidebar.success(
        f"{len(st.session_state.cases)} case(s) recovered."
    )
else:
    st.sidebar.info("New isolated workspace.")


# ==========================================================
# PRESENTATION SETTINGS
# ==========================================================
st.header("📅 Presentation Settings")

presentation_date = st.date_input(
    "Presentation Date",
    date.today()
)

presentation_date_str = (
    presentation_date.strftime(
        "%A, %B %d, %Y"
    )
)

st.markdown("---")


# ==========================================================
# DUTY TEAM
# ==========================================================
st.header("👥 Duty Team")

with st.expander(
    "Consultants and Residents",
    expanded=True
):

    col1, col2 = st.columns(2)

    with col1:

        st.subheader("Consultants")

        eopd_options = [
            "None",
            "Dr. Addisu",
            "Dr. Tilahun",
            "Dr. Samuel",
            "Dr. Mathias",
            "Dr. Tesfatsion",
            "Dr. Kalkidan",
            "Dr. Ashenafi"
        ]

        sport_options = [
            "None",
            "Dr. Mahder",
            "Dr. Mamo"
        ]

        trauma_options = [
            "None",
            "Dr. Abiy",
            "Dr. Milkias",
            "Dr. Beza",
            "Dr. Ibrahim"
        ]

        cons1 = st.selectbox(
            "EOPD/Ward Consultant 1",
            eopd_options
        )

        filtered = [
            x for x in eopd_options
            if x == "None" or x != cons1
        ]

        cons2 = st.selectbox(
            "EOPD/Ward Consultant 2",
            filtered
        )

        cons3 = st.selectbox(
            "Sport Consultant",
            sport_options
        )

        cons4 = st.selectbox(
            "Trauma Consultant",
            trauma_options
        )

        st.text_input(
            "Permanent Consultant",
            value="Dr. Solomon",
            disabled=True
        )

        st.text_input(
            "Permanent Consultant",
            value="Dr. Dawit",
            disabled=True
        )

    with col2:

        st.subheader("Residents")

        res1 = st.text_input("Resident 1")
        res2 = st.text_input("Resident 2")
        res3 = st.text_input("Resident 3")
        res4 = st.text_input("Resident 4")
        res5 = st.text_input("Resident 5")
        res6 = st.text_input("Resident 6")


st.markdown("---")


# ==========================================================
# ADD CASE
# ==========================================================
st.header("📝 Add Patient Case")

idx = st.session_state.case_counter

left, right = st.columns(2)

with left:

    p_name = st.text_input(
        "Patient Initials / Name",
        key=f"name_{idx}"
    )

    raw_mrn = st.text_input(
        "MRN (Numbers Only)",
        key=f"mrn_{idx}"
    )

    mrn_valid = True

    if raw_mrn:

        if not raw_mrn.isdigit():
            st.error(
                "MRN must contain numbers only."
            )

            mrn_valid = False

    p_age = st.text_input(
        "Age",
        key=f"age_{idx}"
    )

    p_sex = st.selectbox(
        "Sex",
        ["M", "F"],
        key=f"sex_{idx}"
    )


with right:

    doi = st.date_input(
        "Date of Injury",
        date.today(),
        key=f"doi_{idx}"
    )

    doi_str = doi.strftime("%b %d, %Y")

    moi_options = [
        "RTA",
        "Fall from Height",
        "FDA",
        "Sports Injury",
        "Direct Blow",
        "Others"
    ]

    selected_moi = st.selectbox(
        "MOI",
        moi_options,
        key=f"moi_{idx}"
    )

    if selected_moi == "Others":
        p_moi = st.text_input(
            "Specify MOI",
            key=f"custom_moi_{idx}"
        )
    else:
        p_moi = selected_moi

    p_duration = st.text_input(
        "Duration",
        key=f"duration_{idx}"
    )

    operated = st.radio(
        "Operated During Shift?",
        ["No", "Yes"],
        key=f"op_{idx}"
    )
    # app.py (Part 2)


# ==========================================================
# NOTES
# ==========================================================
st.subheader("💡 Additional Information")

p_notes = st.text_area(
    "Remarks / Neurovascular Status / Plan",
    placeholder=(
        "Examples:\n"
        "- Open fracture Gustilo IIIA\n"
        "- NV intact\n"
        "- Planned ORIF tomorrow\n"
        "- Awaiting CT scan"
    ),
    key=f"notes_{idx}"
)

# ==========================================================
# RADIOGRAPHS
# ==========================================================
cropped_pre = []
cropped_post = []

st.subheader("📸 Radiographs")

img_col1, img_col2 = st.columns(2)

with img_col1:

    st.markdown("### Injury / Pre-Op")

    pre_files = st.file_uploader(
        "Upload Pre-Op Images",
        type=["jpg", "jpeg", "png"],
        accept_multiple_files=True,
        key=f"pre_files_{idx}"
    )

    if pre_files:

        for i, uploaded in enumerate(pre_files):

            st.write(f"Pre-Op Image {i+1}")

            image = Image.open(uploaded)

            cropped = st_cropper(
                image,
                realtime_update=False,
                box_color="#FF0000",
                aspect_ratio=None,
                key=f"crop_pre_{idx}_{i}"
            )

            st.image(
                cropped,
                caption="Preview",
                use_container_width=True
            )

            cropped_pre.append(cropped)


with img_col2:

    if operated == "Yes":

        st.markdown("### Post-Op")

        post_files = st.file_uploader(
            "Upload Post-Op Images",
            type=["jpg", "jpeg", "png"],
            accept_multiple_files=True,
            key=f"post_files_{idx}"
        )

        if post_files:

            for i, uploaded in enumerate(post_files):

                st.write(f"Post-Op Image {i+1}")

                image = Image.open(uploaded)

                cropped = st_cropper(
                    image,
                    realtime_update=False,
                    box_color="#00FF00",
                    aspect_ratio=None,
                    key=f"crop_post_{idx}_{i}"
                )

                st.image(
                    cropped,
                    caption="Preview",
                    use_container_width=True
                )

                cropped_post.append(cropped)


# ==========================================================
# SAVE CASE
# ==========================================================
save_col1, save_col2 = st.columns([3, 1])

with save_col1:

    if st.button("➕ Save This Case"):

        if not p_name.strip():

            st.error(
                "Patient name/initials are required."
            )

        elif not mrn_valid:

            st.error(
                "Please provide a valid MRN."
            )

        else:

            case = {
                "name": p_name.strip(),
                "mrn": raw_mrn.strip(),
                "age": p_age.strip(),
                "sex": p_sex,
                "doi": doi_str,
                "moi": p_moi,
                "duration": p_duration.strip(),
                "operated": operated,
                "notes": p_notes.strip(),
                "pre_imgs": [
                    image_to_bytes(img)
                    for img in cropped_pre
                ],
                "post_imgs": [
                    image_to_bytes(img)
                    for img in cropped_post
                ]
            }

            st.session_state.cases.append(case)

            st.session_state.case_counter += 1

            save_backup()

            st.success(
                f"Case saved successfully. "
                f"Queue: {len(st.session_state.cases)}"
            )

            st.rerun()


with save_col2:

    if st.button("🗑️ Clear Queue"):

        st.session_state.confirm_clear = True


if st.session_state.get("confirm_clear", False):

    st.warning(
        "Delete ALL cases from this session?"
    )

    yes, no = st.columns(2)

    with yes:

        if st.button("YES"):

            st.session_state.cases = []

            st.session_state.case_counter = 0

            st.session_state.confirm_clear = False

            if os.path.exists(BACKUP_FILE):

                os.remove(BACKUP_FILE)

            st.success("Queue cleared.")

            st.rerun()

    with no:

        if st.button("NO"):

            st.session_state.confirm_clear = False

            st.rerun()


st.markdown("---")


# ==========================================================
# CASE QUEUE
# ==========================================================
st.header("📋 Case Queue")

if not st.session_state.cases:

    st.info("No cases saved yet.")

else:

    st.caption(
        f"{len(st.session_state.cases)} case(s) saved"
    )

    for case_index, case in enumerate(
        st.session_state.cases
    ):

        title = (
            f"{case_index + 1}. "
            f"{case['name']} "
            f"({case['age']}/{case['sex']})"
        )

        with st.expander(title):

            st.write(
                f"**MRN:** {case['mrn']}"
            )

            st.write(
                f"**MOI:** {case['moi']}"
            )

            st.write(
                f"**DOI:** {case['doi']}"
            )

            st.write(
                f"**Duration:** "
                f"{case['duration']}"
            )

            st.write(
                f"**Operated:** "
                f"{case['operated']}"
            )

            if case["notes"]:

                st.write(
                    f"**Notes:** "
                    f"{case['notes']}"
                )

            pre_count = len(case["pre_imgs"])

            post_count = len(case["post_imgs"])

            st.write(
                f"Pre-Op Images: {pre_count}"
            )

            st.write(
                f"Post-Op Images: {post_count}"
            )

            # -------------------------
            # PREVIEW IMAGES
            # -------------------------
            if case["pre_imgs"]:

                st.markdown("**Pre-Op:**")

                cols = st.columns(
                    min(3, len(case["pre_imgs"]))
                )

                for i, img_bytes in enumerate(
                    case["pre_imgs"]
                ):

                    cols[
                        i % len(cols)
                    ].image(
                        bytes_to_image(img_bytes),
                        use_container_width=True
                    )

            if case["post_imgs"]:

                st.markdown("**Post-Op:**")

                cols = st.columns(
                    min(3, len(case["post_imgs"]))
                )

                for i, img_bytes in enumerate(
                    case["post_imgs"]
                ):

                    cols[
                        i % len(cols)
                    ].image(
                        bytes_to_image(img_bytes),
                        use_container_width=True
                    )

            # -------------------------
            # DELETE SINGLE CASE
            # -------------------------
            if st.button(
                f"❌ Delete Case {case_index + 1}",
                key=f"delete_{case_index}"
            ):

                st.session_state.cases.pop(
                    case_index
                )

                st.session_state.case_counter = len(
                    st.session_state.cases
                )

                save_backup()

                st.success("Case deleted.")

                st.rerun()

# app.py (Part 3)


# ==========================================================
# POWERPOINT EXPORT
# ==========================================================
st.markdown("---")
st.header("🚀 Export PowerPoint")

if st.button("📥 Generate Morning Report PPT"):

    if not st.session_state.cases:
        st.warning("No cases available to export.")

    else:

        prs = Presentation()

        # 16:9 Layout
        prs.slide_width = Inches(13.333)
        prs.slide_height = Inches(7.5)

        blank = prs.slide_layouts[6]

        # ==================================================
        # COVER SLIDE
        # ==================================================
        slide = prs.slides.add_slide(blank)

        title_box = slide.shapes.add_textbox(
            Inches(0.8),
            Inches(1.5),
            Inches(12),
            Inches(2)
        )

        tf = title_box.text_frame

        p = tf.paragraphs[0]
        p.text = "ORTHOPEDIC MORNING REPORT"
        p.font.size = Pt(30)
        p.font.bold = True

        p2 = tf.add_paragraph()
        p2.text = presentation_date_str
        p2.font.size = Pt(22)

        p3 = tf.add_paragraph()
        p3.text = (
            "Department of Orthopaedic Surgery"
        )
        p3.font.size = Pt(20)

        # ==================================================
        # DUTY TEAM SLIDE
        # ==================================================
        slide = prs.slides.add_slide(blank)

        title = slide.shapes.add_textbox(
            Inches(0.5),
            Inches(0.3),
            Inches(12),
            Inches(1)
        )

        tf = title.text_frame

        p = tf.paragraphs[0]
        p.text = "Duty Team"
        p.font.size = Pt(28)
        p.font.bold = True

        consultants = [
            cons1,
            cons2,
            cons3,
            cons4,
            "Dr. Solomon",
            "Dr. Dawit"
        ]

        consultants = [
            clean_dr_name(x)
            for x in consultants
            if clean_dr_name(x)
        ]

        residents = [
            res1,
            res2,
            res3,
            res4,
            res5,
            res6
        ]

        residents = [
            clean_dr_name(x)
            for x in residents
            if x.strip()
        ]

        body = slide.shapes.add_textbox(
            Inches(0.7),
            Inches(1.3),
            Inches(12),
            Inches(5)
        )

        tf = body.text_frame

        p = tf.paragraphs[0]
        p.text = (
            "Consultants on Duty:\n"
            + "\n".join(consultants)
        )
        p.font.size = Pt(22)

        p = tf.add_paragraph()
        p.text = (
            "\nResidents on Duty:\n"
            + "\n".join(residents)
        )
        p.font.size = Pt(22)

        # ==================================================
        # PRIORITIZE OPERATED CASES
        # ==================================================
        sorted_cases = (
            [
                c for c in st.session_state.cases
                if c["operated"] == "Yes"
            ]
            +
            [
                c for c in st.session_state.cases
                if c["operated"] == "No"
            ]
        )

        # ==================================================
        # CASE SLIDES
        # ==================================================
        for case_number, case in enumerate(
            sorted_cases,
            start=1
        ):

            all_images = []

            for img in case["pre_imgs"]:
                all_images.append(img)

            for img in case["post_imgs"]:
                all_images.append(img)

            image_chunks = [
                all_images[i:i + 2]
                for i in range(
                    0,
                    len(all_images),
                    2
                )
            ]

            if not image_chunks:
                image_chunks = [[]]

            # ==========================================
            # MULTIPLE SLIDES IF NEEDED
            # ==========================================
            for chunk_index, chunk in enumerate(
                image_chunks
            ):

                slide = prs.slides.add_slide(blank)

                # --------------------------------------
                # FIRST SLIDE OF CASE
                # --------------------------------------
                if chunk_index == 0:

                    header = slide.shapes.add_textbox(
                        Inches(0.4),
                        Inches(0.3),
                        Inches(12.5),
                        Inches(0.8)
                    )

                    tf = header.text_frame

                    p = tf.paragraphs[0]

                    p.text = (
                        f"{case_number}. "
                        f"{case['name']} | "
                        f"{case['age']}/{case['sex']} | "
                        f"MRN: {case['mrn']}"
                    )

                    p.font.size = Pt(26)
                    p.font.bold = True

                    info = slide.shapes.add_textbox(
                        Inches(0.4),
                        Inches(1.0),
                        Inches(12.5),
                        Inches(1)
                    )

                    tf = info.text_frame

                    p = tf.paragraphs[0]

                    p.text = (
                        f"MOI: {case['moi']}   |   "
                        f"DOI: {case['doi']}   |   "
                        f"Duration: {case['duration']}"
                    )

                    p.font.size = Pt(16)

                    if case["operated"] == "Yes":

                        p2 = tf.add_paragraph()

                        p2.text = (
                            "Status: OPERATED"
                        )

                        p2.font.size = Pt(16)

                        p2.font.bold = True

                        p2.font.color.rgb = (
                            RGBColor(0, 100, 0)
                        )

                    if case["notes"]:

                        p3 = tf.add_paragraph()

                        p3.text = case["notes"]

                        p3.font.size = Pt(14)

                        p3.font.italic = True

                    image_top = Inches(2.2)

                else:

                    image_top = Inches(0.8)

                # --------------------------------------
                # IMAGE LAYOUT
                # --------------------------------------
                if len(chunk) == 1:

                    img = bytes_to_image(
                        chunk[0]
                    )

                    buffer = io.BytesIO()

                    img.save(
                        buffer,
                        format="PNG"
                    )

                    buffer.seek(0)

                    slide.shapes.add_picture(
                        buffer,
                        Inches(3.4),
                        image_top,
                        width=Inches(6.5)
                    )

                elif len(chunk) == 2:

                    # Left image
                    img1 = bytes_to_image(
                        chunk[0]
                    )

                    buf1 = io.BytesIO()

                    img1.save(
                        buf1,
                        format="PNG"
                    )

                    buf1.seek(0)

                    slide.shapes.add_picture(
                        buf1,
                        Inches(0.5),
                        image_top,
                        width=Inches(6)
                    )

                    # Right image
                    img2 = bytes_to_image(
                        chunk[1]
                    )

                    buf2 = io.BytesIO()

                    img2.save(
                        buf2,
                        format="PNG"
                    )

                    buf2.seek(0)

                    slide.shapes.add_picture(
                        buf2,
                        Inches(6.8),
                        image_top,
                        width=Inches(6)
                    )

# app.py (Part 4)

        # ==================================================
        # SAVE PRESENTATION TO MEMORY
        # ==================================================
        ppt_buffer = io.BytesIO()

        prs.save(ppt_buffer)

        ppt_buffer.seek(0)

        filename = (
            "Morning_Report_"
            + presentation_date.strftime(
                "%Y_%m_%d"
            )
            + ".pptx"
        )

        st.success(
            "PowerPoint generated successfully."
        )

        st.download_button(
            label="💾 Download Morning Report PPT",
            data=ppt_buffer,
            file_name=filename,
            mime=(
                "application/vnd.openxmlformats-"
                "officedocument.presentationml."
                "presentation"
            )
        )

        # ==================================================
        # EXPORT SUMMARY
        # ==================================================
        operated_cases = len(
            [
                c
                for c in st.session_state.cases
                if c["operated"] == "Yes"
            ]
        )

        non_operated_cases = (
            len(st.session_state.cases)
            - operated_cases
        )

        st.info(
            f"Export Summary\n\n"
            f"Total Cases: "
            f"{len(st.session_state.cases)}\n"
            f"Operated: "
            f"{operated_cases}\n"
            f"Non-Operated: "
            f"{non_operated_cases}"
        )



