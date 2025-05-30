import os
import zipfile
import tempfile
from io import BytesIO

import pandas as pd
import streamlit as st
from docxtpl import DocxTemplate, InlineImage
from docx.shared import Inches
from PIL import Image as PILImage

# ─────────────────────────────────────────────────────────────────────────────
# CONFIG
# ─────────────────────────────────────────────────────────────────────────────

# Path to your constant Word template (must live alongside app.py)
TEMPLATE_PATH = "ApplicationForm.docx"

# Fixed target width for inserted photos (adjust to your frame width)
PHOTO_WIDTH_INCHES = 2

# ─────────────────────────────────────────────────────────────────────────────
# STREAMLIT UI
# ─────────────────────────────────────────────────────────────────────────────

st.set_page_config(page_title="Auto-Fill Forms", layout="wide")
st.title("📄 Auto-Fill Application Forms")
st.markdown(
    """
    1. Upload your **Excel** (.xlsx) with all the data rows.  
    2. Upload a **ZIP** containing a folder per `Application_Number`, each with `photo1.jpg`, `photo2.jpg`, `photo3.jpg`.  
    3. Click **Generate Forms** and download the filled DOCX bundle.
    """
)

excel_file   = st.file_uploader("1️⃣ Upload Excel file", type=["xlsx"])
photos_zip   = st.file_uploader("2️⃣ Upload photos ZIP", type=["zip"])
generate_btn = st.button("🛠️ Generate Forms")

# ─────────────────────────────────────────────────────────────────────────────
# PROCESSING
# ─────────────────────────────────────────────────────────────────────────────

if generate_btn:
    if not excel_file or not photos_zip:
        st.error("Please upload both the Excel and the photos ZIP before generating.")
        st.stop()

    # Read Excel into DataFrame
    try:
        df = pd.read_excel(excel_file)
    except Exception as e:
        st.error(f"Failed to read Excel file: {e}")
        st.stop()

    # Create temp dirs for photos and outputs
    with tempfile.TemporaryDirectory() as tmpdir:
        photos_dir = os.path.join(tmpdir, "photos")
        os.makedirs(photos_dir, exist_ok=True)

        outputs_dir = os.path.join(tmpdir, "outputs")
        os.makedirs(outputs_dir, exist_ok=True)

        # Unzip photos
        with zipfile.ZipFile(photos_zip, "r") as z:
            z.extractall(photos_dir)

        # Loop over each row to build & save a DOCX
        for idx, row in df.iterrows():
            tpl = DocxTemplate(TEMPLATE_PATH)

            # Build context for all non-photo columns
            context = {
                col: row[col]
                for col in df.columns
                if not col.lower().startswith("photo")
            }

            # Insert photos 1–3
            app_no = str(row.get("Application_Number", "")).strip()
            app_folder = os.path.join(photos_dir, app_no)

            for i in (1, 2, 3):
                key = f"photo{i}"
                img_obj = ""  # default blank

                if os.path.isdir(app_folder):
                    # search for jpg / jpeg / png
                    for ext in ("jpg", "jpeg", "png"):
                        img_path = os.path.join(app_folder, f"{key}.{ext}")
                        if os.path.isfile(img_path):
                            # preserve aspect ratio
                            with PILImage.open(img_path) as im:
                                w, h = im.size
                            ratio = h / w
                            tgt_w = Inches(PHOTO_WIDTH_INCHES)
                            tgt_h = tgt_w * ratio

                            img_obj = InlineImage(
                                tpl, img_path, width=tgt_w, height=tgt_h
                            )
                            break

                if not img_obj:
                    st.warning(f"Missing `{key}.*` in folder `{app_no}/`")
                context[key] = img_obj

            # Render and save this document
            tpl.render(context)
            out_name = f"filled_form_{idx+1}.docx"
            out_path = os.path.join(outputs_dir, out_name)
            tpl.save(out_path)

        # Bundle all outputs into an in-memory ZIP
        zip_buffer = BytesIO()
        with zipfile.ZipFile(zip_buffer, "w") as zipf:
            for fname in sorted(os.listdir(outputs_dir)):
                file_path = os.path.join(outputs_dir, fname)
                zipf.write(file_path, arcname=fname)
        zip_buffer.seek(0)

        # Let user download
        st.download_button(
            label="✅ Download All Filled Forms",
            data=zip_buffer,
            file_name="filled_forms.zip",
            mime="application/zip",
        )
