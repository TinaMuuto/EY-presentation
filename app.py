import streamlit as st
import pandas as pd
import pptx
from pptx.util import Inches
import requests
from io import BytesIO
from PIL import Image
import tempfile
import time

# Step-indikator
steps = ["Upload fil", "Behandler data", "Genererer slides", "Færdig - Klar til download"]
progress = st.progress(0)
status_text = st.empty()

def update_progress(step):
    """Opdaterer progress bar og status-tekst"""
    progress.progress((step + 1) / len(steps))
    status_text.text(steps[step])
    time.sleep(1)

# Indlæs Excel-data
def load_excel(file):
    return pd.ExcelFile(file)

# Funktion til at hente matching data fra mapping-file
def get_mapping_data(mapping_df, item_no):
    item_no = str(item_no).strip()
    mapping_df.columns = mapping_df.columns.str.strip().str.replace(r"[{}]", "", regex=True)
    
    col_name = "Product code" if "Product code" in mapping_df.columns else None
    if col_name is None:
        return None
    
    mapping_df[col_name] = mapping_df[col_name].astype(str).str.strip()
    match = mapping_df[mapping_df[col_name] == item_no]
    
    if match.empty and "-" in item_no:
        stripped_item_no = item_no.split("-")[0]
        match = mapping_df[mapping_df[col_name] == stripped_item_no]
    
    return match.iloc[0] if not match.empty else None

# Kopier en slide
def duplicate_slide(prs, slide_index):
    source_slide = prs.slides[slide_index]
    slide_layout = source_slide.slide_layout
    new_slide = prs.slides.add_slide(slide_layout)
    for shape in source_slide.shapes:
        if shape.has_text_frame:
            new_shape = new_slide.shapes.add_textbox(shape.left, shape.top, shape.width, shape.height)
            new_shape.text_frame.text = shape.text_frame.text
    return new_slide

# Indsæt tekst i tekstbokse
def insert_text(slide, field, value):
    value = str(value) if pd.notna(value) else "N/A"
    for shape in slide.shapes:
        if shape.has_text_frame and field in shape.text_frame.text:
            shape.text_frame.text = shape.text_frame.text.replace(field, value)

# Indsæt billeder
def insert_image(slide, placeholder, image_url):
    try:
        if isinstance(image_url, str) and image_url.startswith("http"):
            response = requests.get(image_url)
            img = Image.open(BytesIO(response.content))
            img = img.convert("RGB")
            temp_img = tempfile.NamedTemporaryFile(delete=False, suffix=".jpg")
            img.save(temp_img.name, format="JPEG")
            placeholder.insert_picture(temp_img.name)
    except:
        print(f"Kunne ikke hente billede: {image_url}")

# Generer PowerPoint
def generate_presentation(user_file, mapping_file, stock_file, template_file):
    update_progress(1)
    user_df = pd.read_excel(user_file)
    mapping_df = pd.read_excel(mapping_file, sheet_name=0)
    stock_df = pd.read_excel(stock_file, sheet_name=0)
    
    stock_df.fillna("N/A", inplace=True)
    mapping_df.columns = mapping_df.columns.str.strip().str.replace(r"[{}]", "", regex=True)
    user_df["Item no"] = user_df["Item no"].astype(str).str.strip()
    
    prs = pptx.Presentation(template_file)
    update_progress(2)
    
    for index, row in user_df.iterrows():
        item_no = str(row['Item no']).strip()
        new_slide = duplicate_slide(prs, 0)
        mapping_data = get_mapping_data(mapping_df, item_no)
        
        if mapping_data is None:
            st.warning(f"Ingen match for Item no: {item_no}, prøver med trimmet version...")
            item_no_stripped = item_no.split("-")[0] if "-" in item_no else item_no
            mapping_data = get_mapping_data(mapping_df, item_no_stripped)
        
        if mapping_data is None:
            st.error(f"Ingen data fundet for Item no: {item_no}. Springes over.")
            continue
        
        # Indsæt tekst
        for field in ['Product name', 'Product code', 'Product country of origin']:
            value = str(mapping_data.get(field, "N/A"))
            insert_text(new_slide, field, value)
        
        # Indsæt billeder
        for field in ['Product Packshot1', 'Product Lifestyle1', 'Product Lifestyle2', 'Product Lifestyle3', 'Product Lifestyle4']:
            image_url = str(mapping_data.get(field, "")).strip()
            if image_url:
                insert_image(new_slide, new_slide.shapes[0], image_url)
            else:
                st.warning(f"Billede mangler for {field} på slide {index+1}")
    
    output_file = tempfile.NamedTemporaryFile(delete=False, suffix=".pptx").name
    prs.save(output_file)
    update_progress(3)
    return output_file

# Streamlit UI
st.title("PowerPoint Generator")
user_file = st.file_uploader("Upload din Excel-fil", type=["xlsx"])
mapping_file = "mapping-file.xlsx"
stock_file = "stock.xlsx"
template_file = "template-generator.pptx"

if user_file:
    ppt_file = generate_presentation(user_file, mapping_file, stock_file, template_file)
    st.success("Præsentationen er klar! Download den her:")
    st.download_button(label="Download PowerPoint", 
                       data=open(ppt_file, "rb").read(), 
                       file_name="generated_presentation.pptx",
                       mime="application/vnd.openxmlformats-officedocument.presentationml.presentation")
