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

# Funktion til at loade Excel-data
def load_excel(file):
    return pd.ExcelFile(file)

# Funktion til at hente matching data fra mapping-file
def get_mapping_data(mapping_df, item_no):
    match = mapping_df[mapping_df['{{Product code}}'].astype(str) == item_no]
    if match.empty and "-" in item_no:
        stripped_item_no = item_no.split("-")[0]
        match = mapping_df[mapping_df['{{Product code}}'].astype(str) == stripped_item_no]
    return match.iloc[0] if not match.empty else None

# Funktion til at indsætte tekst i en tabelcelle
def insert_text_in_table(slide, field, value):
    value = str(value) if pd.notna(value) else "N/A"  # Konverter til string og håndter NaN
    for shape in slide.shapes:
        if shape.has_table:
            table = shape.table
            for row in table.rows:
                for cell in row.cells:
                    if field in cell.text:
                        cell.text = cell.text.replace(field, value)

# Funktion til at håndtere billeder
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

# Funktion til at forberede PowerPoint-skabelonen
def prepare_ppt_template(template_file, num_slides):
    prs = pptx.Presentation(template_file)
    slide_layout = prs.slide_layouts[0]
    for _ in range(num_slides):
        prs.slides.add_slide(slide_layout)
    return prs

# Funktion til at indsætte data i slides
def populate_slide(slide, mapping_data):
    text_fields = {
        '{{Product name}}': "Product Name:",
        '{{Product code}}': "Product Code:",
        '{{Product country of origin}}': "Country of origin:",
        '{{Product height}}': "Height:",
        '{{Product width}}': "Width:",
        '{{Product length}}': "Length:",
        '{{Product depth}}': "Depth:",
        '{{Product seat height}}': "Seat Height:",
        '{{Product  diameter}}': "Diameter:",
        '{{CertificateName}}': "Test & certificates for the product:",
        '{{Product Consumption COM}}': "Consumption information for COM:"
    }
    for field, label in text_fields.items():
        value = mapping_data.get(field, "N/A")
        insert_text_in_table(slide, field, str(value))
    
    image_fields = ['{{Product Packshot1}}', '{{Product Lifestyle1}}', '{{Product Lifestyle2}}', '{{Product Lifestyle3}}', '{{Product Lifestyle4}}']
    for field in image_fields:
        for shape in slide.shapes:
            if shape.has_text_frame and field in shape.text_frame.text:
                insert_image(slide, shape, mapping_data.get(field, None))

# Funktion til at generere PowerPoint
def generate_presentation(user_file, mapping_file, stock_file, template_file):
    update_progress(1)
    user_df = pd.read_excel(user_file)
    mapping_df = pd.read_excel(mapping_file, sheet_name=0)
    stock_df = pd.read_excel(stock_file, sheet_name=0)
    
    prs = prepare_ppt_template(template_file, len(user_df) - 1)
    update_progress(2)
    
    for index, row in user_df.iterrows():
        item_no = str(row['Item no'])
        slide = prs.slides[index]
        mapping_data = get_mapping_data(mapping_df, item_no)
        if mapping_data is not None:
            populate_slide(slide, mapping_data)
    
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
