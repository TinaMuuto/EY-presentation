import streamlit as st
import pandas as pd
import pptx
from pptx.util import Inches
import requests
from io import BytesIO
from PIL import Image
import tempfile

# Funktion til at loade Excel-data
def load_excel(file):
    return pd.ExcelFile(file)

# Funktion til at hente matching data fra mapping-file
def get_mapping_data(mapping_df, item_no):
    """ Finder den relevante række baseret på Item no i mapping-file."""
    match = mapping_df[mapping_df['{{Product code}}'].astype(str) == item_no]
    if match.empty and "-" in item_no:
        stripped_item_no = item_no.split("-")[0]
        match = mapping_df[mapping_df['{{Product code}}'].astype(str) == stripped_item_no]
    return match.iloc[0] if not match.empty else None

# Funktion til at indsætte tekst i slide
def insert_text(slide, placeholder, label, value):
    if placeholder and pd.notna(value):
        placeholder.text = f"{label}\n{value}"

# Funktion til at håndtere billeder
def insert_image(slide, placeholder, image_url):
    try:
        if isinstance(image_url, str) and image_url.startswith("http"):
            response = requests.get(image_url)
            img = Image.open(BytesIO(response.content))
            img = img.convert("RGB")  # Konverter fra TIFF hvis nødvendigt
            temp_img = tempfile.NamedTemporaryFile(delete=False, suffix=".jpg")
            img.save(temp_img.name, format="JPEG")
            placeholder.insert_picture(temp_img.name)
    except:
        print(f"Kunne ikke hente billede: {image_url}")

# Funktion til at generere PPTX
def generate_ppt(user_file, mapping_file, stock_file, template_file):
    user_df = pd.read_excel(user_file)
    mapping_df = pd.read_excel(mapping_file, sheet_name=0)
    stock_df = pd.read_excel(stock_file, sheet_name=0)
    
    prs = pptx.Presentation(template_file)
    slide_layout = prs.slide_layouts[0]
    
    for _, row in user_df.iterrows():
        item_no = str(row['Item no'])
        product_name = row['Product name']
        mapping_data = get_mapping_data(mapping_df, item_no)
        
        if mapping_data is None:
            st.warning(f"Kunne ikke finde match for Item no: {item_no}")
            continue  # Spring denne iteration over
        
        slide = prs.slides.add_slide(slide_layout)
        text_placeholders = {shape.text_frame.text: shape.text_frame for shape in slide.shapes if shape.has_text_frame}

        # Indsæt tekstfelter
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
            insert_text(slide, text_placeholders.get(field), label, mapping_data.get(field, "N/A"))
        
        # Indsæt billeder kun hvis de findes
        image_fields = ['{{Product Packshot1}}', '{{Product Lifestyle1}}', '{{Product Lifestyle2}}', '{{Product Lifestyle3}}', '{{Product Lifestyle4}}']
        for field in image_fields:
            insert_image(slide, text_placeholders.get(field), mapping_data.get(field, None))
    
    output_file = tempfile.NamedTemporaryFile(delete=False, suffix=".pptx").name
    prs.save(output_file)
    return output_file

# Streamlit UI
st.title("PowerPoint Generator")
user_file = st.file_uploader("Upload din Excel-fil", type=["xlsx"])
mapping_file = "mapping-file.xlsx"  # Forventet fil på GitHub
stock_file = "stock.xlsx"  # Forventet fil på GitHub
template_file = "template-generator.pptx"  # Forventet fil på GitHub

if user_file:
    ppt_file = generate_ppt(user_file, mapping_file, stock_file, template_file)
    st.download_button(label="Download PowerPoint", data=open(ppt_file, "rb").read(), file_name="generated_presentation.pptx", mime="application/vnd.openxmlformats-officedocument.presentationml.presentation")
