import streamlit as st
import pandas as pd
import os
import io
import re
from pptx import Presentation
from pptx.util import Inches
from urllib.request import urlopen
from PIL import Image

def find_column(df, possible_names):
    """Finder en kolonne uanset variation i navn."""
    for name in possible_names:
        matches = [col for col in df.columns if col.lower().strip() == name.lower().strip()]
        if matches:
            return matches[0]  # Returner den første matchende kolonne
    return None  # Returner None, hvis ingen match findes

def clean_variant_key(value):
    """Fjerner '- config' og gør opslag ikke-case sensitive."""
    if isinstance(value, str):
        value = value.lower().replace(" - config", "").strip()
    return value

def match_item_number(df, item_number_col, item_number):
    """Slår op på Item Nummer og håndterer delvise matches."""
    if item_number is None or pd.isna(item_number):
        return None  # Hvis varenummeret mangler, returner None
    
    df = df.dropna(subset=[item_number_col])  # Fjern rækker, hvor opslag ikke kan ske
    
    match = df[df[item_number_col].str.lower() == str(item_number).lower()]
    if not match.empty:
        return match.iloc[0]
    
    # Hvis ikke eksakt match, prøv at matche første del før "-"
    partial_key = str(item_number).split('-')[0].strip()
    match = df[df[item_number_col].str.lower().str.startswith(partial_key.lower())]
    return match.iloc[0] if not match.empty else None

def insert_image(slide, placeholder_name, image_url):
    """Indsætter et billede fra en URL i et specifikt pladsholderfelt."""
    if image_url and isinstance(image_url, str) and image_url.startswith("http"):
        try:
            with urlopen(image_url) as img_response:
                img = Image.open(img_response)
                img_io = io.BytesIO()
                img.save(img_io, format='PNG')
                img_io.seek(0)
                
                for shape in slide.shapes:
                    if shape.has_text_frame and placeholder_name in shape.text:
                        left, top, width, height = shape.left, shape.top, shape.width, shape.height
                        slide.shapes.add_picture(img_io, left, top, width, height)
                        shape.text = ""  # Fjern placeholder-teksten
        except Exception as e:
            print(f"Fejl ved indsættelse af billede fra {image_url}: {e}")

def generate_ppt(user_data, variant_data, lifestyle_data, line_drawing_data, template_path):
    prs = Presentation(template_path)
    base_slide = prs.slides[0]  # Brug første slide som skabelon
    
    # Hardcoded instruktioner - Mapping mellem PowerPoint felter og datakilder
    instruktioner = [
        {"ppt_field": "Product Name", "source": "User upload sheet", "headline": "Navn: "},
        {"ppt_field": "Product Code", "source": "User upload sheet", "headline": "Kode: "},
        {"ppt_field": "Variant Description", "source": "EY - variant master data", "headline": "Variant: "},
        {"ppt_field": "Product Lifestyle1", "source": "EY - lifestyle", "headline": ""},
        {"ppt_field": "Product Line Drawing1", "source": "EY - line drawing", "headline": ""},
    ]
    
    # Find korrekt kolonnenavn for varenummer
    possible_item_cols = ["Item Nummer", "Item Number", "item number", "Item no", "ITEM NO", "Item No"]
    item_number_col = find_column(user_data, possible_item_cols)
    if not item_number_col:
        st.error("Fejl: Kolonnen med varenummer blev ikke fundet. Sørg for, at din fil har en af følgende kolonnenavne: " + ", ".join(possible_item_cols))
        return None
    
    for _, row in user_data.iterrows():
        item_number = row[item_number_col] if pd.notna(row[item_number_col]) else None
        matched_row = match_item_number(variant_data, "VariantKey", item_number)
        
        slide = prs.slides.add_slide(base_slide.slide_layout)  # Kopi af første slide
        
        for instr in instruktioner:
            ppt_field = instr['ppt_field']
            field_source = instr['source']
            headline = instr['headline']
            
            value = ""
            if "User upload sheet" in field_source:
                value = row.get(ppt_field, "")
            elif "EY - variant master data" in field_source and matched_row is not None:
                value = matched_row.get(ppt_field, "")
            
            if "lifestyle" in ppt_field.lower() or "line drawing" in ppt_field.lower():
                insert_image(slide, ppt_field, value)  # Indsæt billede, hvis det er en URL
            elif value and isinstance(value, str):
                for shape in slide.shapes:
                    if shape.has_text_frame and ppt_field in shape.text:
                        shape.text = f"{headline} {value}" if headline else value
    
    ppt_bytes = io.BytesIO()
    prs.save(ppt_bytes)
    ppt_bytes.seek(0)
    return ppt_bytes

st.title("PowerPoint Generator")

# Fast PowerPoint skabelon (ingen upload-mulighed)
template_path = "Appendix 1 - Ancillary Furniture and Accessories Catalogue _ CLE.pptx"

# Fastlagte datafiler i GitHub
variant_data_path = "EY - variant master data.xlsx"
lifestyle_data_path = "EY - lifestyle.xlsx"
line_drawing_data_path = "EY - line drawing.xlsx"

# Upload kun produktlisten
user_file = st.file_uploader("Upload brugers produktliste (Excel)", type=["xlsx"])

if st.button("Generér PowerPoint") and user_file:
    try:
        user_data = pd.read_excel(user_file)
        variant_data = pd.read_excel(variant_data_path)
        lifestyle_data = pd.read_excel(lifestyle_data_path)
        line_drawing_data = pd.read_excel(line_drawing_data_path)
        
        ppt_bytes = generate_ppt(user_data, variant_data, lifestyle_data, line_drawing_data, template_path)
        
        if ppt_bytes:
            st.success("PowerPoint genereret!")
            st.download_button(
                label="Download PowerPoint",
                data=ppt_bytes,
                file_name="Generated_Presentation.pptx",
                mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
            )
    except Exception as e:
        st.error(f"En fejl opstod: {str(e)}. Tjek din uploadede fil og prøv igen.")
