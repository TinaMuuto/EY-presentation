import streamlit as st
import pandas as pd
import os
import io
import re
from pptx import Presentation
from pptx.util import Inches

def clean_variant_key(value):
    """Fjerner '- config' og gør opslag ikke-case sensitive."""
    if isinstance(value, str):
        value = value.lower().replace(" - config", "").strip()
    return value

def match_item_number(df, item_number):
    """Slår op på Item Nummer og håndterer delvise matches."""
    match = df[df['Item Nummer'].str.lower() == item_number.lower()]
    if not match.empty:
        return match.iloc[0]
    
    # Hvis ikke eksakt match, prøv at matche første del før "-"
    partial_key = item_number.split('-')[0].strip()
    match = df[df['Item Nummer'].str.lower().str.startswith(partial_key.lower())]
    return match.iloc[0] if not match.empty else None

def generate_ppt(user_data, variant_data, lifestyle_data, line_drawing_data, instruktioner, template_path):
    prs = Presentation(template_path)
    
    for _, row in user_data.iterrows():
        item_number = row['Item Nummer']
        matched_row = match_item_number(variant_data, item_number)
        
        slide = prs.slides.add_slide(prs.slide_layouts[5])
        
        for _, instr_row in instruktioner.iterrows():
            ppt_field = instr_row['Field name power point']
            field_source = instr_row['Instruktion']
            headline = instr_row['Field headline']
            
            value = ""
            if "User upload sheet" in field_source:
                value = row.get(ppt_field.replace("{{", "").replace("}}", ""), "")
            elif "EY - variant master data" in field_source and matched_row is not None:
                value = matched_row.get(ppt_field.replace("{{", "").replace("}}", ""), "")
            
            if value and isinstance(value, str):
                for shape in slide.shapes:
                    if shape.has_text_frame and ppt_field in shape.text:
                        shape.text = f"{headline} {value}" if headline else value
    
    ppt_bytes = io.BytesIO()
    prs.save(ppt_bytes)
    ppt_bytes.seek(0)
    return ppt_bytes

st.title("PowerPoint Generator 📊")

user_file = st.file_uploader("Upload brugers Excel-fil", type=["xlsx"])
variant_file = st.file_uploader("Upload 'EY - variant master data'", type=["xlsx"])
lifestyle_file = st.file_uploader("Upload 'EY - lifestyle'", type=["xlsx"])
line_drawing_file = st.file_uploader("Upload 'EY - line drawing'", type=["xlsx"])
instruktioner_file = st.file_uploader("Upload 'Instruktioner'", type=["xlsx"])
template_default = "Appendix 1 - Ancillary Furniture and Accessories Catalogue _ CLE.pptx"
template_file = st.file_uploader("Upload PowerPoint skabelon", type=["pptx"])
if template_file is None:
    template_file = template_default  # Brug standardfilen

if st.button("Generér PowerPoint") and all([user_file, variant_file, lifestyle_file, line_drawing_file, instruktioner_file, template_file]):
    user_data = pd.read_excel(user_file)
    variant_data = pd.read_excel(variant_file)
    lifestyle_data = pd.read_excel(lifestyle_file)
    line_drawing_data = pd.read_excel(line_drawing_file)
    instruktioner = pd.read_excel(instruktioner_file)
    
    ppt_bytes = generate_ppt(user_data, variant_data, lifestyle_data, line_drawing_data, instruktioner, template_file)
    
    st.success("PowerPoint genereret!")
    st.download_button(
        label="Download PowerPoint",
        data=ppt_bytes,
        file_name="Generated_Presentation.pptx",
        mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
    )
