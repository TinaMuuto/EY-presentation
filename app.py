import streamlit as st
import pandas as pd
from pptx import Presentation
from pptx.util import Inches
import io
import re
import copy
import requests
from PIL import Image

# Hjælpefunktion: Find kolonne baseret på nøgleord (ignorér store/små bogstaver)
def find_column(df, keywords):
    for col in df.columns:
        if all(kw in col.lower() for kw in keywords):
            return col
    return None

# Normaliser VariantKey ved at fjerne " - config"
def normalize_variant_key(key):
    return str(key).replace(" - config", "").strip().lower()

# Returner alle rækker i variant data, der matcher et Item no.
def get_variant_matches(item_no, variant_df):
    normalized_item = str(item_no).strip().lower()
    variant_df['VariantKey_norm'] = variant_df['VariantKey'].astype(str).apply(normalize_variant_key)
    matches = variant_df[variant_df['VariantKey_norm'] == normalized_item]
    if matches.empty:
        # Prøv med delvist match: alt før "-"
        part = normalized_item.split("-")[0]
        matches = variant_df[variant_df['VariantKey_norm'].str.startswith(part)]
    return matches

# Enkel opslag-funktion for enkelte felter
def lookup_field(item_no, variant_df, field_name):
    matches = get_variant_matches(item_no, variant_df)
    if matches.empty:
        return ""
    value = matches.iloc[0].get(field_name, "")
    if pd.isna(value):
        return ""
    return str(value).strip()

def lookup_product_RTS(item_no, variant_df):
    matches = get_variant_matches(item_no, variant_df)
    if matches.empty:
        return ""
    rts_rows = matches[matches['VariantIsInStock'].astype(str).str.lower() == "true"]
    names = rts_rows['VariantCommercialName'].dropna().astype(str).replace(to_replace="- All Colors", value="", regex=False)
    if names.empty:
        return ""
    return "\n".join(names.tolist())

def lookup_product_MTO(item_no, variant_df):
    matches = get_variant_matches(item_no, variant_df)
    if matches.empty:
        return ""
    mto_rows = matches[matches['VariantIsInStock'].astype(str).str.lower() != "true"]
    results = []
    for _, row in mto_rows.iterrows():
        name = row.get('VariantCommercialName')
        if pd.isna(name) or name == "":
            name = row.get('VariantName')
        if pd.notna(name) and name != "":
            name = str(name).replace("- All Colors", "").strip()
            results.append(name)
    return "\n".join(results)

def lookup_certificate(item_no, variant_df):
    variant_df_cert = variant_df[variant_df['sys_entitytype'].astype(str).str.lower() == "certificate"]
    matches = get_variant_matches(item_no, variant_df_cert)
    if matches.empty:
        return ""
    certs = matches['CertificateName'].dropna().astype(str)
    if certs.empty:
        return ""
    return "\n".join(certs.tolist())

def lookup_fact_sheet_link(item_no, variant_df):
    return lookup_field(item_no, variant_df, "ProductFactSheetLink")

def lookup_configurator_link(item_no, variant_df):
    return lookup_field(item_no, variant_df, "ProductLinkToConfigurator")

def lookup_website_link(item_no, variant_df):
    return lookup_field(item_no, variant_df, "ProductWebsiteLink")

def lookup_packshot(item_no, variant_df):
    matches = get_variant_matches(item_no, variant_df)
    if matches.empty:
        return ""
    for _, row in matches.iterrows():
        if str(row.get("ResourceDigitalAssetType", "")).strip().lower() == "packshot image":
            url = row.get("ResourceDestinationUrl", "")
            if pd.notna(url) and url != "":
                return str(url).strip()
    return ""

def lookup_lifestyle_images(item_no, variant_df, lifestyle_df):
    matches = get_variant_matches(item_no, variant_df)
    if matches.empty:
        return []
    product_key = matches.iloc[0].get("ProductKey", "")
    if pd.isna(product_key) or product_key == "":
        return []
    rows = lifestyle_df[lifestyle_df['ProductKey'].astype(str).str.lower() == str(product_key).lower()]
    urls = rows['ResourceDestinationUrl'].dropna().astype(str).tolist()
    return urls  # Returnér alle URL'er

def lookup_line_drawing_images(item_no, variant_df, line_drawing_df):
    matches = get_variant_matches(item_no, variant_df)
    if matches.empty:
        return []
    product_key = matches.iloc[0].get("ProductKey", "")
    if pd.isna(product_key) or product_key == "":
        return []
    rows = line_drawing_df[line_drawing_df['ProductKey'].astype(str).str.lower() == str(product_key).lower()]
    urls = rows['ResourceDestinationUrl'].dropna().astype(str).tolist()
    return urls  # Returnér alle URL'er

# Funktion til at duplikere en slide (kopierer indholdet fra en template slide)
def duplicate_slide(pres, slide):
    slide_layout = slide.slide_layout
    new_slide = pres.slides.add_slide(slide_layout)
    for shape in slide.shapes:
        el = shape.element
        new_el = copy.deepcopy(el)
        new_slide.shapes._spTree.insert_element_before(new_el, 'p:extLst')
    return new_slide

# Opdateret funktion der erstatter en billede-placeholder med et billede hentet fra en URL,
# og som bevarer billedets oprindelige forhold
def replace_image_placeholder(slide, placeholder, image_url):
    for shape in slide.shapes:
        if shape.has_text_frame and shape.text.strip() == "{{" + placeholder + "}}":
            left = shape.left
            top = shape.top
            max_width = shape.width
            max_height = shape.height
            try:
                response = requests.get(image_url, timeout=10)
                if response.status_code == 200:
                    image_data = io.BytesIO(response.content)
                    with Image.open(image_data) as img:
                        original_width, original_height = img.size
                        scale = min(max_width / original_width, max_height / original_height)
                        new_width = int(original_width * scale)
                        new_height = int(original_height * scale)
                    # Nulstil pointeren for billed-data
                    image_data.seek(0)
                    slide.shapes.add_picture(image_data, left, top, width=new_width, height=new_height)
                    shape.text = ""
            except Exception as e:
                st.error(f"Fejl ved hentning af billede for {placeholder}: {e}")

# Funktion der erstatter tekst placeholders i en slide
def replace_placeholders(slide, replacements):
    for shape in slide.shapes:
        if shape.has_text_frame:
            for paragraph in shape.text_frame.paragraphs:
                for run in paragraph.runs:
                    for key, value in replacements.items():
                        placeholder = "{{" + key + "}}"
                        if placeholder in run.text:
                            run.text = run.text.replace(placeholder, value)

# Processér ét produkt: udtræk data, erstat felter og indsæt billeder
def process_product(slide, product, variant_df, lifestyle_df, line_drawing_df):
    item_no = str(product.get("Item no", "")).strip()
    product_name = str(product.get("Product name", "")).strip()
    
    # Ordbog til tekstfelter
    replacements = {
        "Product name": "Product Name: " + product_name,
        "Product code": "Product Code: " + item_no,
        "Product country of origin": "Product Country of Origin: " + lookup_field(item_no, variant_df, "VariantCountryOfOrigin"),
        "Product height": "Height: " + lookup_field(item_no, variant_df, "VariantHeight"),
        "Product width": "Width: " + lookup_field(item_no, variant_df, "VariantWidth"),
        "Product length": "Length: " + lookup_field(item_no, variant_df, "VariantLength"),
        "Product depth": "Depth: " + lookup_field(item_no, variant_df, "VariantDepth"),
        "Product seat height": "Seat Height: " + lookup_field(item_no, variant_df, "VariantSeatHeight"),
        "Product diameter": "Diameter: " + lookup_field(item_no, variant_df, "VariantDiameter"),
        "CertificateName": "Test & certificates for the product: " + lookup_certificate(item_no, variant_df),
        "Product Consumption COM": "Consumption information for COM: " + lookup_field(item_no, variant_df, "ProductTextileConsumption_en"),
        "Product RTS": "Product in stock versions: " + lookup_product_RTS(item_no, variant_df),
        "Product MTO": "Product in made to order versions: " + lookup_product_MTO(item_no, variant_df),
        "Product Fact Sheet link": "[Link to Product Fact Sheet](" + lookup_fact_sheet_link(item_no, variant_df) + ")",
        "Product configurator link": "[Configure product here](" + lookup_configurator_link(item_no, variant_df) + ")",
        "Product website link": lookup_website_link(item_no, variant_df)
    }
    
    replace_placeholders(slide, replacements)
    
    # Indsæt packshot billede
    packshot_url = lookup_packshot(item_no, variant_df)
    if packshot_url:
        replace_image_placeholder(slide, "Product Packshot1", packshot_url)
    
    # Indsæt lifestyle-billeder (alle fundne URL'er)
    lifestyle_urls = lookup_lifestyle_images(item_no, variant_df, lifestyle_df)
    for i, url in enumerate(lifestyle_urls):
        placeholder = f"Product Lifestyle{i+1}"
        replace_image_placeholder(slide, placeholder, url)
    
    # Indsæt line drawing-billeder (alle fundne URL'er)
    line_drawing_urls = lookup_line_drawing_images(item_no, variant_df, line_drawing_df)
    for i, url in enumerate(line_drawing_urls):
        placeholder = f"Product line drawing{i+1}"
        replace_image_placeholder(slide, placeholder, url)

##############################################
# Hoveddel af Streamlit-appen
##############################################
st.title("Automatisk Generering af Præsentationer")

uploaded_file = st.file_uploader("Upload din Excel-fil med 'Item no' og 'Product name'", type=["xlsx"])

if uploaded_file is not None:
    try:
        user_df = pd.read_excel(uploaded_file)
    except Exception as e:
        st.error(f"Fejl ved indlæsning af brugerfil: {e}")
    
    # Find de relevante kolonner (prøver at finde varianter af kolonnenavne)
    item_no_col = find_column(user_df, ["item", "no"])
    product_name_col = find_column(user_df, ["product", "name"])
    if item_no_col is None or product_name_col is None:
        st.error("Kunne ikke finde de nødvendige kolonner ('Item no' og 'Product name') i din fil.")
    else:
        # Omdøb kolonnerne for at standardisere navne
        user_df = user_df.rename(columns={item_no_col: "Item no", product_name_col: "Product name"})
        st.write("Brugerdata:")
        st.dataframe(user_df)
        
        # Indlæs de eksterne datafiler – forudsætter at de ligger lokalt
        try:
            variant_df = pd.read_excel("EY - variant master data.xlsx")
            lifestyle_df = pd.read_excel("EY - lifestyle.xlsx")
            line_drawing_df = pd.read_excel("EY - line drawing.xlsx")
        except Exception as e:
            st.error(f"Fejl ved indlæsning af eksterne datafiler: {e}")
        
        # Indlæs PowerPoint-templaten
        try:
            prs = Presentation("template-generator.pptx")
        except Exception as e:
            st.error(f"Fejl ved indlæsning af PowerPoint template: {e}")
        
        # Antag at templaten har én slide, som skal duplikeres for hvert produkt.
        template_slide = prs.slides[0]
        
        # Processér første produkt på templaten
        if len(user_df) > 0:
            product = user_df.iloc[0]
            process_product(template_slide, product, variant_df, lifestyle_df, line_drawing_df)
        
        # For resterende produkter: duplikér templaten og processér
        for idx in range(1, len(user_df)):
            product = user_df.iloc[idx]
            new_slide = duplicate_slide(prs, template_slide)
            process_product(new_slide, product, variant_df, lifestyle_df, line_drawing_df)
        
        # Giv brugeren mulighed for at downloade den genererede præsentation
        ppt_io = io.BytesIO()
        prs.save(ppt_io)
        ppt_io.seek(0)
        st.download_button("Download genereret præsentation", data=ppt_io, file_name="generated_presentation.pptx")
