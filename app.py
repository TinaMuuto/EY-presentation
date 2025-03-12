import streamlit as st
import pandas as pd
from pptx import Presentation
import io
import copy
import requests
from PIL import Image
from pptx.opc.constants import RELATIONSHIP_TYPE as RT

# Undgå decompression bomb-advarsler (vær opmærksom på risikoen, hvis du behandler meget store billeder)
Image.MAX_IMAGE_PIXELS = None

##############################################################################
# 1) Hjælpefunktioner til dataopslag
##############################################################################

def find_column(df, keywords):
    for col in df.columns:
        col_lower = col.lower().replace("_", " ")
        if all(kw in col_lower for kw in keywords):
            return col
    return None

def normalize_variantkey(s):
    return str(s).lower().replace(" - config", "").strip()

def match_variant_rows(item_no, variant_df):
    item_no_norm = str(item_no).strip().lower()
    if 'VariantKey_norm' not in variant_df.columns:
        variant_df['VariantKey_norm'] = variant_df['VariantKey'].apply(normalize_variantkey)
    matches = variant_df[variant_df['VariantKey_norm'] == item_no_norm]
    if not matches.empty:
        return matches
    if '-' in item_no_norm:
        part = item_no_norm.split('-')[0]
        matches = variant_df[variant_df['VariantKey_norm'].str.startswith(part)]
        if not matches.empty:
            return matches
    return pd.DataFrame()

def lookup_single_value(item_no, variant_df, column_name):
    rows = match_variant_rows(item_no, variant_df)
    if rows.empty:
        return ""
    val = rows.iloc[0].get(column_name, "")
    return "" if pd.isna(val) else str(val).strip()

def lookup_certificate(item_no, variant_df):
    df_cert = variant_df[variant_df['sys_entitytype'].astype(str).str.lower() == "certificate"]
    rows = match_variant_rows(item_no, df_cert)
    if rows.empty:
        return ""
    certs = rows['CertificateName'].dropna().astype(str).tolist()
    return "\n".join(certs)

def lookup_rts(item_no, variant_df):
    rows = match_variant_rows(item_no, variant_df)
    if rows.empty:
        return ""
    rts_rows = rows[rows['VariantIsInStock'].astype(str).str.lower() == "true"]
    if rts_rows.empty:
        return ""
    names = [str(val).replace("- All Colors", "").strip() for val in rts_rows['VariantCommercialName'].dropna()]
    return "\n".join(names)

def lookup_mto(item_no, variant_df):
    rows = match_variant_rows(item_no, variant_df)
    if rows.empty:
        return ""
    mto_rows = rows[rows['VariantIsInStock'].astype(str).str.lower() != "true"]
    if mto_rows.empty:
        return ""
    names = []
    for _, row in mto_rows.iterrows():
        val = row.get('VariantCommercialName', "")
        if pd.isna(val) or val.strip() == "":
            val = row.get('VariantName', "")
        name = str(val).replace("- All Colors", "").strip()
        if name:
            names.append(name)
    return "\n".join(names)

def lookup_packshot(item_no, variant_df):
    rows = match_variant_rows(item_no, variant_df)
    if rows.empty:
        return ""
    for _, row in rows.iterrows():
        if str(row.get('ResourceDigitalAssetType', "")).lower().strip() == "packshot image":
            url = row.get('ResourceDestinationUrl', "")
            if pd.notna(url) and url.strip():
                return url.strip()
    return ""

def lookup_lifestyle_images(item_no, variant_df, lifestyle_df):
    rows = match_variant_rows(item_no, variant_df)
    if rows.empty:
        return []
    product_key = rows.iloc[0].get("ProductKey", "")
    if pd.isna(product_key) or not product_key:
        return []
    subset = lifestyle_df[lifestyle_df['ProductKey'].astype(str).str.lower() == str(product_key).lower()]
    urls = subset['ResourceDestinationUrl'].dropna().astype(str).tolist()
    return urls[:3]

def lookup_line_drawings(item_no, variant_df, line_df):
    rows = match_variant_rows(item_no, variant_df)
    if rows.empty:
        return []
    product_key = rows.iloc[0].get("ProductKey", "")
    if pd.isna(product_key) or not product_key:
        return []
    subset = line_df[line_df['ProductKey'].astype(str).str.lower() == str(product_key).lower()]
    urls = subset['ResourceDestinationUrl'].dropna().astype(str).tolist()
    return urls[:8]

##############################################################################
# 2) Duplikeringsfunktion (brug samme layout som originalen)
##############################################################################

def duplicate_slide_in_same_presentation(prs, slide_index=0):
    source_slide = prs.slides[slide_index]
    slide_layout = source_slide.slide_layout  # Brug source slide's eget layout
    new_slide = prs.slides.add_slide(slide_layout)
    for shape in source_slide.shapes:
        el = shape.element
        new_el = copy.deepcopy(el)
        new_slide.shapes._spTree.insert_element_before(new_el, 'p:extLst')
    return new_slide

##############################################################################
# 3) Tekstudskiftning (bevarer run-formattering)
##############################################################################

def replace_text_placeholders(slide, replacements):
    for shape in slide.shapes:
        if shape.has_text_frame:
            for paragraph in shape.text_frame.paragraphs:
                for run in paragraph.runs:
                    for key, val in replacements.items():
                        placeholder = f"{{{{{key}}}}}"
                        if placeholder in run.text:
                            run.text = run.text.replace(placeholder, val)

##############################################################################
# 4) Billedindsættelse med komprimering
##############################################################################

def insert_image_in_placeholder(slide, placeholder, image_url):
    if not image_url:
        return
    try:
        resample_filter = Image.Resampling.LANCZOS
    except AttributeError:
        resample_filter = Image.ANTIALIAS
    for shape in slide.shapes:
        if shape.has_text_frame and shape.text.strip() == f"{{{{{placeholder}}}}}":
            left, top = shape.left, shape.top
            max_w, max_h = shape.width, shape.height
            try:
                resp = requests.get(image_url, timeout=10)
                if resp.status_code == 200:
                    img_data = io.BytesIO(resp.content)
                    with Image.open(img_data) as im:
                        w, h = im.size
                        scale = min(max_w / w, max_h / h)
                        new_w = int(w * scale)
                        new_h = int(h * scale)
                        # Resize billedet med det bestemte resample-filter
                        resized_im = im.resize((new_w, new_h), resample=resample_filter)
                        if resized_im.mode not in ("RGB", "L"):
                            resized_im = resized_im.convert("RGB")
                        # Gem som JPEG med reduceret kvalitet for at komprimere filstørrelsen
                        output_io = io.BytesIO()
                        resized_im.save(output_io, format="JPEG", quality=70)
                        output_io.seek(0)
                    slide.shapes.add_picture(output_io, left, top, width=new_w, height=new_h)
                    shape.text = ""
            except Exception as e:
                st.warning(f"Kunne ikke hente billede fra {image_url}: {e}")

##############################################################################
# 5) Hyperlinkindsættelse (bevarer templatedesign)
##############################################################################

def replace_hyperlink_placeholder(slide, placeholder, display_text, url):
    if not url:
        return
    for shape in slide.shapes:
        if shape.has_text_frame and f"{{{{{placeholder}}}}}" in shape.text:
            for paragraph in shape.text_frame.paragraphs:
                for run in paragraph.runs:
                    placeholder_tag = f"{{{{{placeholder}}}}}"
                    if placeholder_tag in run.text:
                        run.text = run.text.replace(placeholder_tag, "")
                        run.hyperlink.address = url
                        if not run.text:
                            run.text = display_text

##############################################################################
# 6) Udfyld slide for ét produkt
##############################################################################

def fill_slide(slide, product_row, variant_df, lifestyle_df, line_df):
    item_no = str(product_row.get("Item no", "")).strip()
    product_name = str(product_row.get("Product name", "")).strip()
    
    replacements = {
        "Product name": f"Product Name: {product_name}",
        "Product code": f"Product Code: {item_no}",
        "Product country of origin": f"Product Country of Origin: {lookup_single_value(item_no, variant_df, 'VariantCountryOfOrigin')}",
        "Product height": f"Height: {lookup_single_value(item_no, variant_df, 'VariantHeight')}",
        "Product width": f"Width: {lookup_single_value(item_no, variant_df, 'VariantWidth')}",
        "Product length": f"Length: {lookup_single_value(item_no, variant_df, 'VariantLength')}",
        "Product depth": f"Depth: {lookup_single_value(item_no, variant_df, 'VariantDepth')}",
        "Product seat height": f"Seat Height: {lookup_single_value(item_no, variant_df, 'VariantSeatHeight')}",
        "Product diameter": f"Diameter: {lookup_single_value(item_no, variant_df, 'VariantDiameter')}",
        "CertificateName": f"Test & certificates for the product: {lookup_certificate(item_no, variant_df)}",
        "Product Consumption COM": f"Consumption information for COM: {lookup_single_value(item_no, variant_df, 'ProductTextileConsumption_en')}",
        "Product RTS": f"Product in stock versions: {lookup_rts(item_no, variant_df)}",
        "Product MTO": f"Product in made to order versions: {lookup_mto(item_no, variant_df)}",
    }
    
    replace_text_placeholders(slide, replacements)
    
    fact_sheet_url = lookup_single_value(item_no, variant_df, "ProductFactSheetLink")
    replace_hyperlink_placeholder(slide, "Product Fact Sheet link", "Link to Product Fact Sheet", fact_sheet_url)
    
    config_url = lookup_single_value(item_no, variant_df, "ProductLinkToConfigurator")
    replace_hyperlink_placeholder(slide, "Product configurator link", "Configure product here", config_url)
    
    website_url = lookup_single_value(item_no, variant_df, "ProductWebsiteLink")
    replace_hyperlink_placeholder(slide, "Product website link", "See product website", website_url)
    
    packshot_url = lookup_packshot(item_no, variant_df)
    insert_image_in_placeholder(slide, "Product Packshot1", packshot_url)
    
    lifestyle_urls = lookup_lifestyle_images(item_no, variant_df, lifestyle_df)
    for i, url in enumerate(lifestyle_urls):
        placeholder = f"Product Lifestyle{i+1}"
        insert_image_in_placeholder(slide, placeholder, url)
    
    line_urls = lookup_line_drawings(item_no, variant_df, line_df)
    for i, url in enumerate(line_urls):
        placeholder = f"Product line drawing{i+1}"
        insert_image_in_placeholder(slide, placeholder, url)

##############################################################################
# 7) Streamlit-app
##############################################################################

st.title("Automatisk Generering af Præsentationer")

uploaded_file = st.file_uploader("Upload din Excel-fil med 'Item no' og 'Product name'", type=["xlsx"])

if uploaded_file:
    try:
        user_df = pd.read_excel(uploaded_file)
    except Exception as e:
        st.error(f"Fejl ved indlæsning af brugerfil: {e}")
        st.stop()
    
    item_no_col = find_column(user_df, ["item", "no"])
    product_name_col = find_column(user_df, ["product", "name"])
    if not item_no_col or not product_name_col:
        st.error("Kunne ikke finde kolonner for 'Item no' og 'Product name'.")
        st.stop()
    
    user_df = user_df.rename(columns={item_no_col: "Item no", product_name_col: "Product name"})
    # Konverter 'Item no' til string for at undgå typekonverteringsproblemer
    user_df["Item no"] = user_df["Item no"].apply(str)
    
    st.write("Brugerdata (første 10 rækker vist):")
    # Brug st.write for at undgå Arrow-konverteringsfejl
    st.write(user_df.head(10).astype(str))
    
    try:
        variant_df = pd.read_excel("EY - variant master data.xlsx")
        lifestyle_df = pd.read_excel("EY - lifestyle.xlsx")
        line_df = pd.read_excel("EY - line drawing.xlsx")
    except Exception as e:
        st.error(f"Fejl ved indlæsning af eksterne datafiler: {e}")
        st.stop()
    
    try:
        prs = Presentation("template-generator.pptx")
        if not prs.slides:
            st.error("Din template-præsentation har ingen slides.")
            st.stop()
    except Exception as e:
        st.error(f"Fejl ved indlæsning af PowerPoint template: {e}")
        st.stop()
    
    for idx, row in user_df.iterrows():
        if idx == 0:
            slide = prs.slides[0]
            fill_slide(slide, row, variant_df, lifestyle_df, line_df)
        else:
            new_slide = duplicate_slide_in_same_presentation(prs, slide_index=0)
            fill_slide(new_slide, row, variant_df, lifestyle_df, line_df)
    
    ppt_io = io.BytesIO()
    prs.save(ppt_io)
    ppt_io.seek(0)
    st.download_button("Download genereret præsentation", data=ppt_io, file_name="generated_presentation.pptx")
