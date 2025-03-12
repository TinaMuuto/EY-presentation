import streamlit as st
import pandas as pd
from pptx import Presentation
import io
import copy
import requests
from PIL import Image
from pptx.opc.constants import RELATIONSHIP_TYPE as RT

# Undgå advarsler ved store billeder (vær opmærksom på risikoen)
Image.MAX_IMAGE_PIXELS = None

##############################################################################
# Hjælpefunktioner til dataopslag
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
# Funktion til duplikering af slide (brug samme layout som originalen)
##############################################################################

def duplicate_slide_in_same_presentation(prs, slide_index=0):
    source_slide = prs.slides[slide_index]
    slide_layout = source_slide.slide_layout
    new_slide = prs.slides.add_slide(slide_layout)
    for shape in source_slide.shapes:
        el = shape.element
        new_el = copy.deepcopy(el)
        new_slide.shapes._spTree.insert_element_before(new_el, 'p:extLst')
    return new_slide

##############################################################################
# Tekstudskiftning – forenklet (erstatter hele teksten)
##############################################################################

def fill_text_fields(slide, product_row, variant_df):
    item_no = str(product_row.get("Item no", "")).strip()
    product_name = str(product_row.get("Product name", "")).strip()
    
    # Opsætning af tekststrenge – her indsættes alt som ren tekst
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
    
    # For hver shape med tekst, udfør en simpel global erstatning
    for shape in slide.shapes:
        if shape.has_text_frame:
            text = shape.text
            for key, rep in replacements.items():
                placeholder = f"{{{{{key}}}}}"
                text = text.replace(placeholder, rep)
            shape.text = text

    # Hyperlinks indsættes også som ren tekst med link
    insert_hyperlink(slide, "Product Fact Sheet link", "Link to Product Fact Sheet", lookup_single_value(item_no, variant_df, "ProductFactSheetLink"))
    insert_hyperlink(slide, "Product configurator link", "Configure product here", lookup_single_value(item_no, variant_df, "ProductLinkToConfigurator"))
    insert_hyperlink(slide, "Product website link", "See product website", lookup_single_value(item_no, variant_df, "ProductWebsiteLink"))

def insert_hyperlink(slide, placeholder, display_text, url):
    if not url:
        return
    for shape in slide.shapes:
        if shape.has_text_frame and f"{{{{{placeholder}}}}}" in shape.text:
            # Erstat placeholder med display_text
            shape.text = shape.text.replace(f"{{{{{placeholder}}}}}", display_text)
            # Sæt hyperlink for alle runs, der indeholder display_text
            for paragraph in shape.text_frame.paragraphs:
                for run in paragraph.runs:
                    if display_text in run.text:
                        run.hyperlink.address = url

##############################################################################
# Billedindsættelse – skalerer billedet til placeholderens størrelse uden komprimering
##############################################################################

def fill_image_fields(slide, product_row, variant_df, lifestyle_df, line_df):
    item_no = str(product_row.get("Item no", "")).strip()
    
    # Packshot
    packshot_url = lookup_packshot(item_no, variant_df)
    insert_image(slide, "Product Packshot1", packshot_url)
    
    # Lifestyle – indsæt op til 3 billeder
    lifestyle_urls = lookup_lifestyle_images(item_no, variant_df, lifestyle_df)
    for i, url in enumerate(lifestyle_urls):
        placeholder = f"Product Lifestyle{i+1}"
        insert_image(slide, placeholder, url)
    
    # Line drawings – indsæt op til 8 billeder
    line_urls = lookup_line_drawings(item_no, variant_df, line_df)
    for i, url in enumerate(line_urls):
        placeholder = f"Product line drawing{i+1}"
        insert_image(slide, placeholder, url)

def insert_image(slide, placeholder, image_url):
    if not image_url:
        return
    try:
        # Vælg resample-filter (brug LANCZOS hvis muligt)
        try:
            resample_filter = Image.Resampling.LANCZOS
        except AttributeError:
            resample_filter = Image.ANTIALIAS

        for shape in slide.shapes:
            if shape.has_text_frame and shape.text.strip() == f"{{{{{placeholder}}}}}":
                left, top = shape.left, shape.top
                max_w, max_h = shape.width, shape.height
                resp = requests.get(image_url, timeout=10)
                if resp.status_code == 200:
                    img_data = io.BytesIO(resp.content)
                    with Image.open(img_data) as im:
                        w, h = im.size
                        scale = min(max_w / w, max_h / h)
                        new_w = int(w * scale)
                        new_h = int(h * scale)
                        resized_im = im.resize((new_w, new_h), resample=resample_filter)
                        # Her undlades komprimering – gemmes fx som PNG
                        output_io = io.BytesIO()
                        resized_im.save(output_io, format="PNG")
                        output_io.seek(0)
                    slide.shapes.add_picture(output_io, left, top, width=new_w, height=new_h)
                    shape.text = ""
    except Exception as e:
        st.warning(f"Kunne ikke hente billede fra {image_url}: {e}")

##############################################################################
# Hovedprogram – udfyld slides i to passeringer: først tekst/hyperlinks, derefter billeder
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
    user_df["Item no"] = user_df["Item no"].apply(str)  # Sikrer, at Item no er string
    
    st.write("Brugerdata (første 10 rækker vist):")
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
    
    # For hvert produkt, dupliker slide (første slide bruges til det første produkt)
    for idx, row in user_df.iterrows():
        if idx == 0:
            slide = prs.slides[0]
        else:
            slide = duplicate_slide_in_same_presentation(prs, slide_index=0)
        # Først indsættes tekst og hyperlinks
        fill_text_fields(slide, row, variant_df)
        # Derefter indsættes billeder
        fill_image_fields(slide, row, variant_df, lifestyle_df, line_df)
    
    ppt_io = io.BytesIO()
    prs.save(ppt_io)
    ppt_io.seek(0)
    st.download_button("Download genereret præsentation", data=ppt_io, file_name="generated_presentation.pptx")
