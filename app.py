import streamlit as st
import pandas as pd
from pptx import Presentation
from pptx.util import Inches
import io
import copy
import requests
from PIL import Image

##############################################################################
# 1) Hjælpefunktioner til at finde og matche data
##############################################################################

def find_column(df, keywords):
    """
    Finder kolonnen i df, hvis den indeholder alle ord i 'keywords' (ignorerer store/små bogstaver).
    Eksempel: find_column(user_df, ["item", "no"]) -> 'Item No' eller 'ITEM NO.' etc.
    """
    for col in df.columns:
        col_lower = col.lower().replace("_", " ")
        if all(kw in col_lower for kw in keywords):
            return col
    return None

def normalize_variantkey(s):
    """
    Fjerner ' - config' (hvis det findes) og lower-case.
    """
    return str(s).lower().replace(" - config", "").strip()

def match_variant_rows(item_no, variant_df):
    """
    Forsøger at matche item_no mod 'VariantKey' i variant_df.
    1) Eksakt match (efter normalisering).
    2) Hvis intet match, fjern alt efter '-' i item_no og forsøg igen.
    Returnerer de matchende rækker fra variant_df (kan være flere).
    """
    # Normalisér item_no
    item_no_norm = str(item_no).lower().strip()
    
    # Forbered variant_df ved at normalisere VariantKey (kun én gang, hvis ikke gjort før)
    if 'VariantKey_norm' not in variant_df.columns:
        variant_df['VariantKey_norm'] = variant_df['VariantKey'].apply(normalize_variantkey)
    
    # 1) Eksakt match
    matches = variant_df[variant_df['VariantKey_norm'] == item_no_norm]
    if not matches.empty:
        return matches
    
    # 2) Delvist match: alt før '-'
    if '-' in item_no_norm:
        item_no_part = item_no_norm.split('-')[0]
        matches = variant_df[variant_df['VariantKey_norm'].str.startswith(item_no_part)]
        if not matches.empty:
            return matches
    
    # Hvis intet fundet
    return pd.DataFrame()

def lookup_single_value(item_no, variant_df, column_name):
    """
    Returnerer første ikke-tomme værdi fra column_name for de rækker,
    der matcher item_no i variant_df.
    """
    rows = match_variant_rows(item_no, variant_df)
    if rows.empty:
        return ""
    val = rows.iloc[0].get(column_name, "")
    return "" if pd.isna(val) else str(val).strip()

def lookup_certificate(item_no, variant_df):
    """
    Henter certifikater (sys_entitytype = 'Certificate').
    Sammenkæder alle certificate-names med linjeskift.
    """
    # Filtrer variant_df til kun certificate-rækker
    df_cert = variant_df[variant_df['sys_entitytype'].astype(str).str.lower() == "certificate"]
    rows = match_variant_rows(item_no, df_cert)
    if rows.empty:
        return ""
    certs = rows['CertificateName'].dropna().astype(str).tolist()
    return "\n".join(certs)

def lookup_rts(item_no, variant_df):
    """
    Henter 'VariantCommercialName' for rækker med VariantIsInStock = 'True'.
    Fjerner '- All Colors' hvis det findes.
    Sammenkæder med linjeskift.
    """
    rows = match_variant_rows(item_no, variant_df)
    if rows.empty:
        return ""
    # Filtrer rækker, hvor 'VariantIsInStock' = True
    rts_rows = rows[rows['VariantIsInStock'].astype(str).str.lower() == "true"]
    if rts_rows.empty:
        return ""
    names = []
    for val in rts_rows['VariantCommercialName'].dropna():
        name = str(val).replace("- All Colors", "").strip()
        names.append(name)
    return "\n".join(names)

def lookup_mto(item_no, variant_df):
    """
    Henter 'VariantCommercialName' for rækker med VariantIsInStock != 'True'.
    Hvis 'VariantCommercialName' er tom, bruges 'VariantName'.
    Fjerner '- All Colors'.
    Sammenkæder med linjeskift.
    """
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
    """
    Returnerer første 'ResourceDestinationUrl' hvor 'ResourceDigitalAssetType' = 'Packshot image'.
    """
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
    """
    Finder 'ProductKey' fra variant_df, slår op i lifestyle_df og returnerer
    op til 3 'ResourceDestinationUrl'.
    """
    rows = match_variant_rows(item_no, variant_df)
    if rows.empty:
        return []
    product_key = rows.iloc[0].get("ProductKey", "")
    if pd.isna(product_key) or not product_key:
        return []
    # Filtrer lifestyle_df på product_key
    subset = lifestyle_df[lifestyle_df['ProductKey'].astype(str).str.lower() == str(product_key).lower()]
    urls = subset['ResourceDestinationUrl'].dropna().astype(str).tolist()
    return urls[:3]  # Op til 3

def lookup_line_drawings(item_no, variant_df, line_df):
    """
    Finder 'ProductKey' fra variant_df, slår op i line_df og returnerer
    op til 8 'ResourceDestinationUrl'.
    """
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
# 2) Funktion til at kopiere en slide fra en template-præsentation
##############################################################################

def copy_slide_from_template(template_slide, target_presentation):
    """
    Opretter en ny slide i target_presentation og kopierer alle shapes
    fra template_slide over i den nye slide.
    """
    # Vælg en blank layout i target_presentation (layout index 6 er ofte helt blank)
    blank_layout = target_presentation.slide_layouts[6]
    new_slide = target_presentation.slides.add_slide(blank_layout)
    
    # Kopiér shapes
    for shape in template_slide.shapes:
        el = shape.element
        new_el = copy.deepcopy(el)
        new_slide.shapes._spTree.insert_element_before(new_el, 'p:extLst')
    return new_slide

##############################################################################
# 3) Funktioner til at erstatte tekst og indsætte billeder i en slide
##############################################################################

def replace_text_placeholders(slide, replacements):
    """
    Søger efter fx {{Product name}} i slide's tekst,
    og erstatter med replacements["Product name"] osv.
    """
    for shape in slide.shapes:
        if shape.has_text_frame:
            text = shape.text
            for key, val in replacements.items():
                placeholder = f"{{{{{key}}}}}"  # fx "{{Product name}}"
                if placeholder in text:
                    text = text.replace(placeholder, val)
            shape.text = text

def insert_image_in_placeholder(slide, placeholder, image_url):
    """
    Finder shape, hvor shape.text = '{{placeholder}}',
    og indsætter billedet (med bevaret ratio) på samme position/størrelse.
    """
    if not image_url:
        return  # Ingen URL => ingen indsættelse
    for shape in slide.shapes:
        if shape.has_text_frame and shape.text.strip() == f"{{{{{placeholder}}}}}":
            left = shape.left
            top = shape.top
            max_width = shape.width
            max_height = shape.height
            
            # Hent billede
            try:
                resp = requests.get(image_url, timeout=10)
                if resp.status_code == 200:
                    image_data = io.BytesIO(resp.content)
                    with Image.open(image_data) as img:
                        orig_w, orig_h = img.size
                        scale = min(max_width / orig_w, max_height / orig_h)
                        new_w = int(orig_w * scale)
                        new_h = int(orig_h * scale)
                    # Nulstil buffer
                    image_data.seek(0)
                    slide.shapes.add_picture(image_data, left, top, width=new_w, height=new_h)
                    shape.text = ""  # Fjern placeholder-teksten
            except Exception as e:
                st.warning(f"Kunne ikke hente billede fra {image_url}: {e}")

##############################################################################
# 4) Hovedfunktion til at udfylde én slide for ét produkt
##############################################################################

def fill_slide(slide, product_row, variant_df, lifestyle_df, line_df):
    """
    Tager en 'ren' slide, og udfylder dens placeholders baseret på:
      - Data fra brugerens fil (Item no, Product name)
      - Data fra variant_df (master)
      - Billeder fra lifestyle_df og line_df
    """
    item_no = str(product_row.get("Item no", "")).strip()
    product_name = str(product_row.get("Product name", "")).strip()
    
    # Tekstreplacements
    replacements = {
        # 1) {{Product name}} => "Product Name: <fra brugerfil>"
        "Product name": f"Product Name: {product_name}",
        # 2) {{Product code}} => "Product Code: <Item no>"
        "Product code": f"Product Code: {item_no}",
        # 3) {{Product country of origin}} => "Product Country of Origin: <...>"
        "Product country of origin": f"Product Country of Origin: {lookup_single_value(item_no, variant_df, 'VariantCountryOfOrigin')}",
        # 4) {{Product height}}
        "Product height": f"Height: {lookup_single_value(item_no, variant_df, 'VariantHeight')}",
        # 5) {{Product width}}
        "Product width": f"Width: {lookup_single_value(item_no, variant_df, 'VariantWidth')}",
        # 6) {{Product length}}
        "Product length": f"Length: {lookup_single_value(item_no, variant_df, 'VariantLength')}",
        # 7) {{Product depth}}
        "Product depth": f"Depth: {lookup_single_value(item_no, variant_df, 'VariantDepth')}",
        # 8) {{Product seat height}}
        "Product seat height": f"Seat Height: {lookup_single_value(item_no, variant_df, 'VariantSeatHeight')}",
        # 9) {{Product diameter}}
        "Product diameter": f"Diameter: {lookup_single_value(item_no, variant_df, 'VariantDiameter')}",
        # 10) {{CertificateName}}
        "CertificateName": f"Test & certificates for the product: {lookup_certificate(item_no, variant_df)}",
        # 11) {{Product Consumption COM}}
        "Product Consumption COM": f"Consumption information for COM: {lookup_single_value(item_no, variant_df, 'ProductTextileConsumption_en')}",
        # 12) {{Product RTS}}
        "Product RTS": f"Product in stock versions: {lookup_rts(item_no, variant_df)}",
        # 13) {{Product MTO}}
        "Product MTO": f"Product in made to order versions: {lookup_mto(item_no, variant_df)}",
        # 14) {{Product Fact Sheet link}}
        #    => [Link to Product Fact Sheet](URL)
        "Product Fact Sheet link": f"[Link to Product Fact Sheet]({lookup_single_value(item_no, variant_df, 'ProductFactSheetLink')})",
        # 15) {{Product configurator link}}
        "Product configurator link": f"[Configure product here]({lookup_single_value(item_no, variant_df, 'ProductLinkToConfigurator')})",
        # 16) {{Product website link}}
        #    => [See product website](URL)
        #    Hvis du vil have den i hyperlink-format, brug markdown-lignende:
        "Product website link": f"[See product website]({lookup_single_value(item_no, variant_df, 'ProductWebsiteLink')})",
    }
    
    # Erstat tekst
    replace_text_placeholders(slide, replacements)
    
    # 17) Packshot
    packshot_url = lookup_packshot(item_no, variant_df)
    insert_image_in_placeholder(slide, "Product Packshot1", packshot_url)
    
    # 18-20) Lifestyle (op til 3)
    lifestyle_urls = lookup_lifestyle_images(item_no, variant_df, lifestyle_df)
    for i, url in enumerate(lifestyle_urls):
        placeholder = f"Product Lifestyle{i+1}"
        insert_image_in_placeholder(slide, placeholder, url)
    
    # 21-28) Line drawings (op til 8)
    line_urls = lookup_line_drawings(item_no, variant_df, line_df)
    for i, url in enumerate(line_urls):
        placeholder = f"Product line drawing{i+1}"
        insert_image_in_placeholder(slide, placeholder, url)

##############################################################################
# 5) Streamlit-app
##############################################################################

st.title("Automatisk Generering af Præsentationer")

uploaded_file = st.file_uploader("Upload din Excel-fil med 'Item no' og 'Product name'", type=["xlsx"])

if uploaded_file:
    try:
        user_df = pd.read_excel(uploaded_file)
    except Exception as e:
        st.error(f"Fejl ved indlæsning af brugerfil: {e}")
        st.stop()
    
    # Find kolonnenavne
    item_no_col = find_column(user_df, ["item", "no"])
    product_name_col = find_column(user_df, ["product", "name"])
    
    if not item_no_col or not product_name_col:
        st.error("Kunne ikke finde kolonner for 'Item no' og 'Product name' i din fil.")
        st.stop()
    
    # Omdøb kolonner
    user_df = user_df.rename(columns={item_no_col: "Item no", product_name_col: "Product name"})
    st.write("Brugerdata (første 10 rækker vist):")
    st.dataframe(user_df.head(10))
    
    # Indlæs eksterne datafiler
    try:
        variant_df = pd.read_excel("EY - variant master data.xlsx")
        lifestyle_df = pd.read_excel("EY - lifestyle.xlsx")
        line_df = pd.read_excel("EY - line drawing.xlsx")
    except Exception as e:
        st.error(f"Fejl ved indlæsning af eksterne datafiler: {e}")
        st.stop()
    
    # Indlæs template-præsentation
    try:
        template_pres = Presentation("template-generator.pptx")
        if not template_pres.slides:
            st.error("Din template-præsentation har ingen slides.")
            st.stop()
        template_slide = template_pres.slides[0]
    except Exception as e:
        st.error(f"Fejl ved indlæsning af PowerPoint template: {e}")
        st.stop()
    
    # Opret en ny præsentation, hvor vi kopierer en slide for hvert produkt
    final_pres = Presentation()
    # Slet evt. den auto-genererede tomme slide (typisk index 0):
    while len(final_pres.slides) > 0:
        r_id = final_pres.slides[0].slide_id
        final_pres.slides.remove(final_pres.slides.get(slide_id=r_id))
    
    # Gennemgå hver række i brugerens fil, og opret en slide
    for idx, row in user_df.iterrows():
        new_slide = copy_slide_from_template(template_slide, final_pres)
        fill_slide(new_slide, row, variant_df, lifestyle_df, line_df)
    
    # Til sidst: Download-knap
    ppt_io = io.BytesIO()
    final_pres.save(ppt_io)
    ppt_io.seek(0)
    st.download_button(
        label="Download genereret præsentation",
        data=ppt_io,
        file_name="generated_presentation.pptx"
    )
