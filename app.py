import streamlit as st
import pandas as pd
from pptx import Presentation
import io
import copy
import requests
import re
from PIL import Image
from pptx.opc.constants import RELATIONSHIP_TYPE as RT

# Undgå advarsler ved meget store billeder (vær opmærksom på risikoen)
Image.MAX_IMAGE_PIXELS = None

#############################################
# 1) Hjælpefunktioner – Dataopslag
#############################################

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

#############################################
# 2) Slide-håndtering – Fjern slides (workaround)
#############################################

def delete_slide(prs, slide):
    """
    Sletter en slide fra præsentationen ved at fjerne dens XML-element fra _sldIdLst.
    Bemærk: Dette er en workaround, da python-pptx ikke tilbyder en officiel API.
    """
    xml_slides = prs.slides._sldIdLst  
    slide_id = slide.slide_id
    for sld in xml_slides:
        if sld.get("id") == str(slide_id):
            xml_slides.remove(sld)
            break

#############################################
# 3) Funktion til at kopiere en slide fra template til final_pres
#############################################

def copy_slide_from_template(template_slide, target_pres):
    try:
        blank_layout = target_pres.slide_layouts[6]
    except IndexError:
        blank_layout = target_pres.slide_layouts[0]
    new_slide = target_pres.slides.add_slide(blank_layout)
    for shape in template_slide.shapes:
        el = shape.element
        new_el = copy.deepcopy(el)
        new_slide.shapes._spTree.insert_element_before(new_el, 'p:extLst')
    return new_slide

#############################################
# 4) Tekstudskiftning – indsætter tekst som ren tekst
#############################################

def replace_text_placeholders(slide, replacements):
    for shape in slide.shapes:
        if shape.has_text_frame:
            text = shape.text
            for key, val in replacements.items():
                placeholder = f"{{{{{key}}}}}"
                text = text.replace(placeholder, val)
            shape.text = text

#############################################
# 5) Billedindsættelse – nedskalerer store billeder
#############################################

def insert_image(slide, placeholder, image_url):
    if not image_url:
        return
    try:
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
                        # Nedskaler billedet, hvis det er for stort – maks 1920x1080
                        MAX_SIZE = (1920, 1080)
                        im.thumbnail(MAX_SIZE, resample=resample_filter)
                        orig_w, orig_h = im.size
                        scale = min(max_w / orig_w, max_h / orig_h)
                        new_w = int(orig_w * scale)
                        new_h = int(orig_h * scale)
                        resized_im = im.resize((new_w, new_h), resample=resample_filter)
                        output_io = io.BytesIO()
                        # Gem som PNG uden ekstra komprimering
                        resized_im.save(output_io, format="PNG")
                        output_io.seek(0)
                    slide.shapes.add_picture(output_io, left, top, width=new_w, height=new_h)
                    shape.text = ""
    except Exception as e:
        st.warning(f"Kunne ikke hente billede fra {image_url}: {e}")

#############################################
# 6) Hyperlinkindsættelse – forenklet metode
#############################################

def insert_hyperlink(slide, placeholder, display_text, url):
    if not url:
        return
    for shape in slide.shapes:
        if shape.has_text_frame and f"{{{{{placeholder}}}}}" in shape.text:
            shape.text = shape.text.replace(f"{{{{{placeholder}}}}}", display_text)
            for paragraph in shape.text_frame.paragraphs:
                for run in paragraph.runs:
                    if display_text in run.text:
                        run.hyperlink.address = url

#############################################
# 7) Udfyld slide – Fase 1 (Tekst og hyperlinks)
#############################################

def fill_text_fields(slide, product_row, variant_df):
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
        "Product Fact Sheet link": f"[Link to Product Fact Sheet]({lookup_single_value(item_no, variant_df, 'ProductFactSheetLink')})",
        "Product configurator link": f"[Configure product here]({lookup_single_value(item_no, variant_df, 'ProductLinkToConfigurator')})",
        "Product website link": f"[See product website]({lookup_single_value(item_no, variant_df, 'ProductWebsiteLink')})",
    }
    replace_text_placeholders(slide, replacements)
    insert_hyperlink(slide, "Product Fact Sheet link", "Link to Product Fact Sheet", lookup_single_value(item_no, variant_df, "ProductFactSheetLink"))
    insert_hyperlink(slide, "Product configurator link", "Configure product here", lookup_single_value(item_no, variant_df, "ProductLinkToConfigurator"))
    insert_hyperlink(slide, "Product website link", "See product website", lookup_single_value(item_no, variant_df, "ProductWebsiteLink"))

#############################################
# 8) Udfyld slide – Fase 2 (Billeder)
#############################################

def fill_image_fields(slide, product_row, variant_df, lifestyle_df, line_df):
    item_no = str(product_row.get("Item no", "")).strip()
    packshot_url = lookup_packshot(item_no, variant_df)
    insert_image(slide, "Product Packshot1", packshot_url)
    lifestyle_urls = lookup_lifestyle_images(item_no, variant_df, lifestyle_df)
    for i, url in enumerate(lifestyle_urls):
        placeholder = f"Product Lifestyle{i+1}"
        insert_image(slide, placeholder, url)
    line_urls = lookup_line_drawings(item_no, variant_df, line_df)
    for i, url in enumerate(line_urls):
        placeholder = f"Product line drawing{i+1}"
        insert_image(slide, placeholder, url)

#############################################
# 9) Fuld udfyldning af slide (fase 1 og 2)
#############################################

def fill_slide(slide, product_row, variant_df, lifestyle_df, line_df):
    fill_text_fields(slide, product_row, variant_df)
    fill_image_fields(slide, product_row, variant_df, lifestyle_df, line_df)

#############################################
# 10) Main – Two-phase Workflow
#############################################

st.title("Automatisk Generering af Præsentationer – To-trins Workflow")

st.markdown("""
### **Trin 1: Generer Tekstpræsentation**
1. Upload din Excel-fil med kolonnerne **Item no** og **Product name**.
2. Appen genererer en PowerPoint-præsentation baseret på din template (template-generator.pptx), hvor alle tekstfelter og hyperlinks udfyldes.
3. Download den tekstbaserede præsentation.

---

### **Trin 2: Tilføj Billeder**
1. Upload den tekstbaserede præsentation (PPTX).
2. Appen indsætter billeder (packshot, lifestyle og line drawings) baseret på dine eksterne datafiler.
3. Download den opdaterede præsentation med billeder.
""")

phase = st.radio("Vælg procesfase", ("Generer tekstpræsentation", "Tilføj billeder til præsentation"))

#############################################
# Fase 1: Generer tekstpræsentation
#############################################
if phase == "Generer tekstpræsentation":
    st.markdown("### Trin 1: Generer Tekstpræsentation")
    uploaded_excel = st.file_uploader("Upload din Excel-fil med 'Item no' og 'Product name'", type=["xlsx"], key="excel")
    if uploaded_excel:
        try:
            user_df = pd.read_excel(uploaded_excel)
        except Exception as e:
            st.error(f"Fejl ved indlæsning af Excel-fil: {e}")
            st.stop()
        item_no_col = find_column(user_df, ["item", "no"])
        product_name_col = find_column(user_df, ["product", "name"])
        if not item_no_col or not product_name_col:
            st.error("Kunne ikke finde kolonner for 'Item no' og 'Product name'.")
            st.stop()
        user_df = user_df.rename(columns={item_no_col: "Item no", product_name_col: "Product name"})
        user_df["Item no"] = user_df["Item no"].apply(str)
        st.write("Viser de første 10 rækker fra din data:")
        st.dataframe(user_df.head(10))
        try:
            variant_df = pd.read_excel("EY - variant master data.xlsx")
            lifestyle_df = pd.read_excel("EY - lifestyle.xlsx")
            line_df = pd.read_excel("EY - line drawing.xlsx")
        except Exception as e:
            st.error(f"Fejl ved indlæsning af eksterne datafiler: {e}")
            st.stop()
        try:
            # Indlæs templaten fra git
            template_pres = Presentation("template-generator.pptx")
            if not template_pres.slides:
                st.error("Din template-præsentation har ingen slides.")
                st.stop()
            template_slide = template_pres.slides[0]
        except Exception as e:
            st.error(f"Fejl ved indlæsning af PowerPoint template: {e}")
            st.stop()
        # Opret en ny præsentation baseret på templaten
        final_pres = Presentation("template-generator.pptx")
        # Fjern alle eksisterende slides fra final_pres (brug delete_slide workaround)
        for slide in list(final_pres.slides):
            delete_slide(final_pres, slide)
        for idx, row in user_df.iterrows():
            new_slide = copy_slide_from_template(template_slide, final_pres)
            fill_text_fields(new_slide, row, variant_df)
        ppt_io = io.BytesIO()
        final_pres.save(ppt_io)
        ppt_io.seek(0)
        st.download_button("Download tekstpræsentation", data=ppt_io, file_name="text_presentation.pptx")

#############################################
# Fase 2: Tilføj billeder til præsentation
#############################################
elif phase == "Tilføj billeder til præsentation":
    st.markdown("### Trin 2: Tilføj Billeder")
    uploaded_pptx = st.file_uploader("Upload din tekstbaserede præsentation (PPTX)", type=["pptx"], key="pptx")
    if uploaded_pptx:
        try:
            pres = Presentation(uploaded_pptx)
        except Exception as e:
            st.error(f"Fejl ved indlæsning af PowerPoint-fil: {e}")
            st.stop()
        try:
            variant_df = pd.read_excel("EY - variant master data.xlsx")
            lifestyle_df = pd.read_excel("EY - lifestyle.xlsx")
            line_df = pd.read_excel("EY - line drawing.xlsx")
        except Exception as e:
            st.error(f"Fejl ved indlæsning af eksterne datafiler: {e}")
            st.stop()
        # For hver slide, udtræk Item no og indsæt billeder
        for slide in pres.slides:
            item_no = ""
            for shape in slide.shapes:
                if shape.has_text_frame and "Product Code:" in shape.text:
                    match = re.search(r"Product Code:\s*(.*)", shape.text)
                    if match:
                        item_no = match.group(1).split("\n")[0].strip()
                        break
            if item_no:
                fill_image_fields(slide, item_no, variant_df, lifestyle_df, line_df)
            else:
                st.warning("Kunne ikke udtrække 'Product Code' fra en slide. Billeder indsættes ikke for denne slide.")
        ppt_io = io.BytesIO()
        pres.save(ppt_io)
        ppt_io.seek(0)
        st.download_button("Download præsentation med billeder", data=ppt_io, file_name="final_presentation.pptx")

#############################################
# Hjælpefunktion til at slette slides fra en præsentation
#############################################
def delete_slide(prs, slide):
    xml_slides = prs.slides._sldIdLst
    slide_id = slide.slide_id
    for sld in xml_slides:
        if sld.get("id") == str(slide_id):
            xml_slides.remove(sld)
            break
