import streamlit as st
from pptx import Presentation
import requests
from copy import deepcopy

AIRTABLE_TOKEN = st.secrets["AIRTABLE_TOKEN"]
BASE_ID = "app7vq5k1lztBcmNF"
TABLE_NAME = "Informations"

url = f"https://api.airtable.com/v0/{BASE_ID}/{TABLE_NAME}"


headers = {
    "Authorization": f"Bearer {AIRTABLE_TOKEN}",
    "Content-Type": "application/json"
}

params = {
    "maxRecords": 100,
    "view": "Grid view"
}

resp = requests.get(url, headers=headers, params=params)
data = resp.json()

records = data["records"]
ids = [r["id"] for r in records]

rows = []
for r in records:
    row = {}
    row.update(r["fields"])
    rows.append(row)

def chunk_list(lst, size=10):
    for i in range(0, len(lst), size):
        yield lst[i:i+size]

def duplicate_slide(prs, slide_index):
    source_slide = prs.slides[slide_index]

    # Utilise le même layout que la slide source
    layout = source_slide.slide_layout
    new_slide = prs.slides.add_slide(layout)

    # Supprimer les placeholders par défaut
    for shape in list(new_slide.shapes):
        el = shape.element
        el.getparent().remove(el)

    # Copier toutes les shapes
    for shape in source_slide.shapes:
        new_shape = deepcopy(shape.element)
        new_slide.shapes._spTree.insert_element_before(
            new_shape, 'p:extLst'
        )

    return new_slide

def move_slide_to(prs, slide, target_index):
    slides = prs.slides._sldIdLst

    # retrouver l'élément XML de la slide
    for i, sldId in enumerate(slides):
        if prs.slides[i] == slide:
            slide_id = sldId
            current_index = i
            break
    else:
        return  # slide non trouvée

    # déplacer
    slides.remove(slide_id)
    slides.insert(target_index, slide_id)

def update_text_preserve_style(shape, new_text):
    if not shape.has_text_frame:
        return

    tf = shape.text_frame
    if not tf.paragraphs:
        return

    p = tf.paragraphs[0]
    runs = p.runs

    if runs:
        runs[0].text = str(new_text) if new_text is not None else ""
        for r in runs[1:]:
            r.text = ""
    else:
        p.text = str(new_text) if new_text is not None else ""

############### Interface Streamlit #################

st.title("Gestionnaire de formations (Brochure + Site Internet + Gestion)")

with st.sidebar:
    st.title("Aperçu de la brochure", width="content")
    st.pdf("template.pdf")

new_formation = st.data_editor(rows, num_rows="dynamic")

if st.button("Envoyer sur Airtable"):
    airtable_payload = {
        "records": []
    }

    for row in new_formation:
        airtable_payload["records"].append({
            "fields": row
    })

    for batch in chunk_list(ids, 10):
        params = [("records[]", rid) for rid in batch]
        requests.delete(url, headers=headers, params=params)

    for batch in chunk_list(airtable_payload["records"], 10):
        payload = {"records": batch}
        send = requests.post(url, headers=headers, json=payload)

        if send.status_code != 200:
            st.error(send.text)


if st.button("Générer la Brochure"):
    prs = Presentation("template.pptx")

    created_slides = []

    for rows in new_formation:
        mapping = {
            "TextBox 48": rows["Nom"],
            "TextBox 50": rows["Type"],
            "TextBox 34": rows["Langues"],
            "TextBox 38": rows["Stage"],
            "TextBox 42": rows["Description"],
            "TextBox 33": rows["PointFort1"],
            "TextBox 37": rows["PointFort2"],
            "TextBox 40": rows["PointFort3"],
            "TextBox 32": rows["Enseignement1"],
            "TextBox 36": rows["Enseignement2"],
            "TextBox 41": rows["Enseignement3"],
            "TextBox 35": rows["Metier"],
            "TextBox 39": rows["Admission"],
        }

        english_map = {
            "TextBox 51": "2 INTAKES",
            "TextBox 52": "LANGUAGE",
            "TextBox 53": "INTERNSHIP",
            "TextBox 49": "ADMISSION",
            "TextBox 47": "TOP3 - KEY POINTS",
            "TextBox 46": "TEACHING UNITS",
            "TextBox 45": "OCCUPATIONS",
            "TextBox 31": "Fall Spring",
        }
        
        francais_map = {
            "TextBox 51": "2 RENTRÉES",
            "TextBox 52": "LANGUE",
            "TextBox 53": "STAGE",
            "TextBox 49": "ADMISSION",
            "TextBox 47": "TOP3 - POINTS FORTS",
            "TextBox 46": "ENSEIGNEMENTS",
            "TextBox 45": "OPPORTUNITÉS MÉTIERS",
            "TextBox 31": "Automne Printemps",
        }

        if(rows["Type"] == "BACHELOR"):
            slide = duplicate_slide(prs, 12)
            created_slides.append((slide, 12))
        elif(rows["Type"] == "MASTERE"):
            slide = duplicate_slide(prs, 15)
            created_slides.append((slide, 15))
        elif(rows["Type"] == "BAC+6"):
            slide = duplicate_slide(prs, 17)
            created_slides.append((slide, 17))
        elif(rows["Type"] == "DOCTORATE"):
            slide = duplicate_slide(prs, 19)
            created_slides.append((slide, 19))
        elif(rows["Type"] == "BTS"):
            slide = duplicate_slide(prs, 10)
            created_slides.append((slide, 10))
        
        for shape in slide.shapes:
            if shape.name in mapping:
                update_text_preserve_style(shape, mapping[shape.name])
        
        if rows["Langue_Formation"] == "English":
            for shape in slide.shapes:
                if shape.name in english_map:
                    update_text_preserve_style(shape, english_map[shape.name])
        else:
            for shape in slide.shapes:
                if shape.name in francais_map:
                    update_text_preserve_style(shape, francais_map[shape.name])

    by_model = {
        10: [],
        12: [],
        15: [],
        17: [],
        19: [],
    }

    for slide, model_index in created_slides:
        by_model[model_index].append(slide)

    # on les insère dans l'ordre, après chaque slide modèle
    offset = 0
    for model_index in [10, 12, 15, 17, 19]:
        target_base = model_index + 1 + offset

        for s in by_model[model_index]:
            move_slide_to(prs, s, target_base)
            target_base += 1
            offset += 1

    i = 0
    for page in prs.slides:
        for shape in page.shapes:
            if shape.name == "TextBox 99":
                update_text_preserve_style(shape, i)
        i += 1
        
    prs.save("presentation.pptx")
    st.success("PowerPoint mis à jour")

    
    with open("presentation.pptx", "rb") as f:
        st.download_button(
            label="Télécharger la Brochure",
            data=f,
            file_name="presentation.pptx",
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
    )