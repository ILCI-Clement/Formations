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

def delete_slide(prs, slide_index):
    slide_id = prs.slides._sldIdLst[slide_index]
    prs.slides._sldIdLst.remove(slide_id)

############### Interface Streamlit #################

st.title("Gestionnaire de formations (Brochure + Site Internet + Gestion)")
st.header("Editeur de données")
st.warning("La dernière colonne ('Langue_Formation'), désigne la langue de la page de la formation ('English' ou 'Français')")

with st.sidebar:
    st.title("Aperçu de la brochure", width="content")
    st.pdf("template.pdf")

new_formation = st.data_editor(rows, num_rows="dynamic")

st.header("Etape 1:")
st.info("Envoyer les nouvelles données sur Airtable permet de les sauvegarder. Ainsi, si vous rafraîchissez la page, vous pourrez retrouver les modifications que vous avez faites.")

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


st.header("Etape 2:")
st.info("En générant la Brochure, vous allez créer un .pptx contenant toutes les données/formations présentent dans l'éditeur de données au dessus. " \
" Les formations seront agencées automatiquement dans la Brochure." \
" N'hésitez pas à vérifier l'ensemble de la Brochure, car il se peut que certains textes débordent des zones de textes.")
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

    by_model_before_del = {
        10: [],
        12: [],
        15: [],
        17: [],
        19: [],
    }

    model_after_del = {
        10: by_model_before_del[10],
        11: by_model_before_del[12],
        13: by_model_before_del[15],
        14: by_model_before_del[17],
        15: by_model_before_del[19]
    }

    for slide, model_index in created_slides:
        by_model_before_del[model_index].append(slide)


    slides_to_delete = [10, 12, 15, 17, 19]

    # On trie en ordre décroissant
    for index in sorted(slides_to_delete, reverse=True):
        delete_slide(prs, index)
        
    # On insère en ordre inverse pour éviter le décalage des index
    for model_index in [15, 14, 13, 11, 10]:  # ordre inverse des sections
        slides_group = model_after_del[model_index]

        # index d'insertion = juste après la slide modèle supprimée
        target_index = model_index

        # on insère en reverse pour éviter que ça décale
        for s in reversed(slides_group):
            move_slide_to(prs, s, target_index)

    ##### SOMMAIRE #######
    # Index de la slide Sommaire dans ton template
    sommaire_slide = prs.slides[5]

    # Dictionnaire pour stocker les titres
    sommaire_data = {
        "BACHELOR_FR": [],
        "BACHELOR_EN": [],
        "MASTERE_FR": [],
        "MASTERE_EN": [],
        "BAC6_FR": [],
        "BAC6_EN": [],
        "DOCTORATE": [],
        "BTS": [],
    }

    # On parcourt les formations
    for row in new_formation:
        nom = row["Nom"]
        type_form = row["Type"]
        langue = row["Langue_Formation"]

        if type_form == "BACHELOR":
            if langue == "English":
                sommaire_data["BACHELOR_EN"].append(nom)
            else:
                sommaire_data["BACHELOR_FR"].append(nom)

        elif type_form == "MASTERE":
            if langue == "English":
                sommaire_data["MASTERE_EN"].append(nom)
            else:
                sommaire_data["MASTERE_FR"].append(nom)

        elif type_form == "BAC+6":
            if langue == "English":
                sommaire_data["BAC6_EN"].append(nom)
            else:
                sommaire_data["BAC6_FR"].append(nom)

        elif type_form == "DOCTORATE":
            sommaire_data["DOCTORATE"].append(nom)

        elif type_form == "BTS":
            sommaire_data["BTS"].append(nom)


    # Mapping des TextBox du sommaire
    sommaire_mapping = {
        "TextBox 8": "\n".join(sommaire_data["BTS"])+"\n"+"\n".join(sommaire_data["BACHELOR_FR"]),
        "TextBox 5": "\n".join(sommaire_data["BACHELOR_EN"]),
        "TextBox 9": "\n".join(sommaire_data["MASTERE_FR"])+"\n"+"\n".join(sommaire_data["BAC6_FR"]),
        "TextBox 6": "\n".join(sommaire_data["MASTERE_EN"])+"\n"+"\n".join(sommaire_data["BAC6_EN"]),
        "TextBox 7": "\n".join(sommaire_data["DOCTORATE"]),
        "TextBox 10": "\n".join(sommaire_data["DOCTORATE"]),
    }

    # Injection dans la slide
    for shape in sommaire_slide.shapes:
        if shape.name in sommaire_mapping:
            update_text_preserve_style(shape, sommaire_mapping[shape.name])
    
    ###### PAGINATION ######
    i = 0
    for page in prs.slides:
        for shape in page.shapes:
            if shape.name == "TextBox 99":
                update_text_preserve_style(shape, i)
        i += 1
        
    prs.save("presentation.pptx")
    st.toast('PowerPoint mis à jour', icon="🔥")
    
    st.success("Brochure généré avec succès !")

    with open("presentation.pptx", "rb") as f:
        st.download_button(
            label="Télécharger la Brochure",
            data=f,
            file_name="presentation.pptx",
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
    )