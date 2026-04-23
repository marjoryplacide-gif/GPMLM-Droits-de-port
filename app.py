
import streamlit as st
import pandas as pd
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.units import cm
from reportlab.pdfgen import canvas
from reportlab.platypus import Table, TableStyle
import io
import math
import openpyxl
from pyairtable import Api

    
st.set_page_config(
    page_title="Droits de Port — GPMLM",
    page_icon=None,
    layout="wide"
)

@st.cache_data
def charger_navires():
    df = pd.read_excel(
        "Book 1.xlsx",
        sheet_name="Caractéristique navire de pêche"
    )
    df.columns = df.columns.str.strip()
    navires = {}
    for _, row in df.iterrows():
        compagnie = str(row["Compagnie"]).strip()
        navire = str(row["Nom du Navire"]).strip()
        if compagnie not in navires:
            navires[compagnie] = {}
        navires[compagnie][navire] = {
            "longueur": float(row["Longueur hors tout (m)"]),
            "largeur": float(row["Largueur maximal (m)"]),
            "tirant_eau": float(row["Tirant d'eau (m)"])
        }
    return navires

try:
    NAVIRES = charger_navires()
except Exception as e:
    st.error(f" Erreur lecture Excel : {e}")
    NAVIRES = {}

ADRESSES = {
    "DSK Fish": "53 AVENUE DES ARAWAKS, 97200 FORT DE FRANCE",
    "Poissonnerie Bapte": "12 RUE DE LA MER, 97200 FORT DE FRANCE",
    "Delta Transit": "8 BOULEVARD DU PORT, 97200 FORT DE FRANCE",
}
SIGNATAIRES = {
    "DSK Fish": {"nom": "DESCAS MARTHE", "qualite": "DIRECTRICE COMMERCIALE"},
    "Poissonnerie Bapte": {"nom": "SARL BAPTE", "qualite": "QUALITE"},
    "Delta Transit": {"nom": "SARL DELTA TRANSIT", "qualite": "QUALITE"},
}

TAUX_ENTREE = 0.366
REDEVANCE_DÉCHETS = 65
SEUIL_PERCEPTION = 9
MINIMUM_PERCEPTION = 16
def sauvegarder_declaration(data):
    try:
        import requests as req
        headers = {
            "Authorization": "Bearer " + st.secrets["AIRTABLE_TOKEN"],
            "Content-Type": "application/json"
        }
        url = "https://api.airtable.com/v0/" + st.secrets["AIRTABLE_BASE_ID"] + "/tblA6AtKUvtR5geXr"
        payload = {
            "fields": {
               "Date entree": data["date_entree"],
                "Date de sortie": data["date_sortie"],
                "Representant": data["representant"],
                "Nom du navire": data["nom_navire"],
                "Provenance": data["provenance"],
                "Pavillon": data["pavillon"],
                "Zone DN": data["zone_dn"],
                "Tonnage": str(data["tonnage"]),
                "Volume": str(data["volume"]),
                "Taux de base": str(data["taux_base"]),
                "Montant brut": str(data["montant_brut"]),
                "Montant net": str(data["montant_net"]),
                "Montant a percevoir": str(data["montant_percevoir"]),
                "Statut": "En attente"
            }
        }
        response = req.post(url, json=payload, headers=headers)
        response.raise_for_status()
        return True
    except Exception as e:
        st.error("Erreur sauvegarde : " + str(e))
        return False

def calcul_te_retenu(longueur, largeur, te_reel):
    te_theorique = 0.14 * (longueur * largeur) ** 0.5
    return round(max(te_reel, te_theorique), 2)
    
def calcul_volume(longueur, largeur, te_retenu):
    return math.ceil(longueur * largeur * te_retenu)

def calcul_modulation_art2(tonnage, volume):
    if volume == 0 or tonnage == 0:
        return 0
    rapport = tonnage / volume
    if rapport > 2/15: return 0
    elif rapport > 1/10: return -0.10
    elif rapport > 1/20: return -0.30
    elif rapport > 1/40: return -0.50
    elif rapport > 1/100: return -0.60
    elif rapport > 1/250: return -0.70
    elif rapport > 1/500: return -0.80
    else: return -0.95

def calcul_abattement_freq(nb_escales):
    if nb_escales <= 6: return 0
    elif nb_escales <= 15: return -0.10
    elif nb_escales <= 30: return -0.15
    elif nb_escales <= 60: return -0.20
    elif nb_escales <= 120: return -0.25
    else: return -0.30

def calcul_montant_final(redevance, modulation):
    montant_brut = round(redevance * (1 + modulation))
    montant_net = round(montant_brut - modulation)
    if montant_net < SEUIL_PERCEPTION:
        montant_percevoir = 0
    elif montant_net < MINIMUM_PERCEPTION:
        montant_percevoir = MINIMUM_PERCEPTION
    else:
        montant_percevoir = math.ceil(montant_net)
    return montant_brut, montant_net, montant_percevoir

def generer_pdf(data):
    BLEU_PORT = colors.HexColor('#1A5276')
    BLEU_CLAIR = colors.HexColor('#2E86C1')
    GRIS_CLAIR = colors.HexColor('#F2F3F4')
    GRIS_MOYEN = colors.HexColor('#D5D8DC')
    W, H = A4

    buffer = io.BytesIO()
    c = canvas.Canvas(buffer, pagesize=A4)

    def draw_section_title(titre, y):
        bh = 14
        c.setFillColor(BLEU_PORT)
        c.rect(1*cm, y - bh, W - 2*cm, bh, fill=1, stroke=0)
        c.setFillColor(colors.white)
        c.setFont("Helvetica-Bold", 7)
        c.drawString(1.3*cm, y - bh + 3, titre.upper())
        return y - bh

    def draw_field(label, value, x, y, lw=90, vw=130, h=13):
        c.setFillColor(GRIS_CLAIR)
        c.rect(x, y - h, lw, h, fill=1, stroke=0)
        c.setStrokeColor(GRIS_MOYEN)
        c.setLineWidth(0.4)
        c.rect(x, y - h, lw + vw, h, fill=0, stroke=1)
        c.setFillColor(BLEU_PORT)
        c.setFont("Helvetica-Bold", 6.5)
        c.drawString(x + 2, y - h + 3.5, label)
        c.setFillColor(colors.black)
        c.setFont("Helvetica", 7.5)
        c.drawString(x + lw + 3, y - h + 3.5, str(value))

    def draw_table(table_data, col_widths, x, y, last_row_blue=False):
        t = Table(table_data, colWidths=col_widths)
        style = [
            ('BACKGROUND', (0,0), (-1,0), BLEU_CLAIR),
            ('TEXTCOLOR', (0,0), (-1,0), colors.white),
            ('FONTNAME', (0,0), (-1,0), 'Helvetica-Bold'),
            ('FONTSIZE', (0,0), (-1,-1), 7.5),
            ('ALIGN', (0,0), (-1,-1), 'CENTER'),
            ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
            ('GRID', (0,0), (-1,-1), 0.4, GRIS_MOYEN),
            ('BACKGROUND', (0,1), (-1,-1), GRIS_CLAIR),
            ('ROWHEIGHT', (0,0), (-1,-1), 13),
        ]
        if last_row_blue:
            style += [
                ('BACKGROUND', (0,-1), (-1,-1), BLEU_PORT),
                ('TEXTCOLOR', (0,-1), (-1,-1), colors.white),
                ('FONTNAME', (0,-1), (-1,-1), 'Helvetica-Bold'),
            ]
        t.setStyle(TableStyle(style))
        tw, th = t.wrapOn(c, W, H)
        t.drawOn(c, x, y - th)
        return th

    y = H - 10
    c.setFont("Helvetica-Bold", 14)
    c.setFillColor(BLEU_PORT)
    c.drawCentredString(W/2, H - 35, "DÉCLARATION NAVIRE")
    c.setFont("Helvetica", 8)
    c.setFillColor(BLEU_CLAIR)
    c.drawCentredString(W/2, H - 50, "Navires de Pêche — Grand Port Maritime de la Martinique")
    c.setFont("Helvetica-Bold", 8)
    c.setFillColor(BLEU_PORT)
    c.drawRightString(W - 1*cm, H - 25, f"DN N° {data['dn_numero']}")
    c.setFillColor(BLEU_CLAIR)
    c.roundRect(W - 4*cm, H - 60, 3*cm, 18, 4, fill=1, stroke=0)
    c.setFillColor(colors.white)
    c.setFont("Helvetica-Bold", 9)
    c.drawCentredString(W - 2.5*cm, H - 52, "ENTREE")
    c.setStrokeColor(BLEU_PORT)
    c.setLineWidth(1.5)
    c.line(1*cm, H - 70, W - 1*cm, H - 70)
    y = H - 78

    y = draw_section_title("1. Identification du navire et de l'escale", y)
    y -= 3
    draw_field("Nom du navire", data['nom_navire'], 1*cm, y, 75, 120)
    draw_field("Bureau de", "FR06340 FORT DE FRANCE", 9*cm, y, 55, 155)
    y -= 16
    draw_field("Provenance", data['provenance'], 1*cm, y, 65, 120)
    draw_field("Port de", "FORT DE FRANCE", 9*cm, y, 45, 100)
    y -= 16
    draw_field("Date d'entrée", data['date_entree'], 1*cm, y, 70, 90)
    draw_field("Date de sortie", data['date_sortie'], 9*cm, y, 70, 65)
    y -= 16
    draw_field("Zone DN", data['zone_dn'], 1*cm, y, 50, 40)
    draw_field("Type de navire", data['type_navire'], 9*cm, y, 75, 70)
    y -= 16
    y -= 5
    y = draw_section_title("2. Représentant", y)
    y -= 3
    draw_field("Représentant", data['representant'], 1*cm, y, 75, 130)
    y -= 16
    draw_field("Adresse", data['adresse_rep'], 1*cm, y, 50, 380)
    y -= 16
    y -= 5
    y = draw_section_title("3. Tonnage des marchandises", y)
    y -= 3
    draw_field("Marchandises diverses", f"{data['tonnage']} t", 1*cm, y, 120, 80)
    draw_field("TOTAL", f"{data['tonnage']} t", 11*cm, y, 40, 80)
    y -= 16
    y -= 5
    y = draw_section_title("4. Redevance sur le navire (V335)", y)
    y -= 3
    c.setFont("Helvetica", 7)
    c.setFillColor(colors.black)
    c.drawString(1*cm, y - 10, "Seuil de perception : 9,0 €     Minimum de perception : 16,0 €")
    y -= 14
    nav_data = [
        ["Longueur hors tout", "Largeur", "Tirant d'eau", "Volume taxable", "Taux de base"],
        [f"{data['longueur']} m", f"{data['largeur']} m", f"{data['te_retenu']} m",
         f"{data['volume']} m³", str(data['taux_base'])],
    ]
    th = draw_table(nav_data, [3.5*cm, 3*cm, 3*cm, 3.5*cm, 3.5*cm], 1*cm, y)
    y -= (th + 13)
    y = draw_section_title("5. Modulations et abattements", y)
    y -= 3
    montant_mod = round(data['redevance_navire'] - data['montant_apres_mod'], 2)
    mod_data = [
    ["", "Taux", "Montant (€)"],
    ["Modulation selon taux de remplissage (Art. 2)", str(int(data['mod_art2']*100)) + "%", str(montant_mod) + " €"],
    ["Abattement de fréquence (Art. 3)", str(int(data['mod_art3']*100)) + "%", ""],
    ["Modulation environnementale ESI (Art. 4)", "0%", ""],
]
    th = draw_table(mod_data, [10*cm, 3*cm, 3.5*cm], 1*cm, y)
    y -= (th + 13)
    y = draw_section_title("6. Liquidation - Redevance sur navire", y)
    y -= 3
    montant_mod_affiche = round(data['redevance_navire'] - data['montant_apres_mod'])
    montant_net_navire = data['redevance_navire'] - montant_mod_affiche
    liq_data = [
        ["", " (€)"],
        [" brut (Redevance sur navire)", str(data['redevance_navire']) + " €"],
        ["Total des modulations", str(montant_mod_affiche) + " €"],
        ["Montant net", str(data['redevance_navire'] - montant_mod_affiche) + " €"],
        ["MONTANT A PERCEVOIR", str(data['montant_percevoir']) + " €"],
    ]
    th = draw_table(liq_data, [12*cm, 4.5*cm], 1*cm, y, last_row_blue=True)
    y -= (th + 13)
    y = draw_section_title("7. Droits de port a percevoir", y)
    y -= 3
    dp_data = [
        ["Code", "Libelle", "Montant (€)"],
        ["V335", "Redevance sur navire", str(montant_net_navire) + " €"],
        ["V365", "Redevance sur les déchets d'exploitation", "65 €"],
        ["", "TOTAL", str(montant_net_navire + 65) + " €"],
    ]
    th = draw_table(dp_data, [2.5*cm, 10*cm, 4*cm], 1*cm, y, last_row_blue=True)
    y -= (th + 13)
    c.setFont("Helvetica", 8)
    c.setFillColor(colors.black)
    c.drawString(1*cm, y, "Je soussigné(e)")
    c.line(3.8*cm, y, 12*cm, y)
    y -= 18
    c.drawString(1*cm, y, "Qualité")
    c.line(2.5*cm, y, 12*cm, y)
    y -= 18
    c.setFont("Helvetica", 7.5)
    c.drawString(1*cm, y, "certifie sous les peines de droits, l'exactitude des énonciations de la présente déclaration.")
    y -= 20
    c.setFont("Helvetica", 8)
    c.drawString(1*cm, y, "À Fort de France, le")
    c.line(4.5*cm, y, 8*cm, y)
    c.drawString(13*cm, y, "Signature :")
    c.rect(15*cm, y - 25, 3*cm, 30, fill=0, stroke=1)
    c.drawString(3.8*cm, y, data['signataire'])
    y -= 18
    c.drawString(2.5*cm, y, data['qualite'])
    c.setStrokeColor(BLEU_PORT)
    c.setLineWidth(1)
    c.line(1*cm, 20, W - 1*cm, 20)
    c.setFont("Helvetica", 6)
    c.setFillColor(BLEU_PORT)
    c.drawCentredString(W/2, 10, "Grand Port Maritime de la Martinique — Martinique Hub Caraïbe — Our Future is Maritime")
    c.save()
    buffer.seek(0)
    return buffer

st.title("Déclaration des Droits de Port")
st.subheader("Grand Port Maritime de la Martinique — Navires de pêche")
st.markdown("---")

col1, col2 = st.columns(2)

with col2:
    st.header(" Escale")
    date_entree =  st.date_input("Date d'entrée", format="DD/MM/YYYY")
    date_sortie =  st.date_input("Date de sortie", format="DD/MM/YYYY")
    provenance = st.selectbox("Provenance (port d'origine)", ["MARGUARITA","VENEZUELA","GRENADE"])
    zone_dn = st.selectbox("Zone DN", ["A","B","C","D","E","F","G","H","I","J","M","R","Z"], index =9)
    tonnage = st.number_input("Tonnage (tonnes)", min_value=0.0, max_value=10.0, step=0.001, format="%.3f")
    nb_escales = st.number_input("Nombre d'escales dans l'année", min_value=1, step=1)


with col1:
    st.header ("Navire")
    if NAVIRES:
        representant = st.selectbox("Représentant", list(NAVIRES.keys()))
        tous_navires = list(set(
            navire 
            for compagnie in NAVIRES.values() 
            for navire in compagnie.keys()
        ))
        tous_navires.sort()
        nom_navire = st.selectbox("Nom du navire", tous_navires)
        carac = next(
            NAVIRES[c][nom_navire] 
            for c in NAVIRES 
            if nom_navire in NAVIRES[c]
        )
        longueur = carac["longueur"]
        largeur = carac["largeur"]
        tirant_eau = carac["tirant_eau"]
        st.info("Caractéristiques : L=" + str(longueur) + "m | b=" + str(largeur) + "m | Te=" + str(tirant_eau) + "m")
    else:
        st.warning("Aucune donnee de navire disponible.")
        representant = nom_navire = ""
        longueur = largeur = tirant_eau = 0

st.markdown("---")
st.markdown("""
<style>
    .stApp {
        background-color: #EEF2F7;
    }
    .stButton > button {
        background-color: #1A5276;
        color: white;
        border-radius: 6px;
        padding: 12px 20px;
        font-weight: bold;
        border: none;
        width: 100%;
        box-shadow: 0 4px 6px rgba(0,0,0,0.1);
    }
    .stButton > button:hover {
        background-color: #7AB32E;
        box-shadow: 0 6px 12px rgba(0,0,0,0.15);
    }
    .stTextInput > div > div > input {
        background-color: white !important;
        border: 1.5px solid #AED6F1 !important;
        border-radius: 6px !important;
    }
    .stNumberInput > div > div > input {
        background-color: white !important;
        border: 1.5px solid #AED6F1 !important;
        border-radius: 6px !important;
    }
    input[type="text"] {
        background-color: white !important;
    }
    input[type="number"] {
        background-color: white !important;
    }
    .stSelectbox > div > div {
        background-color: white;
        border: 1.5px solid #AED6F1;
        border-radius: 6px;
        box-shadow: 0 2px 4px rgba(0,0,0,0.06);
    }
    h1 {
        color: #1A5276;
        font-weight: 800;
    }
    h2 {
        color: #1A5276;
        background-color: white;
        padding: 10px 15px;
        border-radius: 8px;
        border-left: 4px solid #7AB32E;
        box-shadow: 0 2px 4px rgba(0,0,0,0.08);
    }
    h3 {
        color: #2E86C1;
    }
    div[data-testid="metric-container"] {
        background-color: white;
        border-radius: 10px;
        padding: 15px;
        box-shadow: 0 4px 10px rgba(0,0,0,0.1);
        border-top: 4px solid #7AB32E;
    }
    .stInfo {
        background-color: white;
        border-radius: 8px;
        box-shadow: 0 2px 4px rgba(0,0,0,0.08);
        border-left: 4px solid #1A5276;
    }
</style>
""", unsafe_allow_html=True)

if st.button(" Calculer et générer le DN", type="primary"):
    if not provenance:
        st.error(" Veuillez renseigner la provenance.")
    elif longueur == 0:
        st.error(" Aucune caractéristique de navire disponible.")
    elif tonnage == 0:
        st.error("Veuillez renseigner le tonnage.")
    else:
        te_retenu = calcul_te_retenu(longueur, largeur, tirant_eau)
        volume = calcul_volume(longueur, largeur, te_retenu)
        redevance_navire = round(volume * TAUX_ENTREE)
        mod_art2 = calcul_modulation_art2(tonnage, volume)
        mod_art3 = calcul_abattement_freq(nb_escales)
        modulation_retenue = min(mod_art2, mod_art3)
        montant_brut, montant_net, montant_percevoir = calcul_montant_final(redevance_navire, modulation_retenue)

        st.success("Calculs effectues avec succes !")
        st.markdown("### Resultats")
        c1, c2, c3, c4 = st.columns(4)
        c1.metric("Volume taxable", str(volume) + " m3")
        c2.metric("Redevance navire", str(redevance_navire) + " €")
        c3.metric("Modulation retenue", str(int(modulation_retenue*100)) + "%")
        c4.metric("Montant a percevoir", str(montant_percevoir) + " €")
        col_a, col_b = st.columns(2)
        col_a.info("Art. 2 (taux remplissage) : " + str(int(mod_art2*100)) + "%")
        col_b.info("Art. 3 (fréquence, escale n " + str(nb_escales) + ") : " + str(int(mod_art3*100)) + "%")

        data_pdf = {
            "dn_numero":"",
            "escale_numero":"",
            "nom_navire": nom_navire,
            "pavillon": "",
            "type_navire": "13",
            "provenance": provenance,
            "zone_dn": zone_dn,
            "date_entree": date_entree.strftime("%d/%m/%Y"),
            "date_sortie": date_sortie.strftime("%d/%m/%Y"),
            "representant": representant,
            "adresse_rep": ADRESSES.get(representant, ""),
            "tonnage": tonnage,
            "longueur": longueur,
            "largeur": largeur,
            "te_retenu": te_retenu,
            "volume": volume,
            "taux_base": TAUX_ENTREE,
            "redevance_navire": redevance_navire,
            "mod_art2": mod_art2,
            "mod_art3": mod_art3,
            "montant_brut": redevance_navire,
            "montant_apres_mod": montant_brut,
            "montant_net": montant_net,
            "montant_percevoir": montant_percevoir,
            "signataire": SIGNATAIRES.get(representant, {}).get("nom", ""),
            "qualite": SIGNATAIRES.get(representant, {}).get("qualite", ""),
        }
        sauvegarder_declaration(data_pdf)
        pdf_buffer = generer_pdf(data_pdf)
        st.download_button(
            label=" Télécharger la DN (PDF)",
            data=pdf_buffer,
            file_name=f"DN_{nom_navire}.pdf",
            mime="application/pdf"
        )
