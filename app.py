import streamlit as st
import pandas as pd
from pptx import Presentation
import gspread
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload
import io, os, subprocess, re, shutil, requests, base64
from pypdf import PdfWriter
from datetime import datetime

# --- KONFIGURACJA API AI (Z Twojego kodu React) ---
API_KEY = st.secrets.get("GEMINI_API_KEY", "") # Dodaj klucz do Streamlit Secrets

# --- BAZY DANYCH (Przeniesione z React) ---
CAR_DATABASE = {
    "Renault": {"Scenic E-Tech": ["Crossover"], "Megane E-Tech": ["Hatchback"], "Austral": ["SUV"]},
    "Audi": {"A6": ["Sedan", "Avant"], "RS6": ["Avant"], "Q8": ["SUV"], "e-tron GT": ["Sedan"]},
    "BMW": {"M3": ["Sedan"], "M4": ["Coupe"], "X5": ["SUV"], "Seria 5": ["Sedan"]},
    "Porsche": {"911 (992)": ["Coupe", "Cabriolet"], "Taycan": ["Sedan"], "Cayenne": ["SUV"]}
}

FOIL_GROUPS = {
    "3M 2080 Series": {
        "Satin": ["Satin Black (S12)", "Satin Dark Grey (S162)", "Satin Vampire Red (S208)"],
        "Matte": ["Matte Black (M12)", "Matte Military Green (M26)"],
        "Gloss": ["Gloss Black (G12)", "Gloss Deep Blue"]
    },
    "Avery Dennison SW900": {
        "Satin": ["Satin Khaki Green", "Satin Metallic Grey"],
        "Gloss": ["Gloss Rock Grey", "Gloss Carmine Red"]
    }
}

# --- FUNKCJE SYSTEMOWE ---
def install_fonts():
    font_src = "fonts"
    font_dst = os.path.expanduser("~/.local/share/fonts")
    if os.path.exists(font_src):
        if not os.path.exists(font_dst): os.makedirs(font_dst)
        for f in os.listdir(font_src):
            if f.lower().endswith((".ttf", ".otf")):
                shutil.copy(os.path.join(font_src, f), font_dst)
        subprocess.run(["fc-cache", "-f"], capture_output=True)

def generate_ai_image(prompt):
    url = f"https://generativelanguage.googleapis.com/v1beta/models/imagen-4.0-generate-001:predict?key={API_KEY}"
    payload = {
        "instances": [{"prompt": prompt}],
        "parameters": {"sampleCount": 1}
    }
    response = requests.post(url, json=payload)
    if response.status_code == 200:
        data = response.json()
        img_b64 = data['predictions'][0]['bytesBase64Encoded']
        return base64.b64decode(img_b64)
    return None

def pptx_to_pdf(input_path):
    try:
        subprocess.run(['libreoffice', '--headless', '--convert-to', 'pdf', '--outdir', os.getcwd(), input_path], check=True, capture_output=True)
        pdf_name = os.path.basename(input_path).replace('.pptx', '.pdf')
        return pdf_name if os.path.exists(pdf_name) else None
    except: return None

# --- APLIKACJA ---
st.set_page_config(page_title="Studio Ultimate & Generator Ofert", layout="wide")
install_fonts()

# Autoryzacja Google
creds = Credentials.from_service_account_info(st.secrets["gcp_service_account"], 
        scopes=["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"])
service = build('drive', 'v3', credentials=creds)
client = gspread.authorize(creds)

st.title("Zap & Studio Ultimate")

# --- SIDEBAR: KONFIGURACJA POJAZDU I FOLII ---
with st.sidebar:
    st.header("🚗 Konfigurator AI")
    brand = st.selectbox("Marka", list(CAR_DATABASE.keys()))
    model_list = list(CAR_DATABASE[brand].keys())
    model = st.selectbox("Model", model_list)
    body = st.selectbox("Nadwozie", CAR_DATABASE[brand][model])
    year = st.selectbox("Rocznik", ["2025", "2024", "2023"])
    
    st.header("🎨 Wybór Folii")
    f_brand = st.selectbox("Producent", list(FOIL_GROUPS.keys()))
    f_cat = st.selectbox("Typ", list(FOIL_GROUPS[f_brand].keys()))
    f_color = st.selectbox("Kolor", FOIL_GROUPS[f_brand][f_cat])

    if st.button("🤖 GENERUJ WIZUALIZACJĘ AI"):
        prompt = f"Professional automotive studio photography of a {year} {brand} {model} ({body}) wrapped in {f_brand} {f_color}. STUDIO SETTING: High-end detailing garage. WALLS: Matte black. CEILING: Large white HEXAGONAL HONEYCOMB LED lights. FLOOR: Polished black epoxy with clear reflections. Sharp focus, 8k resolution."
        with st.spinner("AI renderuje Twoje auto..."):
            img_data = generate_ai_image(prompt)
            if img_data:
                st.session_state['generated_img'] = img_data
                st.success("Wizualizacja gotowa!")
            else:
                st.error("Błąd API Imagen. Sprawdź klucz API.")

# --- GŁÓWNY PANEL ---
col1, col2 = st.columns([1, 1])

with col1:
    st.subheader("Dane do oferty")
    klient = st.text_input("Imię i Nazwisko Klienta")
    nr_o = st.text_input("Numer oferty", value=f"IW/{datetime.now().strftime('%Y/%m/%d')}/01")
    
    # Pobieranie cennika z Google Sheets
    url_arkusza = "https://docs.google.com/spreadsheets/d/1iqS6geTNP3Bd_Fj_XdS-wCBrKtnGTMNQZYSso70KIkQ/edit?usp=drive_link"
    sheet = client.open_by_url(url_arkusza).worksheet("Ppf")
    df_prices = pd.DataFrame(sheet.get_all_values()[1:], columns=[c.strip() for c in sheet.get_all_values()[0]])
    
    pakiet = st.selectbox("Wybierz pakiet z cennika", df_prices['Usługa'].tolist())
    rabat = st.number_input("Rabat (PLN)", value=0)

with col2:
    st.subheader("Podgląd wizualizacji")
    if 'generated_img' in st.session_state:
        st.image(st.session_state['generated_img'], use_container_width=True)
    else:
        st.info("Skonfiguruj auto w panelu bocznym i wygeneruj zdjęcie.")

# --- GENEROWANIE PDF ---
if st.button("🚀 GENERUJ I WYŚLIJ OFERTĘ PDF"):
    if 'generated_img' not in st.session_state:
        st.error("Najpierw wygeneruj wizualizację AI!")
    else:
        with st.spinner("Składam ofertę PDF..."):
            # Logika pobierania plików z Drive i podmiany tagów (jak w Twoim poprzednim kodzie)
            # ... (TUTAJ TWOJA DOTYCHCZASOWA LOGIKA POBIERANIA I REPLACEMENTS)
            
            # Kluczowy moment: Wstawianie wygenerowanego zdjęcia AI
            # shape.add_picture(io.BytesIO(st.session_state['generated_img']), ...)
            
            st.success("Oferta wygenerowana!")
            st.download_button("📥 POBIERZ OFERTĘ", data="...", file_name=f"Oferta_{klient}.pdf")
