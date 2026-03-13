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

# --- PEŁNA BAZA SAMOCHODÓW ---
CAR_DATABASE = {
    "Audi": {"A3": ["Hatchback", "Sedan"], "A4": ["Sedan", "Kombi"], "A6": ["Sedan", "Kombi"], "Q3": ["SUV"], "Q5": ["SUV"], "Q8": ["SUV"], "e-tron GT": ["Sedan"], "RS6": ["Kombi"]},
    "BMW": {"Seria 3": ["Sedan", "Kombi"], "Seria 4": ["Coupe", "Gran Coupe"], "Seria 5": ["Sedan", "Kombi"], "X3": ["SUV"], "X5": ["SUV"], "M3": ["Sedan", "Kombi"], "M4": ["Coupe"]},
    "BYD": {"Seal": ["Sedan"], "Atto 3": ["SUV"], "Han": ["Sedan"], "Dolphin": ["Hatchback"]},
    "Ford": {"Focus": ["Hatchback", "Kombi"], "Mustang": ["Coupe", "Cabriolet"], "Mustang Mach-E": ["SUV"], "Puma": ["Crossover"]},
    "Hyundai": {"Tucson": ["SUV"], "Ioniq 5": ["Hatchback/Crossover"], "Ioniq 6": ["Sedan"], "i30": ["Hatchback", "Kombi"], "Kona": ["Crossover"]},
    "Kia": {"EV6": ["Crossover"], "Sportage": ["SUV"], "Ceed": ["Hatchback", "Kombi"], "Stinger": ["Liftback"], "Sorento": ["SUV"]},
    "Lexus": {"NX": ["SUV"], "RX": ["SUV"], "ES": ["Sedan"], "LC": ["Coupe"]},
    "Mercedes-Benz": {"Klasa C": ["Sedan", "Kombi"], "Klasa E": ["Sedan", "Kombi"], "GLC": ["SUV", "Coupe"], "GLE": ["SUV", "Coupe"], "Klasa G": ["SUV"], "AMG GT": ["Coupe"]},
    "MG": {"MG4": ["Hatchback"], "HS": ["SUV"], "ZS": ["SUV"], "Cyberster": ["Roadster"]},
    "NIO": {"ET7": ["Sedan"], "ET5": ["Sedan"], "EL7": ["SUV"]},
    "Porsche": {"911 (992)": ["Coupe", "Cabriolet"], "Taycan": ["Sedan", "Cross Turismo"], "Macan": ["SUV"], "Panamera": ["Sedan"], "Cayenne": ["SUV", "Coupe"]},
    "Renault": {"Scenic E-Tech": ["Crossover"], "Megane E-Tech": ["Hatchback"], "Austral": ["SUV"], "Clio": ["Hatchback"], "Captur": ["Crossover"]},
    "Skoda": {"Octavia": ["Liftback", "Kombi"], "Superb": ["Liftback", "Kombi"], "Kodiaq": ["SUV"], "Enyaq": ["SUV", "Coupe"]},
    "Tesla": {"Model 3": ["Sedan"], "Model Y": ["SUV"], "Model S": ["Sedan"], "Model X": ["SUV"]},
    "Toyota": {"Corolla": ["Hatchback", "Sedan", "Kombi"], "Yaris": ["Hatchback"], "RAV4": ["SUV"], "C-HR": ["Crossover"], "Camry": ["Sedan"]},
    "Volkswagen": {"Golf": ["Hatchback", "Kombi"], "Passat": ["Kombi", "Sedan"], "Arteon": ["Liftback", "Kombi"], "ID.4": ["SUV"], "Tiguan": ["SUV"]},
    "Volvo": {"XC40": ["SUV"], "XC60": ["SUV"], "XC90": ["SUV"], "V60": ["Kombi"]}
}

# --- PEŁNA BAZA FOLII (W TYM XPEL PPF) ---
FOIL_GROUPS = {
    "XPEL (Folie Ochronne PPF)": {
        "Bezbarwne (Twój obecny kolor)": [
            "XPEL Ultimate Plus (Wysoki Połysk)", 
            "XPEL Stealth (Mat/Satyna)"
        ],
        "XPEL Color (Zmiana Koloru PPF)": [
            "Black (Połysk)", "White (Połysk)", "Red (Połysk)", 
            "Nardo Grey (Połysk)", "Miami Blue (Połysk)"
        ]
    },
    "3M 2080 Series": {
        "Matte (Matowe)": ["Matte Black (M12)", "Matte Deep Black (M22)", "Matte Dark Grey (M261)", "Matte White (M10)", "Matte Military Green (M26)"],
        "Satin (Satynowe)": ["Satin Black (S12)", "Satin Dark Grey (S162)", "Satin White (S10)", "Satin Vampire Red (S208)"],
        "Gloss (Połysk)": ["Gloss Black (G12)", "Gloss White (G10)", "Gloss Hot Rod Red (G13)", "Gloss Sky Blue (G77)"],
        "Color Flip (Kameleon)": ["Gloss Flip Electric Wave (GP287)", "Satin Flip Volcanic Flare (SP236)"]
    },
    "Avery Dennison SW900": {
        "Satin": ["Satin Black", "Satin Pearl White", "Satin Carmine Red", "Satin Khaki Green", "Satin Metallic Grey"],
        "Gloss": ["Gloss Black", "Gloss White", "Gloss Obsidian Black", "Gloss Rock Grey", "Gloss Carmine Red"],
        "Matte": ["Matte Black", "Matte White", "Matte Charcoal Metallic", "Matte Olive Green"]
    },
    "Oracal 970RA": {
        "Special": ["Gloss Telegrey", "Gloss Nardo Grey Style", "Matte Nato Olive"],
        "Metallic": ["Gloss Graphite Metallic", "Matte Anthracite Metallic", "Gloss Silver Grey"]
    }
}

# --- FUNKCJE SYSTEMOWE ---
def install_fonts():
    font_src, font_dst = "fonts", os.path.expanduser("~/.local/share/fonts")
    if os.path.exists(font_src):
        if not os.path.exists(font_dst): os.makedirs(font_dst)
        for f in os.listdir(font_src):
            if f.lower().endswith((".ttf", ".otf")): shutil.copy(os.path.join(font_src, f), font_dst)
        subprocess.run(["fc-cache", "-f"], capture_output=True)

def generate_ai_image(prompt):
    api_key = st.secrets["GEMINI_API_KEY"]
    url = f"https://generativelanguage.googleapis.com/v1beta/models/imagen-4.0-generate-001:predict?key={api_key}"
    payload = {"instances": [{"prompt": prompt}], "parameters": {"sampleCount": 1}}
    try:
        response = requests.post(url, json=payload, timeout=60)
        if response.status_code == 200:
            img_b64 = response.json()['predictions'][0]['bytesBase64Encoded']
            return base64.b64decode(img_b64)
    except Exception as e:
        st.error(f"Błąd AI: {e}")
    return None

def download_file(service, file_id):
    request = service.files().get_media(fileId=file_id)
    fh = io.BytesIO(); downloader = MediaIoBaseDownload(fh, request)
    done = False
    while not done: _, done = downloader.next_chunk()
    fh.seek(0); return fh

def pptx_to_pdf(input_path):
    try:
        subprocess.run(['libreoffice', '--headless', '--convert-to', 'pdf', '--outdir', os.getcwd(), input_path], check=True, capture_output=True)
        pdf_name = os.path.basename(input_path).replace('.pptx', '.pdf')
        return pdf_name if os.path.exists(pdf_name) else None
    except: return None

# --- APLIKACJA ---
st.set_page_config(page_title="Zap & Studio Ultimate", layout="wide")
install_fonts()

# Autoryzacja i pobranie bazy plików na start
creds = Credentials.from_service_account_info(st.secrets["gcp_service_account"], 
        scopes=["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"])
service = build('drive', 'v3', credentials=creds)
client = gspread.authorize(creds)

FOLDER_ID = "12HRnKn9KrZy_C1BSgv24PGD-Gl8lTRmn"
q = f"'{FOLDER_ID}' in parents and mimeType='application/vnd.openxmlformats-officedocument.presentationml.presentation' and trashed=false"
results = service.files().list(q=q, fields="files(id, name)").execute()
pliki_na_dysku = results.get('files', [])

# --- PANEL BOCZNY (STUDIO AI + DODATKI) ---
with st.sidebar:
    st.title("🚗 Studio AI")
    brand = st.selectbox("Marka", sorted(list(CAR_DATABASE.keys())))
    model = st.selectbox("Model", sorted(list(CAR_DATABASE[brand].keys())))
    body = st.selectbox("Nadwozie", CAR_DATABASE[brand][model])
    year = st.selectbox("Rocznik", [str(y) for y in range(2026, 1999, -1)])
    
    st.markdown("---")
    st.title("🎨 Folia i Kolor")
    f_brand = st.selectbox("Producent", list(FOIL_GROUPS.keys()))
    f_cat = st.selectbox("Wykończenie", list(FOIL_GROUPS[f_brand].keys()))
    f_color = st.selectbox("Kolor", FOIL_GROUPS[f_brand][f_cat])

    if st.button("🪄 GENERUJ WIZUALIZACJĘ AI"):
        prompt = f"Professional automotive studio photography of a {year} {brand} {model} ({body}) wrapped in {f_brand} {f_color}. High-end detailing garage, HEXAGONAL LED lights, cinematic lighting, 8k resolution, sharp focus. Floor: polished black epoxy with clear reflections."
        with st.spinner("AI renderuje Twoje auto..."):
            img_data = generate_ai_image(prompt)
            if img_data:
                st.session_state['ai_img'] = img_data
                st.success("Render gotowy!")
                
    st.markdown("---")
    st.header("📦 Dodatki do oferty")
    dodatki_dostepne = [f for f in pliki_na_dysku if f['name'].startswith(('4','5'))]
    wybrane_dodatki = []
    for d in sorted(dodatki_dostepne, key=lambda x: x['name']):
        if st.checkbox(d['name'], value=False):
            wybrane_dodatki.append(d)

# --- GŁÓWNY PANEL ---
st.title("🛡️ Generator Ofert ITS WRAP")
col1, col2 = st.columns(2)

with col1:
    klient = st.text_input("Imię i Nazwisko Klienta")
    nr_o = st.text_input("Numer oferty", value=f"IW/{datetime.now().strftime('%Y/%m/%d')}/01")
    
    # Cennik
    sheet = client.open_by_url("https://docs.google.com/spreadsheets/d/1iqS6geTNP3Bd_Fj_XdS-wCBrKtnGTMNQZYSso70KIkQ/edit?usp=drive_link").worksheet("Ppf")
    df = pd.DataFrame(sheet.get_all_values()[1:], columns=[c.strip() for c in sheet.get_all_values()[0]])
    pakiet = st.selectbox("Pakiet z cennika", df['Usługa'].tolist())
    rabat = st.number_input("Rabat (PLN)", value=0)

with col2:
    if 'ai_img' in st.session_state:
        st.image(st.session_state['ai_img'], caption=f"Wizualizacja: {brand} {model} w folii {f_color}", use_container_width=True)
    else:
        st.info("Skonfiguruj auto w panelu bocznym i wygeneruj zdjęcie, aby zobaczyć podgląd.")

# --- GENEROWANIE OFERTY ---
if st.button("🔥 GENERUJ PEŁNĄ OFERTĘ PDF"):
    if 'ai_img' not in st.session_state:
        st.error("Wizualizacja auta jest wymagana. Użyj przycisku w panelu bocznym!")
    else:
        with st.spinner("Składam profesjonalny PDF..."):
            writer = PdfWriter()
            row = df[df['Usługa'] == pakiet].iloc[0]
            cena_num = float(re.sub(r'[^\d,]', '', row['Kwota sprzedaży']).replace(',', '.'))

            replacements = {
                "{{KLIENT}}": klient, "{{MODEL_AUTA}}": f"{brand} {model}",
                "{{RODZAJ_FOLII}}": f_color, "{{USLUGA_NAZWA}}": pakiet,
                "{{NR_OFERTY}}": nr_o,
                "{{CENA_KATALOG}}": f"{cena_num:,.2f} zł".replace(',', ' ').replace('.', ','),
                "{{CENA_KONCOWA}}": f"{(cena_num - rabat):,.2f} zł".replace(',', ' ').replace('.', ',')
            }

            # Składanie klocków (1 -> 2 -> wybrane_dodatki -> 3 -> 6)
            okladka = next((f for f in pliki_na_dysku if f['name'].startswith('1')), None)
            produkt = next((f for f in pliki_na_dysku if f['name'].startswith('2')), None)
            zakres = next((f for f in pliki_na_dysku if f['name'].startswith('3')), None)
            koniec = next((f for f in pliki_na_dysku if f['name'].startswith('6')), None)

            seq = [okladka, produkt] + wybrane_dodatki + [zakres, koniec]
            seq = [f for f in seq if f]

            for f_info in seq:
                prs = Presentation(download_file(service, f_info['id']))
                for slide in prs.slides:
                    # Podmiana zdjęcia AI na okładce
                    if f_info['name'].startswith('1'):
                        for shape in list(slide.shapes):
                            if "{{FOTO_AUTA}}" in shape.name or (shape.has_text_frame and "{{FOTO_AUTA}}" in shape.text):
                                pic = slide.shapes.add_picture(io.BytesIO(st.session_state['ai_img']), shape.left, shape.top, shape.width, shape.height)
                                slide.shapes._spTree.remove(pic._element)
                                slide.shapes._spTree.insert(2, pic._element) # Wysyłamy zdjęcie na spód
                                shape._element.getparent().remove(shape._element)

                    # Podmiana tekstów i wymuszenie czcionki
                    for shape in slide.shapes:
                        if shape.has_text_frame:
                            for p in shape.text_frame.paragraphs:
                                for run in p.runs:
                                    for k, v in replacements.items():
                                        if k in run.text: 
                                            run.text = run.text.replace(k, str(v))
                                            run.font.name = 'URW DIN'

                tmp_p = f"tmp_{f_info['id']}.pptx"
                prs.save(tmp_p)
                pdf = pptx_to_pdf(tmp_p)
                if pdf: writer.append(pdf); os.remove(tmp_p); os.remove(pdf)

            final_io = io.BytesIO(); writer.write(final_io); final_io.seek(0)
            st.balloons()
            st.download_button("📥 POBIERZ OFERTĘ PDF", data=final_io, file_name=f"Oferta_{brand}_{model}.pdf")
