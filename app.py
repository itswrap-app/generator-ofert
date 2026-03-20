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
from PIL import Image
import random

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
    "Volvo": {"XC40": ["SUV"], "XC60": ["SUV"], "XC90": ["SUV"], "V60": ["Kombi"]},
    "Inna marka...": {"Wpisz ręcznie": ["Inne"]}
}

# --- PEŁNA BAZA FOLII ---
FOIL_GROUPS = {
    "XPEL (Folie Ochronne PPF)": {
        "Bezbarwne (Twój obecny kolor)": ["XPEL Ultimate Plus (Wysoki Połysk)", "XPEL Stealth (Mat/Satyna)"],
        "XPEL Color (Zmiana Koloru PPF)": ["Black (Połysk)", "White (Połysk)", "Red (Połysk)", "Nardo Grey (Połysk)", "Miami Blue (Połysk)"]
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

# 1. FUNKCJA GENEROWANIA ZDJĘĆ AI
def generate_ai_image(prompt):
    api_key = st.secrets["GEMINI_API_KEY"]
    url = f"https://generativelanguage.googleapis.com/v1beta/models/imagen-4.0-generate-001:predict?key={api_key}"
    payload = {"instances": [{"prompt": prompt}], "parameters": {"sampleCount": 1}}
    try:
        response = requests.post(url, json=payload, timeout=60)
        if response.status_code == 200:
            img_data = base64.b64decode(response.json()['predictions'][0]['bytesBase64Encoded'])
            
            img = Image.open(io.BytesIO(img_data))
            w, h = img.size
            target_ratio = 21.0 / 18.7
            
            if w / h > target_ratio: 
                new_w = int(h * target_ratio)
                left = (w - new_w) / 2
                img_cropped = img.crop((left, 0, left + new_w, h))
            else: 
                new_h = int(w / target_ratio)
                top = (h - new_h) / 2
                img_cropped = img.crop((0, top, w, top + new_h))
                
            out_bytes = io.BytesIO()
            img_cropped.save(out_bytes, format='PNG')
            return out_bytes.getvalue()
    except Exception as e:
        pass
        
    img_fallback = Image.new('RGB', (2100, 1870), color=(40, 40, 45))
    out_fallback = io.BytesIO()
    img_fallback.save(out_fallback, format='PNG')
    st.info("Brak wsparcia Google Imagen w UE. Użyto idealnie dociętego, eleganckiego tła zastępczego.")
    return out_fallback.getvalue()

# 2. FUNKCJA GENEROWANIA UNIKALNEGO TEKSTU AI
def generate_ai_intro_text(klient, brand, model, pakiet, folia):
    api_key = st.secrets["GEMINI_API_KEY"]
    
    imie = klient.split()[0] if klient.strip() != "" else ""
    czysta_folia = folia.split('(')[0].strip()
    
    prompt = f"""
    Jako Adam Trepka, CEO ITS WRAP, napisz krótki wstęp do oferty oklejania auta.
    
    WYMÓG KREATYWNOŚCI: Ten tekst MUSI być w 100% unikalny i różnić się od poprzednich. Używaj innej struktury zdań, innej formy powitania i innych synonimów, pokazując zaangażowanie i prestiż.
    
    DANE:
    - Imię: {imie}
    - Auto: {brand} {model}
    - Folia: {czysta_folia}
    
    BEZWZGLĘDNE ZASADY:
    1. ZWROT: Odmień imię w wołaczu! (np. Panie Tomaszu, Panie Adamie, Panie Piotrze, Panie Dominiku).
    2. AUTO: Użyj "dla Twojego {brand}". ZABRONIONE jest pisanie "dla pojazdu marki" lub "dla samochodu marki".
    3. STYL: Pisz w 1. osobie liczby pojedynczej ("dobrałem", "zdecydowałem"). Pokaż pasję.
    4. PODPIS: Zakończ tekst DOKŁADNIE w ten sposób:
    
    Z motoryzacyjnym pozdrowieniem,
    Adam Trepka
    CEO It`s Wrap
    
    Długość całego tekstu: max 4 zdania + podpis. Bez pogrubień (*).
    """
    
    # "temperature": 0.9 gwarantuje dużą losowość i unikalność wypowiedzi
    payload = {
        "contents": [{"parts": [{"text": prompt}]}],
        "generationConfig": {"temperature": 0.9}
    }
    
    # Lista modeli do przetestowania (omijamy błędy 404)
    modele_do_sprawdzenia = [
        "gemini-1.5-flash-latest",
        "gemini-1.0-pro-latest"
    ]
    
    for model_name in modele_do_sprawdzenia:
        url = f"https://generativelanguage.googleapis.com/v1beta/models/{model_name}:generateContent?key={api_key}"
        try:
            response = requests.post(url, json=payload, timeout=20)
            if response.status_code == 200:
                return response.json()['candidates'][0]['content']['parts'][0]['text'].strip()
        except Exception:
            continue
    
    # --- PROGRAMISTYCZNY, INTELIGENTNY FALLBACK ---
    wolacz = "Szanowny Kliencie"
    if imie:
        imie_lower = imie.lower()
        if imie_lower.endswith('a'): wolacz = f"Pani {imie}"
        elif imie_lower.endswith(('ik', 'yk')): wolacz = f"Panie {imie}u"
        elif imie_lower.endswith('id'): wolacz = f"Panie {imie}zie"
        elif imie_lower in ["tomasz", "łukasz", "mateusz", "janusz", "mariusz"]: wolacz = f"Panie {imie}u"
        elif imie_lower in ["adam", "michal", "michał", "kamil", "marcin", "adrian", "krystian"]: wolacz = f"Panie {imie}ie"
        elif imie_lower in ["piotr", "kacper", "wiktor"]: wolacz = f"Panie {imie}ze"
        elif imie_lower.endswith('ek'): wolacz = f"Panie {imie[:-2]}ku"
        elif imie_lower.endswith(('i', 'y')): wolacz = f"Panie {imie}"
        else: wolacz = f"Panie {imie}"
            
    marka = brand
    if brand == "Toyota": marka = "Toyoty"
    elif brand == "Skoda": marka = "Skody"
    elif brand == "Kia": marka = "Kii"
    elif brand == "Tesla": marka = "Tesli"
    elif brand == "Porsche": marka = "Porsche"
    elif brand == "Honda": marka = "Hondy"
    elif brand == "Mazda": marka = "Mazdy"

    return f"{wolacz},\n\nDziękuję za wybór naszej firmy. Komponując ofertę dla Twojego {marka}, dobraliśmy bezkompromisowe rozwiązanie, jakim jest folia {czysta_folia}. Dzięki temu mogę zagwarantować Tobie najwyższą jakość ochrony samochodu na długie lata. Serdecznie zapraszam do zapoznania się ze szczegółami przygotowanej wyceny.\n\nZ motoryzacyjnym pozdrowieniem,\nAdam Trepka\nCEO It`s Wrap"

def download_file(service, file_id):
    request = service.files().get_media(fileId=file_id)
    fh = io.BytesIO(); downloader = MediaIoBaseDownload(fh, request)
    done = False
    while not done: _, done = downloader.next_chunk()
    fh.seek(0); return fh

def pptx_to_pdf(input_path):
    try:
        subprocess.run(['libreoffice', '--headless', '--convert-to', 'pdf', '--outdir', os.getcwd(), input_path], check=True, capture_output=True)
        return os.path.basename(input_path).replace('.pptx', '.pdf')
    except: return None

# --- APLIKACJA ---
st.set_page_config(page_title="Zap & Studio Ultimate", layout="wide")
install_fonts()

creds = Credentials.from_service_account_info(st.secrets["gcp_service_account"], scopes=["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"])
service = build('drive', 'v3', credentials=creds)
client = gspread.authorize(creds)

results = service.files().list(q="'12HRnKn9KrZy_C1BSgv24PGD-Gl8lTRmn' in parents and mimeType='application/vnd.openxmlformats-officedocument.presentationml.presentation' and trashed=false", fields="files(id, name)").execute()
pliki_na_dysku = results.get('files', [])

sheet = client.open_by_url("https://docs.google.com/spreadsheets/d/1iqS6geTNP3Bd_Fj_XdS-wCBrKtnGTMNQZYSso70KIkQ/edit?usp=drive_link").worksheet("Ppf")
df = pd.DataFrame(sheet.get_all_values()[1:], columns=[c.strip() for c in sheet.get_all_values()[0]])

# --- PANEL BOCZNY ---
with st.sidebar:
    st.title("🚗 Studio AI")
    brand = st.selectbox("Marka", list(CAR_DATABASE.keys()))
    
    if brand == "Inna marka...":
        custom_brand = st.text_input("Wpisz markę")
        custom_model = st.text_input("Wpisz model")
        final_brand, final_model, body = custom_brand, custom_model, ""
    else:
        final_brand = brand
        final_model = st.selectbox("Model", list(CAR_DATABASE[brand].keys()))
        body = st.selectbox("Nadwozie", CAR_DATABASE[brand][final_model])
        
    year = st.selectbox("Rocznik", [str(y) for y in range(2026, 1999, -1)])
    
    st.markdown("---")
    st.title("🎨 Folia i Kolor")
    f_brand = st.selectbox("Producent", list(FOIL_GROUPS.keys()))
    f_cat = st.selectbox("Wykończenie", list(FOIL_GROUPS[f_brand].keys()))
    f_color = st.selectbox("Kolor", FOIL_GROUPS[f_brand][f_cat])

    paint_color = ""
    if "Bezbarwne" in f_cat:
        paint_color = st.text_input("🚘 Podaj obecny kolor lakieru auta", value="Czarny metallic")

    if st.button("🪄 GENERUJ WIZUALIZACJĘ AI"):
        if "Bezbarwne" in f_cat:
            finish = "matte/satin finish" if "Stealth" in f_color else "high gloss finish"
            prompt = f"Professional automotive studio photography of a {year} {final_brand} {final_model} ({body}). Car paint color: {paint_color}. The car is completely wrapped in clear PPF giving it a {finish}. High-end detailing garage, HEXAGONAL LED lights, cinematic lighting, 8k resolution, sharp focus."
        else:
            prompt = f"Professional automotive studio photography of a {year} {final_brand} {final_model} ({body}) wrapped in {f_brand} {f_color}. High-end detailing garage, HEXAGONAL LED lights, cinematic lighting, 8k resolution, sharp focus."
            
        with st.spinner("AI renderuje Twoje auto..."):
            img_data = generate_ai_image(prompt)
            if img_data:
                st.session_state['ai_img'] = img_data
                
    st.markdown("---")
    st.header("📦 Dodatki do oferty")
    dodatki_dostepne = [f for f in pliki_na_dysku if f['name'].startswith(('4','5'))]
    wybrane_dodatki = [d for d in sorted(dodatki_dostepne, key=lambda x: x['name']) if st.checkbox(d['name'], value=False)]

# --- GŁÓWNY PANEL ---
st.title("🛡️ Generator Ofert ITS WRAP")
col1, col2 = st.columns(2)

with col1:
    klient = st.text_input("Imię i Nazwisko Klienta")
    nr_o = st.text_input("Numer oferty", value=f"IW/{datetime.now().strftime('%Y/%m/%d')}/01")
    
    pakiet = st.selectbox("Pakiet z cennika", df['Usługa'].tolist())
    
    wiersz = df[df['Usługa'] == pakiet].iloc[0]
    try:
        cena_domyslna = float(re.sub(r'[^\d,]', '', wiersz['Kwota sprzedaży']).replace(',', '.'))
    except:
        cena_domyslna = 0.0

    st.markdown("---")
    st.write("💰 **Kalkulacja cenowa**")
    
    cena_manual = st.number_input("Cena bazowa (PLN) - możesz edytować", value=cena_domyslna, step=100.0)
    rabat = st.number_input("Rabat dla klienta (PLN)", value=0.0, step=100.0)
    cena_koncowa = cena_manual - rabat
    
    st.info(f"**Cena do zapłaty (na ofercie): {cena_koncowa:,.2f} zł**".replace(',', ' ').replace('.', ','))

with col2:
    if 'ai_img' in st.session_state:
        st.image(st.session_state['ai_img'], use_container_width=True)
    else:
        st.info("Skonfiguruj auto w panelu bocznym i wygeneruj zdjęcie, aby zobaczyć podgląd.")

# --- GENEROWANIE OFERTY ---
if st.button("🔥 GENERUJ PEŁNĄ OFERTĘ PDF"):
    if 'ai_img' not in st.session_state:
        st.error("Wizualizacja auta jest wymagana. Użyj przycisku w panelu bocznym!")
    else:
        with st.spinner("AI analizuje ofertę i pisze w 100% unikalny list powitalny..."):
            final_foil_text = f"{f_color} (na lakier: {paint_color})" if "Bezbarwne" in f_cat else f_color
            wygenerowany_wstep = generate_ai_intro_text(klient, final_brand, final_model, pakiet, final_foil_text)
            
        with st.spinner("Składam profesjonalny PDF..."):
            writer = PdfWriter()

            replacements = {
                "{{KLIENT}}": klient, 
                "{{MODEL_AUTA}}": f"{final_brand} {final_model}",
                "{{RODZAJ_FOLII}}": final_foil_text, 
                "{{USLUGA_NAZWA}}": pakiet,
                "{{NR_OFERTY}}": nr_o,
                "{{CENA_KATALOG}}": f"{cena_manual:,.2f} zł".replace(',', ' ').replace('.', ','),
                "{{CENA_KONCOWA}}": f"{cena_koncowa:,.2f} zł".replace(',', ' ').replace('.', ','),
                "{{WSTEP_AI}}": wygenerowany_wstep
            }

            okladka = next((f for f in pliki_na_dysku if f['name'].startswith('1_')), None)
            wstep_slide = next((f for f in pliki_na_dysku if f['name'].lower().startswith('1b_')), None)
            
            produkt = None
            if "Ultimate" in f_color:
                produkt = next((f for f in pliki_na_dysku if f['name'].startswith('2') and 'ultimate' in f['name'].lower()), None)
            elif "Stealth" in f_color:
                produkt = next((f for f in pliki_na_dysku if f['name'].startswith('2') and 'stealth' in f['name'].lower()), None)
            elif "Color" in f_cat:
                produkt = next((f for f in pliki_na_dysku if f['name'].startswith('2') and 'color' in f['name'].lower()), None)
            
            if not produkt:
                produkt = next((f for f in pliki_na_dysku if f['name'].startswith('2')), None)

            if rabat > 0:
                zakres = next((f for f in pliki_na_dysku if f['name'].startswith('3') and 'bezrabatu' not in f['name'].lower()), None)
            else:
                zakres = next((f for f in pliki_na_dysku if f['name'].startswith('3') and 'bezrabatu' in f['name'].lower()), None)
            
            if not zakres:
                zakres = next((f for f in pliki_na_dysku if f['name'].startswith('3')), None)

            koniec = next((f for f in pliki_na_dysku if f['name'].startswith('6')), None)

            seq = [okladka, wstep_slide, produkt, zakres] + wybrane_dodatki + [koniec]
            seq = [f for f in seq if f]

            for f_info in seq:
                prs = Presentation(download_file(service, f_info['id']))
                for slide in prs.slides:
                    if f_info['name'].startswith('1_'):
                        for shape in list(slide.shapes):
                            if "{{FOTO_AUTA}}" in shape.name or (shape.has_text_frame and "{{FOTO_AUTA}}" in shape.text):
                                pic = slide.shapes.add_picture(io.BytesIO(st.session_state['ai_img']), shape.left, shape.top, shape.width, shape.height)
                                slide.shapes._spTree.remove(pic._element)
                                slide.shapes._spTree.insert(2, pic._element)
                                shape._element.getparent().remove(shape._element)

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
            st.download_button("📥 POBIERZ OFERTĘ PDF", data=final_io, file_name=f"Oferta_{final_brand}_{final_model}.pdf")
