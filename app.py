import streamlit as st
import pandas as pd
from pptx import Presentation
import gspread
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload, MediaIoBaseUpload
import io, os, subprocess, re, shutil, requests, base64
from pypdf import PdfWriter
from datetime import datetime
from PIL import Image
import random

# --- NOWY LINK DO CENNIKA v3 ---
LINK_DO_ARKUSZA = "https://docs.google.com/spreadsheets/d/1USF81hOinAP_vvz1QZuoNyRCT1ezJcXTDDB6RjuYtrY/edit"
PARENT_FOLDER_ID = "12HRnKn9KrZy_C1BSgv24PGD-Gl8lTRmn"

# --- BAZA OPIEKUNÓW / HANDLOWCÓW ---
HANDLOWCY = {
    "Adam Trepka": {
        "stanowisko": "CEO It`s Wrap",
        "telefon": "+48 111 222 333",
        "email": "adam@itswrap.pl"
    },
    "Jan Kowalski": {
        "stanowisko": "Specjalista ds. Detailingu",
        "telefon": "+48 444 555 666",
        "email": "jan@itswrap.pl"
    }
}

# --- BAZA AUT I SEGMENTÓW ---
CAR_DATABASE = {
    "Audi": {"A3": ["Hatchback", "Sedan"], "A4": ["Sedan", "Kombi"], "A6": ["Sedan", "Kombi"], "Q3": ["SUV"], "Q5": ["SUV"], "Q8": ["SUV"], "e-tron GT": ["Sedan"], "RS6": ["Kombi"]},
    "BMW": {"Seria 3": ["Sedan", "Kombi"], "Seria 4": ["Coupe", "Gran Coupe"], "Seria 5": ["Sedan", "Kombi"], "Seria 7": ["Sedan"], "X3": ["SUV"], "X5": ["SUV"], "M3": ["Sedan", "Kombi"], "M4": ["Coupe"]},
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

SEGMENTY_DOMYSLNE = {
    "Audi": {"A3": "Segment C", "A4": "Segment D", "A6": "Segment D", "Q3": "Segment C", "Q5": "Segment D", "Q8": "Segment E", "e-tron GT": "Segment E", "RS6": "Segment D"},
    "BMW": {"Seria 3": "Segment D", "Seria 4": "Segment D", "Seria 5": "Segment D", "Seria 7": "Segment E", "X3": "Segment D", "X5": "Segment E", "M3": "Segment D", "M4": "Segment D"},
    "Porsche": {"911 (992)": "Segment D", "Taycan": "Segment E", "Macan": "Segment D", "Panamera": "Segment E", "Cayenne": "Segment E"},
    "Tesla": {"Model 3": "Segment D", "Model Y": "Segment D", "Model S": "Segment E", "Model X": "Segment J"}
}

# --- BAZA FOLII ---
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

def generate_ai_image(prompt, reference_image_bytes=None):
    api_key = st.secrets["GEMINI_API_KEY"]
    url = f"https://generativelanguage.googleapis.com/v1beta/models/imagen-4.0-ultra-generate-001:predict?key={api_key}"
    
    wzmocniony_prompt = f"{prompt} Highly detailed, exact same car shape and structure as the reference image, cinematic lighting in a high-end detailing garage, 8k resolution, modern photography."
    
    instance = {"prompt": wzmocniony_prompt}
    
    if reference_image_bytes:
        base64_image = base64.b64encode(reference_image_bytes).decode('utf-8')
        instance["referenceImages"] = [
            {
                "referenceImage": {"bytesBase64Encoded": base64_image}
            }
        ]

    payload = {"instances": [instance], "parameters": {"sampleCount": 1}}
    
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
        else:
            st.error(f"Odrzucenie zlecenia przez API Obrazów: {response.text}")
    except Exception as e:
        st.error(f"Wystąpił błąd komunikacji z modelem: {e}")
        
    img_fallback = Image.new('RGB', (2100, 1870), color=(40, 40, 45))
    out_fallback = io.BytesIO()
    img_fallback.save(out_fallback, format='PNG')
    return out_fallback.getvalue()

def generate_ai_intro_text(klient, brand, model, pakiet, folia, handlowiec_imie, handlowiec_stanowisko):
    imie_surowe = klient.split()[0] if klient.strip() != "" else ""
    czysta_folia = folia.split('(')[0].strip()
    wolacz = "Szanowny Kliencie"
    if imie_surowe:
        imie = imie_surowe.title()
        imie_lower = imie.lower()
        if imie_lower.endswith('a'): wolacz = f"Pani {imie}"
        else:
            wyjatki = {"piotr": "Piotrze", "paweł": "Pawle", "kacper": "Kacprze", "marek": "Marku", "michał": "Michale", "donald": "Donaldzie", "konrad": "Konradzie", "dawid": "Dawidzie"}
            if imie_lower in wyjatki: wolacz = f"Panie {wyjatki[imie_lower]}"
            elif imie_lower.endswith('d'): wolacz = f"Panie {imie}zie"
            elif imie_lower.endswith(('k', 'g', 'ch', 'j', 'sz', 'cz', 'rz', 'l', 'c')):
                if imie_lower.endswith('ek'): wolacz = f"Panie {imie[:-2]}ku"
                else: wolacz = f"Panie {imie}u"
            elif imie_lower.endswith(('n', 'm', 'b', 'w', 'f', 's', 'z', 't', 'p')): wolacz = f"Panie {imie}ie"
            elif imie_lower.endswith('r'): wolacz = f"Panie {imie}ze"
            else: wolacz = f"Panie {imie}"
            
    marka = brand if brand != "Inna marka..." else ""
    szablony = [
        f"{wolacz},\n\nDziękuję za wybór naszej firmy. Komponując ofertę dla Twojego {marka}, dobraliśmy bezkompromisowe rozwiązanie, jakim jest folia {czysta_folia}. Dzięki temu mogę zagwarantować Tobie najwyższą jakość ochrony samochodu na długie lata. Serdecznie zapraszam do zapoznania się ze szczegółami przygotowanej wyceny.",
        f"{wolacz},\n\nW ITS WRAP nie uznajemy kompromisów. Właśnie dlatego, tworząc tę wycenę dla Twojego {marka}, zdecydowałem się na zastosowanie niezawodnej folii {czysta_folia}. To inwestycja, która zapewni Ci spokój ducha i perfekcyjną prezencję auta na drodze. Zachęcam do zapoznania się ze szczegółami."
    ]
    wybrany_tekst = random.choice(szablony)
    return f"{wybrany_tekst}\n\nZ motoryzacyjnym pozdrowieniem,\n{handlowiec_imie}\n{handlowiec_stanowisko}"

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

# --- OBSŁUGA GOOGLE DRIVE DLA OFERT ---
def pobierz_lub_stworz_folder_oferty(service, parent_id):
    query = f"'{parent_id}' in parents and name='Oferty' and mimeType='application/vnd.google-apps.folder' and trashed=false"
    results = service.files().list(q=query, fields="files(id, name)").execute()
    items = results.get('files', [])
    
    if not items:
        file_metadata = {
            'name': 'Oferty',
            'mimeType': 'application/vnd.google-apps.folder',
            'parents': [parent_id]
        }
        folder = service.files().create(body=file_metadata, fields='id').execute()
        return folder.get('id')
    return items[0].get('id')

def wgraj_pdf_na_dysk(service, folder_id, file_name, file_bytes):
    try:
        file_metadata = {
            'name': file_name,
            'parents': [folder_id]
        }
        media = MediaIoBaseUpload(io.BytesIO(file_bytes), mimetype='application/pdf', resumable=True)
        file = service.files().create(body=file_metadata, media_body=media, fields='id, webViewLink').execute()
        
        # Uprawnienia do odczytu dla wszystkich posiadających link
        service.permissions().create(
            fileId=file.get('id'),
            body={'type': 'anyone', 'role': 'reader'}
        ).execute()
        
        return file.get('webViewLink')
    except Exception as e:
        st.error(f"Błąd podczas zapisu pliku PDF na Google Drive: {e}")
        return None

# --- ZAPIS DO REJESTRU ---
def zapisz_do_rejestru(nr_oferty, handlowiec, klient, auto, usluga, folia, cena, pdf_link):
    try:
        sheet_rejestr = client.open_by_url(LINK_DO_ARKUSZA).worksheet("Rejestr")
        dzisiaj = datetime.now().strftime("%Y-%m-%d")
        nowy_wiersz = [dzisiaj, nr_oferty, handlowiec, klient, auto, usluga, folia, f"{cena} zł", "Nowa", pdf_link]
        sheet_rejestr.append_row(nowy_wiersz)
        return True
    except Exception as e:
        st.error(f"Nie udało się zapisać do bazy (Rejestr): {e}")
        return False

def replace_text_in_shape(shape, replacements):
    """Rekurencyjnie wyszukuje i podmienia tagi tekstowe, wchodząc również w tabele i grupy obiektów."""
    if hasattr(shape, "has_text_frame") and shape.has_text_frame:
        for p in shape.text_frame.paragraphs:
            p_text = "".join(run.text for run in p.runs)
            original_text = p_text
            for k, v in replacements.items():
                if k in p_text:
                    p_text = p_text.replace(k, str(v))
            if p_text != original_text and len(p.runs) > 0:
                p.runs[0].text = p_text
                for run in p.runs[1:]:
                    run.text = ""
                    
    if hasattr(shape, "has_table") and shape.has_table:
        for row in shape.table.rows:
            for cell in row.cells:
                replace_text_in_shape(cell, replacements)
                
    if shape.shape_type == 6: # msoGroup
        for subshape in shape.shapes:
            replace_text_in_shape(subshape, replacements)

# --- APLIKACJA ---
st.set_page_config(page_title="Zap & Studio Ultimate", layout="wide")
install_fonts()

creds = Credentials.from_service_account_info(st.secrets["gcp_service_account"], scopes=["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"])
service = build('drive', 'v3', credentials=creds)
client = gspread.authorize(creds)

results = service.files().list(q=f"'{PARENT_FOLDER_ID}' in parents and mimeType='application/vnd.openxmlformats-officedocument.presentationml.presentation' and trashed=false", fields="files(id, name)").execute()
pliki_na_dysku = results.get('files', [])

# POBIERANIE CENNIKA v3
try:
    sheet_cennik = client.open_by_url(LINK_DO_ARKUSZA).worksheet("Cennik usług")
    naglowki = [c.replace('\n', ' ').replace('\r', '').strip() for c in sheet_cennik.get_all_values()[0]]
    df_cennik = pd.DataFrame(sheet_cennik.get_all_values()[1:], columns=naglowki)
except Exception as e:
    st.error(f"Błąd ładowania Cennika v3. Upewnij się, że link i nazwa zakładki 'Cennik usług' są poprawne. Błąd: {e}")
    st.stop()

# --- PANEL BOCZNY ---
with st.sidebar:
    st.title("👤 Opiekun Klienta")
    wybrany_handlowiec = st.selectbox("Kto przygotowuje ofertę?", list(HANDLOWCY.keys()))
    
    st.markdown("---")
    st.title("🚗 Studio AI")
    brand = st.selectbox("Marka", list(CAR_DATABASE.keys()))
    
    if brand == "Inna marka...":
        custom_brand = st.text_input("Wpisz markę")
        custom_model = st.text_input("Wpisz model")
        final_brand, final_model, body, segment_domyslny = custom_brand, custom_model, "", "Segment D"
    else:
        final_brand = brand
        final_model = st.selectbox("Model", list(CAR_DATABASE[brand].keys()))
        body = st.selectbox("Nadwozie", CAR_DATABASE[brand][final_model])
        segment_domyslny = SEGMENTY_DOMYSLNE.get(brand, {}).get(final_model, "Segment D")
        
    segment_final = st.selectbox("Wybierz Segment (do wyceny)", ["Segment A", "Segment B", "Segment C", "Segment D", "Segment E", "Segment J", "Wavecamper"], 
                                 index=["Segment A", "Segment B", "Segment C", "Segment D", "Segment E", "Segment J", "Wavecamper"].index(segment_domyslny))
        
    year = st.selectbox("Rocznik", [str(y) for y in range(2026, 1999, -1)])
    gen_code = st.text_input("Kod karoserii (Opcjonalnie)", help="Np. G70, 992")
    
    st.markdown("---")
    st.title("🎨 Folia i Kolor")
    f_brand = st.selectbox("Producent", list(FOIL_GROUPS.keys()))
    f_cat = st.selectbox("Wykończenie", list(FOIL_GROUPS[f_brand].keys()))
    f_color = st.selectbox("Kolor", FOIL_GROUPS[f_brand][f_cat])

    paint_color = ""
    if "Bezbarwne" in f_cat:
        paint_color = st.text_input("🚘 Podaj obecny kolor lakieru auta", value="Czarny metallic")

    st.markdown("---")
    st.title("📸 Wizualizacja")
    
    uploaded_files = st.file_uploader("Opcjonalnie: Wgraj zdjęcia poglądowe (np. nowej karoserii ze strony producenta)", type=['png', 'jpg', 'jpeg'], accept_multiple_files=True)

    if st.button("🪄 GENERUJ WIZUALIZACJĘ AI"):
        extra = f" {gen_code}," if gen_code else ""
        if "Bezbarwne" in f_cat:
            finish = "matte/satin finish" if "Stealth" in f_color else "high gloss finish"
            prompt = f"Automotive studio photography of the newest {year} {final_brand} {final_model} ({body}). {extra} Car paint color: {paint_color}. The car is completely wrapped in clear PPF giving it a {finish}."
        else:
            prompt = f"Automotive studio photography of the newest {year} {final_brand} {final_model} ({body}). {extra} Wrapped in {f_brand} {f_color}."
            
        ref_image_bytes = None
        if uploaded_files:
            ref_image_bytes = uploaded_files[0].read()
            st.info("Przetwarzam Twoje zdjęcie referencyjne...")

        with st.spinner("AI renderuje Twoje auto..."):
            img_data = generate_ai_image(prompt, ref_image_bytes)
            if img_data:
                st.session_state['ai_img'] = img_data
                
    st.markdown("---")
    st.header("📦 Dodatki do oferty")
    dodatki_dostepne = [f for f in pliki_na_dysku if f['name'].startswith(('4','5'))]
    wybrane_dodatki = [d for d in sorted(dodatki_dostepne, key=lambda x: x['name']) if st.checkbox(d['name'], value=False)]

# ZAKŁADKI W GŁÓWNYM PANELU
tab_kreator, tab_rejestr = st.tabs(["⚙️ Kreator Ofert", "📋 Ewidencja (Rejestr)"])

with tab_kreator:
    st.title("🛡️ Generator Ofert ITS WRAP")
    col1, col2 = st.columns(2)

    with col1:
        klient = st.text_input("Imię i Nazwisko Klienta")
        nr_o = st.text_input("Numer oferty", value=f"IW/{datetime.now().strftime('%Y/%m/%d')}/01")
        
        kategorie = [k for k in df_cennik['Kategoria'].unique() if str(k).strip() != ""]
        kategoria = st.selectbox("Kategoria", kategorie)
        uslugi_kat = [u for u in df_cennik[df_cennik['Kategoria'] == kategoria]['Usługa'].unique() if str(u).strip() != ""]
        pakiet = st.selectbox("Usługa", uslugi_kat)
        
        try:
            wiersz_ceny = df_cennik[(df_cennik['Kategoria'] == kategoria) & (df_cennik['Usługa'] == pakiet) & (df_cennik['Segment'] == segment_final)]
            if not wiersz_ceny.empty:
                cena_str = str(wiersz_ceny['Cena sprzedaży netto PLN'].values[0])
                cena_str = cena_str.replace(' ', '').replace('\xa0', '')
                if ',' in cena_str: cena_str = cena_str.replace('.', '').replace(',', '.')
                cena_domyslna = float(re.sub(r'[^\d.]', '', cena_str))
            else:
                cena_domyslna = 0.0
        except:
            cena_domyslna = 0.0

        st.markdown("---")
        st.write("💰 **Kalkulacja cenowa**")
        
        cena_manual = st.number_input("Cena bazowa NETTO (PLN) - możesz edytować", value=cena_domyslna, step=100.0)
        rabat = st.number_input("Rabat dla klienta (PLN)", value=0.0, step=100.0)
        cena_koncowa = cena_manual - rabat
        
        st.info(f"**Cena do zapłaty netto (na ofercie): {cena_koncowa:,.2f} zł**".replace(',', 'X').replace('.', ',').replace('X', ' '))

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
            with st.spinner("Składam profesjonalny PDF i wgrywam na Dysk Google..."):
                writer = PdfWriter()
                final_foil_text = f"{f_color} (na lakier: {paint_color})" if "Bezbarwne" in f_cat else f_color
                
                dane_handlowca = HANDLOWCY[wybrany_handlowiec]
                wygenerowany_wstep = generate_ai_intro_text(klient, final_brand, final_model, pakiet, final_foil_text, wybrany_handlowiec, dane_handlowca["stanowisko"])

                cena_koncowa_str = f"{cena_koncowa:,.2f} zł".replace(',', 'X').replace('.', ',').replace('X', ' ')
                cena_katalogowa_str = f"{cena_manual:,.2f} zł".replace(',', 'X').replace('.', ',').replace('X', ' ')

                replacements = {
                    "{{KLIENT}}": klient, 
                    "{{MODEL_AUTA}}": f"{final_brand} {final_model}",
                    "{{RODZAJ_FOLII}}": final_foil_text, 
                    "{{USLUGA_NAZWA}}": pakiet,
                    "{{NR_OFERTY}}": nr_o,
                    "{{CENA_KATALOG}}": cena_katalogowa_str,
                    "{{CENA_KONCOWA}}": cena_koncowa_str,
                    "{{WSTEP_AI}}": wygenerowany_wstep,
                    "{{HANDLOWIEC_IMIE}}": wybrany_handlowiec,
                    "{{HANDLOWIEC_TEL}}": dane_handlowca["telefon"],
                    "{{HANDLOWIEC_EMAIL}}": dane_handlowca["email"]
                }

                okladka = next((f for f in pliki_na_dysku if f['name'].startswith('1_')), None)
                wstep_slide = next((f for f in pliki_na_dysku if f['name'].lower().startswith('1b_')), None)
                
                produkt = None
                if "reklam" in pakiet.lower():
                    produkt = next((f for f in pliki_na_dysku if f['name'].startswith('2') and 'reklama' in f['name'].lower()), None)
                elif f_brand == "3M 2080 Series":
                    produkt = next((f for f in pliki_na_dysku if f['name'].startswith('2') and '3m' in f['name'].lower() and 'kolor' in f['name'].lower()), None)
                elif "Ultimate" in f_color: 
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
                            replace_text_in_shape(shape, replacements)

                    tmp_p = f"tmp_{f_info['id']}.pptx"
                    prs.save(tmp_p)
                    pdf = pptx_to_pdf(tmp_p)
                    if pdf: writer.append(pdf); os.remove(tmp_p); os.remove(pdf)

                final_io = io.BytesIO(); writer.write(final_io); final_io.seek(0)
                nazwa_pliku_wyjsciowego = f"Oferta_{final_brand}_{final_model}_{datetime.now().strftime('%H%M%S')}.pdf"
                
                # Odszukanie/Utworzenie folderu Oferty i wgranie pliku
                folder_oferty_id = pobierz_lub_stworz_folder_oferty(service, PARENT_FOLDER_ID)
                utworzony_link = wgraj_pdf_na_dysk(service, folder_oferty_id, nazwa_pliku_wyjsciowego, final_io.getvalue())
                
                link_do_zapisu = utworzony_link if utworzony_link else "Błąd uploadu"

                if zapisz_do_rejestru(nr_o, wybrany_handlowiec, klient, f"{final_brand} {final_model}", pakiet, final_foil_text, cena_koncowa, link_do_zapisu):
                    st.success(f"✅ Oferta zapisana w systemie CRM! Plik został zachowany w folderze 'Oferty'.")
                    
                st.balloons()
                st.download_button("📥 POBIERZ OFERTĘ PDF LOKALNIE", data=final_io, file_name=nazwa_pliku_wyjsciowego)

with tab_rejestr:
    st.header("📋 Ostatnio zapisane oferty")
    try:
        sheet_rejestr_view = client.open_by_url(LINK_DO_ARKUSZA).worksheet("Rejestr")
        dane_rejestru = sheet_rejestr_view.get_all_records()
        if dane_rejestru:
            df_rejestr = pd.DataFrame(dane_rejestru)
            
            # Formatyzacja kolumny, zakładając że kolumna nazywa się podobnie do tego co zostało wysłane.
            # Ważne: w arkuszu "Rejestr" trzeba dopisać w rzędzie z nagłówkami nową kolumnę, np. "Link PDF".
            nazwa_kolumny_link = df_rejestr.columns[-1] # domyślnie ostatnia dodana kolumna to link
            
            st.data_editor(
                df_rejestr,
                column_config={
                    nazwa_kolumny_link: st.column_config.LinkColumn(
                        "Plik PDF", help="Kliknij aby pobrać dokument ofertowy z dysku", display_text="Otwórz dokument"
                    )
                },
                hide_index=True,
                use_container_width=True
            )
            
            csv = df_rejestr.to_csv(index=False).encode('utf-8')
            st.download_button(
                label="⬇️ Pobierz ewidencję jako CSV",
                data=csv,
                file_name=f"Rejestr_Ofert_ITSWRAP_{datetime.now().strftime('%Y%m%d')}.csv",
                mime='text/csv',
            )
        else:
            st.info("Rejestr jest pusty lub arkusz nie zawiera jeszcze wpisów.")
    except Exception as e:
        st.warning(f"Brak możliwości wczytania rejestru. Pamiętaj, by w nagłówkach bazy (Arkusz 'Rejestr', 1 wiersz) dodać kolumnę na link do PDF. Błąd: {e}")
