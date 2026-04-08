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

def generate_ai_image(prompt):
    api_key = st.secrets["GEMINI_API_KEY"]
    # Używamy potężnego modelu Ultra, który ma nowszą wiedzę o pojazdach
    url = f"https://generativelanguage.googleapis.com/v1beta/models/imagen-4.0-ultra-generate-001:predict?key={api_key}"
    
    # Dodajemy tylko uniwersalny dopisek o dokładności detali, rocznik idzie prosto z paska bocznego
    wzmocniony_prompt = f"{prompt} Highly detailed, precise factory body styling, sharp focus, 8k resolution."
    
    payload = {"instances": [{"prompt": wzmocniony_prompt}], "parameters": {"sampleCount": 1}}
    
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
            st.warning(f"Błąd generowania obrazu: {response.text}")
    except Exception as e:
        pass
        
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
            wyjatki = {"piotr": "Piotrze", "paweł": "Pawle", "kacper": "Kacprze", "marek": "Marku", "michał": "Michale", "michal": "Michale", "rafał": "Rafale", "kamil": "Kamilu", "karol": "Karolu", "jerzy": "Jerzy", "igor": "Igorze", "donald": "Donaldzie", "konrad": "Konradzie", "dawid": "Dawidzie", "ryszard": "Ryszardzie", "krzysztof": "Krzysztofie", "maciej": "Macieju", "mikołaj": "Mikołaju", "bartłomiej": "Bartłomieju"}
            if imie_lower in wyjatki: wolacz = f"Panie {wyjatki[imie_lower]}"
            else:
                if imie_lower.endswith('d'): wolacz = f"Panie {imie}zie"
                elif imie_lower.endswith(('k', 'g', 'ch', 'j', 'sz', 'cz', 'rz', 'l', 'c')):
                    if imie_lower.endswith('ek'): wolacz = f"Panie {imie[:-2]}ku"
                    else: wolacz = f"Panie {imie}u"
                elif imie_lower.endswith(('n', 'm', 'b', 'w', 'f', 's', 'z', 't', 'p')): wolacz = f"Panie {imie}ie"
                elif imie_lower.endswith('r'): wolacz = f"Panie {imie}ze"
                else: wolacz = f"Panie {imie}"
            
    marka = brand
    if brand == "Toyota": marka = "Toyoty"
    elif brand == "Skoda": marka = "Skody"
    elif brand == "Kia": marka = "Kii"
    elif brand == "Tesla": marka = "Tesli"
    elif brand == "Porsche": marka = "Porsche"
    elif brand == "Honda": marka = "Hondy"
    elif brand == "Mazda": marka = "Mazdy"

    szablony = [
        f"{wolacz},\n\nDziękuję za wybór naszej firmy. Komponując ofertę dla Twojego {marka}, dobraliśmy bezkompromisowe rozwiązanie, jakim jest folia {czysta_folia}. Dzięki temu mogę zagwarantować Tobie najwyższą jakość ochrony samochodu na długie lata. Serdecznie zapraszam do zapoznania się ze szczegółami przygotowanej wyceny.",
        f"{wolacz},\n\nMotoryzacja to nasza największa pasja, dlatego do ochrony Twojego {marka} podszedłem z najwyższą starannością. Wybrana przez nas folia {czysta_folia} to absolutna czołówka w świecie auto detailingu. Gwarantuje ona, że Twój samochód zachowa nieskazitelny wygląd przez wiele lat. Zapraszam do lektury poniższej oferty.",
        f"{wolacz},\n\nW ITS WRAP nie uznajemy kompromisów. Właśnie dlatego, tworząc tę wycenę dla Twojego {marka}, zdecydowałem się na zastosowanie niezawodnej folii {czysta_folia}. To inwestycja, która zapewni Ci spokój ducha i perfekcyjną prezencję auta na drodze. Zachęcam do zapoznania się ze szczegółami.",
        f"{wolacz},\n\nKażdy samochód traktujemy w naszym studiu całkowicie indywidualnie. Aby wydobyć i trwale zabezpieczyć piękno Twojego {marka}, przygotowałem zestawienie oparte na innowacyjnej technologii folii {czysta_folia}. Z przyjemnością zaprezentuję Ci korzyści płynące z tego wyboru w poniższej ofercie.",
        f"{wolacz},\n\nOddając w nasze ręce swoje auto, oczekujesz perfekcji, a my zamierzamy ją dostarczyć. Z myślą o Twoim {marka} przygotowałem ofertę bazującą na folii {czysta_folia}, która stanowi rynkowy wzór trwałości i estetyki. Poniższa wycena to pierwszy krok do idealnej ochrony Twojego pojazdu.",
        f"{wolacz},\n\nZabezpieczenie lakieru to inwestycja, która wymaga najlepszych materiałów. Dlatego do Twojego {marka} wyselekcjonowałem folię {czysta_folia}. Jestem przekonany, że to rozwiązanie spełni Twoje najwyższe oczekiwania i pozwoli cieszyć się nieskazitelnym autem każdego dnia. Zapraszam do zapoznania się z przygotowaną ofertą."
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

# --- FUNKCJA ZAPISU DO REJESTRU EXCEL ---
def zapisz_do_rejestru(nr_oferty, handlowiec, klient, auto, usluga, folia, cena):
    try:
        sheet_rejestr = client.open_by_url("https://docs.google.com/spreadsheets/d/1iqS6geTNP3Bd_Fj_XdS-wCBrKtnGTMNQZYSso70KIkQ/edit").worksheet("Rejestr")
        dzisiaj = datetime.now().strftime("%Y-%m-%d")
        nowy_wiersz = [dzisiaj, nr_oferty, handlowiec, klient, auto, usluga, folia, f"{cena} zł", "Nowa"]
        sheet_rejestr.append_row(nowy_wiersz)
        return True
    except Exception as e:
        st.error(f"Nie udało się zapisać do bazy: {e}")
        return False

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
    st.title("👤 Opiekun Klienta")
    wybrany_handlowiec = st.selectbox("Kto przygotowuje ofertę?", list(HANDLOWCY.keys()))
    
    st.markdown("---")
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
        # Podłączamy zmienną {year} bezpośrednio do żądania modelu, żeby rysował konkretny rocznik
        if "Bezbarwne" in f_cat:
            finish = "matte/satin finish" if "Stealth" in f_color else "high gloss finish"
            prompt = f"Professional automotive studio photography of a {year} {final_brand} {final_model} ({body}). Exact {year} factory body styling. Car paint color: {paint_color}. The car is completely wrapped in clear PPF giving it a {finish}. High-end detailing garage, cinematic lighting."
        else:
            prompt = f"Professional automotive studio photography of a {year} {final_brand} {final_model} ({body}). Exact {year} factory body styling. Wrapped in {f_brand} {f_color}. High-end detailing garage, cinematic lighting."
            
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

# --- GENEROWANIE OFERTY I ZAPIS DO BAZY ---
if st.button("🔥 GENERUJ PEŁNĄ OFERTĘ PDF"):
    if 'ai_img' not in st.session_state:
        st.error("Wizualizacja auta jest wymagana. Użyj przycisku w panelu bocznym!")
    else:
        with st.spinner("Składam profesjonalny PDF..."):
            writer = PdfWriter()
            final_foil_text = f"{f_color} (na lakier: {paint_color})" if "Bezbarwne" in f_cat else f_color
            
            dane_handlowca = HANDLOWCY[wybrany_handlowiec]
            wygenerowany_wstep = generate_ai_intro_text(klient, final_brand, final_model, pakiet, final_foil_text, wybrany_handlowiec, dane_handlowca["stanowisko"])

            replacements = {
                "{{KLIENT}}": klient, 
                "{{MODEL_AUTA}}": f"{final_brand} {final_model}",
                "{{RODZAJ_FOLII}}": final_foil_text, 
                "{{USLUGA_NAZWA}}": pakiet,
                "{{NR_OFERTY}}": nr_o,
                "{{CENA_KATALOG}}": f"{cena_manual:,.2f} zł".replace(',', ' ').replace('.', ','),
                "{{CENA_KONCOWA}}": f"{cena_koncowa:,.2f} zł".replace(',', ' ').replace('.', ','),
                "{{WSTEP_AI}}": wygenerowany_wstep,
                "{{HANDLOWIEC_IMIE}}": wybrany_handlowiec,
                "{{HANDLOWIEC_TEL}}": dane_handlowca["telefon"],
                "{{HANDLOWIEC_EMAIL}}": dane_handlowca["email"]
            }

            okladka = next((f for f in pliki_na_dysku if f['name'].startswith('1_')), None)
            wstep_slide = next((f for f in pliki_na_dysku if f['name'].lower().startswith('1b_')), None)
            
            # --- ZAKTUALIZOWANA LOGIKA WYBORU STRONY PRODUKTOWEJ ---
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
            
            # --- ZAPIS DO BAZY ---
            if zapisz_do_rejestru(nr_o, wybrany_handlowiec, klient, f"{final_brand} {final_model}", pakiet, final_foil_text, cena_koncowa):
                st.success("✅ Oferta zapisana w systemie CRM!")
                
            st.balloons()
            st.download_button("📥 POBIERZ OFERTĘ PDF", data=final_io, file_name=f"Oferta_{final_brand}_{final_model}.pdf")
