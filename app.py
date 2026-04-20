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

# --- MAPOWANIE SEGMENTÓW (Domyślne przypisanie do modeli) ---
SEGMENTY_AUT = {
    "Audi": {"A3": "Segment C", "A4": "Segment D", "A6": "Segment D", "Q3": "Segment C", "Q5": "Segment D", "Q8": "Segment E", "RS6": "Segment D", "e-tron GT": "Segment E"},
    "BMW": {"Seria 3": "Segment D", "Seria 4": "Segment D", "Seria 5": "Segment D", "Seria 7": "Segment E", "X3": "Segment D", "X5": "Segment E", "M3": "Segment D", "M4": "Segment D"},
    "Porsche": {"911 (992)": "Segment D", "Taycan": "Segment E", "Macan": "Segment D", "Panamera": "Segment E", "Cayenne": "Segment E"},
    "Tesla": {"Model 3": "Segment D", "Model Y": "Segment D", "Model S": "Segment E", "Model X": "Segment J"},
    "Toyota": {"Corolla": "Segment C", "Yaris": "Segment B", "RAV4": "Segment D", "C-HR": "Segment C", "Camry": "Segment D"},
    "Mercedes-Benz": {"Klasa C": "Segment D", "Klasa E": "Segment E", "GLC": "Segment D", "GLE": "Segment E", "Klasa S": "Segment E", "Klasa G": "Segment J"},
    "Inna marka...": {"Wpisz ręcznie": "Segment D"}
}

# --- KONFIGURACJA NADWOZIA DLA PROMPTÓW AI ---
CAR_BODY_TYPES = {
    "Audi": {"A3": ["Hatchback", "Sedan"], "A4": ["Sedan", "Kombi"], "A6": ["Sedan", "Kombi"], "Q3": ["SUV"], "Q5": ["SUV"], "Q8": ["SUV"], "e-tron GT": ["Sedan"], "RS6": ["Kombi"]},
    "BMW": {"Seria 3": ["Sedan", "Kombi"], "Seria 4": ["Coupe", "Gran Coupe"], "Seria 5": ["Sedan", "Kombi"], "Seria 7": ["Sedan"], "X3": ["SUV"], "X5": ["SUV"], "M3": ["Sedan", "Kombi"], "M4": ["Coupe"]},
    "Porsche": {"911 (992)": ["Coupe", "Cabriolet"], "Taycan": ["Sedan", "Cross Turismo"], "Macan": ["SUV"], "Panamera": ["Sedan"], "Cayenne": ["SUV", "Coupe"]},
    "Tesla": {"Model 3": ["Sedan"], "Model Y": ["SUV"], "Model S": ["Sedan"], "Model X": ["SUV"]},
    "Toyota": {"Corolla": ["Hatchback", "Sedan", "Kombi"], "Yaris": ["Hatchback"], "RAV4": ["SUV"], "C-HR": ["Crossover"], "Camry": ["Sedan"]},
    "Inna marka...": {"Wpisz ręcznie": ["Inne"]}
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
    }
}

# --- SYSTEMOWE ---
def install_fonts():
    font_src, font_dst = "fonts", os.path.expanduser("~/.local/share/fonts")
    if os.path.exists(font_src):
        if not os.path.exists(font_dst): os.makedirs(font_dst)
        for f in os.listdir(font_src):
            if f.lower().endswith((".ttf", ".otf")): shutil.copy(os.path.join(font_src, f), font_dst)
        subprocess.run(["fc-cache", "-f"], capture_output=True)

def generate_ai_image(prompt):
    api_key = st.secrets["GEMINI_API_KEY"]
    url = f"https://generativelanguage.googleapis.com/v1beta/models/imagen-4.0-ultra-generate-001:predict?key={api_key}"
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
    return None

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
            elif imie_lower.endswith(('k', 'g', 'ch', 'j')): wolacz = f"Panie {imie}u"
            else: wolacz = f"Panie {imie}ie"
    szablony = [
        f"{wolacz},\n\nDziękuję za wybór naszej firmy. Komponując ofertę dla Twojego {brand}, dobraliśmy bezkompromisowe rozwiązanie, jakim jest folia {czysta_folia}. Dzięki temu mogę zagwarantować Tobie najwyższą jakość ochrony samochodu na długie lata.",
        f"{wolacz},\n\nW ITS WRAP nie uznajemy kompromisów. Właśnie dlatego, tworząc tę wycenę dla Twojego {brand}, zdecydowałem się na zastosowanie niezawodnej folii {czysta_folia}. To inwestycja, która zapewni Ci spokój ducha."
    ]
    return f"{random.choice(szablony)}\n\nZ motoryzacyjnym pozdrowieniem,\n{handlowiec_imie}\n{handlowiec_stanowisko}"

def download_file(service, file_id):
    request = service.files().get_media(fileId=file_id)
    fh = io.BytesIO(); downloader = MediaIoBaseDownload(fh, request); done = False
    while not done: _, done = downloader.next_chunk()
    fh.seek(0); return fh

def pptx_to_pdf(input_path):
    try:
        subprocess.run(['libreoffice', '--headless', '--convert-to', 'pdf', '--outdir', os.getcwd(), input_path], check=True, capture_output=True)
        return os.path.basename(input_path).replace('.pptx', '.pdf')
    except: return None

# --- APLIKACJA START ---
st.set_page_config(page_title="ITS WRAP Generator v3", layout="wide")
install_fonts()

creds = Credentials.from_service_account_info(st.secrets["gcp_service_account"], scopes=["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"])
service = build('drive', 'v3', credentials=creds)
client = gspread.authorize(creds)

results = service.files().list(q="'12HRnKn9KrZy_C1BSgv24PGD-Gl8lTRmn' in parents and mimeType='application/vnd.openxmlformats-officedocument.presentationml.presentation' and trashed=false", fields="files(id, name)").execute()
pliki_na_dysku = results.get('files', [])

# !!! NOWY LINK v3 !!!
LINK_DO_CENNIKA = "https://docs.google.com/spreadsheets/d/1USF81hOinAP_vvz1QZuoNyRCT1ezJcXTDDB6RjuYtrY/edit?"

try:
    sh = client.open_by_url(LINK_DO_CENNIKA)
    # Pancerne czyszczenie nagłówków
    sheet_c = sh.worksheet("Cennik usług")
    nag_c = [c.replace('\n', ' ').replace('\r', '').strip() for c in sheet_c.get_all_values()[0]]
    df_cennik = pd.DataFrame(sheet_c.get_all_values()[1:], columns=nag_c)

    sheet_r = sh.worksheet("Rejestr")
    nag_r = [c.replace('\n', ' ').replace('\r', '').strip() for c in sheet_r.get_all_values()[0]]
    df_rejestr = pd.DataFrame(sheet_r.get_all_values()[1:], columns=nag_r)
except Exception as e:
    st.error(f"⚠️ Błąd arkusza: {e}")
    st.stop()

# --- LOGIKA CENOWA ---
def pobierz_cene_uslugi(kategoria, usluga, segment):
    try:
        row = df_cennik[(df_cennik['Kategoria'] == kategoria) & (df_cennik['Usługa'] == usluga) & (df_cennik['Segment'] == segment)]
        if not row.empty:
            cena_str = str(row['Cena sprzedaży netto PLN'].values[0])
            cena_str = cena_str.replace(' ', '').replace('\xa0', '')
            if ',' in cena_str: cena_str = cena_str.replace('.', '').replace(',', '.')
            cena_czysta = re.sub(r'[^\d.]', '', cena_str)
            return float(cena_czysta) if cena_czysta else 0.0
    except:
        return 0.0
    return 0.0

def generuj_numer_oferty():
    try:
        if df_rejestr.empty: return f"IW/{datetime.now().strftime('%Y/%m')}/001"
        ostatni_nr = df_rejestr.iloc[-1]['Nr Oferty']
        numer = int(ostatni_nr.split('/')[-1]) + 1
        return f"IW/{datetime.now().strftime('%Y/%m')}/{numer:03d}"
    except:
        return f"IW/{datetime.now().strftime('%Y/%m')}/001"

# --- UI BOCZNE ---
with st.sidebar:
    st.title("👤 Konfiguracja")
    wybrany_handlowiec = st.selectbox("Opiekun", list(HANDLOWCY.keys()))
    st.markdown("---")
    brand = st.selectbox("Marka", list(CAR_BODY_TYPES.keys()))
    if brand == "Inna marka...":
        final_brand, final_model, body, segment_domyslny = st.text_input("Marka"), st.text_input("Model"), "Inne", "Segment D"
    else:
        final_brand = brand
        final_model = st.selectbox("Model", list(CAR_BODY_TYPES[brand].keys()))
        body = st.selectbox("Nadwozie", CAR_BODY_TYPES[brand][final_model])
        segment_domyslny = SEGMENTY_AUT.get(brand, {}).get(final_model, "Segment D")
    
    segment_final = st.selectbox("Segment", ["Segment A", "Segment B", "Segment C", "Segment D", "Segment E", "Segment J", "Wavecamper"], 
                                 index=["Segment A", "Segment B", "Segment C", "Segment D", "Segment E", "Segment J", "Wavecamper"].index(segment_domyslny))
    
    year = st.selectbox("Rocznik", [str(y) for y in range(2026, 1999, -1)])
    gen_code = st.text_input("Kod karoserii (np. G70, 992)")

    st.markdown("---")
    f_brand = st.selectbox("Producent folii", list(FOIL_GROUPS.keys()))
    f_cat = st.selectbox("Wykończenie", list(FOIL_GROUPS[f_brand].keys()))
    f_color = st.selectbox("Kolor", FOIL_GROUPS[f_brand][f_cat])
    paint_color = st.text_input("Obecny lakier", value="Czarny metallic") if "Bezbarwne" in f_cat else ""

    if st.button("🪄 GENERUJ WIZUALIZACJĘ AI"):
        extra = f" {gen_code} facelift model," if gen_code else ""
        prompt = f"Professional automotive studio photography of a {year} {final_brand} {final_model} ({body}). Exact factory body styling.{extra} Wrapped in {f_brand} {f_color}. Detailing studio background."
        with st.spinner("Generowanie..."):
            img = generate_ai_image(prompt)
            if img: st.session_state['ai_img'] = img

    st.markdown("---")
    dodatki_pliki = [f for f in pliki_na_dysku if f['name'].startswith(('4','5'))]
    wybrane_dodatki = [d for d in sorted(dodatki_pliki, key=lambda x: x['name']) if st.checkbox(d['name'])]

# --- UI GŁÓWNE ---
st.title("🛡️ Generator Ofert ITS WRAP v3")
t1, t2 = st.tabs(["📄 Nowa Oferta", "🗄️ Rejestr"])

with t1:
    c1, c2 = st.columns(2)
    with c1:
        st.subheader("Klient")
        klient, nip = st.text_input("Imię i Nazwisko / Firma"), st.text_input("NIP")
        nr_o = st.text_input("Numer oferty", value=generuj_numer_oferty())
        
        st.markdown("---")
        st.subheader("Usługa")
        kat_list = df_cennik['Kategoria'].dropna().unique().tolist()
        wybrana_kat = st.selectbox("Kategoria", [k for k in kat_list if k])
        usl_list = df_cennik[df_cennik['Kategoria'] == wybrana_kat]['Usługa'].dropna().unique().tolist()
        wybrany_pakiet = st.selectbox("Pakiet", [u for u in usl_list if u])
        
        cena_bazowa = pobierz_cene_uslugi(wybrana_kat, wybrany_pakiet, segment_final)
        cena_manual = st.number_input("Cena netto (PLN)", value=cena_bazowa)
        rabat = st.number_input("Rabat (PLN)", value=0.0)
        st.info(f"**Do zapłaty: {cena_manual - rabat:,.2f} zł netto**")

    with c2:
        if 'ai_img' in st.session_state: st.image(st.session_state['ai_img'], use_container_width=True)
        else: st.info("Podgląd AI pojawi się tutaj.")

    if st.button("🔥 GENERUJ PDF"):
        if 'ai_img' not in st.session_state: st.error("Najpierw wygeneruj zdjęcie AI!")
        else:
            with st.spinner("Składanie PDF..."):
                writer = PdfWriter()
                foil_text = f"{f_color} (na lakier: {paint_color})" if paint_color else f_color
                d_h = HANDLOWCY[wybrany_handlowiec]
                wstep = generate_ai_intro_text(klient, final_brand, final_model, wybrany_pakiet, foil_text, wybrany_handlowiec, d_h["stanowisko"])

                reps = {
                    "{{KLIENT}}": klient, "{{MODEL_AUTA}}": f"{final_brand} {final_model}",
                    "{{RODZAJ_FOLII}}": foil_text, "{{USLUGA_NAZWA}}": wybrany_pakiet,
                    "{{NR_OFERTY}}": nr_o, "{{CENA_KONCOWA}}": f"{cena_manual - rabat:,.2f} zł",
                    "{{WSTEP_AI}}": wstep, "{{HANDLOWIEC_IMIE}}": wybrany_handlowiec,
                    "{{HANDLOWIEC_TEL}}": d_h["telefon"], "{{HANDLOWIEC_EMAIL}}": d_h["email"]
                }

                okladka = next((f for f in pliki_na_dysku if f['name'].startswith('1_')), None)
                wstep_s = next((f for f in pliki_na_dysku if f['name'].lower().startswith('1b_')), None)
                
                # Logika wyboru slajdu 2 (Produktu)
                p_id = None
                if "reklam" in wybrany_pakiet.lower(): p_id = next((f for f in pliki_na_dysku if 'reklama' in f['name'].lower()), None)
                elif "3M" in f_brand: p_id = next((f for f in pliki_na_dysku if '3m' in f['name'].lower()), None)
                elif "Stealth" in f_color: p_id = next((f for f in pliki_na_dysku if 'stealth' in f['name'].lower()), None)
                else: p_id = next((f for f in pliki_na_dysku if f['name'].startswith('2')), None)

                zakres = next((f for f in pliki_na_dysku if f['name'].startswith('3')), None)
                koniec = next((f for f in pliki_na_dysku if f['name'].startswith('6')), None)

                seq = [okladka, wstep_s, p_id, zakres] + wybrane_dodatki + [koniec]
                for f_inf in [x for x in seq if x]:
                    prs = Presentation(download_file(service, f_inf['id']))
                    for slide in prs.slides:
                        for shape in slide.shapes:
                            if f_inf['name'].startswith('1_') and "{{FOTO_AUTA}}" in (shape.name or ""):
                                slide.shapes.add_picture(io.BytesIO(st.session_state['ai_img']), shape.left, shape.top, shape.width, shape.height)
                            if shape.has_text_frame:
                                for p in shape.text_frame.paragraphs:
                                    for run in p.runs:
                                        for k, v in reps.items():
                                            if k in run.text: run.text = run.text.replace(k, str(v))
                    tmp = f"tmp_{f_inf['id']}.pptx"
                    prs.save(tmp)
                    pdf = pptx_to_pdf(tmp)
                    if pdf: writer.append(pdf); os.remove(tmp); os.remove(pdf)

                out = io.BytesIO(); writer.write(out); out.seek(0)
                sheet_r.append_row([datetime.now().strftime("%Y-%m-%d"), nr_o, wybrany_handlowiec, klient, nip, f"{final_brand} {final_model}", wybrany_pakiet, foil_text, f"{cena_manual-rabat} zł", "Wysłana"])
                st.balloons()
                st.download_button("📥 POBIERZ PDF", out, f"Oferta_{nr_o.replace('/','_')}.pdf")

with t2:
    st.dataframe(df_rejestr.tail(20), use_container_width=True)
