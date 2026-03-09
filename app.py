import streamlit as st
import pandas as pd
from pptx import Presentation
import gspread
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload
import io, os, subprocess, re, shutil
from pypdf import PdfWriter
from datetime import datetime

# --- SYSTEM CZCIONEK URW++ ---
def install_fonts():
    src = "fonts"
    dst = os.path.expanduser("~/.local/share/fonts")
    try:
        if os.path.exists(src):
            if not os.path.exists(dst): os.makedirs(dst)
            font_files = []
            for f in os.listdir(src):
                if f.lower().endswith((".ttf", ".otf")):
                    shutil.copy(os.path.join(src, f), dst)
                    font_files.append(f)
            # Przebudowanie cache czcionek
            subprocess.run(["fc-cache", "-f", "-v"], capture_output=True)
            return font_files
    except Exception as e:
        st.sidebar.error(f"Błąd instalacji czcionek: {e}")
    return []

# --- NARZĘDZIA ---
def get_service():
    creds = Credentials.from_service_account_info(st.secrets["gcp_service_account"], 
            scopes=["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"])
    return build('drive', 'v3', credentials=creds), creds

def download_file(service, file_id):
    request = service.files().get_media(fileId=file_id)
    fh = io.BytesIO(); downloader = MediaIoBaseDownload(fh, request)
    done = False
    while not done: _, done = downloader.next_chunk()
    fh.seek(0); return fh

def pptx_to_pdf(input_path):
    try:
        # Konwersja z wymuszeniem headless (bez okna)
        subprocess.run(['libreoffice', '--headless', '--convert-to', 'pdf', '--outdir', os.getcwd(), input_path], check=True, capture_output=True)
        pdf_name = os.path.basename(input_path).replace('.pptx', '.pdf')
        return pdf_name if os.path.exists(pdf_name) else None
    except Exception as e:
        st.error(f"Błąd LibreOffice: {e}")
        return None

# --- APLIKACJA ---
st.set_page_config(page_title="ITS WRAP v3.2 - URW DIN", layout="wide")
installed = install_fonts()

st.sidebar.title("⚙️ System Czcionek")
if installed:
    st.sidebar.success(f"Aktywne czcionki URW++: {', '.join(installed)}")
else:
    st.sidebar.warning("Brak plików w folderze /fonts")

st.title("🛡️ Generator Ofert ITS WRAP")

try:
    service, creds = get_service()
    client = gspread.authorize(creds)
    
    # Cennik
    url_arkusza = "https://docs.google.com/spreadsheets/d/1iqS6geTNP3Bd_Fj_XdS-wCBrKtnGTMNQZYSso70KIkQ/edit?usp=drive_link"
    sheet = client.open_by_url(url_arkusza).worksheet("Ppf")
    data = sheet.get_all_values()
    df = pd.DataFrame(data[1:], columns=[c.strip() for c in data[0]])

    # Pliki Drive
    FOLDER_ID = "12HRnKn9KrZy_C1BSgv24PGD-Gl8lTRmn"
    q = f"'{FOLDER_ID}' in parents and mimeType='application/vnd.openxmlformats-officedocument.presentationml.presentation' and trashed=false"
    results = service.files().list(q=q, fields="files(id, name)").execute()
    wszystkie_pliki = results.get('files', [])

    # Formularz
    col1, col2 = st.columns(2)
    with col1:
        klient_val = st.text_input("Imię i Nazwisko Klienta")
        model_val = st.text_input("Model Samochodu")
        nr_oferty_val = st.text_input("Numer oferty", value=f"IW/{datetime.now().strftime('%Y/%m/%d')}/01")
    with col2:
        pakiet_val = st.selectbox("Wybierz pakiet", df['Usługa'].tolist())
        rodzaj_folii_manual = st.text_input("Rodzaj Folii (opcjonalnie)")
        rabat_val = st.number_input("Rabat (PLN)", value=0)
        foto_val = st.file_uploader("Zdjęcie auta na okładkę", type=['jpg','png','jpeg'])

    # Dodatki
    st.sidebar.markdown("---")
    st.sidebar.header("Dodatkowe strony")
    dodatki_pliki = [f for f in wszystkie_pliki if not f['name'].startswith(('1','3','6'))]
    wybrane_extra = []
    for d in sorted(dodatki_pliki, key=lambda x: x['name']):
        if st.sidebar.checkbox(d['name'], value=False):
            wybrane_extra.append(d)

    if st.button("🚀 GENERUJ PDF"):
        with st.spinner("Składam ofertę z zachowaniem czcionek URW++..."):
            writer = PdfWriter()
            
            # Pobranie danych z cennika
            row = df[df['Usługa'] == pakiet_val].iloc[0]
            cena_raw = re.sub(r'[^\d,]', '', row['Kwota sprzedaży']).replace(',', '.')
            cena_num = float(cena_raw) if cena_raw else 0.0
            folia_tekst = rodzaj_folii_manual if rodzaj_folii_manual else row['Rodzaj folii']

            replacements = {
                "{{KLIENT}}": klient_val,
                "{{MODEL_AUTA}}": model_val,
                "{{RODZAJ_FOLII}}": folia_tekst,
                "{{USLUGA_NAZWA}}": pakiet_val,
                "{{NR_OFERTY}}": nr_oferty_val,
                "{{CENA_KATALOG}}": f"{cena_num:,.2f} zł".replace(',', ' ').replace('.', ','),
                "{{CENA_KONCOWA}}": f"{(cena_num - rabat_val):,.2f} zł".replace(',', ' ').replace('.', ',')
            }

            # Precyzyjna sekwencja (UNIKALNE)
            okladka_f = next((f for f in wszystkie_pliki if f['name'].startswith('1')), None)
            zakres_f = next((f for f in wszystkie_pliki if f['name'].startswith('3')), None)
            koniec_f = next((f for f in wszystkie_pliki if f['name'].startswith('6')), None)

            seq = []
            if okladka_f: seq.append(okladka_f)
            seq.extend(wybrane_extra)
            if zakres_f: seq.append(zakres_f)
            if koniec_f: seq.append(koniec_f)

            for f_info in seq:
                prs = Presentation(download_file(service, f_info['id']))
                for slide in prs.slides:
                    # Tekst
                    for shape in slide.shapes:
                        if shape.has_text_frame:
                            for p in shape.text_frame.paragraphs:
                                for run in p.runs:
                                    for k, v in replacements.items():
                                        if k in run.text: run.text = run.text.replace(k, str(v))
                    
                    # FOTO (tylko na okładce)
                    if foto_val and f_info['name'].startswith('1'):
                        for shape in slide.shapes:
                            # Szukamy {{FOTO_AUTA}} w nazwie lub tekście alternatywnym
                            alt_text = ""
                            try: alt_text = shape.non_visual_properties.name
                            except: pass
                            
                            if "{{FOTO_AUTA}}" in shape.name or "{{FOTO_AUTA}}" in alt_text:
                                slide.shapes.add_picture(foto_val, shape.left, shape.top, shape.width, shape.height)
                                shape._element.getparent().remove(shape._element)

                tmp_p = f"tmp_{f_info['id']}.pptx"
                prs.save(tmp_p)
                pdf_f = pptx_to_pdf(tmp_p)
                if pdf_f: 
                    writer.append(pdf_f)
                    os.remove(tmp_p); os.remove(pdf_f)

            final_io = io.BytesIO()
            writer.write(final_io); final_io.seek(0)
            st.balloons()
            st.download_button("📥 POBIERZ PDF", data=final_io, file_name=f"Oferta_{model_val}.pdf")

except Exception as e:
    st.error(f"Błąd aplikacji: {e}")
