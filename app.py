import streamlit as st
import pandas as pd
from pptx import Presentation
import gspread
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload
import io
import os
import subprocess
from pypdf import PdfWriter
from datetime import datetime
import re
import shutil

# --- INSTALACJA CZCIONEK ---
def install_custom_fonts():
    """Kopiuje czcionki z folderu /fonts do systemowego folderu czcionek Linuxa."""
    try:
        font_src = "fonts"
        font_dst = os.path.expanduser("~/.local/share/fonts")
        
        if os.path.exists(font_src):
            if not os.path.exists(font_dst):
                os.makedirs(font_dst)
            
            for font_file in os.listdir(font_src):
                if font_file.endswith((".ttf", ".otf")):
                    shutil.copy(os.path.join(font_src, font_file), font_dst)
            
            # Odświeżenie cache czcionek w systemie
            subprocess.run(["fc-cache", "-f", "-v"], capture_output=True)
            return True
    except Exception as e:
        st.sidebar.error(f"Nie udało się zainstalować czcionek: {e}")
    return False

# --- KONFIGURACJA ---
def get_creds():
    scope = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]
    return Credentials.from_service_account_info(st.secrets["gcp_service_account"], scopes=scope)

def download_file(file_id):
    service = build('drive', 'v3', credentials=get_creds())
    request = service.files().get_media(fileId=file_id)
    fh = io.BytesIO()
    downloader = MediaIoBaseDownload(fh, request)
    done = False
    while done is False:
        status, done = downloader.next_chunk()
    fh.seek(0)
    return fh

def pptx_to_pdf(input_pptx_path):
    output_dir = os.getcwd()
    try:
        subprocess.run([
            'libreoffice', '--headless', '--convert-to', 'pdf',
            '--outdir', output_dir, input_pptx_path
        ], check=True, capture_output=True)
        pdf_path = input_pptx_path.replace('.pptx', '.pdf')
        return pdf_path if os.path.exists(pdf_path) else None
    except:
        return None

def clean_price(price_str):
    if not price_str: return 0.0
    cleaned = re.sub(r'[^\d,]', '', str(price_str)).replace(',', '.')
    try: return float(cleaned)
    except: return 0.0

# --- START APLIKACJI ---
st.set_page_config(page_title="ITS WRAP - Generator Ofert", layout="wide")
fonts_ok = install_custom_fonts()

st.title("🛡️ Generator Ofert ITS WRAP")
if fonts_ok:
    st.sidebar.success("✅ Czcionki firmowe URW DIN zainstalowane")

try:
    # 1. Dane z Cennika
    creds = get_creds()
    client = gspread.authorize(creds)
    url_arkusza = "https://docs.google.com/spreadsheets/d/1iqS6geTNP3Bd_Fj_XdS-wCBrKtnGTMNQZYSso70KIkQ/edit?usp=drive_link"
    sheet = client.open_by_url(url_arkusza).worksheet("Ppf")
    data_all = sheet.get_all_values()
    df = pd.DataFrame(data_all[1:], columns=[c.strip() for c in data_all[0]])

    # 2. Pliki z Drive
    FOLDER_ID = "12HRnKn9KrZy_C1BSgv24PGD-Gl8lTRmn"
    service = build('drive', 'v3', credentials=creds)
    
    # TUTAJ DODAJEMY FILTR NA MIME_TYPE (tylko PowerPointy)
    query = (f"'{FOLDER_ID}' in parents and "
             f"mimeType = 'application/vnd.openxmlformats-officedocument.presentationml.presentation' and "
             f"trashed = false")
    
    results = service.files().list(q=query, fields="files(id, name)").execute()
    pliki_na_dysku = results.get('files', [])

    # --- FORMULARZ ---
    col1, col2 = st.columns(2)
    with col1:
        klient = st.text_input("Imię i Nazwisko Klienta")
        model_auta = st.text_input("Model Samochodu")
        rodzaj_folii = st.text_input("Rodzaj Folii (opcjonalnie)")
        nr_oferty = st.text_input("Numer oferty", value=f"IW/{datetime.now().strftime('%Y/%m/%d')}/01")
    
    with col2:
        pakiet = st.selectbox("Wybierz pakiet z cennika", df['Usługa'].tolist())
        rabat = st.number_input("Rabat kwotowy (PLN)", value=0, step=100)
        foto = st.file_uploader("Zdjęcie auta na okładkę", type=['jpg', 'png', 'jpeg'])

    if st.button("🚀 GENERUJ OFERTĘ PDF"):
        with st.spinner("Składam ofertę i renderuję PDF..."):
            writer = PdfWriter()
            
            wiersz = df[df['Usługa'] == pakiet].iloc[0]
            cena_kat_str = wiersz['Kwota sprzedaży']
            folia_cennik = wiersz['Rodzaj folii']
            
            folia_final = rodzaj_folii if rodzaj_folii else folia_cennik
            cena_num = clean_price(cena_kat_str)
            cena_koncowa = cena_num - rabat

            replacements = {
                "{{KLIENT}}": klient,
                "{{MODEL_AUTA}}": model_auta,
                "{{RODZAJ_FOLII}}": folia_final,
                "{{USLUGA_NAZWA}}": pakiet,
                "{{NR_OFERTY}}": nr_oferty,
                "{{CENA_KATALOG}}": f"{cena_num:,.2f} zł".replace(',', ' ').replace('.', ','),
                "{{CENA_KONCOWA}}": f"{cena_koncowa:,.2f} zł".replace(',', ' ').replace('.', ',')
            }

            # KOLEJNOŚĆ
            okladka_f = next((f for f in pliki_na_dysku if f['name'].startswith("1_")), None)
            zakres_f = next((f for f in pliki_na_dysku if f['name'].startswith("3_")), None)
            koniec_f = next((f for f in pliki_na_dysku if f['name'].startswith("6_")), None)
            dodatki = [f for f in pliki_na_dysku if not f['name'].startswith(("1", "3", "6"))]
            
            final_sequence = []
            if okladka_f: final_sequence.append(okladka_f)
            final_sequence.extend(sorted(dodatki, key=lambda x: x['name']))
            if zakres_f: final_sequence.append(zakres_f)
            if koniec_f: final_sequence.append(koniec_f)

            for f_info in final_sequence:
                stream = download_file(f_info['id'])
                prs = Presentation(stream)
                
                for slide in prs.slides:
                    # Tekst
                    for shape in slide.shapes:
                        if shape.has_text_frame:
                            for paragraph in shape.text_frame.paragraphs:
                                for run in paragraph.runs:
                                    for k, v in replacements.items():
                                        if k in run.text:
                                            run.text = run.text.replace(k, str(v))
                    
                    # Zdjęcie na okładce
                    if foto and f_info['name'].startswith("1_"):
                        for shape in slide.shapes:
                            if "{{FOTO_AUTA}}" in shape.name:
                                left, top, width, height = shape.left, shape.top, shape.width, shape.height
                                slide.shapes.add_picture(foto, left, top, width, height)
                                shape_element = shape._element
                                shape_element.getparent().remove(shape_element)

                temp_pptx = f"temp_{f_info['id']}.pptx"
                prs.save(temp_pptx)
                pdf_path = pptx_to_pdf(temp_pptx)
                if pdf_path:
                    writer.append(pdf_path)
                    os.remove(temp_pptx)
                    os.remove(pdf_path)

            final_pdf_stream = io.BytesIO()
            writer.write(final_pdf_stream)
            final_pdf_stream.seek(0)
            st.balloons()
            st.download_button("📥 POBIERZ PDF", data=final_pdf_stream, file_name=f"Oferta_{model_auta}.pdf")

except Exception as e:
    st.error(f"Wystąpił błąd: {e}")
