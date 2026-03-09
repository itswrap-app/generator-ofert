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
    """Zamienia '14 000,00 zł' na liczbę 14000.0"""
    try:
        cleaned = re.sub(r'[^\d,]', '', str(price_str))
        cleaned = cleaned.replace(',', '.')
        return float(cleaned)
    except:
        return 0.0

# --- INTERFEJS ---
st.set_page_config(page_title="ITS WRAP - Generator Ofert", layout="wide")
st.title("🛡️ Profesjonalny Generator Ofert ITS WRAP")

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
    query = f"'{FOLDER_ID}' in parents and mimeType = 'application/vnd.openxmlformats-officedocument.presentationml.presentation' and trashed = false"
    results = service.files().list(q=query, fields="files(id, name)").execute()
    pliki_na_dysku = results.get('files', [])

    # --- FORMULARZ ---
    col1, col2 = st.columns(2)
    with col1:
        klient_auto = st.text_input("Model auta / Klient", placeholder="np. Porsche 911 / Jan Kowalski")
        pakiet = st.selectbox("Wybierz pakiet z cennika", df['Usługa'].tolist())
        nr_oferty = st.text_input("Numer oferty", value=f"IW/{datetime.now().strftime('%Y/%m/%d')}/01")
    
    with col2:
        rabat = st.number_input("Rabat kwotowy (PLN)", value=0, step=100)
        foto = st.file_uploader("Zdjęcie auta na okładkę", type=['jpg', 'png', 'jpeg'])

    if st.button("🔥 GENERUJ OFERTĘ PDF"):
        with st.spinner("Przetwarzam dane i generuję PDF..."):
            writer = PdfWriter()
            
            # Pobranie danych z wiersza arkusza
            wiersz = df[df['Usługa'] == pakiet].iloc[0]
            cena_kat_str = wiersz['Kwota sprzedaży']
            rodzaj_folii = wiersz['Rodzaj folii']
            
            cena_num = clean_price(cena_kat_str)
            cena_koncowa = cena_num - rabat

            replacements = {
                "{{KLIENT}}": klient_auto,
                "{{MODEL_AUTA}}": klient_auto,
                "{{USLUGA_NAZWA}}": pakiet,
                "{{RODZAJ_FOLII}}": rodzaj_folii,
                "{{NR_OFERTY}}": nr_oferty,
                "{{CENA_KATALOG}}": f"{cena_num:,.2f} zł".replace(',', ' ').replace('.', ','),
                "{{CENA_KONCOWA}}": f"{cena_koncowa:,.2f} zł".replace(',', ' ').replace('.', ',')
            }

            # Kolejność składania (1 -> reszta -> 3 -> 6)
            okladka_f = [f for f in pliki_na_dysku if f['name'].startswith("1_")][0]
            zakres_f = [f for f in pliki_na_dysku if f['name'].startswith("3_")][0]
            koniec_f = [f for f in pliki_na_dysku if f['name'].startswith("6_")][0]
            dodatki = [f for f in pliki_na_dysku if not f['name'].startswith(("1_", "3_", "6_"))]
            
            final_files = [okladka_f] + dodatki + [zakres_f] + [koniec_f]

            for f_info in final_files:
                stream = download_file(f_info['id'])
                prs = Presentation(stream)
                
                for slide in prs.slides:
                    # Podmiana tekstu z zachowaniem stylu
                    for shape in slide.shapes:
                        if hasattr(shape, "text_frame") and shape.text_frame:
                            for paragraph in shape.text_frame.paragraphs:
                                for run in paragraph.runs:
                                    for k, v in replacements.items():
                                        if k in run.text:
                                            run.text = run.text.replace(k, str(v))
                    
                    # Podmiana zdjęcia {{FOTO_AUTA}}
                    if foto and f_info['name'].startswith("1_"):
                        for shape in slide.shapes:
                            alt_text = ""
                            try:
                                alt_text = shape._element.xpath('.//p14:nvVisualPropPr/p14:altText')[0]
                            except: pass
                            if "{{FOTO_AUTA}}" in alt_text or "{{FOTO_AUTA}}" in shape.name:
                                left, top, width, height = shape.left, shape.top, shape.width, shape.height
                                slide.shapes.add_picture(foto, left, top, width, height)
                                shape._element.getparent().remove(shape._element)

                temp_pptx = f"temp_{f_info['id']}.pptx"
                prs.save(temp_pptx)
                pdf_file = pptx_to_pdf(temp_pptx)
                if pdf_file:
                    writer.append(pdf_file)
                    os.remove(temp_pptx)
                    os.remove(pdf_file)

            final_pdf = io.BytesIO()
            writer.write(final_pdf)
            final_pdf.seek(0)
            st.download_button("📥 POBIERZ OFERTĘ (PDF)", data=final_pdf, file_name=f"Oferta_{klient_auto}.pdf")

except Exception as e:
    st.error(f"Błąd: {e}")
