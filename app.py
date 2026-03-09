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

# --- AUTORYZACJA ---
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

# --- FUNKCJA PODMIANY ZDJĘCIA ---
def replace_image(slide, placeholder_alt_text, image_stream):
    for shape in slide.shapes:
        try:
            alt_text = shape.non_visual_properties.name
            if not alt_text:
                alt_text = shape._element.xpath('.//p14:nvVisualPropPr/p14:altText')[0]
        except:
            alt_text = shape.name if hasattr(shape, 'name') else ""
        
        if placeholder_alt_text in alt_text:
            left, top, width, height = shape.left, shape.top, shape.width, shape.height
            slide.shapes.add_picture(image_stream, left, top, width, height)
            spTree = shape._element.getparent()
            spTree.remove(shape._element)

# --- INTERFEJS ---
st.set_page_config(page_title="ITS WRAP - Generator Ofert", layout="wide")
st.title("🛡️ Profesjonalny Generator Ofert ITS WRAP")

try:
    # 1. Dane z Cennika
    creds = get_creds()
    client = gspread.authorize(creds)
    url_arkusza = "https://docs.google.com/spreadsheets/d/1iqS6geTNP3Bd_Fj_XdS-wCBrKtnGTMNQZYSso70KIkQ/edit?usp=drive_link"
    sheet = client.open_by_url(url_arkusza).worksheet("Ppf")
    df = pd.DataFrame(sheet.get_all_values()[1:], columns=sheet.get_all_values()[0])

    # 2. Pobieranie plików z Drive
    FOLDER_ID = "12HRnKn9KrZy_C1BSgv24PGD-Gl8lTRmn"
    service = build('drive', 'v3', credentials=creds)
    query = f"'{FOLDER_ID}' in parents and mimeType = 'application/vnd.openxmlformats-officedocument.presentationml.presentation' and trashed = false"
    results = service.files().list(q=query, fields="files(id, name)").execute()
    pliki_na_dysku = results.get('files', [])

    # --- FORMULARZ ---
    col1, col2 = st.columns(2)
    with col1:
        klient = st.text_input("Nazwa Klienta / Auto", placeholder="np. Jan Kowalski / Porsche 911")
        pakiet = st.selectbox("Wybierz pakiet z cennika", df[df.columns[0]].tolist())
        nr_oferty = st.text_input("Numer oferty", value=f"IW/{datetime.now().strftime('%Y/%m/%d')}/01")
    
    with col2:
        rabat = st.number_input("Rabat kwotowy (PLN)", value=0, step=100)
        foto = st.file_uploader("Zdjęcie auta na okładkę", type=['jpg', 'png', 'jpeg'])

    # Wybór stron produktowych (środkowych)
    st.sidebar.header("Dodatkowe strony")
    pliki_srodkowe = [f for f in pliki_na_dysku if not f['name'].startswith(("1_", "3_", "6_"))]
    wybrane_extra = []
    for f in sorted(pliki_srodkowe, key=lambda x: x['name']):
        if st.sidebar.checkbox(f['name'], value=True):
            wybrane_extra.append(f)

    if st.button("🔥 GENERUJ PEŁNĄ OFERTĘ PDF"):
        with st.spinner("Składam ofertę..."):
            writer = PdfWriter()
            
            # Pobranie ceny
            wiersz = df[df[df.columns[0]] == pakiet]
            cena_kat = wiersz[df.columns[1]].values[0]
            cena_num = float(''.join(filter(str.isdigit, str(cena_kat).replace(',','.'))))
            cena_koncowa = cena_num - rabat

            replacements = {
                "{{KLIENT}}": klient,
                "{{USLUGA_NAZWA}}": pakiet,
                "{{NR_OFERTY}}": nr_oferty,
                "{{CENA_KATALOG}}": f"{cena_kat}",
                "{{CENA_KONCOWA}}": f"{cena_koncowa:,.2f} zł".replace(',', ' ').replace('.', ',')
            }

            # KOLEJNOŚĆ SKŁADANIA:
            # 1. Okładka (1_) -> 2. Wybrane dodatki -> 3. Zakres/Oferta (3_) -> 4. Koniec (6_)
            okladka_plik = [f for f in pliki_na_dysku if f['name'].startswith("1_")][0]
            zakres_plik = [f for f in pliki_na_dysku if f['name'].startswith("3_")][0]
            koniec_plik = [f for f in pliki_na_dysku if f['name'].startswith("6_")][0]
            
            finalna_lista = [okladka_plik] + wybrane_extra + [zakres_plik] + [koniec_plik]

            for f_info in finalna_lista:
                stream = download_file(f_info['id'])
                prs = Presentation(stream)
                
                for slide in prs.slides:
                    # Podmiana tekstu
                    for shape in slide.shapes:
                        if hasattr(shape, "text_frame") and shape.text_frame:
                            for paragraph in shape.text_frame.paragraphs:
                                for run in paragraph.runs:
                                    for k, v in replacements.items():
                                        if k in run.text: run.text = run.text.replace(k, str(v))
                    
                    # Podmiana zdjęcia (tylko jeśli wgrano foto i jesteśmy na okładce)
                    if f_info['name'].startswith("1_") and foto:
                        replace_image(slide, "{{FOTO_AUTA}}", foto)

                # Konwersja do PDF
                temp_pptx = f"temp_{f_info['id']}.pptx"
                prs.save(temp_pptx)
                pdf_file = pptx_to_pdf(temp_pptx)
                if pdf_file:
                    writer.append(pdf_file)
                    os.remove(temp_pptx)
                    os.remove(pdf_file)

            # Budowanie finału
            final_pdf = io.BytesIO()
            writer.write(final_pdf)
            final_pdf.seek(0)

            st.balloons()
            st.download_button("📥 POBIERZ PDF", data=final_pdf, file_name=f"Oferta_{klient}.pdf")

except Exception as e:
    st.error(f"Coś poszło nie tak: {e}")
