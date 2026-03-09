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
from PyPDF2 import PdfMerger

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
    """Zamienia PPTX na PDF przy użyciu LibreOffice zainstalowanego na serwerze."""
    output_dir = os.path.dirname(input_pptx_path)
    # Komenda systemowa LibreOffice
    process = subprocess.run([
        'libreoffice', '--headless', '--convert-to', 'pdf',
        '--outdir', output_dir, input_pptx_path
    ], capture_output=True, text=True)
    
    pdf_path = input_pptx_path.replace('.pptx', '.pdf')
    return pdf_path if os.path.exists(pdf_path) else None

# --- INTERFEJS ---
st.set_page_config(page_title="ITS WRAP - PDF Generator", layout="wide")
st.title("🛡️ Profesjonalny Generator PDF")

try:
    # Pobieranie danych z cennika
    creds = get_creds()
    client = gspread.authorize(creds)
    url_arkusza = "https://docs.google.com/spreadsheets/d/1iqS6geTNP3Bd_Fj_XdS-wCBrKtnGTMNQZYSso70KIkQ/edit?usp=drive_link"
    sheet = client.open_by_url(url_arkusza).worksheet("Ppf")
    df = pd.DataFrame(sheet.get_all_values()[1:], columns=sheet.get_all_values()[0])

    # Pobieranie listy plików
    FOLDER_ID = "12HRnKn9KrZy_C1BSgv24PGD-Gl8lTRmn"
    service = build('drive', 'v3', credentials=creds)
    query = f"'{FOLDER_ID}' in parents and mimeType = 'application/vnd.openxmlformats-officedocument.presentationml.presentation' and trashed = false"
    results = service.files().list(q=query, fields="files(id, name)").execute()
    pliki_na_dysku = results.get('files', [])

    st.sidebar.header("Wybierz klocki")
    wybrane_pliki = []
    for f in sorted(pliki_na_dysku, key=lambda x: x['name']):
        if st.sidebar.checkbox(f"{f['name']}", value=True):
            wybrane_pliki.append(f)

    klient = st.text_input("Nazwa Klienta / Auto")
    pakiet = st.selectbox("Wybierz pakiet", df[df.columns[0]].tolist())

    if st.button("🔥 GENERUJ OFERTĘ PDF"):
        with st.spinner("Składam ofertę modułowo..."):
            merger = PdfMerger()
            cena_kat = df[df[df.columns[0]] == pakiet][df.columns[1]].values[0]

            for f_info in wybrane_pliki:
                # 1. Pobierz plik
                stream = download_file(f_info['id'])
                prs = Presentation(stream)
                
                # 2. Podmień tekst w slajdach tego modułu
                for slide in prs.slides:
                    for shape in slide.shapes:
                        if hasattr(shape, "text_frame") and shape.text_frame:
                            for paragraph in shape.text_frame.paragraphs:
                                for run in paragraph.runs:
                                    if "{{KLIENT}}" in run.text: run.text = run.text.replace("{{KLIENT}}", klient)
                                    if "{{USLUGA_NAZWA}}" in run.text: run.text = run.text.replace("{{USLUGA_NAZWA}}", pakiet)
                                    if "{{CENA_KATALOG}}" in run.text: run.text = run.text.replace("{{CENA_KATALOG}}", str(cena_kat))

                # 3. Zapisz tymczasowo jako PPTX
                temp_pptx = f"temp_{f_info['name']}"
                prs.save(temp_pptx)
                
                # 4. Konwertuj na PDF
                temp_pdf = pptx_to_pdf(temp_pptx)
                if temp_pdf:
                    merger.append(temp_pdf)
                    os.remove(temp_pptx)
                    os.remove(temp_pdf)

            # 5. Sklej wszystko w jeden PDF
            final_pdf_io = io.BytesIO()
            merger.write(final_pdf_io)
            final_pdf_io.seek(0)

            st.balloons()
            st.download_button("📥 POBIERZ GOTOWĄ OFERTĘ (PDF)", data=final_pdf_io, file_name=f"Oferta_{klient}.pdf", mime="application/pdf")

except Exception as e:
    st.error(f"Błąd: {e}")
