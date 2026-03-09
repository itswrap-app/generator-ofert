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
    # Próba konwersji przez LibreOffice
    try:
        subprocess.run([
            'libreoffice', '--headless', '--convert-to', 'pdf',
            '--outdir', output_dir, input_pptx_path
        ], check=True, capture_output=True)
        pdf_path = input_pptx_path.replace('.pptx', '.pdf')
        return pdf_path if os.path.exists(pdf_path) else None
    except:
        return None

# --- INTERFEJS ---
st.set_page_config(page_title="ITS WRAP - Generator PDF", layout="wide")
st.title("🛡️ Generator Ofert ITS WRAP (Final Fix)")

try:
    # Pobieranie danych
    creds = get_creds()
    client = gspread.authorize(creds)
    url_arkusza = "https://docs.google.com/spreadsheets/d/1iqS6geTNP3Bd_Fj_XdS-wCBrKtnGTMNQZYSso70KIkQ/edit?usp=drive_link"
    sheet = client.open_by_url(url_arkusza).worksheet("Ppf")
    df = pd.DataFrame(sheet.get_all_values()[1:], columns=sheet.get_all_values()[0])

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

    if st.button("🚀 GENERUJ PDF"):
        if not wybrane_pliki:
            st.warning("Wybierz pliki po lewej!")
        else:
            with st.spinner("Składam ofertę..."):
                writer = PdfWriter()
                cena_kat = df[df[df.columns[0]] == pakiet][df.columns[1]].values[0]

                for f_info in wybrane_pliki:
                    # Pobieranie i podmiana tekstu
                    stream = download_file(f_info['id'])
                    prs = Presentation(stream)
                    for slide in prs.slides:
                        for shape in slide.shapes:
                            if hasattr(shape, "text_frame") and shape.text_frame:
                                for paragraph in shape.text_frame.paragraphs:
                                    for run in paragraph.runs:
                                        if "{{KLIENT}}" in run.text: run.text = run.text.replace("{{KLIENT}}", klient)
                                        if "{{USLUGA_NAZWA}}" in run.text: run.text = run.text.replace("{{USLUGA_NAZWA}}", pakiet)
                                        if "{{CENA_KATALOG}}" in run.text: run.text = run.text.replace("{{CENA_KATALOG}}", str(cena_kat))

                    # Zapis tymczasowy i konwersja
                    temp_name = f"temp_{f_info['id']}.pptx"
                    prs.save(temp_name)
                    pdf_file = pptx_to_pdf(temp_name)
                    
                    if pdf_file:
                        writer.append(pdf_file)
                        os.remove(temp_name)
                        os.remove(pdf_file)

                # Finalny PDF
                final_output = io.BytesIO()
                writer.write(final_output)
                final_output.seek(0)

                st.balloons()
                st.download_button("📥 POBIERZ PDF", data=final_output, file_name=f"Oferta_{klient}.pdf", mime="application/pdf")

except Exception as e:
    st.error(f"Błąd: {e}")
