import streamlit as st
import pandas as pd
from pptx import Presentation
import gspread
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload
import io
import os

# --- KONFIGURACJA I AUTORYZACJA ---
def get_creds():
    scope = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]
    return Credentials.from_service_account_info(st.secrets["gcp_service_account"], scopes=scope)

def get_drive_service():
    return build('drive', 'v3', credentials=get_creds())

def get_data():
    creds = get_creds()
    client = gspread.authorize(creds)
    # TUTAJ TWÓJ LINK DO ARKUSZA (Zostaw ten co miałeś wcześniej)
    url_arkusza = "https://docs.google.com/spreadsheets/d/12HRnKn9KrZy_C1BSgv24PGD-Gl8lTRmn" 
    sheet = client.open_by_url(url_arkusza).worksheet("Ppf")
    data = sheet.get_all_values()
    df = pd.DataFrame(data[1:], columns=data[0])
    df.columns = df.columns.str.strip()
    return df

# --- FUNKCJA POBIERANIA PLIKÓW ---
def download_file(file_id):
    service = get_drive_service()
    request = service.files().get_media(fileId=file_id)
    fh = io.BytesIO()
    downloader = MediaIoBaseDownload(fh, request)
    done = False
    while done is False:
        status, done = downloader.next_chunk()
    fh.seek(0)
    return fh

# --- INTERFEJS ---
st.set_page_config(page_title="ITS WRAP - Generator Live Drive", layout="wide")
st.title("🛡️ Generator Ofert ITS WRAP (Live Drive)")

try:
    df = get_data()
    
    # 1. Pobieramy listę plików z folderu na Drive
    FOLDER_ID = "12HRnKn9KrZy_C1BSgv24PGD-Gl8lTRmn"
    service = get_drive_service()
    results = service.files().list(
        q=f"'{FOLDER_ID}' in parents and trashed = false",
        fields="files(id, name)"
    ).execute()
    pliki_na_dysku = results.get('files', [])

    # Sidebar z wyborem plików
    st.sidebar.header("Wybierz klocki do oferty")
    wybrane_pliki = []
    # Sortujemy, żeby pliki 1_, 2_ były w dobrej kolejności
    for f in sorted(pliki_na_dysku, key=lambda x: x['name']):
        if st.sidebar.checkbox(f"{f['name']}", value=True):
            wybrane_pliki.append(f)

    # Formularz
    col1, col2 = st.columns(2)
    with col1:
        klient = st.text_input("Nazwa Klienta / Auto")
        pakiet = st.selectbox("Wybierz pakiet główny", df[df.columns[0]].tolist())
    with col2:
        foto = st.file_uploader("Zdjęcie na okładkę", type=['jpg', 'png'])
        rabat = st.number_input("Rabat kwotowy (PLN)", value=0)

    if st.button("🚀 GENERUJ OFERTĘ Z DYSKU"):
        with st.spinner("Pobieram i przetwarzam pliki..."):
            for f_info in wybrane_pliki:
                st.write(f"🔄 Wczytuję: {f_info['name']}...")
                plik_binarny = download_file(f_info['id'])
                # Tutaj w następnym kroku dodamy łączenie tych plików w PDF
            
            st.success("Wszystkie pliki pobrane pomyślnie!")

except Exception as e:
    st.error(f"Błąd: {e}")
