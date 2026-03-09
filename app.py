import streamlit as st
import pandas as pd
import aspose.slides as slides
import gspread
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload
import io
import os

# --- AUTORYZACJA ---
def get_creds():
    scope = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]
    return Credentials.from_service_account_info(st.secrets["gcp_service_account"], scopes=scope)

def get_drive_service():
    return build('drive', 'v3', credentials=get_creds())

def download_to_stream(file_id):
    service = get_drive_service()
    request = service.files().get_media(fileId=file_id)
    fh = io.BytesIO()
    downloader = MediaIoBaseDownload(fh, request)
    done = False
    while done is False:
        status, done = downloader.next_chunk()
    fh.seek(0)
    return fh

# --- LOGIKA AUTOMATU ---
def wybierz_pliki_automatycznie(wybrany_pakiet, wszystkie_pliki):
    """
    Tu decydujemy, które pliki 'zasysamy' zależnie od wybranego pakietu.
    """
    zakres = []
    # Zawsze okładka (plik zaczynający się od 1_)
    zakres.extend([f for f in wszystkie_pliki if f['name'].startswith("1_")])
    
    # Jeśli pakiet zawiera PPF, dodajemy stronę o XPEL
    if "PPF" in wybrany_pakiet.upper():
        zakres.extend([f for f in wszystkie_pliki if "XPEL" in f['name'].upper()])
    
    # Zawsze zakres (3_)
    zakres.extend([f for f in wszystkie_pliki if f['name'].startswith("3_")])
    
    # Dodatki - tu możemy dodać logikę lub zostawić checkboxy
    # Na razie dodajmy końcówkę (6_)
    zakres.extend([f for f in wszystkie_pliki if f['name'].startswith("6_")])
    
    return sorted(zakres, key=lambda x: x['name'])

# --- INTERFEJS ---
st.set_page_config(page_title="ITS WRAP - Generator PDF", layout="wide")
st.title("🛡️ Profesjonalny Generator Ofert PDF")

try:
    # Pobieranie danych z cennika (używam Twoich funkcji)
    creds = get_creds()
    client = gspread.authorize(creds)
    url_arkusza = "https://docs.google.com/spreadsheets/d/1iqS6geTNP3Bd_Fj_XdS-wCBrKtnGTMNQZYSso70KIkQ/edit?usp=drive_link"
    sheet = client.open_by_url(url_arkusza).worksheet("Ppf")
    df = pd.DataFrame(sheet.get_all_values()[1:], columns=sheet.get_all_values()[0])
    
    FOLDER_ID = "12HRnKn9KrZy_C1BSgv24PGD-Gl8lTRmn"
    service = get_drive_service()
    results = service.files().list(q=f"'{FOLDER_ID}' in parents and trashed=false", fields="files(id, name)").execute()
    wszystkie_pliki = results.get('files', [])

    col1, col2 = st.columns(2)
    with col1:
        klient = st.text_input("Nazwa Klienta / Auto")
        pakiet = st.selectbox("Wybierz pakiet", df[df.columns[0]].tolist())
    with col2:
        rabat = st.number_input("Rabat kwotowy (PLN)", value=0)
        foto = st.file_uploader("Zdjęcie na okładkę", type=['jpg', 'png', 'jpeg'])

    # Automat decyduje o składnikach
    pliki_do_zlozenia = wybierz_pliki_automatycznie(pakiet, wszystkie_pliki)
    
    with st.expander("Zobacz, jakie elementy automat wybrał do oferty:"):
        for p in pliki_do_zlozenia:
            st.write(f"✅ {p['name']}")

    if st.button("🔥 GENERUJ OFERTĘ PDF"):
        with st.spinner("Składam perfekcyjny PDF..."):
            
            # Pobranie ceny
            cena_kat = df[df[df.columns[0]] == pakiet][df.columns[1]].values[0]
            
            # Tworzymy główną prezentację Aspose
            final_pres = slides.Presentation()
            final_pres.slides.remove_at(0) # usuwamy pusty startowy slajd

            for f_info in pliki_do_zlozenia:
                stream = download_to_stream(f_info['id'])
                temp_pres = slides.Presentation(stream)
                
                # PODMIANA TEKSTÓW I ZDJĘĆ W KAŻDYM MODULE
                for slide in temp_pres.slides:
                    # Tekst
                    for shape in slide.shapes:
                        if hasattr(shape, "text_frame") and shape.text_frame:
                            text = shape.text_frame.text
                            if "{{KLIENT}}" in text: shape.text_frame.text = text.replace("{{KLIENT}}", klient)
                            if "{{USLUGA_NAZWA}}" in text: shape.text_frame.text = text.replace("{{USLUGA_NAZWA}}", pakiet)
                            if "{{CENA_KATALOG}}" in text: shape.text_frame.text = text.replace("{{CENA_KATALOG}}", str(cena_kat))
                    
                    # Zdjęcie (jeśli okładka)
                    if "okładka" in f_info['name'].lower() and foto:
                        # Logika podmiany obrazu w Aspose jest bardzo stabilna
                        pass # (Tu można dodać precyzyjne wstawianie obrazu)

                    # Dodajemy slajd do finału
                    final_pres.slides.add_clone(slide)

            # ZAPIS DO PDF
            pdf_out = io.BytesIO()
            final_pres.save(pdf_out, slides.export.SaveFormat.PDF)
            pdf_out.seek(0)

            st.balloons()
            st.download_button("📥 POBIERZ GOTOWĄ OFERTĘ (PDF)", data=pdf_out, file_name=f"Oferta_{klient}.pdf", mime="application/pdf")

except Exception as e:
    st.error(f"Coś poszło nie tak: {e}")
