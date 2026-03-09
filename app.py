import streamlit as st
import pandas as pd
from pptx import Presentation
from pptx.util import Inches
import gspread
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload
import io

# --- AUTORYZACJA ---
def get_creds():
    scope = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]
    return Credentials.from_service_account_info(st.secrets["gcp_service_account"], scopes=scope)

def get_drive_service():
    return build('drive', 'v3', credentials=get_creds())

def get_data():
    creds = get_creds()
    client = gspread.authorize(creds)
    url_arkusza = "https://docs.google.com/spreadsheets/d/1iqS6geTNP3Bd_Fj_XdS-wCBrKtnGTMNQZYSso70KIkQ/edit?usp=drive_link" 
    sheet = client.open_by_url(url_arkusza).worksheet("Ppf")
    data = sheet.get_all_values()
    df = pd.DataFrame(data[1:], columns=data[0])
    df.columns = df.columns.str.strip()
    return df

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

# --- FUNKCJE NAPRAWCZE ---

def replace_text_in_slide(slide, replacements):
    for shape in slide.shapes:
        if shape.has_text_frame:
            for paragraph in shape.text_frame.paragraphs:
                for run in paragraph.runs:
                    for key, value in replacements.items():
                        if key in run.text:
                            run.text = run.text.replace(key, str(value))

def move_slide_elements(source_slide, target_slide):
    """Kopiuje elementy ze slajdu źródłowego do docelowego, zachowując obrazy."""
    for shape in source_slide.shapes:
        if shape.shape_type == 13: # To jest obraz (Picture)
            img_stream = io.BytesIO(shape.image.blob)
            target_slide.shapes.add_picture(img_stream, shape.left, shape.top, shape.width, shape.height)
        elif shape.has_text_frame:
            # Tworzymy nowe pole tekstowe i kopiujemy treść
            new_shape = target_slide.shapes.add_textbox(shape.left, shape.top, shape.width, shape.height)
            new_shape.text_frame.text = shape.text_frame.text
        # Można tu dodać inne typy kształtów, jeśli ich używasz (np. linie, prostokąty)

# --- INTERFEJS ---
st.title("🛡️ Generator Ofert ITS WRAP v2.0")

try:
    df = get_data()
    FOLDER_ID = "12HRnKn9KrZy_C1BSgv24PGD-Gl8lTRmn"
    query = f"'{FOLDER_ID}' in parents and mimeType = 'application/vnd.openxmlformats-officedocument.presentationml.presentation' and trashed = false"
    service = get_drive_service()
    results = service.files().list(q=query, fields="files(id, name)").execute()
    pliki_na_dysku = results.get('files', [])

    st.sidebar.header("Składniki")
    wybrane_pliki = []
    for f in sorted(pliki_na_dysku, key=lambda x: x['name']):
        if st.sidebar.checkbox(f"{f['name']}", value=True):
            wybrane_pliki.append(f)

    klient = st.text_input("Klient / Auto")
    pakiet = st.selectbox("Pakiet", df[df.columns[0]].tolist())
    foto_okladka = st.file_uploader("Zdjęcie na okładkę", type=['jpg', 'png'])

    if st.button("🚀 GENERUJ OFERTĘ"):
        with st.spinner("Składam ofertę..."):
            # Dane do podmiany
            wiersz = df[df[df.columns[0]] == pakiet]
            cena_kat = wiersz[df.columns[1]].values[0]
            
            replacements = {
                "{{USLUGA_NAZWA}}": pakiet,
                "{{CENA_KATALOG}}": f"{cena_kat}",
                "{{CENA_KONCOWA}}": f"{cena_kat}" # Uproszczone dla testu
            }

            # 1. Tworzymy całkiem nową, czystą prezentację
            final_prs = Presentation()
            # Ustawiamy rozmiar slajdu na 16:9 (standardowy)
            final_prs.slide_width = Inches(13.333)
            final_prs.slide_height = Inches(7.5)

            for f_info in wybrane_pliki:
                stream = download_file(f_info['id'])
                source_prs = Presentation(stream)
                
                for slide in source_prs.slides:
                    # Dodajemy nowy slajd do naszej głównej prezentacji
                    new_slide = final_prs.slides.add_slide(final_prs.slide_layouts[6]) # Pusty layout
                    
                    # Kopiujemy elementy i podmieniamy teksty
                    replace_text_in_slide(slide, replacements)
                    move_slide_elements(slide, new_slide)
                    
                    # Jeśli to okładka i wgrano zdjęcie
                    if "okładka" in f_info['name'].lower() and foto_okladka:
                        new_slide.shapes.add_picture(foto_okladka, Inches(1), Inches(1), width=Inches(5))

            output = io.BytesIO()
            final_prs.save(output)
            output.seek(0)
            
            st.success("Gotowe!")
            st.download_button("📥 POBIERZ PPTX", data=output, file_name=f"Oferta_{klient}.pptx")

except Exception as e:
    st.error(f"Błąd: {e}")
