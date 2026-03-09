import streamlit as st
import pandas as pd
from pptx import Presentation
from pptx.oxml import parse_xml
import gspread
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload
import io

# --- AUTORYZACJA I DRIVE ---
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

# --- OBRÓBKA SLAJDÓW ---

def replace_text_in_prs(prs, replacements):
    for slide in prs.slides:
        for shape in slide.shapes:
            if shape.has_text_frame:
                for paragraph in shape.text_frame.paragraphs:
                    for run in paragraph.runs:
                        for key, value in replacements.items():
                            if key in run.text:
                                run.text = run.text.replace(key, str(value))

def replace_image_in_slide(slide, placeholder_alt_text, new_image_stream):
    for shape in slide.shapes:
        # Próba znalezienia zdjęcia po tekście alternatywnym
        try:
            alt_text = shape.non_visual_properties.name
            if not alt_text: # Jeśli puste, szukamy głębiej w XML
                alt_text = shape._element.xpath('.//p14:nvVisualPropPr/p14:altText')[0]
        except:
            alt_text = ""

        if placeholder_alt_text in alt_text or shape.name == placeholder_alt_text:
            left, top, width, height = shape.left, shape.top, shape.width, shape.height
            # Usuwamy stary kształt i wstawiamy zdjęcie
            spTree = shape._element.getparent()
            spTree.remove(shape._element)
            slide.shapes.add_picture(new_image_stream, left, top, width, height)

# --- INTERFEJS ---
st.set_page_config(page_title="ITS WRAP - Generator LEGO", layout="wide")
st.title("🛡️ Generator Ofert ITS WRAP")

try:
    df = get_data()
    FOLDER_ID = "12HRnKn9KrZy_C1BSgv24PGD-Gl8lTRmn"
    
    query = f"'{FOLDER_ID}' in parents and mimeType = 'application/vnd.openxmlformats-officedocument.presentationml.presentation' and trashed = false"
    service = get_drive_service()
    results = service.files().list(q=query, fields="files(id, name)").execute()
    pliki_na_dysku = results.get('files', [])

    st.sidebar.header("Wybierz klocki")
    wybrane_pliki = []
    for f in sorted(pliki_na_dysku, key=lambda x: x['name']):
        if st.sidebar.checkbox(f"{f['name']}", value=True):
            wybrane_pliki.append(f)

    col1, col2 = st.columns(2)
    with col1:
        klient = st.text_input("Nazwa Klienta / Auto")
        pakiet = st.selectbox("Pakiet z cennika", df[df.columns[0]].tolist())
    with col2:
        foto = st.file_uploader("Zdjęcie na okładkę", type=['jpg', 'png', 'jpeg'])
        rabat = st.number_input("Rabat kwotowy (PLN)", value=0)

    if st.button("🚀 GENERUJ GOTOWĄ OFERTĘ"):
        if not wybrane_pliki:
            st.warning("Zaznacz pliki po lewej!")
        else:
            with st.spinner("Składam ofertę..."):
                # Dane do podmiany
                wiersz_ceny = df[df[df.columns[0]] == pakiet]
                cena_kat = wiersz_ceny[df.columns[1]].values[0]
                cena_num = float(''.join(filter(str.isdigit, str(cena_kat).replace(',','.'))))
                cena_koncowa = cena_num - rabat
                
                replacements = {
                    "{{USLUGA_NAZWA}}": pakiet,
                    "{{CENA_KATALOG}}": f"{cena_kat}",
                    "{{CENA_KONCOWA}}": f"{cena_koncowa:,.2f} zł".replace(',', ' ').replace('.', ',')
                }

                # Startujemy od pierwszego pliku
                base_stream = download_file(wybrane_pliki[0]['id'])
                final_prs = Presentation(base_stream)
                replace_text_in_prs(final_prs, replacements)
                if foto:
                    for slide in final_prs.slides:
                        replace_image_in_slide(slide, "{{FOTO_AUTA}}", foto)

                # Doklejamy resztę
                for f_info in wybrane_pliki[1:]:
                    sub_stream = download_file(f_info['id'])
                    sub_prs = Presentation(sub_stream)
                    replace_text_in_prs(sub_prs, replacements)
                    
                    for slide in sub_prs.slides:
                        # Dodajemy pusty slajd o tym samym layoutcie
                        blank_slide_layout = final_prs.slide_layouts[6] # Zazwyczaj pusty
                        new_slide = final_prs.slides.add_slide(blank_slide_layout)
                        
                        # Kopiujemy kształty bezpieczną metodą XML
                        for shape in slide.shapes:
                            # Używamy parse_xml, aby przenieść definicje namespaces
                            new_shape_xml = parse_xml(shape.element.xml)
                            new_slide.shapes._spTree.append(new_shape_xml)

                output = io.BytesIO()
                final_prs.save(output)
                output.seek(0)

                st.balloons()
                st.download_button(label="📥 POBIERZ OFERTĘ (PPTX)", data=output, file_name=f"Oferta_{klient}.pptx")

except Exception as e:
    st.error(f"Wystąpił problem: {e}")
