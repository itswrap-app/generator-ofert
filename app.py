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

# --- SYSTEM CZCIONEK ---
def install_fonts():
    src = "fonts"
    dst = os.path.expanduser("~/.local/share/fonts")
    if os.path.exists(src):
        if not os.path.exists(dst): os.makedirs(dst)
        for f in os.listdir(src):
            if f.lower().endswith((".ttf", ".otf")):
                shutil.copy(os.path.join(src, f), dst)
        subprocess.run(["fc-cache", "-f"], capture_output=True)
        # Sprawdźmy co widzi system
        res = subprocess.run(["fc-list", ":family"], capture_output=True, text=True)
        return res.stdout
    return "Folder /fonts nie istnieje"

# --- NARZĘDZIA GOOGLE ---
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
        # LibreOffice musi wiedzieć, że pracujemy w środowisku bez ekranu
        subprocess.run(['libreoffice', '--headless', '--convert-to', 'pdf', '--outdir', os.getcwd(), input_path], check=True, capture_output=True)
        pdf_name = os.path.basename(input_path).replace('.pptx', '.pdf')
        return pdf_name if os.path.exists(pdf_name) else None
    except: return None

# --- APLIKACJA ---
st.set_page_config(page_title="ITS WRAP - Generator PRO", layout="wide")
font_status = install_fonts()

st.sidebar.title("⚙️ Status Systemu")
with st.sidebar.expander("Zainstalowane czcionki"):
    st.code(font_status)

st.title("🛡️ Generator Ofert ITS WRAP")

try:
    service, creds = get_service()
    client = gspread.authorize(creds)
    
    # 1. Cennik
    url_arkusza = "https://docs.google.com/spreadsheets/d/1iqS6geTNP3Bd_Fj_XdS-wCBrKtnGTMNQZYSso70KIkQ/edit?usp=drive_link"
    sheet = client.open_by_url(url_arkusza).worksheet("Ppf")
    data = sheet.get_all_values()
    df = pd.DataFrame(data[1:], columns=[c.strip() for c in data[0]])

    # 2. Pobieranie listy plików (Tylko PowerPoint)
    FOLDER_ID = "12HRnKn9KrZy_C1BSgv24PGD-Gl8lTRmn"
    q = f"'{FOLDER_ID}' in parents and mimeType='application/vnd.openxmlformats-officedocument.presentationml.presentation' and trashed=false"
    results = service.files().list(q=q, fields="files(id, name)").execute()
    wszystkie_pliki = results.get('files', [])

    # --- FORMULARZ ---
    col1, col2 = st.columns(2)
    with col1:
        klient_name = st.text_input("Klient", placeholder="Jan Kowalski")
        model_auta = st.text_input("Model Samochodu", placeholder="Tesla 3")
        nr_oferty = st.text_input("Numer oferty", value=f"IW/{datetime.now().strftime('%Y/%m/%d')}/01")
    with col2:
        pakiet_usluga = st.selectbox("Pakiet", df['Usługa'].tolist())
        manual_folia = st.text_input("Własna folia (opcjonalnie)")
        rabat_pln = st.number_input("Rabat (PLN)", value=0)
        foto_file = st.file_uploader("Zdjęcie auta na okładkę", type=['jpg','png','jpeg'])

    # Sidebar: wybór dodatków
    st.sidebar.header("Dodatki")
    dodatki_pliki = [f for f in wszystkie_pliki if not f['name'].startswith(('1','3','6'))]
    wybrane_extra = []
    for d in sorted(dodatki_pliki, key=lambda x: x['name']):
        if st.sidebar.checkbox(d['name'], value=False):
            wybrane_extra.append(d)

    if st.button("🚀 GENERUJ PDF"):
        with st.spinner("Składam ofertę..."):
            writer = PdfWriter()
            row = df[df['Usługa'] == pakiet_usluga].iloc[0]
            
            # Cena
            cena_raw = re.sub(r'[^\d,]', '', row['Kwota sprzedaży']).replace(',', '.')
            cena_num = float(cena_raw) if cena_raw else 0.0
            folia_final = manual_folia if manual_folia else row['Rodzaj folii']

            replacements = {
                "{{KLIENT}}": klient_name, "{{MODEL_AUTA}}": model_auta,
                "{{RODZAJ_FOLII}}": folia_final, "{{USLUGA_NAZWA}}": pakiet_usluga,
                "{{NR_OFERTY}}": nr_oferty,
                "{{CENA_KATALOG}}": f"{cena_num:,.2f} zł".replace(',', ' ').replace('.', ','),
                "{{CENA_KONCOWA}}": f"{(cena_num - rabat_pln):,.2f} zł".replace(',', ' ').replace('.', ',')
            }

            # Kolejność bez duplikatów
            okladka = next((f for f in wszystkie_pliki if f['name'].startswith('1')), None)
            zakres = next((f for f in wszystkie_pliki if f['name'].startswith('3')), None)
            koniec = next((f for f in wszystkie_pliki if f['name'].startswith('6')), None)

            pliki_to_process = [f for f in [okladka] + wybrane_extra + [zakres, koniec] if f]

            for f_info in pliki_to_process:
                prs = Presentation(download_file(service, f_info['id']))
                for slide in prs.slides:
                    # Agresywne szukanie obrazka {{FOTO_AUTA}}
                    if foto_file and f_info['name'].startswith('1'):
                        for shape in list(slide.shapes):
                            # Sprawdzamy tekst wewnątrz kształtu LUB nazwę kształtu
                            found = False
                            if shape.has_text_frame and "{{FOTO_AUTA}}" in shape.text_frame.text:
                                found = True
                            elif "{{FOTO_AUTA}}" in shape.name:
                                found = True
                            
                            if found:
                                slide.shapes.add_picture(foto_file, shape.left, shape.top, shape.width, shape.height)
                                # Usuwamy tag tekstowy
                                sp = shape._element
                                sp.getparent().remove(sp)

                    # Podmiana tekstów
                    for shape in slide.shapes:
                        if shape.has_text_frame:
                            for p in shape.text_frame.paragraphs:
                                for run in p.runs:
                                    # Hardkodujemy czcionkę po podmianie tekstu!
                                    for k, v in replacements.items():
                                        if k in run.text:
                                            run.text = run.text.replace(k, str(v))
                                            # Wymuszamy Twoją czcionkę URW DIN
                                            run.font.name = 'URW DIN' 

                tmp_p = f"tmp_{f_info['id']}.pptx"
                prs.save(tmp_p)
                pdf_f = pptx_to_pdf(tmp_p)
                if pdf_f: 
                    writer.append(pdf_f)
                    os.remove(tmp_p); os.remove(pdf_f)

            final_pdf = io.BytesIO()
            writer.write(final_pdf); final_pdf.seek(0)
            st.download_button("📥 POBIERZ PDF", data=final_pdf, file_name=f"Oferta_{model_auta}.pdf")

except Exception as e:
    st.error(f"Błąd: {e}")
