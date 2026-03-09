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
    font_src = "fonts"
    font_dst = os.path.expanduser("~/.local/share/fonts")
    try:
        if os.path.exists(font_src):
            if not os.path.exists(font_dst): os.makedirs(font_dst)
            for f in os.listdir(font_src):
                if f.lower().endswith((".ttf", ".otf")):
                    shutil.copy(os.path.join(font_src, f), font_dst)
            subprocess.run(["fc-cache", "-f"], capture_output=True)
            return True
    except: pass
    return False

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
        subprocess.run(['libreoffice', '--headless', '--convert-to', 'pdf', '--outdir', os.getcwd(), input_path], check=True, capture_output=True)
        pdf_name = os.path.basename(input_path).replace('.pptx', '.pdf')
        return pdf_name if os.path.exists(pdf_name) else None
    except: return None

# --- APLIKACJA ---
st.set_page_config(page_title="ITS WRAP - PDF PRO", layout="wide")
install_fonts()
st.title("🛡️ Generator Ofert ITS WRAP v3.0")

try:
    service, creds = get_service()
    client = gspread.authorize(creds)
    
    # 1. Pobieranie Cennika
    url_arkusza = "https://docs.google.com/spreadsheets/d/1iqS6geTNP3Bd_Fj_XdS-wCBrKtnGTMNQZYSso70KIkQ/edit?usp=drive_link"
    sheet = client.open_by_url(url_arkusza).worksheet("Ppf")
    data = sheet.get_all_values()
    df = pd.DataFrame(data[1:], columns=[c.strip() for c in data[0]])

    # 2. Pobieranie Plików PPTX
    FOLDER_ID = "12HRnKn9KrZy_C1BSgv24PGD-Gl8lTRmn"
    q = f"'{FOLDER_ID}' in parents and mimeType='application/vnd.openxmlformats-officedocument.presentationml.presentation' and trashed=false"
    results = service.files().list(q=q, fields="files(id, name)").execute()
    wszystkie_pliki = results.get('files', [])

    # --- FORMULARZ ---
    col1, col2 = st.columns(2)
    with col1:
        klient_name = st.text_input("Imię i Nazwisko Klienta")
        model_auta = st.text_input("Model Samochodu (np. BMW M3)")
        nr_oferty = st.text_input("Numer oferty", value=f"IW/{datetime.now().strftime('%Y/%m/%d')}/01")
    with col2:
        pakiet_usluga = st.selectbox("Wybierz pakiet z cennika", df['Usługa'].tolist())
        manual_folia = st.text_input("Rodzaj Folii (zostaw puste, by wziąć z cennika)")
        rabat_pln = st.number_input("Rabat (PLN)", value=0, step=100)
        foto_file = st.file_uploader("Zdjęcie auta na okładkę", type=['jpg','png','jpeg'])

    # Sidebar: wybór dodatków (pliki 2_, 4_, 5_)
    st.sidebar.header("Wybierz dodatkowe strony")
    dodatki_pliki = [f for f in wszystkie_pliki if not f['name'].startswith(('1','3','6'))]
    wybrane_extra = []
    for d in sorted(dodatki_pliki, key=lambda x: x['name']):
        if st.sidebar.checkbox(d['name'], value=False):
            wybrane_extra.append(d)

    if st.button("🚀 GENERUJ PDF"):
        with st.spinner("Składam ofertę bez duplikatów..."):
            writer = PdfWriter()
            
            # Pobranie danych z cennika
            row = df[df['Usługa'] == pakiet_usluga].iloc[0]
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

            # Precyzyjna kolejność (tylko UNIKALNE pliki)
            okladka = next((f for f in wszystkie_pliki if f['name'].startswith('1')), None)
            zakres = next((f for f in wszystkie_pliki if f['name'].startswith('3')), None)
            koniec = next((f for f in wszystkie_pliki if f['name'].startswith('6')), None)

            pliki_to_process = []
            if okladka: pliki_to_process.append(okladka)
            pliki_to_process.extend(wybrane_extra)
            if zakres: pliki_to_process.append(zakres)
            if koniec: pliki_to_process.append(koniec)

            for f_info in pliki_to_process:
                prs = Presentation(download_file(service, f_info['id']))
                for slide in prs.slides:
                    # Tekst
                    for shape in slide.shapes:
                        if shape.has_text_frame:
                            for p in shape.text_frame.paragraphs:
                                for run in p.runs:
                                    for k, v in replacements.items():
                                        if k in run.text: run.text = run.text.replace(k, str(v))
                    
                    # Zdjęcie (tylko na okładce)
                    if foto_file and f_info['name'].startswith('1'):
                        for shape in slide.shapes:
                            # Szukamy po nazwie kształtu w Selection Pane
                            if "{{FOTO_AUTA}}" in shape.name:
                                slide.shapes.add_picture(foto_file, shape.left, shape.top, shape.width, shape.height)
                                shape._element.getparent().remove(shape._element)

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
