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
        res = subprocess.run(["fc-list", ":family"], capture_output=True, text=True)
        return res.stdout
    return "Brak folderu fonts"

# --- NARZĘDZIA ---
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
st.set_page_config(page_title="ITS WRAP v4.2", layout="wide")
font_status = install_fonts()

st.title("🛡️ Generator Ofert ITS WRAP")

try:
    service, creds = get_service()
    client = gspread.authorize(creds)
    
    # Cennik
    url_arkusza = "https://docs.google.com/spreadsheets/d/1iqS6geTNP3Bd_Fj_XdS-wCBrKtnGTMNQZYSso70KIkQ/edit?usp=drive_link"
    sheet = client.open_by_url(url_arkusza).worksheet("Ppf")
    data = sheet.get_all_values()
    df = pd.DataFrame(data[1:], columns=[c.strip() for c in data[0]])

    # Pliki Drive
    FOLDER_ID = "12HRnKn9KrZy_C1BSgv24PGD-Gl8lTRmn"
    q = f"'{FOLDER_ID}' in parents and mimeType='application/vnd.openxmlformats-officedocument.presentationml.presentation' and trashed=false"
    results = service.files().list(q=q, fields="files(id, name)").execute()
    wszystkie_pliki = results.get('files', [])

    # Formularz
    c1, c2 = st.columns(2)
    with c1:
        klient = st.text_input("Klient")
        model = st.text_input("Model Samochodu")
        nr_o = st.text_input("Numer oferty", value=f"IW/{datetime.now().strftime('%Y/%m/%d')}/01")
    with c2:
        pakiet = st.selectbox("Wybierz pakiet", df['Usługa'].tolist())
        rabat = st.number_input("Rabat (PLN)", value=0)
        foto = st.file_uploader("Zdjęcie auta na okładkę", type=['jpg','png','jpeg'])

    # Dodatki w sidebarze (tylko pliki 4_ i 5_)
    st.sidebar.header("Dodatki (Szyba/Powłoka)")
    dodatki_pliki = [f for f in wszystkie_pliki if f['name'].startswith(('4','5'))]
    wybrane_extra = []
    for d in sorted(dodatki_pliki, key=lambda x: x['name']):
        if st.sidebar.checkbox(d['name'], value=False):
            wybrane_extra.append(d)

    if st.button("🚀 GENERUJ PDF"):
        with st.spinner("Składam ofertę..."):
            writer = PdfWriter()
            row = df[df['Usługa'] == pakiet].iloc[0]
            cena_raw = re.sub(r'[^\d,]', '', row['Kwota sprzedaży']).replace(',', '.')
            cena_num = float(cena_raw) if cena_raw else 0.0

            replacements = {
                "{{KLIENT}}": klient, "{{MODEL_AUTA}}": model,
                "{{RODZAJ_FOLII}}": row['Rodzaj folii'],
                "{{USLUGA_NAZWA}}": pakiet, "{{NR_OFERTY}}": nr_o,
                "{{CENA_KATALOG}}": f"{cena_num:,.2f} zł".replace(',', ' ').replace('.', ','),
                "{{CENA_KONCOWA}}": f"{(cena_num - rabat):,.2f} zł".replace(',', ' ').replace('.', ',')
            }

            # --- KOLEJNOŚĆ AUTOMATYCZNA ---
            okladka_f = next((f for f in wszystkie_pliki if f['name'].startswith('1')), None)
            produkt_f = next((f for f in wszystkie_pliki if f['name'].startswith('2')), None) # XPEL Ultimate
            zakres_f = next((f for f in wszystkie_pliki if f['name'].startswith('3')), None) # Zakres prac
            koniec_f = next((f for f in wszystkie_pliki if f['name'].startswith('6')), None) # Stopka

            seq = [okladka_f, produkt_f] + wybrane_extra + [zakres_f, koniec_f]
            seq = [f for f in seq if f] # Usuwamy puste

            for f_info in seq:
                prs = Presentation(download_file(service, f_info['id']))
                for slide in prs.slides:
                    # 1. ZDJĘCIE (Wysyłamy na spód)
                    if foto and f_info['name'].startswith('1'):
                        for shape in list(slide.shapes):
                            if shape.has_text_frame and "{{FOTO_AUTA}}" in shape.text_frame.text:
                                left, top, width, height = shape.left, shape.top, shape.width, shape.height
                                new_pic = slide.shapes.add_picture(foto, left, top, width, height)
                                
                                # --- MAGIA: Przesunięcie na spód ---
                                slide.shapes._spTree.remove(new_pic._element)
                                slide.shapes._spTree.insert(2, new_pic._element) # 2 to pierwsza warstwa pod tekstem
                                
                                sp = shape._element
                                sp.getparent().remove(sp)

                    # 2. TEKSTY
                    for shape in slide.shapes:
                        if shape.has_text_frame:
                            for p in shape.text_frame.paragraphs:
                                for run in p.runs:
                                    for k, v in replacements.items():
                                        if k in run.text:
                                            run.text = run.text.replace(k, str(v))
                                            run.font.name = 'URW DIN'

                tmp_p = f"tmp_{f_info['id']}.pptx"
                prs.save(tmp_p)
                pdf = pptx_to_pdf(tmp_p)
                if pdf:
                    writer.append(pdf)
                    os.remove(tmp_p); os.remove(pdf)

            final = io.BytesIO()
            writer.write(final); final.seek(0)
            st.download_button("📥 POBIERZ PDF", data=final, file_name=f"Oferta_{model}.pdf")

except Exception as e:
    st.error(f"Błąd: {e}")
