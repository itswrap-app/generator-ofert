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

# --- SYSTEM CZCIONEK (Wersja Pro) ---
def install_fonts():
    src = "fonts"
    dst = os.path.expanduser("~/.local/share/fonts")
    installed = []
    try:
        if os.path.exists(src):
            if not os.path.exists(dst): os.makedirs(dst)
            for f in os.listdir(src):
                if f.lower().endswith((".ttf", ".otf")):
                    target = os.path.join(dst, f)
                    shutil.copy(os.path.join(src, f), target)
                    installed.append(f)
            # Rejestracja czcionek w systemie Linux
            subprocess.run(["fc-cache", "-f", "-v"], capture_output=True)
            return installed
    except: pass
    return installed

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
        # Headless conversion via LibreOffice
        subprocess.run(['libreoffice', '--headless', '--convert-to', 'pdf', '--outdir', os.getcwd(), input_path], check=True, capture_output=True)
        pdf_name = os.path.basename(input_path).replace('.pptx', '.pdf')
        return pdf_name if os.path.exists(pdf_name) else None
    except: return None

# --- APLIKACJA ---
st.set_page_config(page_title="ITS WRAP v4.0", layout="wide")
installed_fonts = install_fonts()

with st.sidebar:
    st.title("⚙️ Panel Techniczny")
    if installed_fonts:
        st.success(f"Wykryto czcionki: {len(installed_fonts)}")
        for f in installed_fonts: st.text(f"• {f}")
    else:
        st.error("Brak folderu 'fonts'!")

st.title("🛡️ Generator Ofert ITS WRAP")

try:
    service, creds = get_service()
    client = gspread.authorize(creds)
    
    # Dane
    url_arkusza = "https://docs.google.com/spreadsheets/d/1iqS6geTNP3Bd_Fj_XdS-wCBrKtnGTMNQZYSso70KIkQ/edit?usp=drive_link"
    sheet = client.open_by_url(url_arkusza).worksheet("Ppf")
    data = sheet.get_all_values()
    df = pd.DataFrame(data[1:], columns=[c.strip() for c in data[0]])

    # Formularz
    c1, c2 = st.columns(2)
    with c1:
        klient = st.text_input("Imię i Nazwisko Klienta")
        model = st.text_input("Model Samochodu")
        nr_o = st.text_input("Numer oferty", value=f"IW/{datetime.now().strftime('%Y/%m/%d')}/01")
    with c2:
        pakiet = st.selectbox("Wybierz pakiet", df['Usługa'].tolist())
        manual_folia = st.text_input("Własny rodzaj folii (opcjonalnie)")
        rabat = st.number_input("Rabat (PLN)", value=0)
        foto = st.file_uploader("Wgraj zdjęcie auta", type=['jpg','png','jpeg'])

    # Pobieranie listy plików z Drive
    FOLDER_ID = "12HRnKn9KrZy_C1BSgv24PGD-Gl8lTRmn"
    results = service.files().list(q=f"'{FOLDER_ID}' in parents and trashed=false", fields="files(id, name)").execute()
    pliki = results.get('files', [])

    if st.button("🚀 GENERUJ FINALNY PDF"):
        with st.spinner("Przetwarzam grafikę i czcionki..."):
            writer = PdfWriter()
            row = df[df['Usługa'] == pakiet].iloc[0]
            
            cena_raw = re.sub(r'[^\d,]', '', row['Kwota sprzedaży']).replace(',', '.')
            cena_num = float(cena_raw) if cena_raw else 0.0

            replacements = {
                "{{KLIENT}}": klient, "{{MODEL_AUTA}}": model,
                "{{RODZAJ_FOLII}}": manual_folia if manual_folia else row['Rodzaj folii'],
                "{{USLUGA_NAZWA}}": pakiet, "{{NR_OFERTY}}": nr_o,
                "{{CENA_KATALOG}}": f"{cena_num:,.2f} zł".replace(',', ' ').replace('.', ','),
                "{{CENA_KONCOWA}}": f"{(cena_num - rabat):,.2f} zł".replace(',', ' ').replace('.', ',')
            }

            # Filtrowanie plików bez duplikatów
            base_names = ['1', '3', '6']
            main_files = {n: next((f for f in pliki if f['name'].startswith(n)), None) for n in base_names}
            dodatki = [f for f in pliki if not f['name'].startswith(('1','3','6')) and 'pptx' in f['name'].lower()]

            kolejnosc = []
            if main_files['1']: kolejnosc.append(main_files['1'])
            kolejnosc.extend(sorted(dodatki, key=lambda x: x['name']))
            if main_files['3']: kolejnosc.append(main_files['3'])
            if main_files['6']: kolejnosc.append(main_files['6'])

            for f_info in kolejnosc:
                prs = Presentation(download_file(service, f_info['id']))
                
                for slide in prs.slides:
                    # Tekst (zachowujemy styl)
                    for shape in slide.shapes:
                        if shape.has_text_frame:
                            for p in shape.text_frame.paragraphs:
                                for run in p.runs:
                                    for k, v in replacements.items():
                                        if k in run.text: run.text = run.text.replace(k, str(v))
                    
                    # ZDJĘCIE (szukanie pancerne)
                    if foto and f_info['name'].startswith('1'):
                        for shape in slide.shapes:
                            # Sprawdzamy wszystko: nazwę, tekst alternatywny, opisy
                            alt_text = ""
                            try: alt_text = shape.non_visual_properties.name + shape._element.xpath('.//p14:nvVisualPropPr/p14:altText')[0]
                            except: pass
                            
                            if "{{FOTO_AUTA}}" in shape.name or "{{FOTO_AUTA}}" in alt_text:
                                slide.shapes.add_picture(foto, shape.left, shape.top, shape.width, shape.height)
                                shape._element.getparent().remove(shape._element)

                tmp_pptx = f"tmp_{f_info['id']}.pptx"
                prs.save(tmp_pptx)
                pdf = pptx_to_pdf(tmp_pptx)
                if pdf:
                    writer.append(pdf)
                    os.remove(tmp_pptx); os.remove(pdf)

            final_pdf = io.BytesIO()
            writer.write(final_pdf); final_pdf.seek(0)
            st.download_button("📥 POBIERZ OFERTĘ PDF", data=final_pdf, file_name=f"Oferta_{model}.pdf")

except Exception as e:
    st.error(f"Coś poszło nie tak: {e}")
