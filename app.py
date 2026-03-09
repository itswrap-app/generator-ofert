import streamlit as st
import pandas as pd
from pptx import Presentation
import gspread
from google.oauth2.service_account import Credentials
import traceback

st.set_page_config(page_title="ITS WRAP - Generator Ofert", layout="centered")

def get_data():
    scope = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]
    creds = Credentials.from_service_account_info(st.secrets["gcp_service_account"], scopes=scope)
    client = gspread.authorize(creds)
    
    # TUTAJ WKLEJ SWÓJ LINK DO ARKUSZA (Zostaw cudzysłowy!)
    url_arkusza = "https://docs.google.com/spreadsheets/d/1iqS6geTNP3Bd_Fj_XdS-wCBrKtnGTMNQZYSso70KIkQ/edit?usp=drive_link"
    
    sheet = client.open_by_url(url_arkusza).worksheet("Ppf")
    return pd.DataFrame(sheet.get_all_records())

st.title("🚀 Generator Ofert ITS WRAP")

try:
    df = get_data()
    st.success("Połączono z cennikiem!")
    
    klient = st.text_input("Nazwa Klienta / Model auta")
    pakiet = st.selectbox("Wybierz pakiet z cennika", df['Usługa'].tolist())
    
    if st.button("Generuj Ofertę (Test)"):
        wybrana_cena = df[df['Usługa'] == pakiet]['Kwota sprzedaży'].values[0]
        st.write(f"Wybrałeś: {pakiet} za {wybrana_cena} zł.")
        st.info("Kolejny krok: Składanie plików PPTX.")

except Exception as e:
    st.error(f"Wystąpił błąd: {e}")
    st.code(traceback.format_exc()) # To pokaże nam dokładne źródło problemu
