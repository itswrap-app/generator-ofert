import streamlit as st
import pandas as pd
from pptx import Presentation
import gspread
from google.oauth2.service_account import Credentials
import io

# Konfiguracja strony
st.set_page_config(page_title="ITS WRAP - Generator Ofert", layout="centered")

# Funkcja łączenia z Google Sheets
def get_data():
    scope = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]
    # Tu pobierzemy klucze z "Secrets" Streamlita (później)
    creds = Credentials.from_service_account_info(st.secrets["gcp_service_account"], scopes=scope)
    client = gspread.authorize(creds)
    # Wpisz dokładnie nazwę swojego arkusza:
    sheet = client.open("Ceny i marża produktów Its Wrap 2025 (2)").worksheet("Ppf")
    return pd.DataFrame(sheet.get_all_records())

st.title("🚀 Generator Ofert ITS WRAP")

try:
    df = get_data()
    st.success("Połączono z cennikiem!")
    
    # Formularz dla handlowca
    klient = st.text_input("Nazwa Klienta / Model auta")
    pakiet = st.selectbox("Wybierz pakiet z cennika", df['Usługa'].tolist())
    
    # Przycisk generowania (na razie jako test)
    if st.button("Generuj Ofertę (Test)"):
        wybrana_cena = df[df['Usługa'] == pakiet]['Kwota sprzedaży'].values[0]
        st.write(f"Wybrałeś: {pakiet} za {wybrana_cena} zł.")
        st.info("Kolejny krok: Składanie plików PPTX z Twojego Google Drive.")

except Exception as e:
    st.error(f"Czekam na konfigurację klucza JSON... {e}")
