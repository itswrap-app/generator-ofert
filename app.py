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
    url_arkusza = "https://docs.google.com/spreadsheets/d/1iqS6geTNP3Bd_Fj_XdS-wCBrKtnGTMNQZYSso70KIkQ/edit?usp=sharing" # PAMIĘTAJ O ZMIANIE TEGO LINKU NA SWÓJ
    
    sheet = client.open_by_url(url_arkusza).worksheet("Ppf")
    
    # Pobieramy wszystko jako surową tabelę
    data = sheet.get_all_values()
    
    # Tworzymy czysty zbiór danych: bierzemy pierwszy wiersz jako nagłówki i resztę jako dane
    df = pd.DataFrame(data[1:], columns=data[0])
    
    # Standardyzujemy nazwy kolumn, by uniknąć problemu ze spacjami
    df.columns = df.columns.str.strip()
    
    return df

st.title("🚀 Generator Ofert ITS WRAP")

try:
    df = get_data()
    st.success("Połączono z cennikiem!")
    
    # Bierzemy dane niezależnie od niewidocznych spacji
    nazwa_kolumny_uslugi = df.columns[0] # Pierwsza kolumna
    nazwa_kolumny_ceny = df.columns[1] # Druga kolumna
    
    klient = st.text_input("Nazwa Klienta / Model auta")
    pakiet = st.selectbox("Wybierz pakiet z cennika", df[nazwa_kolumny_uslugi].tolist())
    
    if st.button("Generuj Ofertę (Test)"):
        wybrana_cena = df[df[nazwa_kolumny_uslugi] == pakiet][nazwa_kolumny_ceny].values[0]
        st.write(f"Wybrałeś: **{pakiet}**")
        st.write(f"Cena: **{wybrana_cena}**")
        st.info("Kolejny krok: Składanie plików PPTX.")

except Exception as e:
    st.error(f"Wystąpił błąd: {e}")
    st.code(traceback.format_exc())
