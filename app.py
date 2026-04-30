import streamlit as st
import pandas as pd
from pptx import Presentation
import gspread
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload, MediaIoBaseUpload
import io, os, subprocess, re, shutil, requests, base64
from pypdf import PdfWriter
from datetime import datetime
from PIL import Image
import random

# --- NOWY LINK DO CENNIKA v3 ---
LINK_DO_ARKUSZA = "https://docs.google.com/spreadsheets/d/1USF81hOinAP_vvz1QZuoNyRCT1ezJcXTDDB6RjuYtrY/edit"
PARENT_FOLDER_ID = "12HRnKn9KrZy_C1BSgv24PGD-Gl8lTRmn"

# --- BAZA OPIEKUNÓW / HANDLOWCÓW ---
HANDLOWCY = {
    "Adam Trepka": {
        "stanowisko": "CEO It`s Wrap",
        "telefon": "+48 111 222 333",
        "email": "adam@itswrap.pl"
    },
    "Jan Kowalski": {
        "stanowisko": "Specjalista ds. Detailingu",
        "telefon": "+48 444 555 666",
        "email": "jan@itswrap.pl"
    }
}

# --- BAZA AUT I SEGMENTÓW ---
CAR_DATABASE = {
    "Audi": {"A3": ["Hatchback", "Sedan"], "A4": ["Sedan", "Kombi"], "A6": ["Sedan", "Kombi"], "Q3": ["SUV"], "Q5": ["SUV"], "Q8": ["SUV"], "e-tron GT": ["Sedan"], "RS6": ["Kombi"]},
    "BMW": {"Seria 3": ["Sedan", "Kombi"], "Seria 4": ["Coupe", "Gran Coupe"], "Seria 5": ["Sedan", "Kombi"], "Seria 7": ["Sedan"], "X3": ["SUV"], "X5": ["SUV"], "M3": ["Sedan", "Kombi"], "M4": ["Coupe"]},
    "BYD": {"Seal": ["Sedan"], "Atto 3": ["SUV"], "Han": ["Sedan"], "Dolphin": ["Hatchback"]},
    "Ford": {"Focus": ["Hatchback", "Kombi"], "Mustang": ["Coupe", "Cabriolet"], "Mustang Mach-E": ["SUV"], "Puma": ["Crossover"]},
    "Hyundai": {"Tucson": ["SUV"], "Ioniq 5": ["Hatchback/Crossover"], "Ioniq 6": ["Sedan"], "i30": ["Hatchback", "Kombi"], "Kona": ["Crossover"]},
    "Kia": {"EV6": ["Crossover"], "Sportage": ["SUV"], "Ceed": ["Hatchback", "Kombi"], "Stinger": ["Liftback"], "Sorento": ["SUV"]},
    "Lexus": {"NX": ["SUV"], "RX": ["SUV"], "ES": ["Sedan"], "LC": ["Coupe"]},
    "Mercedes-Benz": {"Klasa C": ["Sedan", "Kombi"], "Klasa E": ["Sedan", "Kombi"], "GLC": ["SUV", "Coupe"], "GLE": ["SUV", "Coupe"], "Klasa G": ["SUV"], "AMG GT": ["Coupe"]},
    "MG": {"MG4": ["Hatchback"], "HS": ["SUV"], "ZS": ["SUV"], "Cyberster": ["Roadster"]},
    "NIO": {"ET7": ["Sedan"], "ET5": ["Sedan"], "EL7": ["SUV"]},
    "Porsche": {"911 (992)": ["Coupe", "Cabriolet"], "Taycan": ["Sedan", "Cross Turismo"], "Macan": ["SUV"], "Panamera": ["Sedan"], "Cayenne": ["SUV", "Coupe"]},
    "Renault": {"Scenic E-Tech": ["Crossover"], "Megane E-Tech": ["Hatchback"], "Austral": ["SUV"], "Clio": ["Hatchback"], "Captur": ["Crossover"]},
    "Skoda": {"Octavia": ["Liftback", "Kombi"], "Superb": ["Liftback", "Kombi"], "Kodiaq": ["SUV"], "Enyaq": ["SUV", "Coupe"]},
    "Tesla": {"Model 3": ["Sedan"], "Model Y": ["SUV"], "Model S": ["Sedan"], "Model X": ["SUV"]},
    "Toyota": {"Corolla": ["Hatchback", "Sedan", "Kombi"], "Yaris": ["Hatchback"], "RAV4": ["SUV"], "C-HR": ["Crossover"], "Camry": ["Sedan"]},
    "Volkswagen": {"Golf": ["Hatchback", "Kombi"], "Passat": ["Kombi", "Sedan"], "Arteon": ["Liftback", "Kombi"], "ID.4": ["SUV"], "Tiguan": ["SUV"]},
    "Volvo": {"XC40": ["SUV"], "XC60": ["SUV"], "XC90": ["SUV"], "V60": ["Kombi"]},
    "Inna marka...": {"Wpisz ręcznie": ["Inne"]}
}

SEGMENTY_DOMYSLNE = {
    "Audi": {"A3": "Segment C", "A4": "Segment D", "A6": "Segment D", "Q3": "Segment C", "Q5": "Segment D", "Q8": "Segment E", "e-tron GT": "Segment E", "RS6": "Segment D"},
    "BMW": {"Seria 3": "Segment D", "Seria 4": "Segment D", "Seria 5": "Segment D", "Seria 7": "Segment E", "X3": "Segment D", "X5": "Segment E", "M3": "Segment D", "M4": "Segment D"},
    "Porsche": {"911 (992)": "Segment D", "Taycan": "Segment E", "Macan": "Segment D", "Panamera": "Segment E", "Cayenne": "Segment E"},
    "Tesla": {"Model 3": "Segment D", "Model Y": "Segment D", "Model S": "Segment E", "Model X": "Segment J"}
}

# --- BAZA FOLII ---
FOIL_GROUPS = {
    "XPEL (Folie Ochronne PPF)": {
        "Bezbarwne (Twój obecny kolor)": ["XPEL Ultimate Plus (Wysoki Połysk)", "XPEL Stealth (Mat/Satyna)"],
        "XPEL Color (Zmiana Koloru PPF)": ["Black (Połysk)", "White (Połysk)", "Red (Połysk)", "Nardo Grey (Połysk)", "Miami Blue (Połysk)"]
    },
    "3M 2080 Series": {
        "Matte (Matowe)": ["Matte Black (M12)", "Matte Deep Black (M22)", "Matte Dark Grey (M261)", "Matte White (M10)", "Matte Military Green (M26)"],
        "Satin (Satynowe)": ["Satin Black (S12)", "Satin Dark Grey (S162)", "Satin White (S10)", "Satin Vampire Red (S208)"],
        "Gloss (Połysk)": ["Gloss Black (G12)", "Gloss White (G10)", "Gloss Hot Rod Red (G13)", "Gloss Sky Blue (G77)"],
        "Color Flip (Kameleon)": ["Gloss Flip Electric Wave (GP287)", "Satin Flip Volcanic Flare (SP236)"]
    },
    "Avery Dennison SW900": {
        "Satin": ["Satin Black", "Satin Pearl White", "Satin Carmine Red", "Satin Khaki Green", "Satin Metallic Grey"],
        "Gloss": ["Gloss Black", "Gloss White", "Gloss Obsidian Black", "Gloss Rock Grey", "Gloss Carmine Red"],
        "Matte": ["Matte Black", "Matte White", "Matte Charcoal Metallic", "Matte Olive Green"]
    },
    "Oracal 970RA": {
        "Special": ["Gloss Telegrey", "Gloss Nardo Grey Style", "Matte Nato Olive"],
        "Metallic": ["Gloss Graphite Metallic", "Matte Anthracite Metallic", "Gloss Silver Grey"]
    }
}

# ==========================================================================
# OPIS SCENY - CIEMNY GARAŻ Z OŚWIETLENIEM LED
# Jedno miejsce definicji żeby łatwo edytować styl wszystkich wizualizacji
# ==========================================================================
SCENE_DESCRIPTION = (
    "The car is photographed inside a completely dark modern detailing garage. "
    "Pitch black walls, black polished concrete floor with subtle reflections of the car body. "
    "The only light sources are modern cool-white LED strip lights integrated into the ceiling and along the floor edges, "
    "creating dramatic rim lighting on the car's silhouette and sharp highlights on the body panels. "
    "No other objects, no tools, no people, no text, no windows, no doors visible - only the car in pure darkness with LED lighting. "
    "Shot with professional automotive photography technique: 3/4 front angle, 35mm lens, cinematic composition, "
    "ultra sharp focus on the car, 8k resolution, photorealistic."
)

# --- FUNKCJE SYSTEMOWE ---
def install_fonts():
    font_src, font_dst = "fonts", os.path.expanduser("~/.local/share/fonts")
    if os.path.exists(font_src):
        if not os.path.exists(font_dst): os.makedirs(font_dst)
        for f in os.listdir(font_src):
            if f.lower().endswith((".ttf", ".otf")): shutil.copy(os.path.join(font_src, f), font_dst)
        subprocess.run(["fc-cache", "-f"], capture_output=True)


def _crop_to_target_ratio(img_bytes):
    """Przycina obraz do proporcji 21:18.7 (jak w oryginalnej aplikacji)."""
    img = Image.open(io.BytesIO(img_bytes))
    if img.mode != 'RGB':
        img = img.convert('RGB')
    w, h = img.size
    target_ratio = 21.0 / 18.7
    if w / h > target_ratio:
        new_w = int(h * target_ratio)
        left = (w - new_w) / 2
        img_cropped = img.crop((left, 0, left + new_w, h))
    else:
        new_h = int(w / target_ratio)
        top = (h - new_h) / 2
        img_cropped = img.crop((0, top, w, top + new_h))
    out_bytes = io.BytesIO()
    img_cropped.save(out_bytes, format='PNG')
    return out_bytes.getvalue()


def generate_ai_image(car_description, color_description, reference_images=None):
    """
    Generuje wizualizację auta przez Gemini 2.5 Flash Image (Nano Banana).
    
    car_description: opis auta (marka, model, rocznik, nadwozie) - dla trybu text-to-image
    color_description: opis docelowego koloru/folii
    reference_images: lista tupli [(bytes, mime_type), ...] - zdjęcia referencyjne auta
    
    Gdy przekazane są reference_images:
      - Model TRAKTUJE je jako referencje geometrii/proporcji tego samego auta
      - Zachowuje bryłę, zmienia TYLKO kolor karoserii
      - Scenę zastępuje ciemnym garażem z LED
    
    Gdy NIE ma referencji:
      - Model generuje auto od zera na podstawie opisu
      - Scena: ciemny garaż z LED
    """
    api_key = st.secrets["GEMINI_API_KEY"]
    
    # Poprawny endpoint Gemini 2.5 Flash Image
    url = f"https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash-image:generateContent"
    headers = {
        "Content-Type": "application/json",
        "x-goog-api-key": api_key
    }
    
    parts = []
    
    if reference_images and len(reference_images) > 0:
        # TRYB EDYCJI I2I - mamy zdjęcia referencyjne
        # KLUCZOWE: jawnie mówimy modelowi że to wszystkie zdjęcia tego SAMEGO auta
        # (różne ujęcia), a on ma wygenerować JEDNO nowe ujęcie w garażu z tym autem.
        # Inaczej Nano Banana próbuje scalać zdjęcia jak fusion.
        
        count = len(reference_images)
        if count == 1:
            ref_instruction = (
                "The attached image is a reference photograph of the EXACT car model I want to visualize. "
                "Your task: render this SAME car (identical body shape, identical headlights, identical wheels, "
                "identical proportions, identical generation/year) but place it in a new scene and change its color."
            )
        else:
            ref_instruction = (
                f"The {count} attached images are ALL reference photographs of the SAME car model taken from different angles. "
                "Use them TOGETHER to understand the exact 3D geometry, body shape, and design details of this car. "
                "Your task: render this EXACT same car model (identical generation, identical body shape, identical headlights, "
                "identical wheels, identical proportions) as ONE NEW photograph in a new scene with a new color. "
                "Do NOT try to merge or combine the images into a collage. Treat them as 3D reference of one single car."
            )
        
        full_prompt = (
            f"{ref_instruction}\n\n"
            f"NEW SCENE AND COLOR REQUIREMENTS:\n"
            f"- Car body wrap/paint color: {color_description}\n"
            f"- Scene: {SCENE_DESCRIPTION}\n\n"
            f"CRITICAL RULES:\n"
            f"- Preserve the EXACT car model from the reference(s) - same generation, same body, same face, same wheels\n"
            f"- Only change: the body color and the environment around the car\n"
            f"- Do NOT modify the car's geometry, proportions, or design details\n"
            f"- Output: ONE single photograph of the car in the described dark garage with LED lighting\n"
            f"- Framing: 3/4 front view, car centered, full body visible"
        )
        
        # Prompt tekstowy NAJPIERW
        parts.append({"text": full_prompt})
        
        # Potem wszystkie zdjęcia referencyjne
        for img_bytes, mime_type in reference_images:
            base64_image = base64.b64encode(img_bytes).decode('utf-8')
            parts.append({
                "inline_data": {
                    "mime_type": mime_type,
                    "data": base64_image
                }
            })
    else:
        # TRYB GENERACJI od zera - brak referencji
        # Wzmocniony prompt aby wymusić aktualną generację modelu
        full_prompt = (
            f"Generate a photorealistic image of the following car:\n"
            f"{car_description}\n\n"
            f"Make sure this is the CURRENT generation of this model as sold in showrooms today - "
            f"the latest facelift or newest generation available in 2025/2026, NOT older generations. "
            f"If uncertain about the newest generation, prefer showing a generic modern body style "
            f"matching the described segment rather than an outdated specific model.\n\n"
            f"Color: the car's body is wrapped/painted in {color_description}.\n\n"
            f"Scene: {SCENE_DESCRIPTION}"
        )
        parts.append({"text": full_prompt})
    
    payload = {
        "contents": [{"parts": parts}],
        "generationConfig": {
            "responseModalities": ["IMAGE"],
            "temperature": 0.4  # Niższa temperatura = bardziej przewidywalne, mniej halucynacji
        }
    }
    
    try:
        response = requests.post(url, headers=headers, json=payload, timeout=120)
        
        if response.status_code != 200:
            st.error(f"API zwróciło błąd ({response.status_code}): {response.text[:800]}")
            return _fallback_image()
        
        data = response.json()
        
        # Obsługa braku kandydatów (np. blokada safety)
        if 'candidates' not in data or len(data['candidates']) == 0:
            st.error(f"API nie zwróciło obrazu. Pełna odpowiedź: {str(data)[:800]}")
            return _fallback_image()
        
        candidate = data['candidates'][0]
        
        # Sprawdzenie czy nie było blokady
        if candidate.get('finishReason') in ('SAFETY', 'PROHIBITED_CONTENT', 'BLOCKLIST'):
            st.error(f"Generacja zablokowana przez safety filter: {candidate.get('finishReason')}")
            return _fallback_image()
        
        # Ekstrakcja obrazu z parts (może być camelCase lub snake_case)
        content_parts = candidate.get('content', {}).get('parts', [])
        for part in content_parts:
            inline = part.get('inlineData') or part.get('inline_data')
            if inline and inline.get('data'):
                img_data = base64.b64decode(inline['data'])
                return _crop_to_target_ratio(img_data)
        
        # Gdy nie ma obrazu - pokazujemy co zwrócił model (często tekst z wyjaśnieniem)
        text_response = ""
        for part in content_parts:
            if part.get('text'):
                text_response += part['text'] + "\n"
        
        st.error(f"Model nie zwrócił obrazu. Odpowiedź modelu: {text_response[:500] or 'brak'}")
        return _fallback_image()
        
    except requests.exceptions.Timeout:
        st.error("Timeout - model nie odpowiedział w 120 sekund. Spróbuj ponownie.")
        return _fallback_image()
    except Exception as e:
        st.error(f"Błąd komunikacji z modelem: {e}")
        return _fallback_image()


def _fallback_image():
    """Ciemny placeholder jeśli generacja się nie uda."""
    img_fallback = Image.new('RGB', (2100, 1870), color=(15, 15, 18))
    out_fallback = io.BytesIO()
    img_fallback.save(out_fallback, format='PNG')
    return out_fallback.getvalue()


def _zbuduj_wolacz_po_polsku(klient):
    """
    Zwraca poprawną formę wołacza po polsku ("Panie Piotrze", "Pani Anno", itp.).
    Odporna na puste pole - wtedy zwraca "Szanowny Kliencie".
    Ta logika leci przed AI bo Gemini słabo radzi sobie z polską deklinacją.
    """
    imie_surowe = klient.split()[0] if klient.strip() != "" else ""
    if not imie_surowe:
        return "Szanowny Kliencie"
    
    imie = imie_surowe.title()
    imie_lower = imie.lower()
    
    if imie_lower.endswith('a'):
        return f"Pani {imie}"
    
    wyjatki = {
        "piotr": "Piotrze", "paweł": "Pawle", "kacper": "Kacprze",
        "marek": "Marku", "michał": "Michale", "donald": "Donaldzie",
        "konrad": "Konradzie", "dawid": "Dawidzie"
    }
    if imie_lower in wyjatki:
        return f"Panie {wyjatki[imie_lower]}"
    
    if imie_lower.endswith('d'):
        return f"Panie {imie}zie"
    if imie_lower.endswith(('k', 'g', 'ch', 'j', 'sz', 'cz', 'rz', 'l', 'c')):
        if imie_lower.endswith('ek'):
            return f"Panie {imie[:-2]}ku"
        return f"Panie {imie}u"
    if imie_lower.endswith(('n', 'm', 'b', 'w', 'f', 's', 'z', 't', 'p')):
        return f"Panie {imie}ie"
    if imie_lower.endswith('r'):
        return f"Panie {imie}ze"
    return f"Panie {imie}"


def _fallback_intro_text(wolacz, marka, czysta_folia, handlowiec_imie, handlowiec_stanowisko):
    """
    Awaryjny szablon - używany gdy Gemini nie odpowie.
    UWAGA: Tekst CELOWO jest oznaczony "[WSTĘP AWARYJNY]" żeby od razu było widać
    że generacja AI padła i trzeba to wyłapać. W produkcji jeśli to się pokaże,
    znaczy że jest problem z API Gemini.
    """
    marka_txt = marka if marka and marka != "Inna marka..." else "samochodu"
    tresc = (
        f"{wolacz},\n\n"
        f"[WSTĘP AWARYJNY - generacja AI niedostępna]\n"
        f"Dziękuję za zaufanie przy wyborze zabezpieczenia Pańskiego {marka_txt}. "
        f"Zastosowana folia {czysta_folia} zapewni skuteczną ochronę lakieru. "
        f"Zapraszam do zapoznania się ze szczegółami wyceny."
    )
    return f"{tresc}\n\nZ motoryzacyjnym pozdrowieniem,\n{handlowiec_imie}\n{handlowiec_stanowisko}"


def generate_ai_intro_text(klient, brand, model, pakiet, folia, handlowiec_imie, handlowiec_stanowisko):
    """
    Generuje spersonalizowany wstęp do oferty przez Gemini 2.5 Flash.
    Ton: profesjonalny, ciepły. Długość: 3-4 zdania. Zawsze świeży tekst.
    
    W razie błędu API wraca do bezpiecznego szablonu fallback (oferta musi się
    wygenerować ZAWSZE, nawet jeśli Gemini akurat nie działa).
    """
    wolacz = _zbuduj_wolacz_po_polsku(klient)
    czysta_folia = folia.split('(')[0].strip()
    marka_model = f"{brand} {model}".strip() if brand and brand != "Inna marka..." else "samochodu"
    
    # Kontekst dla modelu - wszystko co może pomóc w personalizacji
    kontekst = (
        f"Autor wiadomości: {handlowiec_imie}, {handlowiec_stanowisko}\n"
        f"Firma: ITS WRAP (premium detailing, folie ochronne PPF, oklejanie samochodów)\n"
        f"Klient: {klient if klient.strip() else 'NIEZNANE IMIĘ (zwracaj się ogólnie)'}\n"
        f"Samochód klienta: {marka_model}\n"
        f"Zamówiona usługa: {pakiet}\n"
        f"Wybrana folia: {czysta_folia}"
    )
    
    prompt = f"""Napisz krótki, spersonalizowany wstęp do oferty handlowej od specjalisty detailingu do klienta.

KONTEKST:
{kontekst}

STRUKTURA WYJŚCIA (OBOWIĄZKOWA):
Linia 1: "{wolacz},"
Linia 2: (pusta)
Linia 3+: Treść główna - DOKŁADNIE 3-4 zdania.

ZASADY TECHNICZNE:
- Odpowiedz WYŁĄCZNIE treścią wiadomości. Bez preambuły, bez komentarzy, bez markdownu, bez cudzysłowów.
- NIE dodawaj podpisu, nazwiska, stanowiska ani formuły pożegnalnej ("Pozdrawiam", "Z poważaniem" itp.) - podpis zostanie dodany automatycznie.
- NIE używaj emoji.

WYMAGANIA TREŚCIOWE (KAŻDA JEST OBOWIĄZKOWA):
1. MUSISZ wprost wspomnieć markę i model auta: {marka_model}
2. MUSISZ wprost wspomnieć nazwę folii: {czysta_folia}
3. MUSISZ podać jedną konkretną korzyść z zastosowania tej folii (ochrona lakieru, efekt wizualny, trwałość - wybierz jedną i rozwiń jednym zdaniem).
4. Ton: profesjonalny, ciepły, ludzki. Bez sprzedażowej nachalności.
5. Forma: "Pan/Pani" (ty kolumnowe). Nigdy nie przechodź na "ty".
6. ZAKAZANE frazy (są zużyte i brzmią sztucznie): "bezkompromisowe rozwiązanie", "pasja do motoryzacji", "najwyższa jakość", "na długie lata", "spokój ducha", "perfekcyjna prezencja".

PRZYKŁADY DOBREGO STYLU (do inspiracji, NIE kopiuj):
- "Dziękuję za zaufanie przy wyborze zabezpieczenia Pańskiego Porsche 911. Folia XPEL Ultimate Plus, którą zaproponowałem, skutecznie ochroni lakier przed odpryskami i zachowa fabryczny połysk przez wiele sezonów użytkowania. Zapraszam do szczegółów wyceny."
- "Miło mi przesłać wycenę zabezpieczenia BMW Serii 3. Matowe wykończenie 3M 2080 Matte Black nada autu unikalny charakter, jednocześnie chroniąc oryginalny lakier pod spodem. Wszystkie szczegóły znajduje Pan w dalszej części oferty."

Napisz teraz wstęp (zaczynając od "{wolacz},"):"""
    
    try:
        # Sprawdzenie czy klucz API istnieje w sekretach
        try:
            api_key = st.secrets["GEMINI_API_KEY"]
        except KeyError:
            st.error("🔑 Brak klucza GEMINI_API_KEY w secrets.toml - używam szablonu awaryjnego.")
            return _fallback_intro_text(wolacz, brand, czysta_folia, handlowiec_imie, handlowiec_stanowisko)
        
        if not api_key or len(str(api_key).strip()) < 10:
            st.error("🔑 Klucz GEMINI_API_KEY jest pusty lub nieprawidłowy - używam szablonu awaryjnego.")
            return _fallback_intro_text(wolacz, brand, czysta_folia, handlowiec_imie, handlowiec_stanowisko)
        
        url = "https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent"
        headers = {
            "Content-Type": "application/json",
            "x-goog-api-key": api_key
        }
        payload = {
            "contents": [{"parts": [{"text": prompt}]}],
            "generationConfig": {
                "temperature": 0.9,  # wysoka -> za każdym razem świeży, unikalny tekst
                "maxOutputTokens": 1024,  # spory zapas - thinking może jeszcze odgryzać część
                # KLUCZOWE: wyłączamy "thinking mode" w Gemini 2.5 Flash.
                # Bez tego model zjada nawet kilkaset tokenów na wewnętrzne myślenie,
                # które liczy się do maxOutputTokens - efekt: odpowiedź urwana w połowie słowa.
                # Do prostego zadania (4 zdania tekstu) myślenie jest zbędne.
                "thinkingConfig": {
                    "thinkingBudget": 0
                }
            }
        }
        
        response = requests.post(url, headers=headers, json=payload, timeout=30)
        
        if response.status_code != 200:
            st.error(
                f"❌ Gemini API zwrócił błąd {response.status_code}.\n\n"
                f"Pełna odpowiedź API:\n```\n{response.text[:1500]}\n```\n\n"
                f"Używam szablonu awaryjnego."
            )
            return _fallback_intro_text(wolacz, brand, czysta_folia, handlowiec_imie, handlowiec_stanowisko)
        
        data = response.json()
        
        if 'candidates' not in data or not data['candidates']:
            st.error(
                f"❌ Gemini nie zwrócił żadnego kandydata odpowiedzi.\n\n"
                f"Pełna odpowiedź:\n```\n{str(data)[:1500]}\n```\n\n"
                f"Używam szablonu awaryjnego."
            )
            return _fallback_intro_text(wolacz, brand, czysta_folia, handlowiec_imie, handlowiec_stanowisko)
        
        candidate = data['candidates'][0]
        finish_reason = candidate.get('finishReason')
        
        if finish_reason in ('SAFETY', 'PROHIBITED_CONTENT', 'BLOCKLIST', 'RECITATION'):
            st.error(f"❌ Gemini zablokował odpowiedź (powód: {finish_reason}) - używam szablonu awaryjnego.")
            return _fallback_intro_text(wolacz, brand, czysta_folia, handlowiec_imie, handlowiec_stanowisko)
        
        if finish_reason == 'MAX_TOKENS':
            # Model uderzył w limit - zwykle winne są "thinking tokens" w Gemini 2.5.
            # Mamy już thinkingBudget:0 i maxOutputTokens:1024, więc to NIE powinno się zdarzać.
            # Jeśli jednak się zdarzy - pokazujemy ile tokenów model zużył na thinking.
            usage = data.get('usageMetadata', {})
            thoughts = usage.get('thoughtsTokenCount', 'n/a')
            output = usage.get('candidatesTokenCount', 'n/a')
            st.error(
                f"❌ Gemini uderzył w limit tokenów. Thinking: {thoughts}, Output: {output}. "
                f"Używam szablonu awaryjnego."
            )
            return _fallback_intro_text(wolacz, brand, czysta_folia, handlowiec_imie, handlowiec_stanowisko)
        
        parts = candidate.get('content', {}).get('parts', [])
        wygenerowany_tekst = ""
        for part in parts:
            if part.get('text'):
                wygenerowany_tekst += part['text']
        
        wygenerowany_tekst = wygenerowany_tekst.strip()
        
        if not wygenerowany_tekst:
            st.error(
                f"❌ Gemini zwrócił pustą odpowiedź (finishReason: {finish_reason}).\n\n"
                f"Pełna odpowiedź kandydata:\n```\n{str(candidate)[:1500]}\n```\n\n"
                f"Używam szablonu awaryjnego."
            )
            return _fallback_intro_text(wolacz, brand, czysta_folia, handlowiec_imie, handlowiec_stanowisko)
        
        # Sanity check - jeśli model zwrócił coś podejrzanie krótkiego, fallback
        if len(wygenerowany_tekst) < 80:
            st.warning(f"⚠️ Gemini zwrócił zbyt krótki tekst ({len(wygenerowany_tekst)} znaków): '{wygenerowany_tekst}' - używam szablonu awaryjnego.")
            return _fallback_intro_text(wolacz, brand, czysta_folia, handlowiec_imie, handlowiec_stanowisko)
        
        # Obrona przed modelem, który mimo instrukcji sam doda podpis.
        # KLUCZOWE: szukamy fraz TYLKO na początku linii (po \n), nigdy w środku zdania.
        linie = wygenerowany_tekst.split('\n')
        frazy_podpisu_start_linii = [
            'z motoryzacyjnym', 'z poważaniem', 'z wyrazami',
            'pozdrawiam', 'serdecznie pozdr', 'łączę pozdr',
            'z uszanowaniem', 'pozdrowienia', 'z najlepszymi'
        ]
        linie_wynikowe = []
        for linia in linie:
            linia_lower = linia.strip().lower()
            if any(linia_lower.startswith(fraza) for fraza in frazy_podpisu_start_linii):
                break  # od tej linii w dół - to podpis, odcinamy
            linie_wynikowe.append(linia)
        
        wygenerowany_tekst = '\n'.join(linie_wynikowe).strip()
        
        if len(wygenerowany_tekst) < 80:
            st.warning(f"⚠️ Po obcięciu podpisu tekst zbyt krótki ({len(wygenerowany_tekst)} znaków) - używam szablonu awaryjnego.")
            return _fallback_intro_text(wolacz, brand, czysta_folia, handlowiec_imie, handlowiec_stanowisko)
        
        return f"{wygenerowany_tekst}\n\nZ motoryzacyjnym pozdrowieniem,\n{handlowiec_imie}\n{handlowiec_stanowisko}"
        
    except requests.exceptions.Timeout:
        st.error("⏱️ Gemini - timeout po 30s - używam szablonu awaryjnego.")
        return _fallback_intro_text(wolacz, brand, czysta_folia, handlowiec_imie, handlowiec_stanowisko)
    except Exception as e:
        import traceback
        st.error(
            f"❌ Wyjątek podczas generacji wstępu:\n```\n{type(e).__name__}: {e}\n\n{traceback.format_exc()[:1000]}\n```\n\n"
            f"Używam szablonu awaryjnego."
        )
        return _fallback_intro_text(wolacz, brand, czysta_folia, handlowiec_imie, handlowiec_stanowisko)


def download_file(service, file_id):
    request = service.files().get_media(fileId=file_id)
    fh = io.BytesIO(); downloader = MediaIoBaseDownload(fh, request)
    done = False
    while not done: _, done = downloader.next_chunk()
    fh.seek(0); return fh


def pptx_to_pdf(input_path):
    try:
        subprocess.run(['libreoffice', '--headless', '--convert-to', 'pdf', '--outdir', os.getcwd(), input_path], check=True, capture_output=True)
        return os.path.basename(input_path).replace('.pptx', '.pdf')
    except: return None


# --- OBSŁUGA GOOGLE DRIVE DLA OFERT (OAuth - prywatne konto Google) ---
# Service Accounts nie mają miejsca na Drive, więc do uploadu PDF-ów używamy
# OAuth delegation - token odświeżający pozwala aplikacji działać w imieniu
# prywatnego konta Google użytkownika (i zużywać jego limit 15GB).
#
# Konfiguracja: sekret [oauth_drive] w secrets.toml z client_id, client_secret,
# refresh_token, token_uri (wygenerowane jednorazowo skryptem generuj_token.py).
#
# Service account (zmienna 'service') dalej obsługuje odczyt szablonów PPTX
# i Google Sheets (rejestr, cennik) - bo tam problemu z miejscem nie ma.

from google.oauth2.credentials import Credentials as OAuthCredentials

# UWAGA: NIE używamy @st.cache_resource bo cache trzymałby unieważniony token
# nawet po odnowieniu sekretów. Tworzenie klienta jest tanie - pomija się tylko
# milisekundy, a zyskujemy pewność że zawsze idzie świeży request do Google.
def get_oauth_drive_service():
    """Tworzy klienta Drive API działającego w imieniu prywatnego konta Google."""
    try:
        oauth_conf = st.secrets["oauth_drive"]
    except KeyError:
        st.error(
            "🔑 Brak konfiguracji [oauth_drive] w secrets.toml. "
            "Uruchom skrypt generuj_token.py lokalnie aby wygenerować refresh token."
        )
        return None
    
    # Walidacja - czy wszystkie 4 wartości są obecne i niepuste
    wymagane_klucze = ['client_id', 'client_secret', 'refresh_token', 'token_uri']
    brakujace = [k for k in wymagane_klucze if not oauth_conf.get(k, '').strip()]
    if brakujace:
        st.error(f"🔑 W [oauth_drive] brakuje wartości: {', '.join(brakujace)}")
        return None
    
    # Sanity check refresh tokena - powinien zaczynać się od "1//" i mieć min 50 znaków
    rt = oauth_conf['refresh_token'].strip()
    if not rt.startswith('1//') or len(rt) < 50:
        st.error(
            f"🔑 refresh_token wygląda nieprawidłowo (długość: {len(rt)} znaków, "
            f"początek: '{rt[:10]}...'). Poprawny token zaczyna się od '1//' i ma 100+ znaków. "
            f"Wygeneruj nowy używając generuj_token.py."
        )
        return None
    
    try:
        creds = OAuthCredentials(
            token=None,  # zostanie odświeżony automatycznie z refresh_token
            refresh_token=rt,
            token_uri=oauth_conf['token_uri'].strip(),
            client_id=oauth_conf['client_id'].strip(),
            client_secret=oauth_conf['client_secret'].strip(),
            scopes=['https://www.googleapis.com/auth/drive.file']
        )
        # Próbujemy od razu odświeżyć - jeśli token jest unieważniony, dowiemy się TERAZ
        # zamiast w środku uploadu PDF.
        from google.auth.transport.requests import Request as AuthRequest
        creds.refresh(AuthRequest())
        return build('drive', 'v3', credentials=creds)
    except Exception as e:
        err_str = str(e).lower()
        if 'invalid_grant' in err_str or 'token has been expired' in err_str or 'revoked' in err_str:
            st.error(
                "🔑 **Refresh token został unieważniony przez Google.** Możliwe przyczyny:\n\n"
                "1. **OAuth consent screen jest w trybie 'Testing'** → tokeny wygasają po 7 dniach. "
                "   Wejdź w Google Cloud Console → Ekran zgody OAuth → kliknij **'Publish app'**.\n"
                "2. **Cofnąłeś dostęp aplikacji** na [myaccount.google.com/permissions](https://myaccount.google.com/permissions)\n"
                "3. **OAuth Client ID został usunięty/zmieniony** w Google Cloud Console\n"
                "4. **Hasło konta Google zostało zmienione** - inwaliduje wszystkie tokeny\n\n"
                "**Rozwiązanie:** uruchom ponownie `generuj_token.py` lokalnie, "
                "skopiuj nowe wartości do Streamlit Secrets."
            )
        else:
            st.error(f"🔑 Błąd inicjalizacji OAuth Drive: {e}")
        return None


def pobierz_lub_stworz_folder_oferty_oauth(oauth_service):
    """
    Szuka folderu 'Oferty ITS WRAP' w korzeniu prywatnego Drive użytkownika.
    Tworzy go jeśli nie istnieje.
    
    UWAGA: scope drive.file pozwala operować TYLKO na plikach/folderach
    stworzonych przez tę aplikację - więc nie zobaczymy ewentualnego istniejącego
    folderu 'Oferty' stworzonego ręcznie. Dlatego używamy własnej unikalnej nazwy.
    """
    folder_name = "Oferty ITS WRAP"
    query = f"name='{folder_name}' and mimeType='application/vnd.google-apps.folder' and trashed=false"
    results = oauth_service.files().list(
        q=query,
        fields="files(id, name)",
        spaces='drive'
    ).execute()
    items = results.get('files', [])
    
    if not items:
        file_metadata = {
            'name': folder_name,
            'mimeType': 'application/vnd.google-apps.folder'
        }
        folder = oauth_service.files().create(
            body=file_metadata,
            fields='id'
        ).execute()
        return folder.get('id')
    return items[0].get('id')


def wgraj_pdf_na_dysk(oauth_service, folder_id, file_name, file_bytes):
    """Upload PDF na prywatny Drive właściciela konta OAuth."""
    try:
        file_metadata = {
            'name': file_name,
            'parents': [folder_id]
        }
        media = MediaIoBaseUpload(io.BytesIO(file_bytes), mimetype='application/pdf', resumable=True)
        file = oauth_service.files().create(
            body=file_metadata,
            media_body=media,
            fields='id, webViewLink'
        ).execute()
        
        # Uprawnienia: każdy z linkiem może otworzyć (read-only)
        oauth_service.permissions().create(
            fileId=file.get('id'),
            body={'type': 'anyone', 'role': 'reader'}
        ).execute()
        
        return file.get('webViewLink')
    except Exception as e:
        st.error(f"Błąd podczas zapisu pliku PDF na Google Drive: {e}")
        return None


def zapisz_do_rejestru(nr_oferty, handlowiec, klient, auto, usluga, folia, cena, pdf_link):
    try:
        sheet_rejestr = client.open_by_url(LINK_DO_ARKUSZA).worksheet("Rejestr")
        dzisiaj = datetime.now().strftime("%Y-%m-%d")
        nowy_wiersz = [dzisiaj, nr_oferty, handlowiec, klient, auto, usluga, folia, f"{cena} zł", "Nowa", pdf_link]
        sheet_rejestr.append_row(nowy_wiersz)
        return True
    except Exception as e:
        st.error(f"Nie udało się zapisać do bazy (Rejestr): {e}")
        return False


def replace_text_in_shape(shape, replacements):
    """Rekurencyjnie wyszukuje i podmienia tagi tekstowe."""
    if hasattr(shape, "has_text_frame") and shape.has_text_frame:
        for p in shape.text_frame.paragraphs:
            p_text = "".join(run.text for run in p.runs)
            original_text = p_text
            for k, v in replacements.items():
                if k in p_text:
                    p_text = p_text.replace(k, str(v))
            if p_text != original_text and len(p.runs) > 0:
                p.runs[0].text = p_text
                for run in p.runs[1:]:
                    run.text = ""
                    
    if hasattr(shape, "has_table") and shape.has_table:
        for row in shape.table.rows:
            for cell in row.cells:
                replace_text_in_shape(cell, replacements)
                
    if shape.shape_type == 6:  # msoGroup
        for subshape in shape.shapes:
            replace_text_in_shape(subshape, replacements)


# --- APLIKACJA ---
st.set_page_config(
    page_title="IT'S WRAP - Generator Ofert",
    page_icon="🛡️",
    layout="wide",
    initial_sidebar_state="expanded"
)
install_fonts()

# ==========================================================================
# STYL CI IT'S WRAP - zgodny z Brand Manual 2026
# 
# Paleta kolorów (z brand manual str. 16, 23):
# - IT'S WRAP BLUE: #007DC5 (primary, akcenty, przyciski)
# - IT'S WRAP NAVY BLUE: #042643 (sidebar, ciemne tło)
# - SEMI 1-4: #003559, #004470, #005489, #0064A3 (gradienty/tła)
# - Biel #FFFFFF, Czerń #000000
#
# Typografia (str. 26):
# - URW DIN (bazowy) - font komercyjny, trudno dostępny
# - Roboto (Google Font) - DOPUSZCZALNY ZAMIENNIK na www - używamy tego
# ==========================================================================

# Ładujemy logo (jeśli jest) - wyświetlane w sidebar
import base64 as _b64_css
def _logo_base64(sciezka):
    try:
        with open(sciezka, 'rb') as f:
            return _b64_css.b64encode(f.read()).decode('utf-8')
    except FileNotFoundError:
        return None

# Szukamy logo w typowych lokalizacjach repo
LOGO_PATH_CANDIDATES = ['logo.png', 'assets/logo.png', 'static/logo.png', 'images/logo.png']
LOGO_B64 = None
LOGO_EXT = None
for _cand in LOGO_PATH_CANDIDATES:
    if os.path.exists(_cand):
        LOGO_B64 = _logo_base64(_cand)
        LOGO_EXT = _cand.split('.')[-1]
        break

st.markdown("""
<link rel="preconnect" href="https://fonts.googleapis.com">
<link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
<link href="https://fonts.googleapis.com/css2?family=Roboto:wght@300;400;500;700;900&display=swap" rel="stylesheet">

<style>
/* ==========================================================================
   PALETA MARKI IT'S WRAP
   ========================================================================== */
:root {
    --iw-blue: #007DC5;
    --iw-blue-dark: #0064A3;
    --iw-blue-darker: #005489;
    --iw-navy: #042643;
    --iw-navy-light: #003559;
    --iw-navy-lighter: #004470;
    --iw-white: #FFFFFF;
    --iw-black: #000000;
    --iw-gray-light: #F5F7FA;
    --iw-gray-border: #E1E7ED;
    --iw-gray-text: #4A5568;
}

/* ==========================================================================
   TYPOGRAFIA - Roboto jako font bazowy (zgodny z brand manualem)
   ========================================================================== */
html, body, [class*="css"], .stApp, .stMarkdown, .stText,
.stSelectbox, .stTextInput, .stNumberInput, .stButton,
.stRadio, .stCheckbox, .stFileUploader, .stDataFrame {
    font-family: 'Roboto', -apple-system, BlinkMacSystemFont, sans-serif !important;
}

/* Nagłówki - pogrubione, Roboto 700/900 */
h1, h2, h3, h4, h5, h6,
.stMarkdown h1, .stMarkdown h2, .stMarkdown h3 {
    font-family: 'Roboto', sans-serif !important;
    font-weight: 700 !important;
    letter-spacing: -0.01em;
    color: var(--iw-navy) !important;
}

/* H1 jak tytuły sekcji w brand manualu - uppercase, w niebieskim */
h1, .stMarkdown h1 {
    color: var(--iw-blue) !important;
    text-transform: uppercase;
    letter-spacing: 0.02em;
    font-weight: 900 !important;
    border-bottom: 3px solid var(--iw-blue);
    padding-bottom: 12px;
    margin-bottom: 24px;
}

/* ==========================================================================
   GŁÓWNY PANEL - jasne tło (papier firmowy)
   ========================================================================== */
.stApp {
    background-color: var(--iw-white);
}

.main .block-container {
    padding-top: 2rem;
    padding-bottom: 3rem;
    max-width: 1400px;
}

/* ==========================================================================
   SIDEBAR - ciemny navy (inspirowany belką mailową w CI)
   ========================================================================== */
section[data-testid="stSidebar"] {
    background-color: var(--iw-navy) !important;
    border-right: 4px solid var(--iw-blue);
}

section[data-testid="stSidebar"] * {
    color: var(--iw-white) !important;
}

section[data-testid="stSidebar"] h1,
section[data-testid="stSidebar"] h2,
section[data-testid="stSidebar"] h3,
section[data-testid="stSidebar"] .stMarkdown h1,
section[data-testid="stSidebar"] .stMarkdown h2 {
    color: var(--iw-white) !important;
    text-transform: uppercase;
    letter-spacing: 0.05em;
    font-weight: 700 !important;
    font-size: 1rem;
    border-bottom: 2px solid var(--iw-blue);
    padding-bottom: 6px;
    margin-top: 16px;
}

/* Separator w sidebarze - subtelna kreska w blue */
section[data-testid="stSidebar"] hr {
    border-color: rgba(0, 125, 197, 0.3) !important;
    margin: 1.5rem 0 !important;
}

/* Kontrolki w sidebarze - ciemne pola z niebieskim borderem */
section[data-testid="stSidebar"] input,
section[data-testid="stSidebar"] select,
section[data-testid="stSidebar"] textarea,
section[data-testid="stSidebar"] [data-baseweb="select"] > div {
    background-color: var(--iw-navy-light) !important;
    color: var(--iw-white) !important;
    border: 1px solid rgba(0, 125, 197, 0.4) !important;
}

section[data-testid="stSidebar"] input:focus,
section[data-testid="stSidebar"] [data-baseweb="select"]:focus-within > div {
    border-color: var(--iw-blue) !important;
    box-shadow: 0 0 0 1px var(--iw-blue) !important;
}

section[data-testid="stSidebar"] label {
    color: rgba(255, 255, 255, 0.85) !important;
    font-size: 0.85rem !important;
    font-weight: 500 !important;
}

/* File uploader w sidebarze */
section[data-testid="stSidebar"] [data-testid="stFileUploadDropzone"] {
    background-color: var(--iw-navy-light) !important;
    border: 2px dashed rgba(0, 125, 197, 0.5) !important;
}

section[data-testid="stSidebar"] [data-testid="stFileUploadDropzone"]:hover {
    border-color: var(--iw-blue) !important;
    background-color: var(--iw-navy-lighter) !important;
}

/* ==========================================================================
   PRZYCISKI - w kolorze IT'S WRAP BLUE
   ========================================================================== */
.stButton > button {
    background-color: var(--iw-blue) !important;
    color: var(--iw-white) !important;
    border: none !important;
    border-radius: 4px !important;
    font-weight: 700 !important;
    letter-spacing: 0.03em;
    text-transform: uppercase;
    font-size: 0.9rem !important;
    padding: 10px 24px !important;
    transition: all 0.2s ease;
    box-shadow: 0 2px 4px rgba(4, 38, 67, 0.15);
}

.stButton > button:hover {
    background-color: var(--iw-blue-dark) !important;
    box-shadow: 0 4px 12px rgba(0, 125, 197, 0.3);
    transform: translateY(-1px);
}

.stButton > button:active {
    transform: translateY(0);
    box-shadow: 0 2px 4px rgba(4, 38, 67, 0.15);
}

/* Primary button (oznaczony type="primary") - pełna bryła blue z intensywniejszym hover */
.stButton > button[kind="primary"] {
    background-color: var(--iw-blue) !important;
    font-weight: 900 !important;
    padding: 12px 28px !important;
}

.stButton > button[kind="primary"]:hover {
    background-color: var(--iw-navy) !important;
}

/* Download button */
.stDownloadButton > button {
    background-color: var(--iw-navy) !important;
    color: var(--iw-white) !important;
    border: 2px solid var(--iw-blue) !important;
    font-weight: 700 !important;
    text-transform: uppercase;
}

.stDownloadButton > button:hover {
    background-color: var(--iw-blue) !important;
    border-color: var(--iw-blue) !important;
}

/* ==========================================================================
   POLA TEKSTOWE W GŁÓWNYM PANELU
   ========================================================================== */
.main input, .main select, .main textarea,
.main [data-baseweb="select"] > div,
.main [data-baseweb="input"] {
    border: 1px solid var(--iw-gray-border) !important;
    border-radius: 4px !important;
    background-color: var(--iw-white) !important;
}

.main input:focus, .main textarea:focus,
.main [data-baseweb="select"]:focus-within > div,
.main [data-baseweb="input"]:focus-within {
    border-color: var(--iw-blue) !important;
    box-shadow: 0 0 0 2px rgba(0, 125, 197, 0.15) !important;
}

/* ==========================================================================
   ZAKŁADKI (TABS)
   ========================================================================== */
.stTabs [data-baseweb="tab-list"] {
    gap: 0;
    border-bottom: 2px solid var(--iw-gray-border);
    background: transparent;
}

.stTabs [data-baseweb="tab"] {
    background: transparent !important;
    border: none !important;
    border-radius: 0 !important;
    padding: 12px 24px !important;
    font-weight: 700 !important;
    text-transform: uppercase;
    letter-spacing: 0.03em;
    font-size: 0.9rem !important;
    color: var(--iw-gray-text) !important;
}

.stTabs [data-baseweb="tab"][aria-selected="true"] {
    color: var(--iw-blue) !important;
    border-bottom: 3px solid var(--iw-blue) !important;
    margin-bottom: -2px;
}

/* ==========================================================================
   ALERTY / KOMUNIKATY - spójne z paletą marki
   ========================================================================== */
div[data-testid="stAlert"] {
    border-radius: 4px !important;
    border-left-width: 4px !important;
}

/* INFO - niebieski akcent */
div[data-baseweb="notification"][kind="info"],
div[data-testid="stAlert"][class*="info"] {
    background-color: rgba(0, 125, 197, 0.08) !important;
    border-left-color: var(--iw-blue) !important;
}

/* SUCCESS - spójna zieleń brand-adjacent */
div[data-testid="stAlert"][class*="success"] {
    background-color: rgba(34, 139, 89, 0.08) !important;
    border-left-color: #228B59 !important;
}

/* ==========================================================================
   PODGLĄD WIZUALIZACJI - subtelna ramka
   ========================================================================== */
.main [data-testid="stImage"] img {
    border-radius: 4px;
    border: 1px solid var(--iw-gray-border);
    box-shadow: 0 4px 16px rgba(4, 38, 67, 0.08);
}

/* ==========================================================================
   DATA EDITOR / TABELE REJESTRU
   ========================================================================== */
[data-testid="stDataFrame"], [data-testid="stDataEditor"] {
    border: 1px solid var(--iw-gray-border) !important;
    border-radius: 4px !important;
}

/* Nagłówki tabel */
[data-testid="stDataFrame"] thead th,
[data-testid="stDataEditor"] thead th {
    background-color: var(--iw-navy) !important;
    color: var(--iw-white) !important;
    font-weight: 700 !important;
    text-transform: uppercase;
    font-size: 0.8rem !important;
    letter-spacing: 0.02em;
}

/* ==========================================================================
   STOPKA / DOPISKI
   ========================================================================== */
.iw-footer {
    margin-top: 3rem;
    padding: 1.5rem 0;
    border-top: 2px solid var(--iw-gray-border);
    text-align: center;
    color: var(--iw-gray-text);
    font-size: 0.8rem;
    text-transform: uppercase;
    letter-spacing: 0.05em;
}

.iw-footer strong {
    color: var(--iw-blue);
    font-weight: 900;
}

/* ==========================================================================
   NAGŁÓWEK W SIDEBARZE - logo + pasek
   ========================================================================== */
.iw-sidebar-logo {
    text-align: center;
    padding: 1rem 0.5rem 1.5rem 0.5rem;
    margin-bottom: 0.5rem;
    border-bottom: 2px solid var(--iw-blue);
}

.iw-sidebar-logo img {
    max-width: 180px;
    height: auto;
}

.iw-sidebar-claim {
    text-align: center;
    font-size: 0.7rem;
    letter-spacing: 0.15em;
    color: var(--iw-blue) !important;
    text-transform: uppercase;
    font-weight: 700;
    margin-top: 0.5rem;
}

/* ==========================================================================
   NAGŁÓWEK GŁÓWNEGO PANELU - styl jak strona tytułowa brand manualu
   ========================================================================== */
.iw-main-header {
    display: flex;
    align-items: center;
    gap: 20px;
    padding: 24px 0 20px 0;
    border-bottom: 3px solid var(--iw-blue);
    margin-bottom: 32px;
}

.iw-main-header-text {
    flex: 1;
}

.iw-main-header-title {
    font-family: 'Roboto', sans-serif;
    font-weight: 900;
    font-size: 2rem;
    color: var(--iw-navy);
    text-transform: uppercase;
    letter-spacing: 0.02em;
    line-height: 1.1;
    margin: 0;
}

.iw-main-header-subtitle {
    font-family: 'Roboto', sans-serif;
    font-weight: 400;
    font-size: 0.85rem;
    color: var(--iw-blue);
    text-transform: uppercase;
    letter-spacing: 0.2em;
    margin-top: 6px;
}

/* Hide default Streamlit branding for cleaner look */
#MainMenu {visibility: hidden;}
footer {visibility: hidden;}
header[data-testid="stHeader"] {
    background: transparent;
}
</style>
""", unsafe_allow_html=True)

creds = Credentials.from_service_account_info(st.secrets["gcp_service_account"], scopes=["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"])
service = build('drive', 'v3', credentials=creds)
client = gspread.authorize(creds)

results = service.files().list(
    q=f"'{PARENT_FOLDER_ID}' in parents and mimeType='application/vnd.openxmlformats-officedocument.presentationml.presentation' and trashed=false",
    fields="files(id, name)",
    supportsAllDrives=True,
    includeItemsFromAllDrives=True
).execute()
pliki_na_dysku = results.get('files', [])

try:
    sheet_cennik = client.open_by_url(LINK_DO_ARKUSZA).worksheet("Cennik usług")
    naglowki = [c.replace('\n', ' ').replace('\r', '').strip() for c in sheet_cennik.get_all_values()[0]]
    df_cennik = pd.DataFrame(sheet_cennik.get_all_values()[1:], columns=naglowki)
except Exception as e:
    st.error(f"Błąd ładowania Cennika v3. Upewnij się, że link i nazwa zakładki 'Cennik usług' są poprawne. Błąd: {e}")
    st.stop()


# --- PANEL BOCZNY ---
with st.sidebar:
    # --- NAGŁÓWEK SIDEBARA: LOGO IT'S WRAP + CLAIM ---
    if LOGO_B64:
        st.markdown(f"""
        <div class="iw-sidebar-logo">
            <img src="data:image/{LOGO_EXT};base64,{LOGO_B64}" alt="IT'S WRAP">
            <div class="iw-sidebar-claim">Make It Change</div>
        </div>
        """, unsafe_allow_html=True)
    else:
        # Fallback tekstowy jeśli logo nie zostało wgrane do repo
        st.markdown("""
        <div class="iw-sidebar-logo">
            <div style="font-size: 1.5rem; font-weight: 900; letter-spacing: 0.05em; color: #FFFFFF;">
                IT'S WRAP
            </div>
            <div class="iw-sidebar-claim">Make It Change</div>
        </div>
        """, unsafe_allow_html=True)
    
    st.markdown("### Opiekun Klienta")
    wybrany_handlowiec = st.selectbox("Kto przygotowuje ofertę?", list(HANDLOWCY.keys()))
    
    st.markdown("---")
    st.markdown("### Studio AI")
    brand = st.selectbox("Marka", list(CAR_DATABASE.keys()))
    
    if brand == "Inna marka...":
        custom_brand = st.text_input("Wpisz markę")
        custom_model = st.text_input("Wpisz model")
        final_brand, final_model, body, segment_domyslny = custom_brand, custom_model, "", "Segment D"
    else:
        final_brand = brand
        final_model = st.selectbox("Model", list(CAR_DATABASE[brand].keys()))
        body = st.selectbox("Nadwozie", CAR_DATABASE[brand][final_model])
        segment_domyslny = SEGMENTY_DOMYSLNE.get(brand, {}).get(final_model, "Segment D")
        
    segment_final = st.selectbox("Wybierz Segment (do wyceny)", ["Segment A", "Segment B", "Segment C", "Segment D", "Segment E", "Segment J", "Wavecamper"], 
                                 index=["Segment A", "Segment B", "Segment C", "Segment D", "Segment E", "Segment J", "Wavecamper"].index(segment_domyslny))
        
    year = st.selectbox("Rocznik", [str(y) for y in range(2026, 1999, -1)])
    gen_code = st.text_input("Kod karoserii (Opcjonalnie)", help="Np. G70, 992")
    
    st.markdown("---")
    st.markdown("### Folia i Kolor")
    f_brand = st.selectbox("Producent", list(FOIL_GROUPS.keys()))
    f_cat = st.selectbox("Wykończenie", list(FOIL_GROUPS[f_brand].keys()))
    f_color = st.selectbox("Kolor", FOIL_GROUPS[f_brand][f_cat])

    paint_color = ""
    if "Bezbarwne" in f_cat:
        paint_color = st.text_input("🚘 Podaj obecny kolor lakieru auta", value="Czarny metallic")

    st.markdown("---")
    st.markdown("### Wizualizacja AI")
    
    st.caption(
        "💡 **Rekomendacja:** wgraj 1-3 zdjęcia auta z konfiguratora producenta lub press-kitu "
        "(różne ujęcia tego samego modelu). Model AI zachowa bryłę i zmieni tylko kolor na ciemny garaż."
    )
    
    uploaded_files = st.file_uploader(
        "Zdjęcia referencyjne (opcjonalnie, ale mocno zalecane dla nowych modeli)", 
        type=['png', 'jpg', 'jpeg'], 
        accept_multiple_files=True
    )
    
    # Pole URL - alternatywa dla uploadu (np. prosto z konfiguratora producenta)
    url_zdjecia = st.text_input(
        "...lub wklej URL zdjęcia (np. z konfiguratora)",
        help="Bezpośredni link do obrazka .jpg/.png ze strony producenta"
    )

    if st.button("🪄 GENERUJ WIZUALIZACJĘ AI", type="primary"):
        # --- BUDOWA OPISU AUTA (car_description) ---
        extra_code = f" ({gen_code})" if gen_code else ""
        car_description = (
            f"{year} {final_brand} {final_model}{extra_code}, body type: {body}. "
            f"This is a {segment_final.replace('Segment ', 'segment ')} vehicle."
        )
        
        # --- BUDOWA OPISU KOLORU (color_description) ---
        if "Bezbarwne" in f_cat:
            if "Stealth" in f_color:
                color_description = (
                    f"the original {paint_color} factory paint covered with matte/satin transparent PPF film "
                    f"(XPEL Stealth) giving the paint a deep matte finish while keeping the original color visible"
                )
            else:
                color_description = (
                    f"the original {paint_color} factory paint covered with high-gloss transparent PPF film "
                    f"(XPEL Ultimate Plus) giving the paint extra depth and shine"
                )
        else:
            color_description = f"{f_color} vinyl wrap by {f_brand}"
        
        # --- ZEBRANIE ZDJĘĆ REFERENCYJNYCH ---
        reference_images = []
        
        # Z uploadu
        if uploaded_files:
            for uf in uploaded_files:
                img_bytes = uf.read()
                mime = uf.type if uf.type else "image/jpeg"
                reference_images.append((img_bytes, mime))
        
        # Z URL (jeśli nie ma uploadu ale jest URL)
        if url_zdjecia and url_zdjecia.strip():
            try:
                resp = requests.get(url_zdjecia.strip(), timeout=15, headers={
                    'User-Agent': 'Mozilla/5.0'
                })
                if resp.status_code == 200:
                    mime = resp.headers.get('content-type', 'image/jpeg').split(';')[0].strip()
                    if 'image' in mime:
                        reference_images.append((resp.content, mime))
                    else:
                        st.warning(f"URL nie wskazuje na obrazek (typ: {mime}).")
                else:
                    st.warning(f"Nie mogę pobrać obrazka z URL (kod: {resp.status_code}).")
            except Exception as e:
                st.warning(f"Nie udało się pobrać zdjęcia z URL: {e}")
        
        # --- KOMUNIKATY DLA UŻYTKOWNIKA ---
        if reference_images:
            st.info(f"🎯 Tryb edycji: używam {len(reference_images)} zdjęcia/zdjęć referencyjnych auta.")
        else:
            st.info("🎨 Tryb generacji: tworzę auto od zera na podstawie opisu (jakość może być niższa dla bardzo nowych modeli - rozważ dodanie zdjęcia referencyjnego).")
        
        with st.spinner("Gemini 2.5 Flash Image renderuje Twoje auto w ciemnym garażu LED..."):
            img_data = generate_ai_image(
                car_description=car_description,
                color_description=color_description,
                reference_images=reference_images if reference_images else None
            )
            if img_data:
                st.session_state['ai_img'] = img_data
                
    st.markdown("---")
    st.markdown("### Dodatki do oferty")
    dodatki_dostepne = [f for f in pliki_na_dysku if f['name'].startswith(('4','5'))]
    wybrane_dodatki = [d for d in sorted(dodatki_dostepne, key=lambda x: x['name']) if st.checkbox(d['name'], value=False)]


# ZAKŁADKI W GŁÓWNYM PANELU
tab_kreator, tab_rejestr = st.tabs(["⚙️ Kreator Ofert", "📋 Ewidencja (Rejestr)"])

with tab_kreator:
    # Nagłówek w stylu okładek brand manualu
    st.markdown("""
    <div class="iw-main-header">
        <div class="iw-main-header-text">
            <div class="iw-main-header-title">Generator Ofert</div>
            <div class="iw-main-header-subtitle">Professional Car Wrapping &amp; PPF · IT'S WRAP</div>
        </div>
    </div>
    """, unsafe_allow_html=True)
    col1, col2 = st.columns(2)

    with col1:
        klient = st.text_input("Imię i Nazwisko Klienta")
        nr_o = st.text_input("Numer oferty", value=f"IW/{datetime.now().strftime('%Y/%m/%d')}/01")
        
        kategorie = [k for k in df_cennik['Kategoria'].unique() if str(k).strip() != ""]
        kategoria = st.selectbox("Kategoria", kategorie)
        uslugi_kat = [u for u in df_cennik[df_cennik['Kategoria'] == kategoria]['Usługa'].unique() if str(u).strip() != ""]
        pakiet = st.selectbox("Usługa", uslugi_kat)
        
        try:
            wiersz_ceny = df_cennik[(df_cennik['Kategoria'] == kategoria) & (df_cennik['Usługa'] == pakiet) & (df_cennik['Segment'] == segment_final)]
            if not wiersz_ceny.empty:
                cena_str = str(wiersz_ceny['Cena sprzedaży netto PLN'].values[0])
                cena_str = cena_str.replace(' ', '').replace('\xa0', '')
                if ',' in cena_str: cena_str = cena_str.replace('.', '').replace(',', '.')
                cena_domyslna = float(re.sub(r'[^\d.]', '', cena_str))
            else:
                cena_domyslna = 0.0
        except:
            cena_domyslna = 0.0

        st.markdown("---")
        st.write("💰 **Kalkulacja cenowa**")
        
        cena_manual = st.number_input("Cena bazowa NETTO (PLN) - możesz edytować", value=cena_domyslna, step=100.0)
        rabat = st.number_input("Rabat dla klienta (PLN)", value=0.0, step=100.0)
        cena_koncowa = cena_manual - rabat
        
        st.info(f"**Cena do zapłaty netto (na ofercie): {cena_koncowa:,.2f} zł**".replace(',', 'X').replace('.', ',').replace('X', ' '))

    with col2:
        if 'ai_img' in st.session_state:
            st.image(st.session_state['ai_img'], use_container_width=True)
        else:
            st.info("Skonfiguruj auto w panelu bocznym i wygeneruj zdjęcie, aby zobaczyć podgląd.")

    # --- GENEROWANIE OFERTY ---
    if st.button("🔥 GENERUJ PEŁNĄ OFERTĘ PDF"):
        if 'ai_img' not in st.session_state:
            st.error("Wizualizacja auta jest wymagana. Użyj przycisku w panelu bocznym!")
        else:
            with st.spinner("Składam profesjonalny PDF i wgrywam na Dysk Google..."):
                writer = PdfWriter()
                final_foil_text = f"{f_color} (na lakier: {paint_color})" if "Bezbarwne" in f_cat else f_color
                
                dane_handlowca = HANDLOWCY[wybrany_handlowiec]
                wygenerowany_wstep = generate_ai_intro_text(klient, final_brand, final_model, pakiet, final_foil_text, wybrany_handlowiec, dane_handlowca["stanowisko"])

                cena_koncowa_str = f"{cena_koncowa:,.2f} zł".replace(',', 'X').replace('.', ',').replace('X', ' ')
                cena_katalogowa_str = f"{cena_manual:,.2f} zł".replace(',', 'X').replace('.', ',').replace('X', ' ')

                replacements = {
                    "{{KLIENT}}": klient, 
                    "{{MODEL_AUTA}}": f"{final_brand} {final_model}",
                    "{{RODZAJ_FOLII}}": final_foil_text, 
                    "{{USLUGA_NAZWA}}": pakiet,
                    "{{NR_OFERTY}}": nr_o,
                    "{{CENA_KATALOG}}": cena_katalogowa_str,
                    "{{CENA_KONCOWA}}": cena_koncowa_str,
                    "{{WSTEP_AI}}": wygenerowany_wstep,
                    "{{HANDLOWIEC_IMIE}}": wybrany_handlowiec,
                    "{{HANDLOWIEC_TEL}}": dane_handlowca["telefon"],
                    "{{HANDLOWIEC_EMAIL}}": dane_handlowca["email"]
                }

                okladka = next((f for f in pliki_na_dysku if f['name'].startswith('1_')), None)
                wstep_slide = next((f for f in pliki_na_dysku if f['name'].lower().startswith('1b_')), None)
                
                produkt = None
                if "reklam" in pakiet.lower():
                    produkt = next((f for f in pliki_na_dysku if f['name'].startswith('2') and 'reklama' in f['name'].lower()), None)
                elif f_brand == "3M 2080 Series":
                    produkt = next((f for f in pliki_na_dysku if f['name'].startswith('2') and '3m' in f['name'].lower() and 'kolor' in f['name'].lower()), None)
                elif "Ultimate" in f_color: 
                    produkt = next((f for f in pliki_na_dysku if f['name'].startswith('2') and 'ultimate' in f['name'].lower()), None)
                elif "Stealth" in f_color: 
                    produkt = next((f for f in pliki_na_dysku if f['name'].startswith('2') and 'stealth' in f['name'].lower()), None)
                elif "Color" in f_cat: 
                    produkt = next((f for f in pliki_na_dysku if f['name'].startswith('2') and 'color' in f['name'].lower()), None)
                
                if not produkt: 
                    produkt = next((f for f in pliki_na_dysku if f['name'].startswith('2')), None)
                
                if rabat > 0: 
                    zakres = next((f for f in pliki_na_dysku if f['name'].startswith('3') and 'bezrabatu' not in f['name'].lower()), None)
                else: 
                    zakres = next((f for f in pliki_na_dysku if f['name'].startswith('3') and 'bezrabatu' in f['name'].lower()), None)
                
                if not zakres: 
                    zakres = next((f for f in pliki_na_dysku if f['name'].startswith('3')), None)

                koniec = next((f for f in pliki_na_dysku if f['name'].startswith('6')), None)

                seq = [okladka, wstep_slide, produkt, zakres] + wybrane_dodatki + [koniec]
                seq = [f for f in seq if f]

                for f_info in seq:
                    prs = Presentation(download_file(service, f_info['id']))
                    for slide in prs.slides:
                        if f_info['name'].startswith('1_'):
                            for shape in list(slide.shapes):
                                if "{{FOTO_AUTA}}" in shape.name or (shape.has_text_frame and "{{FOTO_AUTA}}" in shape.text):
                                    pic = slide.shapes.add_picture(io.BytesIO(st.session_state['ai_img']), shape.left, shape.top, shape.width, shape.height)
                                    slide.shapes._spTree.remove(pic._element)
                                    slide.shapes._spTree.insert(2, pic._element)
                                    shape._element.getparent().remove(shape._element)

                        for shape in slide.shapes:
                            replace_text_in_shape(shape, replacements)

                    tmp_p = f"tmp_{f_info['id']}.pptx"
                    prs.save(tmp_p)
                    pdf = pptx_to_pdf(tmp_p)
                    if pdf: writer.append(pdf); os.remove(tmp_p); os.remove(pdf)

                final_io = io.BytesIO(); writer.write(final_io); final_io.seek(0)
                nazwa_pliku_wyjsciowego = f"Oferta_{final_brand}_{final_model}_{datetime.now().strftime('%H%M%S')}.pdf"
                
                # Upload PDF na prywatny Drive użytkownika (OAuth, nie service account)
                oauth_drive = get_oauth_drive_service()
                if oauth_drive:
                    folder_oferty_id = pobierz_lub_stworz_folder_oferty_oauth(oauth_drive)
                    utworzony_link = wgraj_pdf_na_dysk(oauth_drive, folder_oferty_id, nazwa_pliku_wyjsciowego, final_io.getvalue())
                else:
                    utworzony_link = None
                    st.warning("OAuth Drive niedostępny - PDF nie został zapisany w chmurze. Pobierz plik lokalnie poniżej.")
                
                link_do_zapisu = utworzony_link if utworzony_link else "Błąd uploadu"

                if zapisz_do_rejestru(nr_o, wybrany_handlowiec, klient, f"{final_brand} {final_model}", pakiet, final_foil_text, cena_koncowa, link_do_zapisu):
                    st.success(f"✅ Oferta zapisana w systemie CRM! Plik został zachowany w folderze 'Oferty'.")
                    
                st.balloons()
                st.download_button("📥 POBIERZ OFERTĘ PDF LOKALNIE", data=final_io, file_name=nazwa_pliku_wyjsciowego)

with tab_rejestr:
    st.markdown("""
    <div class="iw-main-header">
        <div class="iw-main-header-text">
            <div class="iw-main-header-title">Ewidencja Ofert</div>
            <div class="iw-main-header-subtitle">Rejestr · Archiwum · Kontrola</div>
        </div>
    </div>
    """, unsafe_allow_html=True)
    try:
        sheet_rejestr_view = client.open_by_url(LINK_DO_ARKUSZA).worksheet("Rejestr")
        dane_rejestru = sheet_rejestr_view.get_all_records()
        if dane_rejestru:
            df_rejestr = pd.DataFrame(dane_rejestru)
            
            nazwa_kolumny_link = df_rejestr.columns[-1]
            
            st.data_editor(
                df_rejestr,
                column_config={
                    nazwa_kolumny_link: st.column_config.LinkColumn(
                        "Plik PDF", help="Kliknij aby pobrać dokument ofertowy z dysku", display_text="Otwórz dokument"
                    )
                },
                hide_index=True,
                use_container_width=True
            )
            
            csv = df_rejestr.to_csv(index=False).encode('utf-8')
            st.download_button(
                label="⬇️ Pobierz ewidencję jako CSV",
                data=csv,
                file_name=f"Rejestr_Ofert_ITSWRAP_{datetime.now().strftime('%Y%m%d')}.csv",
                mime='text/csv',
            )
        else:
            st.info("Rejestr jest pusty lub arkusz nie zawiera jeszcze wpisów.")
    except Exception as e:
        st.warning(f"Brak możliwości wczytania rejestru. Pamiętaj, by w nagłówkach bazy (Arkusz 'Rejestr', 1 wiersz) dodać kolumnę na link do PDF. Błąd: {e}")

# --- STOPKA ---
st.markdown("""
<div class="iw-footer">
    <strong>IT'S WRAP</strong> · ONDRE.PL · Rynek Śródecki 4, 61-126 Poznań · +48 602 494 133
    <br/>
    <span style="opacity: 0.6; letter-spacing: 0.03em;">www.itswrap.pl · Professional Car Wrapping &amp; PPF · Certyfikowana jakość 3M</span>
</div>
""", unsafe_allow_html=True)
