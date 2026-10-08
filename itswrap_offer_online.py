"""Additional online offer renderer/publisher. Never changes PDF, prices or CRM."""
import base64
import hashlib
import html
import io
import json
import re
from decimal import Decimal, ROUND_HALF_UP
from pathlib import Path
from urllib.parse import quote

PUBLISH_URL = 'https://pacanowski.pro/api/oferty'


def esc(value):
    return html.escape(str(value if value is not None else ''), quote=True)


def money(value):
    value = Decimal(str(value)).quantize(Decimal('.01'), rounding=ROUND_HALF_UP)
    return f'{value:,.2f}'.replace(',', ' ').replace('.', ',') + ' zł'


def paragraphs(value):
    return ''.join('<p>' + esc(p).replace('\n', '<br>') + '</p>' for p in re.split(r'\n\s*\n', str(value or '').strip()) if p.strip())


def image_uri(raw):
    if not isinstance(raw, (bytes, bytearray)) or not raw:
        return ''
    # Keep the actual generated image, never borrow an image from another offer.
    from PIL import Image, ImageOps
    im = ImageOps.exif_transpose(Image.open(io.BytesIO(raw))).convert('RGB')
    im.thumbnail((1920, 1920))
    out = io.BytesIO(); im.save(out, format='JPEG', quality=88)
    return 'data:image/jpeg;base64,' + base64.b64encode(out.getvalue()).decode()


def pdf_pages(raw):
    if not raw:
        return []
    from pypdf import PdfReader
    return [(p.extract_text() or '').strip() for p in PdfReader(io.BytesIO(raw)).pages]


# Qualitative descriptions based on the current "Folie i karty" brief.
# Unconfirmed technical numbers, warranties and evidence claims are not promoted to facts.
CATALOG = [
    ('3M PPF Series 100', ('3m', '100'), 'Bezbarwna folia PPF do ochrony lakieru z zachowaniem jego koloru.', 'Nowe auto lub lakier zakwalifikowany do aplikacji.', 'Ochrona nie oznacza odporności na wszystkie uszkodzenia. Jeśli celem jest zmiana koloru, porównaj folię kolorową lub kolorowy PPF.'),
    ('XPEL Ultimate Plus', ('xpel', 'ultimate'), 'Bezbarwna folia PPF dla klienta, który chce zachować wygląd lakieru i ograniczyć ślady eksploatacji.', 'Zachowanie połysku i koloru lakieru.', 'Parametry i gwarancję potwierdzamy dla dokładnej wersji. Ultimate Plus, Fusion i Ultimate Plus 10 nie są tym samym produktem.'),
    ('XPEL Xtreme', ('xpel', 'xtreme'), 'Bezbarwna folia PPF jako wariant ochrony lakieru.', 'Klient porównujący warianty ochrony.', 'Porównanie z innymi liniami wymaga potwierdzenia konkretnej specyfikacji i dostępności.'),
    ('XPEL Stealth', ('xpel', 'stealth'), 'Matowe lub satynowe wykończenie PPF łączy zmianę wyglądu z ochroną powierzchni.', 'Efekt matu lub satyny.', 'Docelowy efekt potwierdzamy na próbce dla lakieru klienta.'),
    ('XPEL Color', ('xpel', 'color'), 'Kolorowy PPF łączy zmianę koloru z funkcją ochronną.', 'Kolor i ochrona w jednym rozwiązaniu.', 'Kolor, wykończenie i warunki ochrony wymagają potwierdzenia dla wybranej wersji. Nie przenosimy warunków bezbarwnego PPF na kolorowy.'),
    ('3M Wrap Film 2080', ('3m', '2080'), 'Folia do zmiany koloru i wykończenia pojazdu.', 'Metamorfoza wyglądu samochodu.', 'Efekt wybieramy na próbniku. Folia do zmiany koloru nie zastępuje ochrony PPF.'),
    ('PWF Style', ('pwf', 'style'), 'Folia do zmiany koloru z wybranej linii PWF.', 'Indywidualny kolor i wykończenie.', 'Oczekiwany efekt potwierdzamy na próbniku. PWF Style i PWF Performance wymagają osobnych specyfikacji.'),
    ('PWF Performance', ('pwf', 'performance'), 'Wariant PWF łączący kolor z funkcją ochronną.', 'Kolor i ochrona po kwalifikacji.', 'Zakres dobieramy po sprawdzeniu dokładnej wersji i przeznaczenia produktu.'),
    ('Orajet 3551 + Oraguard 215', ('3551',), 'Folia do drukowanej grafiki reklamowej z dobranym laminatem.', 'Grafika na zakwalifikowanych powierzchniach.', 'Dobór zależy od kształtu powierzchni i czasu kampanii. Głębokie przetłoczenia wymagają osobnej kwalifikacji. Program 3M MCS nie dotyczy materiałów Orafol.'),
    ('Oracal 551', ('551',), 'Folia barwiona do wycinanych napisów, logotypów i elementów grafiki.', 'Grafika wektorowa i napisy.', 'Dobór koloru i zakresu cięcia zależy od projektu. To osobny produkt od folii do druku 3551.'),
    ('3M IJ180-10 + laminat', ('ij180-10',), 'Biała folia do drukowanej grafiki reklamowej, dobierana jako część kompletnego systemu.', 'Grafika z kryjącym tłem.', 'Rekomendacja zależy od powierzchni, projektu i użytkowania. Gwarancja systemowa wymaga zgodności folii, laminatu, drukarki i tuszu.'),
    ('3M IJ180-114 + laminat', ('ij180-114',), 'Przezroczysta folia pozwala wykorzystać kolor lakieru jako tło grafiki. Biały poddruk, jeśli wybrany, pozwala sterować kryciem.', 'Projekt wykorzystujący tło lakieru lub biały poddruk.', 'Efekt druku potwierdzamy dla koloru pojazdu i projektu. Bez bieli kolor podłoża wpływa na odbiór grafiki.'),
    ('Avery Dennison V-4000', ('v-4000',), 'Materiał odblaskowy pozwala uzyskać grafikę reagującą na padające światło.', 'Projekt wymagający efektu odblaskowego.', 'Dobór zależy od przeznaczenia i kompatybilności systemu druku. Nie stanowi automatycznie homologowanego oznakowania.'),
    ('OWV', ('owv',), 'Perforowana folia do grafiki na zakwalifikowanej szybie.', 'Grafika na odpowiedniej szybie.', 'Konkretny produkt, widoczność i zakres aplikacji potwierdzamy przed realizacją.'),
]

CSS = '''
:root{--blue:#007dc5;--navy:#172a46;--ink:#182029;--muted:#65717c;--line:#e1e7ed}*{box-sizing:border-box}html{scroll-behavior:smooth;scroll-padding-top:95px}body{margin:0;background:#eef2f5;color:var(--ink);font:16px/1.65 Arial,Helvetica,sans-serif}a{color:inherit}main{max-width:1280px;margin:auto;background:white}header{padding:22px 6%;display:flex;align-items:center;justify-content:space-between;gap:20px;border-bottom:3px solid var(--blue);background:white}header img{width:200px;max-width:45vw}.ref{font-size:12px;color:var(--muted)}.nav{display:flex;gap:24px;padding:15px 6%;position:sticky;top:0;background:#fffefc;z-index:4;border-bottom:1px solid var(--line);overflow:auto;font-size:13px;white-space:nowrap}.nav a{text-decoration:none;font-weight:bold}.hero{position:relative;background:#12223a;min-height:610px;display:flex;align-items:end;color:#fff;overflow:hidden;border-radius:0 0 36px 36px}.hero>img{position:absolute;width:100%;height:100%;object-fit:cover;inset:0;opacity:.76}.hero:after{content:'';position:absolute;inset:0;background:linear-gradient(90deg,#07111cf2,#07111c66 75%),linear-gradient(0deg,#07111cbb,transparent)}.hero-content{position:relative;z-index:1;padding:75px 7%;max-width:900px}h1{font-size:clamp(42px,6vw,76px);line-height:1.06;letter-spacing:-.045em;margin:18px 0 25px}h2{font-size:clamp(28px,3.5vw,44px);line-height:1.15;letter-spacing:-.035em;margin:8px 0 25px}h3{font-size:21px;line-height:1.3;margin:8px 0 14px}.kicker{font-size:11px;letter-spacing:.17em;text-transform:uppercase;font-weight:bold;color:var(--blue)}.hero .kicker,.dark .kicker{color:#72caff}.line{width:100px;height:4px;background:var(--blue);margin:24px 0}.section{padding:70px 7%}.soft{background:#f3f7fa}.dark{background:var(--navy);color:white}.lead{font-size:21px}.muted,figcaption{color:var(--muted);font-size:13px}.dark .muted{color:#c5d3e2}.grid{display:grid;grid-template-columns:1fr 1fr;gap:28px}.cards{display:grid;grid-template-columns:repeat(3,1fr);gap:22px}.card{background:#f2f6f9;border:1px solid #e3ebf1;border-radius:22px;padding:28px}.dark .card{background:#ffffff09;border-color:#ffffff24}.pricecard{border-radius:26px;overflow:hidden;border:1px solid var(--line);display:grid;grid-template-columns:1.15fr .85fr;background:white}.price-info{padding:35px}.price-total{background:var(--navy);color:white;padding:35px;display:flex;flex-direction:column;justify-content:center}.amount{font-size:clamp(32px,4vw,52px);font-weight:bold;line-height:1.2;letter-spacing:-.04em}.price-total small{color:#c9d7e6}.old{color:#b5c7d9;text-decoration:line-through}.tag{display:inline-block;background:#eaf5fc;color:#076897;border-radius:20px;padding:5px 12px;font-size:12px;font-weight:bold;margin:4px 8px 4px 0}.visual{width:100%;max-height:760px;object-fit:contain;background:#edf1f4;border-radius:24px}figure{margin:25px 0}figcaption{margin-top:10px}dl{margin:0}dt{font-size:12px;text-transform:uppercase;letter-spacing:.04em;color:var(--muted);margin-top:18px}dd{margin:4px 0 14px;font-size:17px}.btn{display:inline-block;padding:14px 22px;background:var(--blue);color:white;text-decoration:none;font-weight:bold;border-radius:5px;margin:10px 10px 0 0}.btn.light{background:white;color:var(--navy)}.flow{counter-reset:step;display:grid;grid-template-columns:repeat(3,1fr);gap:20px}.flow>div{counter-increment:step;padding:24px;border-top:2px solid var(--blue)}.flow>div:before{content:'0' counter(step);font-size:30px;color:var(--blue);font-weight:bold}.flow p{margin-bottom:0}.source{white-space:pre-wrap;line-height:1.75;font-size:15px}details{border-bottom:1px solid var(--line);padding:20px 0}summary{cursor:pointer;font-size:18px;font-weight:bold}.notice{border-left:3px solid var(--blue);padding:16px 22px;background:#f3f7fa}.contact{background:linear-gradient(110deg,#123e63,#172a46);color:white;border-radius:32px 32px 0 0}.contact a{overflow-wrap:anywhere}footer{padding:25px 7%;background:#0a1523;color:#a9bacb;font-size:12px}p{margin:0 0 18px}table{border-collapse:collapse;width:100%}td,th{text-align:left;padding:12px;border-bottom:1px solid var(--line)}.table{overflow-x:auto}@media(max-width:760px){.section{padding:48px 6%}.hero{min-height:520px}.hero-content{padding:55px 7%}.grid,.pricecard,.cards,.flow{grid-template-columns:1fr}.ref{max-width:45%;overflow-wrap:anywhere}.nav{gap:18px}.price-info,.price-total{padding:26px}.hero>img{object-position:65% center}h1{font-size:43px}}@media print{.nav,.btn{display:none}.hero{min-height:350px}.section{padding:30px}.card,figure,details{break-inside:avoid}details>*{display:block}body{background:white}}
'''


def generate_offer_html(offer, logo_path='', *, vat_rate='23', conditions=''):
    data = offer.get('dane_oferty') or {}
    if isinstance(data, str): data = json.loads(data)
    snap = offer.get('_html_snapshot') or {}
    fields = snap.get('fields') or {}
    get = lambda key, default='': fields.get('{{' + key + '}}') or default
    b2b = (offer.get('tryb') or data.get('tryb')) == 'B2B'
    mode = 'B2B · FLEET' if b2b else 'B2C · PREMIUM'
    client = offer.get('klient', '')
    car = offer.get('auto', '')
    service = offer.get('pakiet', '')
    foil = get('RODZAJ_FOLII') or ' · '.join(filter(None, [data.get('f_brand'), data.get('f_cat'), data.get('f_color')]))
    intro = data.get('wstep') or get('WSTEP_AI')
    visual = image_uri(snap.get('visual'))
    pages = pdf_pages(offer.get('pdf_bytes'))
    rate = Decimal(str(vat_rate)); net = Decimal(str(offer['cena']))
    if not rate.is_finite() or not 0 <= rate <= 100 or not net.is_finite(): raise ValueError('Nieprawidłowa kwota lub VAT')
    gross = (net * (1 + rate / 100)).quantize(Decimal('.01'), rounding=ROUND_HALF_UP)
    shown = net if b2b else gross
    catalog = snap.get('catalog_net')
    old = ''
    if catalog is not None and Decimal(str(catalog)) > net:
        old_value = Decimal(str(catalog)) * (1 if b2b else 1 + rate/100)
        old = f'<div class="old">{money(old_value)}</div><small>Cena przed rabatem</small>'
    logo = ''
    if logo_path and Path(logo_path).is_file():
        logo = base64.b64encode(Path(logo_path).read_bytes()).decode()
    else:
        from itswrap_html_export_dev import _LOGO_BASE64
        logo = _LOGO_BASE64
    title = 'Twoja marka w ruchu.' if b2b else ('Ochrona z dowodem.' if 'ppf' in (service + foil).lower() else 'Zmiana z przewagą.')
    sections = []
    def section(key, label, title, body, cls=''):
        sections.append(f'<section id="{key}" class="section {cls}"><span class="kicker">{esc(label)}</span><h2>{esc(title)}</h2><div class="line"></div>{body}</section>')
    name = get('HANDLOWIEC_IMIE', data.get('handlowiec', ''))
    phone = get('HANDLOWIEC_TEL'); email = get('HANDLOWIEC_EMAIL')
    contact_links = ''
    if phone: contact_links += f'<a class="btn" href="tel:{esc(re.sub(r"[^+0-9]", "", phone))}">{esc(phone)}</a>'
    if email and re.fullmatch(r'[^\s<>@]+@[^\s<>@]+\.[^\s<>@]+', email):
        contact_links += f'<a class="btn light" href="mailto:{esc(email)}?subject={quote(str(offer.get("nr_o", "Oferta")))}">Napisz do doradcy</a>'
    section('rekomendacja', '01 · Rekomendacja doradcy', 'Rozwiązanie dobrane do Twojego auta.' if not b2b else 'Rozwiązanie dobrane do Twojej firmy.', paragraphs(intro) + f'<p><b>Kontakt w sprawie oferty: {esc(name)}</b></p>')
    scope = get('ZAKRES_OPIS') or service
    quantity = int(data.get('liczba_pojazdow') or 1)
    fleet = f'<dt>Liczba pojazdów</dt><dd>{quantity}</dd>' if b2b else ''
    if b2b and snap.get('unit_net') is not None:
        fleet += f'<dt>Cena jednostkowa przed rabatem łącznym</dt><dd>{money(snap["unit_net"])} netto</dd>'
    vat_label = f'{money(net)} netto · VAT {esc(rate)}%' if not b2b else f'VAT {esc(rate)}% · {money(gross)} brutto'
    summary = f'<div class="pricecard"><div class="price-info"><span class="tag">{esc(mode)}</span><h3>{esc(service)}</h3><dl><dt>Pojazd</dt><dd>{esc(car)}</dd><dt>Materiał / wykończenie</dt><dd>{esc(foil)}</dd><dt>Zakres</dt><dd>{esc(scope)}</dd>{fleet}</dl></div><div class="price-total">{old}<p>Cena łączna {"netto" if b2b else "brutto"}</p><div class="amount">{money(shown)}</div><small>{vat_label}</small><a class="btn" href="#kontakt">Zapytaj o termin →</a></div></div>'
    condition_text = conditions.strip() or 'Termin: do uzgodnienia\nCzas realizacji: do uzgodnienia\nWażność oferty: do potwierdzenia\nPłatność / rezerwacja: do uzgodnienia\nGwarancja producenta: zgodnie z warunkami konkretnego produktu — do potwierdzenia\nOdpowiedzialność za montaż: do potwierdzenia przez doradcę\nDeklarowana trwałość: odrębna od gwarancji, zależna od produktu i użytkowania'
    if b2b and not conditions.strip(): condition_text += '\nPartie realizacji i przestoje: do uzgodnienia\nDemontaż / zwrot leasingu: po ocenie materiału i podłoża'
    summary += '<h3 style="margin-top:35px">Warunki realizacji</h3><div class="notice">' + paragraphs(condition_text) + '</div>'
    section('skrot', '02 · Oferta w skrócie', 'Wszystko, co potrzebne do decyzji.', summary, 'soft')
    if b2b: section('potrzeba', '03 · Potrzeba klienta', 'Spójny wizerunek w codziennym ruchu.', paragraphs(intro))
    caption = 'Wizualizacja poglądowa. Przed produkcją potwierdzamy projekt produkcyjny.' if b2b else 'Wizualizacja poglądowa. Kolor i wykończenie potwierdzamy na próbniku.'
    visual_body = f'<figure><a href="{visual}" target="_blank" rel="noopener"><img class="visual" src="{visual}" alt="Wizualizacja: {esc(car)} · {esc(foil)}"></a><figcaption>{caption}</figcaption></figure>' if visual else '<p>Ta oferta nie zawiera zapisanej wizualizacji. Doradca może dołączyć ją po ustaleniu projektu.</p>'
    def material():
        match_text = (foil + ' ' + str(data.get('zestaw_b2b',''))).lower()
        matches = [c for c in CATALOG if all(x in match_text for x in c[1])]
        # Never guess a different product because a name is missing or inconsistent.
        cards = ''.join(f'<article class="card"><h3>{esc(c[0])}</h3><p>{esc(c[2])}</p><h4>Dla kogo?</h4><p>{esc(c[3])}</p><h4>Dobór i ograniczenia</h4><p>{esc(c[4])}</p></article>' for c in matches)
        if not cards: cards = f'<article class="card"><h3>{esc(foil)}</h3><p>Dokładna specyfikacja wybranego materiału znajduje się w kartach przygotowanej oferty poniżej. Dobór produktu i warunki ochrony potwierdza doradca.</p></article>'
        # Product copy from the actual PDF, not a generic replacement.
        product_pages = [p for p in pages[2:] if re.search(r'foli[aeiyę]|laminat|powłok', p, re.I) and not re.search(r'cena (katalog|specjal|końc)', p, re.I)]
        details = ''.join('<details><summary>Opis materiału i informacje z oferty</summary><div class="source">' + esc(p) + '</div></details>' for p in product_pages)
        section('material', 'Materiał · zastosowanie · ograniczenia', 'Co proponujemy i dlaczego.', '<div class="grid">'+cards+'</div>'+details)
    if b2b:
        section('wizualizacja', 'Koncepcja i wizualizacja', 'Zobacz proponowany efekt.', visual_body)
        material()
        if re.search(r'druk|latex|lateks', foil + ' '.join(pages), re.I):
            section('technologia', 'Technologia produkcji', 'Od pliku do gotowej grafiki.', '<div class="cards"><div class="card"><h3>Projekt i pliki</h3><p>Weryfikacja projektu względem bryły pojazdu i przygotowanie do produkcji.</p></div><div class="card"><h3>System materiałowy</h3><p>Folia, druk i laminat dobrane do projektu oraz planowanego użytkowania. Biały poddruk wyłącznie, jeśli jest częścią wybranego zakresu.</p></div><div class="card"><h3>Kontrola i aplikacja</h3><p>Sprawdzenie gotowej grafiki, przygotowanie powierzchni i montaż zgodnie z uzgodnionym zakresem.</p></div></div>', 'dark')
    else:
        material(); section('wizualizacja', 'Wizualizacja', 'Twój samochód. Wybrany efekt.', visual_body)
    scope_pages = [p for p in pages if re.search(r'w cenie|wybrany zakres|zakres oferty', p, re.I)]
    scope_body = paragraphs(scope) + ''.join('<div class="card source">'+esc(p)+'</div>' for p in scope_pages)
    scope_body += '<h3 style="margin-top:30px">Cena końcowa</h3><p class="amount">'+money(shown)+(' netto' if b2b else ' brutto')+'</p><p class="muted">'+vat_label+'</p>'
    scope_body += '<p class="muted">Szczegóły poniżej zachowują treść przygotowanego PDF. Kwoty opisane tam jako netto pozostają netto. Nie doliczamy ponownie dodatków.</p>'
    section('zakres', 'Zakres i cena', 'Co obejmuje propozycja.', scope_body, 'soft')
    section('dlaczego', 'Dlaczego IT’S WRAP', 'Dobór. Wykonanie. Opieka.', '<div class="cards"><article class="card"><h3>Dobór do zastosowania</h3><p>Łączymy oczekiwany efekt z zakresem, stanem podłoża i sposobem użytkowania pojazdu.</p></article><article class="card"><h3>Ustalony zakres</h3><p>Przed pracą potwierdzamy materiał, kolor lub projekt oraz zakres montażu. Rozdzielamy trwałość, gwarancję producenta i odpowiedzialność za wykonanie.</p></article><article class="card"><h3>Kontakt z doradcą</h3><p>Masz wskazaną osobę do omówienia oferty, harmonogramu oraz zasad późniejszej pielęgnacji.</p></article></div>', 'dark')
    steps = [('Projekt','Ustalamy zakres i sposób oznakowania.'),('Akceptacja','Potwierdzamy projekt produkcyjny i materiał.'),('Produkcja','Przygotowujemy grafikę zgodnie z przyjętym systemem.'),('Montaż partiami','Uzgadniamy rotację i czas wyłączenia pojazdów.'),('Odbiór','Sprawdzamy wykonanie i przekazujemy auto.'),('Użytkowanie i demontaż','Warunki pielęgnacji i rozklejania ustalamy dla materiału oraz lakieru.')] if b2b else [('Oględziny','Oceniamy lakier i kwalifikujemy powierzchnie.'),('Przygotowanie','Przygotowujemy auto w zakresie wskazanym w ofercie.'),('Montaż','Aplikujemy wybrany materiał.'),('Kontrola','Sprawdzamy krawędzie i jakość wykończenia.'),('Odbiór','Omawiamy wykonany zakres.'),('Pielęgnacja','Przekazujemy zalecenia właściwe dla zastosowanej folii.')]
    section('proces', 'Proces i opieka', 'Wiesz, co dzieje się dalej.', '<div class="flow">'+''.join('<div><h3>'+a+'</h3><p>'+b+'</p></div>' for a,b in steps)+'</div>')
    addons = data.get('dodatki') or []
    if addons:
        section('dodatki','Rozszerzenia' if b2b else 'Dodatki','Wybrane elementy dodatkowe.','<ul>'+''.join('<li>'+esc(str(x).removesuffix('.pptx'))+'</li>' for x in addons)+'</ul><p>Ceny i zakres wybranych dodatków potwierdzamy w kalkulacji. Sam wybór karty nie oznacza dodatkowej opłaty ani wliczenia usługi w cenę.</p>', 'soft')
    # Preserve all existing PDF information, including non-text diagrams via original PDF.
    pdf_uri = 'data:application/pdf;base64,' + base64.b64encode(offer['pdf_bytes']).decode() if offer.get('pdf_bytes') else ''
    appendix = ''.join(f'<details><summary>Strona {i+1} · szczegółowe informacje</summary><div class="source">{esc(p)}</div></details>' for i,p in enumerate(pages) if p)
    if pdf_uri: appendix = f'<a class="btn" download="oferta.pdf" href="{pdf_uri}">Pobierz pełną ofertę PDF</a><p class="muted">Oryginalny dokument z wszystkimi kartami, grafikami i warunkami.</p>'+appendix
    section('szczegoly','Informacje dodatkowe','Pełna dokumentacja oferty.',appendix)
    section('kontakt','Twój doradca','Potwierdź zakres i harmonogram.' if b2b else 'Potwierdź wariant i termin.',f'<p class="lead">{esc(name)}</p><p>Porozmawiajmy o ofercie {esc(offer.get("nr_o"))} dla {esc(car)}.</p>'+contact_links, 'contact')
    hero_image = f'<img src="{visual}" alt="">' if visual else ''
    doc = f'''<!doctype html><html lang="pl"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><meta name="robots" content="noindex,nofollow,noarchive"><meta name="referrer" content="no-referrer"><title>Twoja oferta · IT’S WRAP</title><style>{CSS}</style></head><body><main><header><img alt="IT’S WRAP · MAKE IT CHANGE" src="data:image/png;base64,{logo}"><span class="ref">{esc(offer.get('nr_o'))}<br>{esc(client)}</span></header><nav class="nav" aria-label="Sekcje oferty"><a href="#skrot">Oferta w skrócie</a><a href="#wizualizacja">Wizualizacja</a><a href="#material">Materiał</a><a href="#zakres">Zakres i cena</a><a href="#kontakt">Kontakt</a></nav><div class="hero">{hero_image}<div class="hero-content"><span class="kicker">MAKE IT CHANGE · {mode}</span><h1>{title}</h1><p class="lead">{esc(car)}</p><p>{esc(service)}<br>Propozycja dla: {esc(client)}</p><a class="btn" href="#skrot">Zobacz swoją ofertę →</a></div></div>{''.join(sections)}<footer>IT’S WRAP · MAKE IT CHANGE · {esc(offer.get('nr_o'))}</footer></main></body></html>'''
    return doc.encode('utf-8')


def publish_offer(document, key):
    import requests
    if not key: raise ValueError('Połącz publikację kluczem w sekcji oferty online.')
    if len(document) > 20 * 1024 * 1024: raise ValueError('Oferta przekracza limit 20 MB.')
    response = requests.post(PUBLISH_URL, data=document, headers={'Authorization':'Bearer '+key, 'Content-Type':'text/html; charset=utf-8', 'Idempotency-Key':hashlib.sha256(document).hexdigest()}, timeout=60)
    if response.status_code == 401: raise ValueError('Klucz publikacji jest nieprawidłowy.')
    if response.status_code == 413: raise ValueError('Oferta przekracza limit publikacji.')
    response.raise_for_status()
    url = response.json().get('url', '')
    if not re.fullmatch(r'https://pacanowski\.pro/oferty/[a-f0-9]{32}', url): raise ValueError('Serwer nie potwierdził adresu oferty.')
    return url


def render_online_offer(st, offer):
    """All optional work happens here; exceptions must never break the legacy UI."""
    with st.expander('Oferta online · HTML', expanded=False):
        st.caption('Pełna oferta z wizualizacją i indywidualnym linkiem na pacanowski.pro.')
        try: key = str(st.secrets.get('ITSWRAP_PUBLISH_KEY', ''))
        except Exception: key = ''
        if not key:
            st.markdown('Jednorazowe połączenie: [pobierz klucz publikacji](https://pacanowski.pro/oferty/polaczenie). Klucz możesz zapisać w Streamlit Secrets jako `ITSWRAP_PUBLISH_KEY` lub wkleić poniżej na czas tej sesji.')
            key = st.text_input('Klucz publikacji', type='password', key='itswrap_publish_key_session')
        offer_key = hashlib.sha256(str(offer.get('nr_o')).encode()+offer.get('pdf_bytes',b'')).hexdigest()[:16]
        rate = st.number_input('VAT do prezentacji ceny brutto (%)',min_value=0.0,max_value=100.0,value=23.0,step=1.0,key='html_vat_'+offer_key)
        conditions = st.text_area('Warunki realizacji (opcjonalne uzupełnienie oferty online)',value='',placeholder='Termin:\nCzas realizacji:\nWażność:\nPłatność / rezerwacja:\nGwarancja producenta:\nOdpowiedzialność za montaż:\nDeklarowana trwałość:',key='html_conditions_'+offer_key)
        st.caption('Puste warunki pojawią się jako „do uzgodnienia / potwierdzenia”. Dotychczasowy PDF i kalkulacja pozostają bez zmian.')
        current_key = offer_key + hashlib.sha256((str(rate)+conditions).encode()).hexdigest()[:12]
        if st.button('Generuj link do oferty',key='html_publish_'+offer_key,disabled=not bool(key)):
            with st.spinner('Przygotowuję treści, wizualizację i link…'):
                document = generate_offer_html(offer,vat_rate=rate,conditions=conditions)
                url = publish_offer(document,key)
                st.session_state['html_published_'+current_key] = url
        url = st.session_state.get('html_published_'+current_key)
        if url:
            st.success('Oferta opublikowana. Link możesz przekazać klientowi.')
            st.code(url,language=None)
            st.link_button('Otwórz ofertę online',url)
