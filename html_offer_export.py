"""Eksperymentalny eksport HTML, niezależny od istniejącego generowania PDF.

Nie publikuje oferty, nie komunikuje się z zewnętrznymi usługami.
"""
from html import escape
from datetime import datetime


def export_offer_html(offer):
    """Zwraca samodzielny dokument HTML jako bytes UTF-8."""
    def safe(key, default=""):
        value = offer.get(key, default)
        return escape(str(value if value is not None else ""), quote=True)

    number = safe("nr_o", "Oferta testowa")
    customer = safe("klient", "Klient")
    vehicle = safe("auto", "Pojazd")
    package = safe("pakiet", "Usługa")
    mode = safe("tryb", "B2C")
    price = offer.get("cena")
    try:
        price_label = f"{float(price):,.2f}".replace(",", " ").replace(".", ",") + " zł"
    except (TypeError, ValueError):
        price_label = "Wycena indywidualna"
    price_label = escape(price_label)
    price_type = "netto" if mode == "B2B" else "— sprawdź sposób naliczania VAT w kalkulacji"
    created = datetime.now().strftime("%d.%m.%Y")
    html = """<!doctype html>
<html lang="pl"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<meta name="robots" content="noindex, nofollow">
<title>IT'S WRAP — {number}</title>
<style>
:root{{--blue:#007dc5;--ink:#091827;--muted:#688092}}
*{{box-sizing:border-box}} body{{margin:0;font-family:Arial,Helvetica,sans-serif;background:#f0f4f7;color:var(--ink)}}
header{{background:#071b2d;color:white;padding:24px max(24px,calc((100vw - 1080px)/2));display:flex;justify-content:space-between;gap:24px;align-items:center}}
.brand{{font-weight:900;font-size:27px;letter-spacing:-1px}} .brand span{{color:#27a9f0}}
main{{max-width:1080px;margin:0 auto;padding:48px 24px 80px}}
.hero{{background:#071b2d;color:white;border-radius:28px;padding:64px 48px;margin-bottom:30px}}
.eyebrow{{font-size:13px;letter-spacing:2px;text-transform:uppercase;color:#70caff;font-weight:700}}
h1{{font-size:clamp(36px,6vw,70px);line-height:1.03;letter-spacing:-2px;margin:18px 0 24px}}
h2{{font-size:26px;margin:0 0 20px}}.hero p{{font-size:19px;color:#d1e1ec}}
.grid{{display:grid;grid-template-columns:1fr 1fr;gap:20px}}
.card{{background:white;border-radius:22px;padding:32px;margin-bottom:20px}}
.label{{font-size:12px;font-weight:bold;letter-spacing:1.5px;color:var(--muted);text-transform:uppercase}}
.value{{font-size:23px;font-weight:bold;margin-top:12px}}
.price{{font-size:40px;font-weight:800;color:var(--blue);margin:16px 0 4px}}
footer{{max-width:1080px;margin:auto;padding:0 24px 40px;color:#63788b;font-size:14px}}
@media(max-width:700px){{header{{padding:20px}}.hero{{padding:38px 26px}}.grid{{grid-template-columns:1fr}}.card{{padding:24px}}}}
</style></head><body>
<header><div class="brand">IT'S <span>WRAP</span></div><small>MAKE IT CHANGE</small></header>
<main>
<section class="hero"><div class="eyebrow">Oferta indywidualna · {number}</div>
<h1>{vehicle}</h1><p>Przygotowaliśmy propozycję dla: <strong>{customer}</strong>.</p></section>
<div class="grid">
<section class="card"><div class="label">Proponowana usługa</div><div class="value">{package}</div></section>
<section class="card"><div class="label">Rodzaj oferty</div><div class="value">{mode}</div></section>
</div>
<section class="card"><h2>Podsumowanie wyceny</h2>
<div class="label">Cena z obecnego generatora</div><div class="price">{price_label}</div>
<p>{price_type}</p><small>Data eksportu: {created}</small></section>
<section class="card"><h2>Co dalej?</h2>
<p>Ta strona jest wersją testową eksportu HTML z generatora IT'S WRAP. Warianty, wizualizacje, zatwierdzone opisy produktowe oraz publikację unikalnego linku dodamy w kolejnych etapach.</p></section>
</main><footer>IT'S WRAP · wersja testowa HTML · dokument nieopublikowany</footer>
</body></html>""".format(number=number, vehicle=vehicle, customer=customer,
                    package=package, mode=mode, price_label=price_label,
                    price_type=price_type, created=created)
    return html.encode("utf-8")
