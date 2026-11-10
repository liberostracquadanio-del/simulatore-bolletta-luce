"""Aggiorna il simulatore della bolletta della luce con l'ultimo trimestre ARERA.

1. trova sulla pagina ARERA «Prezzi e tariffe» il file delle condizioni economiche per i clienti
   vulnerabili (E<anno>-smt.xlsx) e lo scarica (oppure usa --xlsx con un file locale);
2. legge il primo foglio (il trimestre più recente) e controlla struttura e valori;
3. rigenera index.html dal modello e salva i corrispettivi in tariffe.json;
4. scrive un riepilogo (Markdown) con i numeri vecchi e nuovi del cliente tipo ARERA.

Esce con errore, senza toccare index.html, se un controllo non passa.

Uso: python generatore/aggiorna.py [--xlsx FILE] [--riepilogo FILE]
"""
import argparse
import json
import re
import sys
import tempfile
import urllib.request
from pathlib import Path

import openpyxl

QUI = Path(__file__).resolve().parent
RADICE = QUI.parent
PAGINA_ARERA = "https://www.arera.it/area-operatori/prezzi-e-tariffe"
UA = {"User-Agent": "Mozilla/5.0 (simulatore-bolletta-luce)"}
MESI = {"gen": "gennaio", "feb": "febbraio", "mar": "marzo", "apr": "aprile", "mag": "maggio", "giu": "giugno",
        "lug": "luglio", "ago": "agosto", "set": "settembre", "ott": "ottobre", "nov": "novembre", "dic": "dicembre"}
FONTS = "https://fonts.googleapis.com/css2?family=Source+Serif+4:opsz,wght@8..60,300;8..60,400;8..60,600;8..60,700&display=swap"
COLORI = {"__BLU__": "#009cdf", "__BLU_M__": "#83c3ef", "__BLU_C__": "#b2d8f6", "__SCURO__": "#005f8a"}


class ErroreControllo(Exception):
    pass


def scarica():
    with urllib.request.urlopen(urllib.request.Request(PAGINA_ARERA, headers=UA), timeout=60) as r:
        pagina = r.read().decode("utf8", "replace")
    link = sorted(set(re.findall(r'href="([^"]*/E(\d{4})-smt\.xlsx)"', pagina)), key=lambda x: x[1])
    if not link:
        raise ErroreControllo(f"Nella pagina {PAGINA_ARERA} non c'è più un link a un file E<anno>-smt.xlsx")
    url = link[-1][0]
    url = url if url.startswith("http") else "https://www.arera.it" + url
    dest = Path(tempfile.mkdtemp()) / url.rsplit("/", 1)[1]
    with urllib.request.urlopen(urllib.request.Request(url, headers=UA), timeout=120) as r:
        dest.write_bytes(r.read())
    return dest, url


def numero(ws, cella):
    v = ws[cella].value
    if v in (None, "- ", "-", ""):
        return 0.0
    if not isinstance(v, (int, float)):
        raise ErroreControllo(f"La cella {cella} del foglio «{ws.title}» non contiene un numero: {v!r}")
    return float(v)


def leggi(xlsx):
    wb = openpyxl.load_workbook(xlsx, data_only=True)
    ws = wb.worksheets[0]
    # struttura attesa (uguale in tutti i fogli da luglio 2024 a dicembre 2026)
    attese = {"C17": "PE", "F17": "PD", "G17": "PCV", "H17": "DISPbt", "I17": "PPE", "M17": "σ1", "N17": "σ2",
              "O17": "σ3", "P17": "UC3", "Q17": "UC6", "S17": "ASOS", "T17": "ARIM",
              "B20": "Quota energia (euro/kWh)", "B21": "Quota fissa (euro/anno)", "B22": "Quota potenza (euro/kW/anno)",
              "B29": "Quota energia (euro/kWh)", "B30": "Quota fissa (euro/anno)"}
    for c, v in attese.items():
        if str(ws[c].value or "").strip() != v:
            raise ErroreControllo(f"Struttura cambiata: nel foglio «{ws.title}» la cella {c} vale {ws[c].value!r}, atteso {v!r}")
    m = re.fullmatch(r"\s*([a-z]{3})-([a-z]{3})\s+(\d{4})\s*", ws.title)
    if not m or m[1] not in MESI or m[2] not in MESI:
        raise ErroreControllo(f"Nome del foglio inatteso: «{ws.title}»")
    periodo = f"{MESI[m[1]]}-{MESI[m[2]]} {m[3]}"
    n = lambda c: numero(ws, c)
    res = dict(pe=n("C20"), pd=n("F20"), ppe=n("I20"), pcv=n("G21"), disp=n("H21"),
               sigma3=n("O20"), uc3=n("P20"), uc6_e=n("Q20"), sigma1=n("M21"), sigma2=n("N22"), uc6_p=n("Q22"),
               asos_e=n("S20"), arim_e=n("T20"), oneri_fissi=0.0)
    nonres = dict(res, asos_e=n("S29"), arim_e=n("T29"), oneri_fissi=n("S30") + n("T30"))
    # plausibilità
    if not 0.02 <= res["pe"] <= 1.0:
        raise ErroreControllo(f"Prezzo dell'energia PE fuori scala: {res['pe']} euro/kWh")
    if not (0 < res["pcv"] < 200 and 0 < res["sigma1"] < 200 and 0 < res["sigma2"] < 200):
        raise ErroreControllo("Quote fisse o di potenza fuori scala")
    return periodo, ws.title, res, nonres


def kwh_tassati(kwh, kw, residente):
    if not residente or kw > 3:
        return kwh
    m = kwh / 12
    lim = 220 if kw > 1.5 else 150
    es = min(m, 150) if m <= lim else max(0.0, 150 - (m - lim))
    return (m - es) * 12


def spesa(kwh, kw, residente, dpun, c, p):
    pe = c["pe"] + dpun / 1000 * (1 + p["perdite_bt"])
    energia = (pe + c["pd"] + c["ppe"]) * kwh + c["pcv"] + c["disp"]
    trasporto = (c["sigma3"] + c["uc3"] + c["uc6_e"]) * kwh + c["sigma1"] + (c["sigma2"] + c["uc6_p"]) * kw
    oneri = (c["asos_e"] + c["arim_e"]) * kwh + c["oneri_fissi"]
    accisa = p["accisa_eur_kwh"] * kwh_tassati(kwh, kw, residente)
    imp = energia + trasporto + oneri + accisa
    return dict(energia=energia, trasporto=trasporto, oneri=oneri, imposte=accisa + imp * p["iva"], totale=imp * (1 + p["iva"]))


def cliente_tipo(c, p, kwh=2000):
    a = spesa(kwh, 3, True, 0, c, p)
    lordo = a["totale"] - spesa(kwh, 3, True, -p["taglio_eur_mwh"], c, p)["totale"]
    netto = lordo * (1 - p["onere_quota"])
    return dict(a, netto=netto, perc=netto / a["totale"] * 100)


def it(x, d=2):
    return f"{x:,.{d}f}".replace(",", "X").replace(".", ",").replace("X", ".")


def genera_html(periodo, url, res, nonres, p):
    link = lambda href, testo: f"<a target='_blank' rel='noopener' href='{href}'>{testo}</a>"
    fonte = ("Fonte: ARERA, " + link(url, f"condizioni economiche per i clienti vulnerabili in maggior tutela, {periodo}")
             + "; " + link("https://www.normattiva.it/uri-res/N2Ls?urn:nir:stato:decreto.legislativo:1995-10-26;504~art52",
                           "D.Lgs. 504/1995, art. 52") + " e Agenzia delle Dogane (accisa); ARERA, "
             + link("https://www.arera.it/fileadmin/allegati/relaz_ann/26/RA26_Volume_1.pdf",
                    "Relazione annuale 2026, vol. 1") + " (utenze per tipo di prezzo, pp. 156, 199 e 201) e "
             + link("https://www.arera.it/fileadmin/allegati/docs/25/204-2025-R-com-ALLEGATO_A.pdf",
                    "glossario della bolletta"))
    arr = lambda d: {k: round(v, 6) for k, v in d.items()}
    parametri = {"res": arr(res), "nonres": arr(nonres), "accisa": p["accisa_eur_kwh"], "iva": p["iva"],
                 "perdite": p["perdite_bt"], "taglio": p["taglio_eur_mwh"], "onere": p["onere_quota"]}
    html = (QUI / "modello.html").read_text(encoding="utf8")
    html = html.replace("__FONTS__", FONTS).replace("__FONTE__", fonte)
    for k, v in COLORI.items():
        html = html.replace(k, v)
    html = html.replace("__PARAMETRI__", json.dumps(parametri))
    if "__" in re.sub(r"<script>.*?</script>", "", html, flags=re.S):
        raise ErroreControllo("Segnaposto non sostituito nel modello HTML")
    # solo ASCII: entità HTML fuori dallo script, \uXXXX dentro
    i = html.index("<script>\nconst P")
    j = html.index("</script>", i)
    html = (html[:i].encode("ascii", "xmlcharrefreplace").decode() + html[i:j].encode("ascii", "backslashreplace").decode()
            + html[j:].encode("ascii", "xmlcharrefreplace").decode())
    return html


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--xlsx")
    ap.add_argument("--riepilogo", default=str(RADICE / "riepilogo.md"))
    a = ap.parse_args()
    p = json.loads((QUI / "parametri.json").read_text(encoding="utf8"))
    if a.xlsx:
        xlsx, url = Path(a.xlsx), "https://www.arera.it/fileadmin//area_operatori/prezzi_e_tariffe/" + Path(a.xlsx).name
    else:
        xlsx, url = scarica()
    periodo, foglio, res, nonres = leggi(xlsx)
    nuovo = cliente_tipo(res, p)
    if not 300 <= nuovo["totale"] <= 2500:
        raise ErroreControllo(f"Spesa del cliente tipo fuori scala: {nuovo['totale']:.2f} euro")

    f_tar = QUI / "tariffe.json"
    vecchio_dati = json.loads(f_tar.read_text(encoding="utf8")) if f_tar.exists() else None
    (RADICE / "index.html").write_text(genera_html(periodo, url, res, nonres, p), encoding="ascii")
    f_tar.write_text(json.dumps({"periodo": periodo, "foglio": foglio, "url": url, "res": res, "nonres": nonres},
                                ensure_ascii=False, indent=1), encoding="utf8")

    # riepilogo per la pull request
    righe = [f"## Tariffe ARERA: {periodo}", "",
             f"File: {url}, foglio «{foglio}»", "",
             "Cliente tipo ARERA (2.000 kWh l'anno, 3 kW, residente), centesimi per kWh. "
             "**Confrontali con il comunicato ARERA del trimestre**: devono coincidere.", "",
             "| Voce | " + (vecchio_dati["periodo"] if vecchio_dati else "prima") + f" | {periodo} |", "|---|---|---|"]
    vecchio = cliente_tipo(vecchio_dati["res"], p) if vecchio_dati else None
    voci = [("Materia energia (con la vendita)", "energia"), ("Trasporto e contatore", "trasporto"),
            ("Oneri di sistema", "oneri"), ("Imposte", "imposte"), ("**Totale**", "totale")]
    for nome, k in voci:
        v = it(vecchio[k] / 20) if vecchio else "–"
        righe.append(f"| {nome} | {v} | {it(nuovo[k] / 20)} |")
    righe += ["", "| Simulatore | prima | ora |", "|---|---|---|",
              f"| Spesa annua | {it(vecchio['totale']) + ' €' if vecchio else '–'} | {it(nuovo['totale'])} € |",
              f"| Risparmio netto con il decreto | {it(vecchio['netto']) + ' €' if vecchio else '–'} | {it(nuovo['netto'])} € |",
              f"| Risparmio netto in % della bolletta | {it(vecchio['perc'], 1) + '%' if vecchio else '–'} | {it(nuovo['perc'], 1)}% |",
              "", "Se la percentuale cambia, va controllato il testo nel CMS (titolo e introduzione citano il cliente tipo).",
              "", "Comunicati ARERA: https://www.arera.it/comunicati-stampa"]
    Path(a.riepilogo).write_text("\n".join(righe) + "\n", encoding="utf8")
    print("\n".join(righe))


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    try:
        main()
    except ErroreControllo as e:
        print(f"CONTROLLO NON SUPERATO: {e}")
        sys.exit(2)
