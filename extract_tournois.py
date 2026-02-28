#!/usr/bin/env python3
"""
Extracteur de tournois FFT depuis un PDF et le site tenup.fft.fr

Usage:
    python extract_tournois.py <fichier_pdf> [--output fichier.xlsx] [--delay 1.0]
    python extract_tournois.py <fichier_pdf> --selenium   # si le site nécessite JavaScript
    python extract_tournois.py --resume sauvegarde.json   # reprendre une extraction interrompue
"""

import argparse
import json
import os
import re
import sys
import time

from PyPDF2 import PdfReader
import requests
from bs4 import BeautifulSoup
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill


# ---------------------------------------------------------------------------
# 1. Extraction des codes tournoi depuis le PDF
# ---------------------------------------------------------------------------

def extraire_codes_tournois(chemin_pdf):
    """Extrait les 6 derniers chiffres de chaque code tournoi du PDF."""
    codes = []
    reader = PdfReader(chemin_pdf)
    for page in reader.pages:
        texte = page.extract_text()
        if not texte:
            continue
        matches = re.findall(r"CODE\s*:\s*T\s*(\d+)", texte)
        for match in matches:
            code_6 = match[-6:]
            if code_6 not in codes:
                codes.append(code_6)
    return codes


# ---------------------------------------------------------------------------
# 2. Récupération du HTML d'une page tournoi
# ---------------------------------------------------------------------------

def _fetch_html_requests(url, session, retries=3):
    """Récupère le HTML via requests avec retry."""
    for attempt in range(retries):
        try:
            response = session.get(url, timeout=30)
            response.raise_for_status()
            return response.text
        except requests.RequestException as e:
            if attempt < retries - 1:
                time.sleep(2 ** attempt)
                continue
            raise e


def _fetch_html_selenium(url, driver):
    """Récupère le HTML via Selenium (pour pages rendues en JS)."""
    driver.get(url)
    time.sleep(2)  # attendre le rendu JS
    return driver.page_source


def _creer_driver_selenium():
    """Crée un driver Selenium Chrome en mode headless."""
    try:
        from selenium import webdriver
        from selenium.webdriver.chrome.options import Options
    except ImportError:
        print("ERREUR: selenium n'est pas installé.")
        print("Installez-le avec : pip install selenium")
        print("Et installez ChromeDriver : https://chromedriver.chromium.org/")
        sys.exit(1)

    options = Options()
    options.add_argument("--headless")
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")
    options.add_argument(
        "user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
    )
    return webdriver.Chrome(options=options)


# ---------------------------------------------------------------------------
# 3. Parsing du HTML d'un tournoi
# ---------------------------------------------------------------------------

def parser_page_tournoi(html, code):
    """Parse le HTML d'une page tournoi et retourne les infos structurées."""
    url = f"https://tenup.fft.fr/tournoi/{code}"
    info = {
        "code": code,
        "url": url,
        "nom": "",
        "club": "",
        "debut": "",
        "fin": "",
        "surface": "",
        "juge_arbitre": "",
        "lieu": "",
        "mail": "",
        "tel": "",
        "epreuves": [],
    }

    soup = BeautifulSoup(html, "html.parser")
    texte_complet = soup.get_text(separator="\n")

    # --- Infos générales ---
    info["nom"] = _extraire_texte_label(soup, "Nom")
    info["club"] = _extraire_texte_label(soup, "Club")
    info["debut"] = (
        _extraire_texte_label(soup, "Début")
        or _extraire_texte_label(soup, "Debut")
    )
    info["fin"] = _extraire_texte_label(soup, "Fin")
    info["surface"] = _extraire_texte_label(soup, "Surface")
    info["juge_arbitre"] = (
        _extraire_texte_label(soup, "Juge Arbitre")
        or _extraire_texte_label(soup, "Juge-Arbitre")
    )
    info["lieu"] = _extraire_texte_label(soup, "Lieu")
    info["mail"] = _extraire_mail(soup)
    info["tel"] = _extraire_tel(soup)

    # Fallback via regex sur le texte complet
    _fallback_texte(info, texte_complet)

    # --- Épreuves ---
    info["epreuves"] = _extraire_epreuves(soup, texte_complet)

    return info


def _extraire_texte_label(soup, label):
    """Cherche un label suivi de sa valeur dans le HTML."""
    # Stratégie 1 : élément contenant exactement le label
    elements = soup.find_all(
        string=re.compile(rf"^\s*{re.escape(label)}\s*$", re.IGNORECASE)
    )
    for el in elements:
        parent = el.find_parent()
        if parent:
            sibling = parent.find_next_sibling()
            if sibling:
                texte = sibling.get_text(strip=True)
                if texte:
                    return texte
            texte_parent = parent.get_text(strip=True)
            texte_parent = re.sub(
                rf"^\s*{re.escape(label)}\s*:?\s*", "", texte_parent
            )
            if texte_parent:
                return texte_parent

    # Stratégie 2 : "label : valeur" dans un même noeud texte
    pattern = re.compile(rf"{re.escape(label)}\s*:?\s*(.+)", re.IGNORECASE)
    elements = soup.find_all(string=pattern)
    for el in elements:
        match = pattern.search(el)
        if match:
            return match.group(1).strip()

    return ""


def _extraire_mail(soup):
    """Extrait l'adresse email (mailto ou texte)."""
    mailto = soup.find("a", href=re.compile(r"mailto:", re.IGNORECASE))
    if mailto:
        email = mailto["href"].replace("mailto:", "").strip()
        return email.split("?")[0]
    texte = soup.get_text()
    match = re.search(r"[\w.+-]+@[\w.-]+\.\w+", texte)
    return match.group(0) if match else ""


def _extraire_tel(soup):
    """Extrait le numéro de téléphone."""
    tel_link = soup.find("a", href=re.compile(r"tel:", re.IGNORECASE))
    if tel_link:
        return (
            tel_link.get_text(strip=True)
            or tel_link["href"].replace("tel:", "").strip()
        )
    return ""


def _chercher_dans_texte(texte, pattern):
    """Cherche un pattern regex dans le texte et retourne le 1er groupe."""
    match = re.search(pattern, texte, re.IGNORECASE)
    return match.group(1).strip() if match else ""


def _fallback_texte(info, texte):
    """Complète les champs vides en cherchant dans le texte brut."""
    fallbacks = {
        "nom": r"Nom\s*[:\-]?\s*(.+)",
        "club": r"Club\s*[:\-]?\s*(.+)",
        "debut": r"D[ée]but\s*[:\-]?\s*(\d{2}/\d{2}/\d{2,4})",
        "fin": r"Fin\s*[:\-]?\s*(\d{2}/\d{2}/\d{2,4})",
        "surface": r"Surface\s*[:\-]?\s*(.+)",
        "juge_arbitre": r"Juge[\s-]?Arbitre\s*[:\-]?\s*(.+)",
        "lieu": r"Lieu\s*[:\-]?\s*(.+)",
        "mail": r"Mail\s*[:\-]?\s*([\w.+-]+@[\w.-]+\.\w+)",
        "tel": r"T[ée]l\.?\s*[:\-]?\s*([\d\s]{10,})",
    }
    for champ, pattern in fallbacks.items():
        if not info[champ]:
            info[champ] = _chercher_dans_texte(texte, pattern)


# ---------------------------------------------------------------------------
# 4. Extraction des épreuves correspondantes
# ---------------------------------------------------------------------------

def _extraire_epreuves(soup, texte_complet):
    """Extrait les épreuves Simple Messieurs (TS) avec Âge 11/12 ans ou 11 ans."""
    epreuves = []

    # Cherche dans le DOM les sections "Simple Messieurs"
    epreuve_sections = soup.find_all(
        string=re.compile(r"Simple\s+Messieurs", re.IGNORECASE)
    )

    for section_text in epreuve_sections:
        parent = section_text.find_parent()
        bloc = parent
        for _ in range(10):
            if bloc is None or bloc.name in ("body", "html", "[document]"):
                break
            bloc_text = bloc.get_text()
            has_ts = re.search(r"\(TS\)", bloc_text)
            has_age = re.search(
                r"[ÂA]ge\s*:\s*(11/12\s*ans|11\s*ans)", bloc_text, re.IGNORECASE
            )
            if has_ts and has_age:
                epreuve = _extraire_details_epreuve(bloc)
                if epreuve and epreuve not in epreuves:
                    epreuves.append(epreuve)
                break
            bloc = bloc.parent

    # Fallback : analyse du texte brut
    if not epreuves:
        epreuves = _extraire_epreuves_depuis_texte(texte_complet)

    return epreuves


def _extraire_details_epreuve(bloc):
    """Extrait tarif, classement, format depuis un bloc HTML d'épreuve."""
    texte = bloc.get_text(separator="\n")
    epreuve = {}

    # Tarif jeune
    match = re.search(r"Tarif\s+jeune\s*:?\s*([\d,]+\s*€)", texte, re.IGNORECASE)
    if match:
        epreuve["tarif_jeune"] = match.group(1).strip()
    else:
        match = re.search(
            r"Tarif\s+jeune\s*:?\s*(.+?)(?:\n|$)", texte, re.IGNORECASE
        )
        if match:
            epreuve["tarif_jeune"] = match.group(1).strip()

    # Classement
    match = re.search(r"Classement\s*:?\s*(.+?)(?:\n|$)", texte, re.IGNORECASE)
    if match:
        epreuve["classement"] = match.group(1).strip()

    # Format
    match = re.search(r"Format\s*:?\s*(.+?)(?:\n|$)", texte, re.IGNORECASE)
    if match:
        epreuve["format"] = match.group(1).strip()

    # Âge
    match = re.search(r"[ÂA]ge\s*:\s*(.+?)(?:\n|$)", texte, re.IGNORECASE)
    if match:
        epreuve["age"] = match.group(1).strip()

    # Nom de l'épreuve
    match = re.search(
        r"(Simple\s+Messieurs\s*\(TS\).*?)(?:\n|$)", texte, re.IGNORECASE
    )
    if match:
        epreuve["nom_epreuve"] = match.group(1).strip()

    return epreuve if epreuve else None


def _extraire_epreuves_depuis_texte(texte):
    """Fallback : extraction depuis le texte brut de la page."""
    epreuves = []
    lines = texte.split("\n")

    i = 0
    while i < len(lines):
        line = lines[i]
        if re.search(r"Simple\s+Messieurs\s*\(TS\)", line, re.IGNORECASE):
            bloc_debut = max(0, i - 5)
            bloc_fin = min(len(lines), i + 30)
            bloc = "\n".join(lines[bloc_debut:bloc_fin])

            if re.search(
                r"[ÂA]ge\s*:\s*(11/12\s*ans|11\s*ans)", bloc, re.IGNORECASE
            ):
                epreuve = {}
                match = re.search(
                    r"Tarif\s+jeune\s*:?\s*([\d,]+\s*€)", bloc, re.IGNORECASE
                )
                if match:
                    epreuve["tarif_jeune"] = match.group(1).strip()
                else:
                    match = re.search(
                        r"Tarif\s+jeune\s*:?\s*(.+?)(?:\n|$)", bloc, re.IGNORECASE
                    )
                    if match:
                        epreuve["tarif_jeune"] = match.group(1).strip()

                match = re.search(
                    r"Classement\s*:?\s*(.+?)(?:\n|$)", bloc, re.IGNORECASE
                )
                if match:
                    epreuve["classement"] = match.group(1).strip()

                match = re.search(
                    r"Format\s*:?\s*(.+?)(?:\n|$)", bloc, re.IGNORECASE
                )
                if match:
                    epreuve["format"] = match.group(1).strip()

                match = re.search(
                    r"[ÂA]ge\s*:\s*(.+?)(?:\n|$)", bloc, re.IGNORECASE
                )
                if match:
                    epreuve["age"] = match.group(1).strip()

                epreuve["nom_epreuve"] = line.strip()
                if epreuve:
                    epreuves.append(epreuve)
        i += 1

    return epreuves


# ---------------------------------------------------------------------------
# 5. Génération du fichier Excel
# ---------------------------------------------------------------------------

def generer_excel(tournois, chemin_sortie):
    """Génère le fichier Excel final."""
    wb = Workbook()
    ws = wb.active
    ws.title = "Tournois"

    headers = [
        "N°", "Code", "URL", "Nom du tournoi", "Club",
        "Début", "Fin", "Surface", "Juge Arbitre", "Lieu",
        "Mail", "Tél", "Épreuve", "Âge", "Tarif jeune",
        "Classement", "Format", "Erreur",
    ]

    header_font = Font(bold=True, color="FFFFFF")
    header_fill = PatternFill(
        start_color="2E75B6", end_color="2E75B6", fill_type="solid"
    )

    for col, header in enumerate(headers, 1):
        cell = ws.cell(row=1, column=col, value=header)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center")

    row = 2
    for i, tournoi in enumerate(tournois, 1):
        epreuves = tournoi.get("epreuves", [])
        if not epreuves:
            _ecrire_ligne_tournoi(ws, row, i, tournoi, None)
            ws.cell(row=row, column=13, value="Aucune épreuve correspondante")
            row += 1
        else:
            for epreuve in epreuves:
                _ecrire_ligne_tournoi(ws, row, i, tournoi, epreuve)
                row += 1

    # Largeur des colonnes
    largeurs = {
        "A": 6, "B": 10, "C": 40, "D": 35, "E": 35, "F": 12, "G": 12,
        "H": 30, "I": 25, "J": 50, "K": 30, "L": 18, "M": 35,
        "N": 15, "O": 15, "P": 20, "Q": 45, "R": 30,
    }
    for col_letter, width in largeurs.items():
        ws.column_dimensions[col_letter].width = width

    ws.freeze_panes = "A2"
    wb.save(chemin_sortie)
    return row - 2


def _ecrire_ligne_tournoi(ws, row, num, tournoi, epreuve):
    """Écrit une ligne de données dans la feuille Excel."""
    ws.cell(row=row, column=1, value=num)
    ws.cell(row=row, column=2, value=tournoi.get("code", ""))
    ws.cell(row=row, column=3, value=tournoi.get("url", ""))
    ws.cell(row=row, column=4, value=tournoi.get("nom", ""))
    ws.cell(row=row, column=5, value=tournoi.get("club", ""))
    ws.cell(row=row, column=6, value=tournoi.get("debut", ""))
    ws.cell(row=row, column=7, value=tournoi.get("fin", ""))
    ws.cell(row=row, column=8, value=tournoi.get("surface", ""))
    ws.cell(row=row, column=9, value=tournoi.get("juge_arbitre", ""))
    ws.cell(row=row, column=10, value=tournoi.get("lieu", ""))
    ws.cell(row=row, column=11, value=tournoi.get("mail", ""))
    ws.cell(row=row, column=12, value=tournoi.get("tel", ""))
    if epreuve:
        ws.cell(row=row, column=13, value=epreuve.get("nom_epreuve", ""))
        ws.cell(row=row, column=14, value=epreuve.get("age", ""))
        ws.cell(row=row, column=15, value=epreuve.get("tarif_jeune", ""))
        ws.cell(row=row, column=16, value=epreuve.get("classement", ""))
        ws.cell(row=row, column=17, value=epreuve.get("format", ""))
    ws.cell(row=row, column=18, value=tournoi.get("erreur", ""))


# ---------------------------------------------------------------------------
# 6. Sauvegarde / reprise JSON
# ---------------------------------------------------------------------------

def sauvegarder_progres(tournois, codes_restants, chemin_json):
    """Sauvegarde le progrès pour pouvoir reprendre plus tard."""
    data = {
        "tournois_traites": tournois,
        "codes_restants": codes_restants,
    }
    with open(chemin_json, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)


def charger_progres(chemin_json):
    """Charge une sauvegarde précédente."""
    with open(chemin_json, "r", encoding="utf-8") as f:
        data = json.load(f)
    return data["tournois_traites"], data["codes_restants"]


# ---------------------------------------------------------------------------
# 7. Fonction principale
# ---------------------------------------------------------------------------

def main():
    parser = argparse.ArgumentParser(
        description="Extrait les informations de tournois FFT depuis un PDF"
    )
    parser.add_argument(
        "pdf", nargs="?", help="Chemin vers le fichier PDF"
    )
    parser.add_argument(
        "--output", "-o", default="tournois_extraits.xlsx",
        help="Fichier Excel de sortie (défaut: tournois_extraits.xlsx)",
    )
    parser.add_argument(
        "--delay", "-d", type=float, default=1.0,
        help="Délai en secondes entre chaque requête (défaut: 1.0)",
    )
    parser.add_argument(
        "--selenium", action="store_true",
        help="Utiliser Selenium (Chrome headless) au lieu de requests",
    )
    parser.add_argument(
        "--resume", metavar="FICHIER_JSON",
        help="Reprendre une extraction depuis une sauvegarde JSON",
    )
    parser.add_argument(
        "--save", metavar="FICHIER_JSON", default="progres_tournois.json",
        help="Fichier de sauvegarde intermédiaire (défaut: progres_tournois.json)",
    )
    args = parser.parse_args()

    # Vérifications des arguments
    if not args.pdf and not args.resume:
        parser.error("Fournissez un fichier PDF ou utilisez --resume")

    # Charger ou extraire les codes
    if args.resume and os.path.exists(args.resume):
        print(f"Reprise depuis : {args.resume}")
        tournois, codes = charger_progres(args.resume)
        print(f"  -> {len(tournois)} tournoi(s) déjà traité(s)")
        print(f"  -> {len(codes)} tournoi(s) restant(s)")
    else:
        print(f"Lecture du PDF : {args.pdf}")
        codes = extraire_codes_tournois(args.pdf)
        print(f"  -> {len(codes)} codes de tournoi trouvés")
        tournois = []

    if not codes and not tournois:
        print("Aucun code de tournoi trouvé.")
        sys.exit(1)

    total = len(tournois) + len(codes)

    # Préparer le mode de fetch
    driver = None
    session = None
    if args.selenium:
        print("Mode Selenium (Chrome headless)")
        driver = _creer_driver_selenium()
    else:
        print("Mode requests")
        session = requests.Session()
        session.headers.update({
            "User-Agent": (
                "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                "AppleWebKit/537.36 (KHTML, like Gecko) "
                "Chrome/120.0.0.0 Safari/537.36"
            ),
            "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
            "Accept-Language": "fr-FR,fr;q=0.9,en;q=0.5",
        })

    # Traiter les tournois restants
    print(f"\nExtraction des tournois (délai: {args.delay}s)...")
    try:
        while codes:
            code = codes[0]
            i = len(tournois) + 1
            print(f"  [{i:3d}/{total}] Tournoi {code}...", end=" ", flush=True)

            time.sleep(args.delay)

            try:
                if args.selenium:
                    html = _fetch_html_selenium(
                        f"https://tenup.fft.fr/tournoi/{code}", driver
                    )
                else:
                    html = _fetch_html_requests(
                        f"https://tenup.fft.fr/tournoi/{code}", session
                    )

                info = parser_page_tournoi(html, code)
            except Exception as e:
                info = {
                    "code": code,
                    "url": f"https://tenup.fft.fr/tournoi/{code}",
                    "nom": "", "club": "", "debut": "", "fin": "",
                    "surface": "", "juge_arbitre": "", "lieu": "",
                    "mail": "", "tel": "", "epreuves": [],
                    "erreur": str(e),
                }

            tournois.append(info)
            codes.pop(0)

            # Affichage du résultat
            nb_epreuves = len(info.get("epreuves", []))
            if info.get("erreur"):
                print(f"ERREUR: {info['erreur']}")
            elif nb_epreuves > 0:
                print(
                    f"OK - {info['nom']} - "
                    f"{nb_epreuves} épreuve(s) correspondante(s)"
                )
            else:
                print(f"OK - {info['nom']} - aucune épreuve 11/12 ans SM(TS)")

            # Sauvegarde intermédiaire tous les 10 tournois
            if len(tournois) % 10 == 0:
                sauvegarder_progres(tournois, codes, args.save)

    except KeyboardInterrupt:
        print(f"\n\nInterruption ! Sauvegarde en cours...")
        sauvegarder_progres(tournois, codes, args.save)
        print(f"Progrès sauvegardé dans : {args.save}")
        print(f"Reprenez avec : python extract_tournois.py --resume {args.save}")
        sys.exit(0)
    finally:
        if driver:
            driver.quit()

    # Sauvegarde finale
    sauvegarder_progres(tournois, [], args.save)

    # Générer le fichier Excel
    print(f"\nGénération du fichier Excel : {args.output}")
    nb_lignes = generer_excel(tournois, args.output)
    print(f"  -> {nb_lignes} ligne(s) écrite(s)")

    # Résumé
    nb_avec_epreuves = sum(1 for t in tournois if t.get("epreuves"))
    nb_erreurs = sum(1 for t in tournois if t.get("erreur"))
    print(f"\nRésumé :")
    print(f"  Tournois traités : {len(tournois)}")
    print(f"  Avec épreuve(s) correspondante(s) : {nb_avec_epreuves}")
    print(f"  Erreurs : {nb_erreurs}")
    print(f"  Fichier Excel : {args.output}")


if __name__ == "__main__":
    main()
