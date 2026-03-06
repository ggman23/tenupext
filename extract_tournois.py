#!/usr/bin/env python3
"""
Extracteur de tournois FFT depuis un PDF et le site tenup.fft.fr

Usage:
    python extract_tournois.py <fichier_pdf> [--output fichier.xlsx] [--delay 1.0]
    python extract_tournois.py <fichier_pdf> --playwright  # RECOMMANDÉ (le site nécessite JS)
    python extract_tournois.py <fichier_pdf> --selenium    # alternative avec Selenium
    python extract_tournois.py --resume sauvegarde.json    # reprendre une extraction interrompue

    # Mode mise à jour : n'extraire que les nouveaux tournois
    python extract_tournois.py <fichier_pdf> --playwright --tag idf-oct-dec-2026 --mode maj
    # Mode complet (défaut) : tout extraire, met à jour l'historique
    python extract_tournois.py <fichier_pdf> --playwright --tag idf-oct-dec-2026 --mode complet

Installation (Playwright - recommandé):
    pip install playwright
    playwright install chromium
"""

import argparse
import json
import os
import re
import sys
import time

from PyPDF2 import PdfReader
import requests
from bs4 import BeautifulSoup, NavigableString
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill


# ---------------------------------------------------------------------------
# 1. Extraction des codes tournoi depuis le PDF
# ---------------------------------------------------------------------------

def extraire_codes_tournois(chemin_pdf):
    """Extrait les 6 derniers chiffres de chaque code tournoi du PDF.

    Supporte les codes T (tournoi) et C (championnat).
    Ex: CODE : T 20265777048200197275  -> 197275
        CODE : C 202657L0077000169679  -> 169679
    """
    codes = []
    reader = PdfReader(chemin_pdf)
    for page in reader.pages:
        texte = page.extract_text()
        if not texte:
            continue
        # Capture les codes T et C (peuvent contenir des lettres)
        matches = re.findall(r"CODE\s*:\s*[TC]\s*(\S+)", texte)
        for match in matches:
            # Extrait les 6 derniers chiffres (ignore les lettres en fin)
            digits = re.findall(r"\d", match)
            if len(digits) >= 6:
                code_6 = "".join(digits[-6:])
                if code_6 not in codes:
                    codes.append(code_6)
    return codes


def extraire_tournois_du_pdf(chemin_pdf):
    """Extrait les codes ET les infos de base de chaque tournoi depuis le PDF.

    Le PDF contient : nom, club, dates, surface, juge-arbitre, lieu, contact, tarifs.
    Ces infos sont plus fiables et plus rapides à obtenir que le scraping URL.
    Le classement du PDF est ignoré (toujours « Open », souvent faux).
    Le format n'est pas dans le PDF (uniquement sur l'URL).
    """
    reader = PdfReader(chemin_pdf)
    texte_complet = ""
    for page in reader.pages:
        texte = page.extract_text()
        if texte:
            texte_complet += texte + "\n\n"

    texte_complet = _normaliser_espaces(texte_complet)

    # Trouver tous les CODE
    pattern_code = r"CODE\s*:\s*[TC]\s*(\S+)"
    matches = list(re.finditer(pattern_code, texte_complet))

    tournois = []
    seen_codes = set()

    for idx, match in enumerate(matches):
        raw = match.group(1)
        digits = re.findall(r"\d", raw)
        if len(digits) < 6:
            continue
        code = "".join(digits[-6:])

        if code in seen_codes:
            continue
        seen_codes.add(code)

        # Bloc AVANT le CODE : header du tournoi (club, nom, dates, juge, surface)
        debut_avant = matches[idx - 1].end() if idx > 0 else 0
        bloc_avant = texte_complet[debut_avant:match.start()]

        # Bloc APRÈS le CODE : installations, engagements, tableau des épreuves
        debut_apres = match.end()
        fin_apres = (
            matches[idx + 1].start()
            if idx < len(matches) - 1
            else len(texte_complet)
        )
        bloc_apres = texte_complet[debut_apres:fin_apres]

        info = _extraire_infos_pdf_bloc(bloc_avant, bloc_apres, code)
        tournois.append(info)

    return tournois


def _extraire_infos_pdf_bloc(bloc_avant, bloc_apres, code):
    """Extrait les infos d'un tournoi depuis les blocs PDF.

    Format réel du PDF FFT :
        CLUB NAME                         (pas de label)
        Nom du tournoi                    (pas de label)
        DD/MM/YYYY au DD/MM/YYYY          (dates sans « Du »)
        JUGE-ARBITRE : Prénom NOM
        SURFACE(S) : Type
        PRIX EN ESPÈCE : …  /  PRIX EN LOTS : …
        INSCRIPTIONS / PAIEMENT EN LIGNE : Oui / Oui
        CODE : T …

        INSTALLATIONS :  CodePostal_VILLE
                         Adresse
                         CodePostal VILLE
                         Téléphone

        ENGAGEMENTS :    email@...

        Catégorie  Epreuve           Classement  Droits
        11/12      Simple Messieurs  Open        16€

    bloc_avant = texte entre le CODE précédent et ce CODE
    bloc_apres = texte entre ce CODE et le CODE suivant
    """
    info = {
        "code": code,
        "url": f"https://tenup.fft.fr/tournoi/{code}",
        "nom": "",
        "club": "",
        "debut": "",
        "fin": "",
        "surface": "",
        "juge_arbitre": "",
        "lieu": "",
        "mail": "",
        "tel": "",
        "tarif_jeune_pdf": "",
        "epreuves": [],
    }

    # =========================================================
    # BLOC AVANT : header du tournoi
    # =========================================================

    # Dates : "DD/MM/YYYY au DD/MM/YYYY" (dernière occurrence du bloc)
    date_matches = re.findall(
        r"(\d{2}/\d{2}/\d{2,4})\s+au\s+(\d{2}/\d{2}/\d{2,4})", bloc_avant
    )
    if date_matches:
        info["debut"], info["fin"] = date_matches[-1]

    # Club et Nom : les 2 lignes non-vides juste avant la ligne de dates
    m_date = None
    for m_date in re.finditer(
        r"(\d{2}/\d{2}/\d{2,4})\s+au\s+(\d{2}/\d{2}/\d{2,4})", bloc_avant
    ):
        pass  # on veut la dernière occurrence
    if m_date:
        text_before_date = bloc_avant[:m_date.start()]
        lignes = [l.strip() for l in text_before_date.split("\n") if l.strip()]
        # Filtrer les lignes de bruit (numéros de page, séparateurs, lignes trop courtes)
        lignes = [
            l for l in lignes
            if len(l) > 3
            and not re.match(r"^\d+$", l)
            and not re.match(r"^[-=_]+$", l)
            and not re.match(r"^Page\s", l, re.IGNORECASE)
        ]
        if len(lignes) >= 2:
            info["club"] = lignes[-2]
            info["nom"] = lignes[-1]
        elif len(lignes) == 1:
            info["nom"] = lignes[-1]

    # JUGE-ARBITRE : Prénom NOM (dernière occurrence)
    juge_matches = re.findall(
        r"JUGE[\s-]?ARBITRE\s*:\s*(.+)", bloc_avant, re.IGNORECASE
    )
    if juge_matches:
        info["juge_arbitre"] = juge_matches[-1].strip()

    # SURFACE(S) : Type (dernière occurrence)
    surf_matches = re.findall(
        r"SURFACE\(?S?\)?\s*:\s*(.+)", bloc_avant, re.IGNORECASE
    )
    if surf_matches:
        info["surface"] = surf_matches[-1].strip()

    # =========================================================
    # BLOC APRÈS : installations, engagements, tableau
    # =========================================================

    # INSTALLATIONS : adresse, ville, téléphone
    m = re.search(
        r"INSTALLATIONS?\s*:\s*(.+?)(?=ENGAGEMENTS?|Cat[ée]gorie|\Z)",
        bloc_apres, re.DOTALL | re.IGNORECASE,
    )
    if m:
        bloc_install = m.group(1)

        # Lieu : code postal + ville (prendre la version la plus complète)
        lieux = re.findall(r"(\d{5})\s+([A-ZÀ-Ü][A-ZÀ-Ü\s]+)", bloc_install)
        if lieux:
            # Prendre le match le plus long (souvent le 2e, non tronqué)
            cp, ville = max(lieux, key=lambda x: len(x[1]))
            info["lieu"] = f"{ville.strip()} ({cp})"

        # Téléphone dans le bloc installations
        m_tel = re.search(
            r"((?:0[1-9])[\s.]?\d{2}[\s.]?\d{2}[\s.]?\d{2}[\s.]?\d{2})",
            bloc_install,
        )
        if m_tel:
            info["tel"] = m_tel.group(1).strip()

    # ENGAGEMENTS : email
    m = re.search(
        r"ENGAGEMENTS?\s*:?\s*([\w.+-]+@[\w.-]+\.\w+)",
        bloc_apres, re.IGNORECASE,
    )
    if m:
        info["mail"] = m.group(1).strip()
    if not info["mail"]:
        m = re.search(r"[\w.+-]+@[\w.-]+\.\w+", bloc_apres)
        if m:
            info["mail"] = m.group(0)

    # Tableau des épreuves : tarif (colonne « Droits »)
    # Chercher une ligne contenant "11" + "Simple Messieurs" + montant €
    for line in bloc_apres.split("\n"):
        if re.search(r"\b11\b", line) and re.search(
            r"Simple\s+Messieurs|SM\b", line, re.IGNORECASE
        ):
            m = re.search(r"(\d+[.,]?\d*\s*\u20ac)", line)
            if m:
                info["tarif_jeune_pdf"] = m.group(1).strip()
                break
    # Fallback : première ligne avec "11" et un montant €
    if not info["tarif_jeune_pdf"]:
        for line in bloc_apres.split("\n"):
            if re.search(r"\b11\b", line):
                m = re.search(r"(\d+[.,]?\d*\s*\u20ac)", line)
                if m:
                    info["tarif_jeune_pdf"] = m.group(1).strip()
                    break

    return info


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


# --- Playwright ---

def _creer_playwright():
    """Crée un navigateur Playwright Chromium en mode headless."""
    try:
        from playwright.sync_api import sync_playwright
    except ImportError:
        print("ERREUR: playwright n'est pas installé.")
        print("Installez-le avec :")
        print("  pip install playwright")
        print("  playwright install chromium")
        sys.exit(1)

    pw = sync_playwright().start()
    browser = pw.chromium.launch(headless=True)
    context = browser.new_context(
        user_agent=(
            "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
            "AppleWebKit/537.36 (KHTML, like Gecko) "
            "Chrome/120.0.0.0 Safari/537.36"
        ),
        locale="fr-FR",
    )
    page = context.new_page()
    return pw, browser, page


def _fetch_html_playwright(url, page, timeout=15000):
    """Récupère le HTML via Playwright avec attente du contenu dynamique."""
    page.goto(url, wait_until="domcontentloaded", timeout=30000)
    # Attendre que le réseau soit calme (plus de requêtes XHR en cours).
    try:
        page.wait_for_load_state("networkidle", timeout=timeout)
    except Exception:
        pass

    # Attendre que les épreuves soient chargées (contenu dynamique JS).
    # On cherche "Simple Messieurs" ou "Âge" qui indiquent que les cartes
    # d'épreuves sont rendues.
    selectors_to_try = [
        "text=/^SM$/",               # Badge SM
        "text=/Simple/i",
        "text=/11.?12/i",
        "text=/ge.*:/i",             # "Âge :" (avec ou sans accent)
    ]
    content_found = False
    for selector in selectors_to_try:
        try:
            page.wait_for_selector(selector, timeout=5000)
            content_found = True
            break
        except Exception:
            continue
    if not content_found:
        # Scroller vers le bas pour déclencher le lazy loading
        try:
            page.evaluate("window.scrollTo(0, document.body.scrollHeight)")
            page.wait_for_timeout(2000)
            page.evaluate("window.scrollTo(0, 0)")
            page.wait_for_timeout(1000)
        except Exception:
            pass
        # Dernier recours : attendre 5 secondes supplémentaires
        page.wait_for_timeout(5000)

    # Cliquer sur d'éventuels onglets/accordéons pour révéler du contenu caché
    _cliquer_onglets_epreuves(page)

    # Attendre un peu après les clics pour que le contenu se charge
    page.wait_for_timeout(1000)

    return page.content()


def _cliquer_onglets_epreuves(page):
    """Clique sur les onglets/tabs/accordéons pour charger les épreuves cachées."""
    tab_selectors = [
        "[aria-expanded='false']",
        "button:has-text('preuve')",
        "a:has-text('preuve')",
        "button:has-text('Voir')",
        "a:has-text('Voir plus')",
        "button:has-text('Afficher')",
        # Sélecteurs spécifiques tenup.fft.fr
        ".tab-link",
        ".nav-link:not(.active)",
        "[role='tab'][aria-selected='false']",
        "button:has-text('SM')",
        "button:has-text('Simple')",
        "a:has-text('Simple Messieurs')",
    ]
    for selector in tab_selectors:
        try:
            elements = page.query_selector_all(selector)
            for el in elements:
                try:
                    el.click()
                    page.wait_for_timeout(500)
                except Exception:
                    continue
        except Exception:
            continue


# ---------------------------------------------------------------------------
# 3. Parsing du HTML d'un tournoi
# ---------------------------------------------------------------------------

def _normaliser_espaces(texte):
    """Remplace les espaces insécables et autres espaces Unicode par des espaces normaux."""
    # \xa0 = espace insécable, \u202f = espace fine insécable, etc.
    return re.sub(r"[\xa0\u202f\u2007\u2009\u200a]", " ", texte)


def _est_valeur_valide(texte):
    """Vérifie que la valeur extraite n'est pas du bruit (JS, template, pub)."""
    if not texte or len(texte.strip()) < 2:
        return False
    # Rejeter le JavaScript / publicitaire
    if re.search(
        r"googletag|eSlot|pubads|gpt-ad|addService|defineSl|\.js\b|"
        r"adsbygoogle|__webpack|function\s*\(|var\s+\w|window\.",
        texte, re.IGNORECASE,
    ):
        return False
    # Rejeter les fragments de template / placeholder connus de tenup
    if re.search(
        r"du club.*ville|Ville.*Date.*debut|^\s*/\s*fin\s*$",
        texte, re.IGNORECASE,
    ):
        return False
    return True


def _est_classement_valide(texte):
    """Vérifie que le texte ressemble à un classement FFT.

    Exemples valides : "NC - N1", "NC - 30/5", "40 - 30/1", "30/5 - 30/1",
    "NC", "N1", "30/5", "NC à 30/5", "Classé(e) ou NC".
    Rejette les phrases descriptives comme "proximité et date d'inscription...".
    """
    if not texte:
        return False
    texte = texte.strip()
    # Trop long pour être un classement (> 40 chars)
    if len(texte) > 40:
        return False
    # Doit contenir au moins un élément typique d'un classement FFT
    return bool(re.search(
        r"\bNC\b|\bN[1-4]\b|\b[1-4]0\b|\b[23]0/[1-5]\b|\b15/[1-5]\b",
        texte,
    ))


# Textes de navigation / bruit à exclure du nom d'épreuve
_TEXTES_NAVIGATION = {
    "retour aux résultats",
    "retour",
    "voir plus",
    "voir les épreuves",
    "afficher",
    "aucune épreuve correspondante",
    "aucune épreuve",
    "chargement",
    "loading",
}


def parser_page_tournoi(html, code):
    """Parse le HTML d'une page tournoi et retourne les infos structurées."""
    # Normaliser les espaces insécables dans tout le HTML
    html = _normaliser_espaces(html)

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

    # Supprimer les scripts, styles et pubs pour ne pas matcher du JS/CSS
    for tag in soup(["script", "style", "noscript"]):
        tag.decompose()

    texte_complet = soup.get_text(separator="\n")

    # --- Infos générales via sélecteurs CSS (structure tenup.fft.fr) ---

    # Nom du tournoi : <h1 itemprop="name">...</h1>
    el = soup.select_one(".tournoi-detail-page-title h1")
    if not el:
        el = soup.select_one("h1[itemprop='name']")
    if el:
        info["nom"] = el.get_text(strip=True)

    # Club / Ville : <h2 class="tournoi-detail-page-club">CLUB / VILLE</h2>
    el = soup.select_one(".tournoi-detail-page-club")
    if el:
        info["club"] = el.get_text(strip=True)

    # Dates : <span class="tournoi-detail-page-date-debut">JJ/MM/AA</span>
    el = soup.select_one(".tournoi-detail-page-date-debut")
    if el:
        info["debut"] = el.get_text(strip=True)
    el = soup.select_one(".tournoi-detail-page-date-fin")
    if el:
        info["fin"] = el.get_text(strip=True)

    # Surface : <span class="tournoi-detail-page-competition-surfaces-content">
    el = soup.select_one(".tournoi-detail-page-competition-surfaces-content")
    if el:
        info["surface"] = el.get_text(strip=True)

    # Lieu : addr1 (nom du stade) + addr2 (code postal + ville)
    lieu_parts = []
    el = soup.select_one(".tournoi-detail-page-lieu-addr1")
    if el:
        lieu_parts.append(el.get_text(strip=True))
    el = soup.select_one(".tournoi-detail-page-lieu-addr2")
    if el:
        lieu_parts.append(el.get_text(strip=True))
    if lieu_parts:
        info["lieu"] = ", ".join(lieu_parts)

    # Juge-arbitre : chercher dans le texte (pas de classe CSS dédiée)
    info["juge_arbitre"] = (
        _extraire_texte_label(soup, "Juge Arbitre")
        or _extraire_texte_label(soup, "Juge-Arbitre")
    )

    # Email et téléphone
    info["mail"] = _extraire_mail(soup)
    info["tel"] = _extraire_tel(soup)

    # Valider les dates (format attendu : JJ/MM/AAAA ou JJ/MM/AA)
    for champ in ("debut", "fin"):
        if info[champ] and not re.match(r"\d{2}/\d{2}/\d{2,4}", info[champ]):
            info[champ] = ""

    # Fallback via regex sur le texte complet (pour les champs non trouvés)
    _fallback_texte(info, texte_complet)

    # --- Épreuves ---
    info["epreuves"] = _extraire_epreuves(soup, texte_complet)

    return info


def _extraire_texte_label(soup, label):
    """Cherche un label suivi de sa valeur dans le HTML.

    Ignore les valeurs qui ressemblent à du bruit (JS, pubs, template).
    """
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
                if texte and _est_valeur_valide(texte):
                    return texte
            texte_parent = parent.get_text(strip=True)
            texte_parent = re.sub(
                rf"^\s*{re.escape(label)}\s*:?\s*", "", texte_parent
            )
            if texte_parent and _est_valeur_valide(texte_parent):
                return texte_parent

    # Stratégie 2 : "label : valeur" dans un même noeud texte
    pattern = re.compile(rf"{re.escape(label)}\s*:?\s*(.+)", re.IGNORECASE)
    elements = soup.find_all(string=pattern)
    for el in elements:
        match = pattern.search(el)
        if match:
            val = match.group(1).strip()
            if _est_valeur_valide(val):
                return val

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
    """Cherche un pattern regex dans le texte et retourne le 1er groupe.

    Valide que le résultat n'est pas du bruit.
    """
    match = re.search(pattern, texte, re.IGNORECASE)
    if match:
        val = match.group(1).strip()
        if _est_valeur_valide(val):
            return val
    return ""


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


def _fusionner_donnees(pdf_info, url_info):
    """Fusionne les données PDF et URL.

    Stratégie :
    - Infos générales (nom, club, dates, lieu, etc.) : PDF prioritaire, URL en fallback
    - Classement, format : URL prioritaire (le PDF dit toujours « Open »)
    - Tarif : PDF en priorité, URL en fallback
    - Épreuves : toujours depuis l'URL (seule source pour classement/format)
    """
    info = {
        "code": url_info.get("code") or pdf_info.get("code", ""),
        "url": url_info.get("url") or pdf_info.get("url", ""),
        "epreuves": url_info.get("epreuves", []),
    }

    # Champs généraux : PDF prioritaire, URL en fallback
    for champ in (
        "nom", "club", "debut", "fin", "surface",
        "juge_arbitre", "lieu", "mail", "tel",
    ):
        info[champ] = pdf_info.get(champ, "") or url_info.get(champ, "")

    # Tarif : utiliser le tarif PDF comme fallback pour les épreuves
    tarif_pdf = pdf_info.get("tarif_jeune_pdf", "")
    for ep in info["epreuves"]:
        if not ep.get("tarif_jeune") and tarif_pdf:
            ep["tarif_jeune"] = tarif_pdf

    # Conserver l'erreur éventuelle
    if url_info.get("erreur"):
        info["erreur"] = url_info["erreur"]

    return info


# ---------------------------------------------------------------------------
# 4. Extraction des épreuves correspondantes
# ---------------------------------------------------------------------------

def _bloc_contient_age_11(bloc_text):
    """Vérifie si un bloc contient un âge avec '11' (11, 11/12, 11-12)."""
    # Âge : 11/12 ans, Âge : 11 ans, Âge:11/12, etc.
    if re.search(r"[ÂA]ge\s*:?\s*[^\n]*11", bloc_text, re.IGNORECASE):
        return True
    # "11/12 ans" ou "11-12 ans" dans le texte libre (titre d'épreuve)
    if re.search(r"11\s*[/\-]\s*12", bloc_text, re.IGNORECASE):
        return True
    # "11 ans" isolé
    if re.search(r"\b11\s*ans\b", bloc_text, re.IGNORECASE):
        return True
    return False


def _extraire_epreuves(soup, texte_complet):
    """Extrait l'épreuve avec badge SM et âge contenant 11.

    Règle simple : badge "SM" + âge avec "11".
    Couvre tous les types (Simple Messieurs, Simple Mixte, TMC Garçons...).
    Retourne une liste avec au plus 1 épreuve par tournoi.
    """
    # Stratégie 1 DOM : Chercher le badge "SM" puis remonter au plus haut
    # ancêtre valide (nb_sm <= 1 et contient âge 11).
    # On prend le plus haut pour avoir la carte complète (avec Classement/Format).
    # Chercher "SM" isolé (badge peut contenir espaces Unicode)
    sm_badges = soup.find_all(
        string=re.compile(r"^[\s\u00a0\u202f]*SM[\s\u00a0\u202f]*$")
    )
    # Fallback : chercher les éléments courts contenant uniquement "SM"
    if not sm_badges:
        for el in soup.find_all(["span", "div", "strong", "b", "p", "a"]):
            if el.get_text(strip=True) == "SM":
                sm_badges.append(el)

    for badge_or_text in sm_badges:
        # NavigableString -> parent est le Tag contenant le texte
        # Tag (fallback) -> on commence directement depuis ce Tag
        if isinstance(badge_or_text, NavigableString):
            bloc = badge_or_text.parent
        else:
            bloc = badge_or_text
        if bloc is None:
            continue
        best_bloc = None
        for _ in range(15):
            if bloc is None or bloc.name in ("body", "html", "[document]"):
                break
            bloc_text = bloc.get_text()
            nb_sm = len(re.findall(r"\bSM\b", bloc_text))
            if nb_sm > 1:
                break
            if _bloc_contient_age_11(bloc_text):
                best_bloc = bloc
            bloc = bloc.parent

        if best_bloc:
            epreuve = _extraire_details_epreuve(best_bloc)
            if epreuve:
                return [epreuve]

    # Stratégie 2 DOM : Chercher "Simple Messieurs" (pour les pages où le
    # badge SM n'est pas dans un noeud texte isolé).
    sm_elements = soup.find_all(
        string=re.compile(r"Simple\s+Messieurs", re.IGNORECASE)
    )
    for el_text in sm_elements:
        bloc = el_text.find_parent()
        best_bloc = None
        for _ in range(10):
            if bloc is None or bloc.name in ("body", "html", "[document]"):
                break
            bloc_text = bloc.get_text()
            if _bloc_contient_age_11(bloc_text):
                nb_simple = len(re.findall(
                    r"Simple\s+Messieurs", bloc_text, re.IGNORECASE
                ))
                if nb_simple <= 1:
                    best_bloc = bloc
                elif nb_simple > 3:
                    break
            bloc = bloc.parent

        if best_bloc:
            epreuve = _extraire_details_epreuve(best_bloc)
            if epreuve:
                return [epreuve]

    # Stratégie 3 DOM : Chercher les cartes d'épreuves par classe CSS,
    # puis filtrer sur âge 11 et type SM/Simple Messieurs.
    for cls in (".epreuve-card", ".epreuve-detail",
                "[class*='epreuve']", ".card"):
        cards = soup.select(cls)
        for card in cards:
            card_text = card.get_text()
            if not _bloc_contient_age_11(card_text):
                continue
            if re.search(r"\bSM\b|Simple\s+Messieurs", card_text, re.IGNORECASE):
                epreuve = _extraire_details_epreuve(card)
                if epreuve:
                    return [epreuve]

    # Stratégie 4 : Fallback texte brut
    epreuves = _extraire_epreuves_depuis_texte(texte_complet)
    if epreuves:
        return [epreuves[0]]

    return []


def _extraire_details_epreuve(bloc):
    """Extrait tarif, classement, format, âge et nom depuis un bloc HTML d'épreuve."""
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
        if match and match.group(1).strip():
            epreuve["tarif_jeune"] = match.group(1).strip()
    if "tarif_jeune" not in epreuve:
        match = re.search(r"Tarif\s*:?\s*([\d,]+\s*€)", texte, re.IGNORECASE)
        if match:
            epreuve["tarif_jeune"] = match.group(1).strip()

    # Classement - collecter TOUS les candidats puis prendre le premier valide.
    # On utilise finditer car le texte peut contenir "Classement demandé (...)"
    # avant le vrai "Classement : NC - 30/5".
    _classement_candidats = []
    # Stratégie 1 : CSS selector direct (classe tenup.fft.fr)
    cls_el = bloc.select_one(".epreuve-detail-classement-detail")
    if cls_el:
        val = cls_el.get_text(strip=True)
        cleaned = re.sub(r"^Classement\s*:?\s*", "", val, flags=re.IGNORECASE)
        if cleaned:
            _classement_candidats.append(cleaned)
    # Stratégie 2 : regex sur TOUTES les occurrences de "Classement" dans le texte
    for m in re.finditer(r"Classement\s*:?\s*(.+?)(?:\n|$)", texte, re.IGNORECASE):
        val = m.group(1).strip()
        if val:
            _classement_candidats.append(val)
    # Stratégie 3 : label "Classement" seul sur une ligne, valeur sur la suivante
    for m in re.finditer(
        r"Classement\s*:?\s*\n+\s*(.+?)(?:\n|$)", texte, re.IGNORECASE
    ):
        val = m.group(1).strip()
        if val:
            _classement_candidats.append(val)
    # Stratégie 4 : élément DOM contenant exactement "Classement", valeur en sibling
    classement_el = bloc.find(
        string=re.compile(r"^\s*Classement\s*:?\s*$", re.IGNORECASE)
    )
    if classement_el:
        parent = classement_el.find_parent()
        if parent:
            sibling = parent.find_next_sibling()
            if sibling:
                val = sibling.get_text(strip=True)
                if val:
                    _classement_candidats.append(val)
            parent_text = parent.get_text(strip=True)
            cleaned = re.sub(
                r"^Classement\s*:?\s*", "", parent_text,
                flags=re.IGNORECASE,
            )
            if cleaned:
                _classement_candidats.append(cleaned)
    # Prendre le premier candidat qui ressemble à un vrai classement
    for candidat in _classement_candidats:
        if _est_classement_valide(candidat):
            epreuve["classement"] = candidat
            break

    # Format - CSS selector direct + regex + DOM traversal
    # Stratégie 1 : CSS selector (classe tenup.fft.fr) - variantes
    for cls in (".epreuve-detail-format", ".epreuve-format",
                "[class*='format']"):
        if "format" in epreuve:
            break
        fmt_el = bloc.select_one(cls)
        if fmt_el:
            val = fmt_el.get_text(strip=True)
            cleaned = re.sub(r"^Format\s*:?\s*", "", val, flags=re.IGNORECASE)
            if cleaned:
                epreuve["format"] = cleaned
    # Stratégie 2 : regex sur le texte (avec séparateur souple : / - / rien)
    if "format" not in epreuve:
        match = re.search(
            r"Format\s*[:\-–]?\s*(.+?)(?:\n|$)", texte, re.IGNORECASE
        )
        if match and match.group(1).strip():
            epreuve["format"] = match.group(1).strip()
    if "format" not in epreuve:
        # Label "Format" seul sur une ligne, valeur sur la ligne suivante
        match = re.search(
            r"Format\s*[:\-–]?\s*\n+\s*(.+?)(?:\n|$)", texte, re.IGNORECASE
        )
        if match and match.group(1).strip():
            epreuve["format"] = match.group(1).strip()
    if "format" not in epreuve:
        # Élément DOM contenant "Format" (avec ou sans :), valeur en sibling
        format_el = bloc.find(
            string=re.compile(r"^\s*Format\s*[:\-–]?\s*$", re.IGNORECASE)
        )
        if format_el:
            parent = format_el.find_parent()
            if parent:
                sibling = parent.find_next_sibling()
                if sibling:
                    val = sibling.get_text(strip=True)
                    if val:
                        epreuve["format"] = val
                if "format" not in epreuve:
                    parent_text = parent.get_text(strip=True)
                    cleaned = re.sub(
                        r"^Format\s*[:\-–]?\s*", "", parent_text,
                        flags=re.IGNORECASE,
                    )
                    if cleaned:
                        epreuve["format"] = cleaned
    # Stratégie 5 : chercher des patterns typiques de format de match
    if "format" not in epreuve:
        match = re.search(
            r"(\d\s*set(?:s)?\s*(?:de\s*\d+\s*jeux?|gagnant).*?)(?:\n|$)",
            texte, re.IGNORECASE,
        )
        if match:
            epreuve["format"] = match.group(1).strip()
    if "format" not in epreuve:
        match = re.search(
            r"((?:super[- ]?)?tie[- ]?break.*?)(?:\n|$)",
            texte, re.IGNORECASE,
        )
        if match:
            epreuve["format"] = match.group(1).strip()

    # Âge
    match = re.search(r"[ÂA]ge\s*:?\s*(.+?)(?:\n|$)", texte, re.IGNORECASE)
    if match and match.group(1).strip():
        epreuve["age"] = match.group(1).strip()
    if "age" not in epreuve:
        # Chercher "11/12 ans" ou "11 ans" directement
        match = re.search(r"(11\s*[/\-]\s*12\s*ans)", texte, re.IGNORECASE)
        if match:
            epreuve["age"] = match.group(1).strip()
        else:
            match = re.search(r"(11\s*ans)", texte, re.IGNORECASE)
            if match:
                epreuve["age"] = match.group(1).strip()

    # Nom de l'épreuve : première ligne significative (pas le badge, pas un champ,
    # pas un texte de navigation)
    for line in texte.split("\n"):
        line_s = line.strip()
        if (line_s
                and line_s not in ("SM", "SD")
                and line_s.lower() not in _TEXTES_NAVIGATION
                and not re.match(
                    r"^([ÂA]ge|Format|Classement|Tarif|Inscription)\s*:",
                    line_s, re.IGNORECASE)):
            epreuve["nom_epreuve"] = line_s
            break

    return epreuve if epreuve else None


def _extraire_epreuves_depuis_texte(texte):
    """Fallback texte brut : cherche les blocs avec badge SM et âge 11."""
    lines = texte.split("\n")

    i = 0
    while i < len(lines):
        line = lines[i]
        # Chercher le badge SM (mot isolé, pas SD)
        if re.search(r"\bSM\b", line) and not re.search(r"\bSD\b", line):
            # Trouver la fin du bloc : prochain badge SM ou SD
            bloc_fin = min(len(lines), i + 50)
            for j in range(i + 1, bloc_fin):
                if re.search(r"\b(?:SM|SD)\b", lines[j]):
                    bloc_fin = j
                    break

            bloc = "\n".join(lines[i:bloc_fin])

            if _bloc_contient_age_11(bloc):
                # Nom : première ligne significative (pas le badge, pas un champ,
                # pas un texte de navigation)
                nom = ""
                for k in range(i, bloc_fin):
                    l_s = lines[k].strip()
                    if (l_s and l_s not in ("SM", "SD")
                            and l_s.lower() not in _TEXTES_NAVIGATION
                            and not re.match(
                                r"^([ÂA]ge|Format|Classement|Tarif|Inscription)\s*:",
                                l_s, re.IGNORECASE)):
                        nom = l_s
                        break

                epreuve = {"nom_epreuve": nom}

                match = re.search(
                    r"Tarif\s+jeune\s*:?\s*([\d,]+\s*€)", bloc, re.IGNORECASE
                )
                if match:
                    epreuve["tarif_jeune"] = match.group(1).strip()

                # Classement : collecter TOUS les candidats (finditer)
                _cls_candidats = []
                for m in re.finditer(
                    r"Classement\s*:?\s*(.+?)(?:\n|$)", bloc, re.IGNORECASE
                ):
                    val = m.group(1).strip()
                    if val:
                        _cls_candidats.append(val)
                for m in re.finditer(
                    r"Classement\s*:?\s*\n+\s*(.+?)(?:\n|$)", bloc,
                    re.IGNORECASE,
                ):
                    val = m.group(1).strip()
                    if val:
                        _cls_candidats.append(val)
                for candidat in _cls_candidats:
                    if _est_classement_valide(candidat):
                        epreuve["classement"] = candidat
                        break

                match = re.search(
                    r"Format\s*[:\-–]?\s*(.+?)(?:\n|$)", bloc, re.IGNORECASE
                )
                if match and match.group(1).strip():
                    epreuve["format"] = match.group(1).strip()
                else:
                    match = re.search(
                        r"Format\s*[:\-–]?\s*\n+\s*(.+?)(?:\n|$)", bloc,
                        re.IGNORECASE,
                    )
                    if match and match.group(1).strip():
                        epreuve["format"] = match.group(1).strip()
                if "format" not in epreuve:
                    match = re.search(
                        r"(\d\s*set(?:s)?\s*(?:de\s*\d+\s*jeux?|gagnant).*?)(?:\n|$)",
                        bloc, re.IGNORECASE,
                    )
                    if match:
                        epreuve["format"] = match.group(1).strip()
                if "format" not in epreuve:
                    match = re.search(
                        r"((?:super[- ]?)?tie[- ]?break.*?)(?:\n|$)",
                        bloc, re.IGNORECASE,
                    )
                    if match:
                        epreuve["format"] = match.group(1).strip()

                match = re.search(
                    r"[ÂA]ge\s*:?\s*(.+?)(?:\n|$)", bloc, re.IGNORECASE
                )
                if match and match.group(1).strip():
                    epreuve["age"] = match.group(1).strip()
                if "age" not in epreuve:
                    match = re.search(
                        r"(11\s*[/\-]\s*12\s*ans)", bloc, re.IGNORECASE
                    )
                    if match:
                        epreuve["age"] = match.group(1).strip()

                return [epreuve]
        i += 1

    return []


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
# 5b. Rapport de validation post-extraction
# ---------------------------------------------------------------------------

def _rapport_validation(tournois):
    """Affiche un rapport de validation identifiant les tournois avec des champs manquants."""

    # Champs critiques au niveau tournoi
    champs_tournoi = [
        ("nom", "Nom"),
        ("club", "Club"),
        ("debut", "Début"),
        ("fin", "Fin"),
        ("surface", "Surface"),
        ("lieu", "Lieu"),
    ]
    # Champs critiques au niveau épreuve
    champs_epreuve = [
        ("classement", "Classement"),
        ("format", "Format"),
        ("tarif_jeune", "Tarif jeune"),
    ]

    # Filtrer uniquement les tournois qui ont des épreuves 11/12 ans
    tournois_avec_epreuves = [t for t in tournois if t.get("epreuves")]

    if not tournois_avec_epreuves:
        return

    problemes = []  # Liste de (code, nom, [champs_manquants])

    for t in tournois_avec_epreuves:
        code = t.get("code", "?")
        nom = t.get("nom", "(sans nom)")
        manquants = []

        # Vérifier les champs tournoi
        for champ, label in champs_tournoi:
            val = t.get(champ, "")
            if not val or not str(val).strip():
                manquants.append(label)

        # Vérifier les champs épreuve
        for ep in t.get("epreuves", []):
            for champ, label in champs_epreuve:
                val = ep.get(champ, "")
                if not val or not str(val).strip():
                    ep_nom = ep.get("nom_epreuve", "épreuve")
                    tag = f"{label} ({ep_nom})"
                    if tag not in manquants:
                        manquants.append(tag)

        if manquants:
            problemes.append((code, nom, manquants))

    # Affichage du rapport
    print(f"\n{'='*70}")
    print(f"  RAPPORT DE VALIDATION")
    print(f"{'='*70}")
    print(f"  Tournois avec épreuve(s) 11/12 ans : {len(tournois_avec_epreuves)}")

    if not problemes:
        print(f"  ✓ Tous les champs importants sont renseignés !")
    else:
        print(f"  ⚠ {len(problemes)} tournoi(s) avec champ(s) manquant(s) :\n")
        for code, nom, manquants in problemes:
            print(f"  [{code}] {nom}")
            for m in manquants:
                print(f"           → {m} : VIDE")
            print()

    # Résumé par type de champ manquant
    if problemes:
        compteur = {}
        for _, _, manquants in problemes:
            for m in manquants:
                # Extraire le nom du champ (avant la parenthèse si épreuve)
                champ_base = m.split(" (")[0]
                compteur[champ_base] = compteur.get(champ_base, 0) + 1

        print(f"  --- Synthèse des champs manquants ---")
        for champ, nb in sorted(compteur.items(), key=lambda x: -x[1]):
            print(f"    {champ:15s} : {nb} tournoi(s)")

    # Tournois en erreur
    erreurs = [t for t in tournois if t.get("erreur")]
    if erreurs:
        print(f"\n  --- Tournois en erreur ({len(erreurs)}) ---")
        for t in erreurs:
            print(f"  [{t.get('code', '?')}] {t.get('erreur', '')}")

    print(f"{'='*70}")


# ---------------------------------------------------------------------------
# 6. Sauvegarde / reprise JSON
# ---------------------------------------------------------------------------

# ---------------------------------------------------------------------------
# 6b. Historique des extractions (pour mode mise à jour)
# ---------------------------------------------------------------------------

HISTORIQUE_DIR = "historique"


def _formater_duree(secondes):
    """Formate une durée en secondes en chaîne lisible (ex: 1h02min37s)."""
    secondes = int(secondes)
    if secondes < 60:
        return f"{secondes}s"
    minutes, sec = divmod(secondes, 60)
    if minutes < 60:
        return f"{minutes}min{sec:02d}s"
    heures, minutes = divmod(minutes, 60)
    return f"{heures}h{minutes:02d}min{sec:02d}s"


def _chemin_historique(tag):
    """Retourne le chemin du fichier historique pour un tag donné."""
    # Nettoyer le tag pour en faire un nom de fichier valide
    tag_safe = re.sub(r'[^\w\-]', '_', tag)
    return os.path.join(HISTORIQUE_DIR, f"{tag_safe}.json")


def charger_historique(tag):
    """Charge l'historique des codes déjà extraits pour un tag.

    Retourne un set de codes tournoi.
    """
    chemin = _chemin_historique(tag)
    if not os.path.exists(chemin):
        return set()
    with open(chemin, "r", encoding="utf-8") as f:
        data = json.load(f)
    return set(data.get("codes", []))


def sauvegarder_historique(tag, codes):
    """Sauvegarde les codes tournoi extraits dans l'historique du tag.

    Les nouveaux codes sont ajoutés aux codes existants (union).
    """
    os.makedirs(HISTORIQUE_DIR, exist_ok=True)
    codes_existants = charger_historique(tag)
    codes_tous = codes_existants | set(codes)
    chemin = _chemin_historique(tag)
    data = {
        "tag": tag,
        "codes": sorted(codes_tous),
        "nb_codes": len(codes_tous),
        "derniere_maj": time.strftime("%Y-%m-%d %H:%M:%S"),
    }
    with open(chemin, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)
    return chemin


def lister_historiques():
    """Liste tous les tags disponibles dans l'historique."""
    if not os.path.exists(HISTORIQUE_DIR):
        return []
    tags = []
    for f in sorted(os.listdir(HISTORIQUE_DIR)):
        if f.endswith(".json"):
            chemin = os.path.join(HISTORIQUE_DIR, f)
            with open(chemin, "r", encoding="utf-8") as fh:
                data = json.load(fh)
            tags.append({
                "tag": data.get("tag", f[:-5]),
                "nb_codes": data.get("nb_codes", 0),
                "derniere_maj": data.get("derniere_maj", "?"),
            })
    return tags


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

def _mode_test(args):
    """Mode test : analyse un seul tournoi et affiche un diagnostic complet."""
    code = args.test
    url = f"https://tenup.fft.fr/tournoi/{code}"
    print(f"=== MODE TEST - Tournoi {code} ===")
    print(f"URL : {url}")
    print()

    # Préparer le fetch
    pw = browser = pw_page = driver = session = None
    try:
        if args.playwright:
            print("Mode : Playwright (Chromium headless)")
            pw, browser, pw_page = _creer_playwright()
            html = _fetch_html_playwright(url, pw_page)
        elif args.selenium:
            print("Mode : Selenium")
            driver = _creer_driver_selenium()
            html = _fetch_html_selenium(url, driver)
        else:
            print("Mode : requests")
            print("ATTENTION : tenup.fft.fr nécessite JavaScript !")
            print("           Utilisez --playwright pour un résultat fiable.")
            session = requests.Session()
            session.headers.update({
                "User-Agent": (
                    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                    "AppleWebKit/537.36"
                ),
            })
            html = _fetch_html_requests(url, session)
    finally:
        if browser:
            browser.close()
        if pw:
            pw.stop()
        if driver:
            driver.quit()

    print(f"\nHTML récupéré : {len(html)} caractères")

    # Sauvegarder le HTML
    debug_file = f"debug_tournoi_{code}.html"
    with open(debug_file, "w", encoding="utf-8") as f:
        f.write(html)
    print(f"HTML sauvegardé dans : {debug_file}")

    # Analyser le contenu brut
    soup = BeautifulSoup(html, "html.parser")
    texte = _normaliser_espaces(soup.get_text(separator="\n"))

    # Diagnostic : chercher les mots-clés attendus
    print("\n--- DIAGNOSTIC DU CONTENU ---")
    mots_cles = [
        "SM",
        "SD",
        "Simple Messieurs",
        "Simple Mixte",
        "TMC",
        "11/12",
        "11 ans",
        "Âge",
        "Tarif jeune",
        "Classement",
        "Format",
    ]
    for mot in mots_cles:
        count = texte.lower().count(mot.lower())
        status = f"TROUVÉ ({count}x)" if count > 0 else "ABSENT"
        print(f"  '{mot}' : {status}")

    # Montrer les lignes contenant "Simple" ou "11/12"
    print("\n--- LIGNES CONTENANT 'SM', 'SD', '11' ou 'Âge' ---")
    lines = texte.split("\n")
    shown = 0
    for i, line in enumerate(lines):
        line_stripped = line.strip()
        if not line_stripped:
            continue
        if re.search(r"\bSM\b|\bSD\b|simple|TMC|11.?12|\b11\s*ans|[âa]ge\s*:", line, re.IGNORECASE):
            print(f"  L{i}: {line_stripped[:120]}")
            shown += 1
            if shown >= 30:
                print("  ... (tronqué)")
                break
    if shown == 0:
        print("  (aucune ligne trouvée - le contenu JS n'a probablement pas été chargé)")

    # --- Données PDF (si un PDF est fourni) ---
    pdf_info = {}
    if args.pdf:
        print("\n--- DONNÉES EXTRAITES DU PDF ---")
        pdf_tournois = extraire_tournois_du_pdf(args.pdf)
        pdf_data = {t["code"]: t for t in pdf_tournois}
        pdf_info = pdf_data.get(code, {})
        if pdf_info:
            for champ in ("nom", "club", "debut", "fin", "surface",
                          "juge_arbitre", "lieu", "mail", "tel",
                          "tarif_jeune_pdf"):
                val = pdf_info.get(champ, "")
                label = champ.replace("_", " ").capitalize()
                print(f"  {label} : {val or '(vide)'}")
        else:
            print(f"  Code {code} non trouvé dans le PDF")

    # Lancer le parsing URL
    print("\n--- RÉSULTAT DU PARSING URL ---")
    url_info = parser_page_tournoi(html, code)
    print(f"  Nom : {url_info.get('nom', '(vide)')}")
    print(f"  Club : {url_info.get('club', '(vide)')}")
    print(f"  Début : {url_info.get('debut', '(vide)')}")
    print(f"  Fin : {url_info.get('fin', '(vide)')}")
    print(f"  Lieu : {url_info.get('lieu', '(vide)')}")

    # Fusionner PDF + URL
    info = _fusionner_donnees(pdf_info, url_info)

    print("\n--- RÉSULTAT FUSIONNÉ (PDF + URL) ---")
    for champ in ("nom", "club", "debut", "fin", "surface",
                  "juge_arbitre", "lieu", "mail", "tel"):
        val = info.get(champ, "")
        label = champ.replace("_", " ").capitalize()
        source = ""
        if pdf_info.get(champ) and pdf_info[champ] == val:
            source = " [PDF]"
        elif url_info.get(champ) and url_info[champ] == val:
            source = " [URL]"
        print(f"  {label} : {val or '(vide)'}{source}")

    epreuves = info.get("epreuves", [])
    if epreuves:
        print(f"  Épreuves 11/12 ans trouvées : {len(epreuves)}")
        for j, ep in enumerate(epreuves, 1):
            print(f"    [{j}] {ep.get('nom_epreuve', '?')}")
            print(f"        Âge : {ep.get('age', '?')}")
            print(f"        Tarif jeune : {ep.get('tarif_jeune', '?')}")
            print(f"        Classement : {ep.get('classement', '?')} [URL]")
            print(f"        Format : {ep.get('format', '?')} [URL]")
    else:
        print("  AUCUNE ÉPREUVE 11/12 ANS TROUVÉE")
        print()
        print("  Conseils :")
        print("  - Vérifiez le fichier debug_tournoi_{}.html".format(code))
        print("  - Si le fichier HTML est presque vide, le JS n'a pas été chargé")
        print("  - Essayez avec --playwright si ce n'est pas déjà fait")


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
        "--playwright", action="store_true",
        help="Utiliser Playwright (Chromium headless) - RECOMMANDÉ pour tenup.fft.fr",
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
    parser.add_argument(
        "--debug", action="store_true",
        help="Sauvegarder le HTML de chaque page dans le dossier debug_html/",
    )
    parser.add_argument(
        "--debug-pdf", action="store_true",
        help="Afficher le texte brut extrait du PDF (pour diagnostic)",
    )
    parser.add_argument(
        "--test", metavar="CODE",
        help="Tester avec un seul code tournoi et afficher un diagnostic détaillé",
    )
    parser.add_argument(
        "--tag", metavar="NOM",
        help=(
            "Identifiant de la recherche (ex: idf-oct-dec-2026, loire-atlantique-jan-2027). "
            "Permet de séparer les historiques par région/période pour le mode mise à jour."
        ),
    )
    parser.add_argument(
        "--mode", choices=["complet", "maj"],
        default="complet",
        help=(
            "Mode d'extraction : 'complet' extrait tous les tournois (défaut), "
            "'maj' n'extrait que les nouveaux tournois par rapport à l'historique du --tag."
        ),
    )
    parser.add_argument(
        "--list-tags", action="store_true",
        help="Lister tous les tags d'historique disponibles et quitter.",
    )
    args = parser.parse_args()

    # Mode liste des tags
    if args.list_tags:
        tags = lister_historiques()
        if not tags:
            print("Aucun historique trouvé.")
        else:
            print(f"{'Tag':<35} {'Tournois':>10}   Dernière MAJ")
            print("-" * 70)
            for t in tags:
                print(f"{t['tag']:<35} {t['nb_codes']:>10}   {t['derniere_maj']}")
        return

    # Mode test : tester un seul tournoi avec diagnostic détaillé
    if args.test:
        _mode_test(args)
        return

    # Vérification : --mode maj nécessite --tag
    if args.mode == "maj" and not args.tag:
        parser.error("Le mode 'maj' nécessite --tag pour identifier l'historique à comparer.")

    # Vérifications des arguments
    if not args.pdf and not args.resume:
        parser.error("Fournissez un fichier PDF ou utilisez --resume")

    # Charger ou extraire les codes + infos PDF
    pdf_data = {}  # code -> info extraite du PDF
    if args.resume and os.path.exists(args.resume):
        print(f"Reprise depuis : {args.resume}")
        tournois, codes = charger_progres(args.resume)
        print(f"  -> {len(tournois)} tournoi(s) déjà traité(s)")
        print(f"  -> {len(codes)} tournoi(s) restant(s)")
        # Ré-extraire les données PDF si le fichier est fourni
        if args.pdf:
            pdf_tournois = extraire_tournois_du_pdf(args.pdf)
            pdf_data = {t["code"]: t for t in pdf_tournois}
            print(f"  -> {len(pdf_data)} fiche(s) PDF chargée(s)")
    else:
        print(f"Lecture du PDF : {args.pdf}")
        pdf_tournois = extraire_tournois_du_pdf(args.pdf)
        pdf_data = {t["code"]: t for t in pdf_tournois}
        codes = [t["code"] for t in pdf_tournois]
        print(f"  -> {len(codes)} tournoi(s) trouvé(s) dans le PDF")
        # Afficher les infos PDF extraites en aperçu
        for t in pdf_tournois[:3]:
            nom = t.get("nom", "?") or "?"
            club = t.get("club", "") or ""
            debut = t.get("debut", "") or ""
            apercu = f"     {t['code']}: {nom}"
            if club:
                apercu += f" ({club})"
            if debut:
                apercu += f" - {debut}"
            print(apercu)
        if len(pdf_tournois) > 3:
            print(f"     ... et {len(pdf_tournois) - 3} autre(s)")
        tournois = []

    # Mode debug PDF : afficher le texte brut extrait
    if args.debug_pdf and args.pdf:
        reader = PdfReader(args.pdf)
        print("\n=== DEBUG PDF - TEXTE BRUT ===")
        for i, page in enumerate(reader.pages, 1):
            texte = page.extract_text()
            if texte:
                print(f"\n--- PAGE {i} ---")
                print(_normaliser_espaces(texte))
        print("\n=== FIN DEBUG PDF ===\n")
        print("Données extraites par tournoi :")
        for t in pdf_data.values():
            print(f"\n  Code: {t['code']}")
            for champ in ("nom", "club", "debut", "fin", "surface",
                          "juge_arbitre", "lieu", "mail", "tel",
                          "tarif_jeune_pdf"):
                val = t.get(champ, "")
                label = champ.replace("_", " ").capitalize()
                print(f"    {label}: {val or '(vide)'}")

    # Filtrage en mode mise à jour : ne garder que les nouveaux codes
    if args.tag and args.mode == "maj" and not args.resume:
        codes_connus = charger_historique(args.tag)
        nb_avant = len(codes)
        codes = [c for c in codes if c not in codes_connus]
        # Filtrer aussi pdf_data pour ne garder que les nouveaux
        pdf_data = {c: v for c, v in pdf_data.items() if c not in codes_connus}
        nb_nouveaux = len(codes)
        print(f"\n  Mode mise à jour (tag: {args.tag}) :")
        print(f"    Tournois dans le PDF      : {nb_avant}")
        print(f"    Déjà dans l'historique     : {nb_avant - nb_nouveaux}")
        print(f"    Nouveaux à extraire        : {nb_nouveaux}")
        if nb_nouveaux == 0:
            print("\n  Aucun nouveau tournoi à extraire.")
            sys.exit(0)
    elif args.tag and args.mode == "complet":
        print(f"\n  Mode complet (tag: {args.tag}) : extraction de tous les tournois.")

    if not codes and not tournois:
        print("Aucun code de tournoi trouvé.")
        sys.exit(1)

    total = len(tournois) + len(codes)

    # Préparer le mode de fetch
    driver = None
    session = None
    pw = None
    browser = None
    pw_page = None
    if args.playwright:
        print("Mode Playwright (Chromium headless) - recommandé")
        pw, browser, pw_page = _creer_playwright()
    elif args.selenium:
        print("Mode Selenium (Chrome headless)")
        driver = _creer_driver_selenium()
    else:
        print("Mode requests (ATTENTION: tenup.fft.fr nécessite JS, utilisez --playwright)")
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

    # Créer le dossier debug si nécessaire
    if args.debug:
        os.makedirs("debug_html", exist_ok=True)
        print("Mode debug activé : HTML sauvegardé dans debug_html/")

    # Traiter les tournois restants
    t_debut = time.time()
    nb_deja_traites = len(tournois)
    print(f"\nExtraction des tournois (délai: {args.delay}s)...")
    try:
        while codes:
            code = codes[0]
            i = len(tournois) + 1
            # Calcul du temps écoulé et ETA
            elapsed = time.time() - t_debut
            nb_faits = len(tournois) - nb_deja_traites
            if nb_faits > 0:
                temps_par_tournoi = elapsed / nb_faits
                restants = len(codes)
                eta = temps_par_tournoi * restants
                eta_str = f" | ETA: {_formater_duree(eta)}"
            else:
                eta_str = ""
            print(
                f"  [{i:3d}/{total}] ({_formater_duree(elapsed)}{eta_str}) "
                f"Tournoi {code}...",
                end=" ", flush=True,
            )

            time.sleep(args.delay)

            try:
                url = f"https://tenup.fft.fr/tournoi/{code}"
                if args.playwright:
                    html = _fetch_html_playwright(url, pw_page)
                elif args.selenium:
                    html = _fetch_html_selenium(url, driver)
                else:
                    html = _fetch_html_requests(url, session)

                # Sauvegarder le HTML en mode debug
                if args.debug:
                    debug_path = f"debug_html/tournoi_{code}.html"
                    with open(debug_path, "w", encoding="utf-8") as f:
                        f.write(html)

                url_info = parser_page_tournoi(html, code)
                # Fusionner PDF (infos générales) + URL (classement/format)
                info = _fusionner_donnees(pdf_data.get(code, {}), url_info)
            except Exception as e:
                # Utiliser les données PDF même si l'URL échoue
                info = pdf_data.get(code, {}).copy()
                if not info:
                    info = {
                        "code": code,
                        "url": f"https://tenup.fft.fr/tournoi/{code}",
                        "nom": "", "club": "", "debut": "", "fin": "",
                        "surface": "", "juge_arbitre": "", "lieu": "",
                        "mail": "", "tel": "", "epreuves": [],
                    }
                info["erreur"] = f"URL: {e}"

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
                print(f"OK - {info['nom']} - aucune épreuve 11/12 ans trouvée")

            # Sauvegarde intermédiaire tous les 10 tournois
            if len(tournois) % 10 == 0:
                sauvegarder_progres(tournois, codes, args.save)

    except KeyboardInterrupt:
        duree_interr = time.time() - t_debut
        print(f"\n\nInterruption après {_formater_duree(duree_interr)} ! Sauvegarde en cours...")
        sauvegarder_progres(tournois, codes, args.save)
        print(f"Progrès sauvegardé dans : {args.save}")
        print(f"Reprenez avec : python extract_tournois.py --resume {args.save}")
        sys.exit(0)
    finally:
        if driver:
            driver.quit()
        if browser:
            browser.close()
        if pw:
            pw.stop()

    # Sauvegarde finale
    sauvegarder_progres(tournois, [], args.save)

    # Générer le fichier Excel
    print(f"\nGénération du fichier Excel : {args.output}")
    nb_lignes = generer_excel(tournois, args.output)
    print(f"  -> {nb_lignes} ligne(s) écrite(s)")

    # Sauvegarder l'historique si un tag est fourni
    if args.tag:
        codes_extraits = [t["code"] for t in tournois if t.get("code")]
        chemin_hist = sauvegarder_historique(args.tag, codes_extraits)
        codes_total = len(charger_historique(args.tag))
        print(f"\n  Historique mis à jour : {chemin_hist} ({codes_total} tournoi(s) au total)")

    # Résumé
    duree_totale = time.time() - t_debut
    nb_avec_epreuves = sum(1 for t in tournois if t.get("epreuves"))
    nb_erreurs = sum(1 for t in tournois if t.get("erreur"))
    print(f"\nRésumé :")
    print(f"  Tournois traités : {len(tournois)}")
    print(f"  Avec épreuve(s) correspondante(s) : {nb_avec_epreuves}")
    print(f"  Erreurs : {nb_erreurs}")
    print(f"  Fichier Excel : {args.output}")
    print(f"  Script exécuté en {_formater_duree(duree_totale)}")

    # Rapport de validation
    _rapport_validation(tournois)


if __name__ == "__main__":
    main()
