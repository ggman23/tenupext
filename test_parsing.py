#!/usr/bin/env python3
"""Tests unitaires pour le parsing des tournois (sans réseau)."""

import sys
import os

# Ajouter le répertoire courant au path
sys.path.insert(0, os.path.dirname(__file__))

from extract_tournois import (
    parser_page_tournoi,
    _extraire_epreuves,
    _extraire_details_epreuve,
    _extraire_epreuves_depuis_texte,
    _bloc_contient_age_11,
    _normaliser_espaces,
    _est_valeur_valide,
    _est_classement_valide,
    _fusionner_donnees,
)
from bs4 import BeautifulSoup

# ────────────────────────────────────────────────────────────────────
# HTML simulé d'une page tournoi tenup.fft.fr
# ────────────────────────────────────────────────────────────────────

HTML_TOURNOI_COMPLET = """
<html>
<head><title>Tournoi Test</title></head>
<body>
<div class="tournoi-detail-page-title">
    <h1 itemprop="name">TOURNOI JEUNES DE PRINTEMPS</h1>
</div>
<h2 class="tournoi-detail-page-club">TC EXAMPLE / PARIS</h2>
<span class="tournoi-detail-page-date-debut">15/03/2026</span>
<span class="tournoi-detail-page-date-fin">22/03/2026</span>
<span class="tournoi-detail-page-competition-surfaces-content">Terre battue</span>
<div class="tournoi-detail-page-lieu-addr1">Stade Roland Garros</div>
<div class="tournoi-detail-page-lieu-addr2">75016 PARIS</div>

<!-- Épreuves -->
<div class="epreuve-card">
    <div class="badge">SM</div>
    <div class="epreuve-nom">Simple Messieurs 11/12 ans</div>
    <div class="epreuve-details">
        <div>Âge : 11/12 ans</div>
        <div>Tarif jeune : 12,00 €</div>
        <div>Classement : NC à 30/5</div>
        <div>Format : 2 sets + super tie-break</div>
    </div>
</div>

<!-- Épreuve SD (ne doit PAS être retournée) -->
<div class="epreuve-card-sd">
    <div class="badge">SD</div>
    <div>Simple Dames 11/12 ans</div>
    <div>Âge : 11/12 ans</div>
    <div>Tarif jeune : 10,00 €</div>
</div>
</body>
</html>
"""

HTML_SM_TEXTE_LONG = """
<html><body>
<div class="epreuve-card">
    <span>SM</span>
    <div>Simple Messieurs TMC Garçons</div>
    <div>Âge : 11 ans</div>
    <div>Tarif jeune : 8,00 €</div>
    <div>Classement</div>
    <div>NC à 30/4</div>
    <div>Format</div>
    <div>3 sets</div>
</div>
</body></html>
"""

HTML_CLASSEMENT_LIGNE_SEPAREE = """
<html><body>
<div class="epreuve-wrapper">
    <div class="badge">SM</div>
    <div>Épreuve Garçons 11/12 ans</div>
    <div>Âge : 11/12 ans</div>
    <div>Tarif jeune : 15,00 €</div>
    <span>Classement</span>
    <span>NC à 30/3</span>
    <span>Format</span>
    <span>1 set</span>
</div>
</body></html>
"""

HTML_SANS_EPREUVE_11 = """
<html><body>
<div class="tournoi-detail-page-title">
    <h1 itemprop="name">TOURNOI ADULTES</h1>
</div>
<div class="epreuve-card">
    <div class="badge">SM</div>
    <div>Simple Messieurs Seniors</div>
    <div>Âge : 35 ans et plus</div>
    <div>Classement : NC à 15/4</div>
</div>
</body></html>
"""

HTML_SIMPLE_MESSIEURS_FALLBACK = """
<html><body>
<div class="epreuve-wrapper">
    <div class="epreuve-header">Simple Messieurs - 11/12 ans</div>
    <div>Âge : 11/12 ans</div>
    <div>Tarif jeune : 11,00 €</div>
    <div>Classement : NC à 30/5</div>
    <div>Format : 2 sets</div>
</div>
</body></html>
"""


def run_tests():
    passed = 0
    failed = 0
    errors = []

    def check(name, condition, detail=""):
        nonlocal passed, failed
        if condition:
            passed += 1
            print(f"  OK  {name}")
        else:
            failed += 1
            msg = f"  FAIL  {name}"
            if detail:
                msg += f" — {detail}"
            print(msg)
            errors.append(name)

    # ──────────────────────────────────────────────────────────────
    print("=== TEST 1 : Parsing complet d'un tournoi ===")
    info = parser_page_tournoi(HTML_TOURNOI_COMPLET, "123456")
    check("Code", info["code"] == "123456")
    check("Nom", info["nom"] == "TOURNOI JEUNES DE PRINTEMPS", f"got: {info['nom']!r}")
    check("Club", info["club"] == "TC EXAMPLE / PARIS", f"got: {info['club']!r}")
    check("Début", info["debut"] == "15/03/2026", f"got: {info['debut']!r}")
    check("Fin", info["fin"] == "22/03/2026", f"got: {info['fin']!r}")
    check("Surface", info["surface"] == "Terre battue", f"got: {info['surface']!r}")
    check("Lieu", "Stade Roland Garros" in info["lieu"], f"got: {info['lieu']!r}")
    check("Épreuves trouvées", len(info["epreuves"]) >= 1, f"got: {len(info['epreuves'])} épreuves")
    if info["epreuves"]:
        ep = info["epreuves"][0]
        check("Tarif jeune", "12" in ep.get("tarif_jeune", ""), f"got: {ep.get('tarif_jeune')!r}")
        check("Classement", "30/5" in ep.get("classement", ""), f"got: {ep.get('classement')!r}")
        check("Format", "super tie-break" in ep.get("format", "").lower() or "2 sets" in ep.get("format", "").lower(), f"got: {ep.get('format')!r}")
        check("Âge", "11" in ep.get("age", ""), f"got: {ep.get('age')!r}")

    # ──────────────────────────────────────────────────────────────
    print("\n=== TEST 2 : Classement et Format sur lignes séparées (DOM sibling) ===")
    info2 = parser_page_tournoi(HTML_SM_TEXTE_LONG, "222222")
    check("Épreuves trouvées", len(info2["epreuves"]) >= 1, f"got: {len(info2['epreuves'])}")
    if info2["epreuves"]:
        ep = info2["epreuves"][0]
        check("Tarif", "8" in ep.get("tarif_jeune", ""), f"got: {ep.get('tarif_jeune')!r}")
        check("Classement", "30/4" in ep.get("classement", ""), f"got: {ep.get('classement')!r}")
        check("Format", "3 sets" in ep.get("format", "").lower(), f"got: {ep.get('format')!r}")
        check("Âge", "11" in ep.get("age", ""), f"got: {ep.get('age')!r}")

    # ──────────────────────────────────────────────────────────────
    print("\n=== TEST 3 : Classement/Format avec <span> siblings ===")
    info3 = parser_page_tournoi(HTML_CLASSEMENT_LIGNE_SEPAREE, "333333")
    check("Épreuves trouvées", len(info3["epreuves"]) >= 1, f"got: {len(info3['epreuves'])}")
    if info3["epreuves"]:
        ep = info3["epreuves"][0]
        check("Classement", "30/3" in ep.get("classement", ""), f"got: {ep.get('classement')!r}")
        check("Format", "1 set" in ep.get("format", "").lower(), f"got: {ep.get('format')!r}")

    # ──────────────────────────────────────────────────────────────
    print("\n=== TEST 4 : Pas d'épreuve 11/12 ans -> liste vide ===")
    info4 = parser_page_tournoi(HTML_SANS_EPREUVE_11, "444444")
    check("Aucune épreuve 11/12", len(info4["epreuves"]) == 0, f"got: {len(info4['epreuves'])}")

    # ──────────────────────────────────────────────────────────────
    print("\n=== TEST 5 : Fallback 'Simple Messieurs' (pas de badge SM isolé) ===")
    info5 = parser_page_tournoi(HTML_SIMPLE_MESSIEURS_FALLBACK, "555555")
    check("Épreuves trouvées", len(info5["epreuves"]) >= 1, f"got: {len(info5['epreuves'])}")
    if info5["epreuves"]:
        ep = info5["epreuves"][0]
        check("Tarif", "11" in ep.get("tarif_jeune", ""), f"got: {ep.get('tarif_jeune')!r}")

    # ──────────────────────────────────────────────────────────────
    print("\n=== TEST 6 : _bloc_contient_age_11 ===")
    check("11/12 ans", _bloc_contient_age_11("Âge : 11/12 ans"))
    check("11 ans", _bloc_contient_age_11("Catégorie 11 ans"))
    check("11-12 ans", _bloc_contient_age_11("Âge : 11-12 ans"))
    check("Pas 11", not _bloc_contient_age_11("Âge : 13/14 ans"))
    check("Pas 11 (adulte)", not _bloc_contient_age_11("Seniors plus"))

    # ──────────────────────────────────────────────────────────────
    print("\n=== TEST 7 : _normaliser_espaces ===")
    check("Espace insécable", _normaliser_espaces("hello\u00a0world") == "hello world")
    check("Espace fine", _normaliser_espaces("test\u202fvalue") == "test value")

    # ──────────────────────────────────────────────────────────────
    print("\n=== TEST 8 : _est_valeur_valide ===")
    check("Valeur OK", _est_valeur_valide("NC à 30/5"))
    check("Valeur OK date", _est_valeur_valide("15/03/2026"))
    check("Vide", not _est_valeur_valide(""))
    check("JS function", not _est_valeur_valide("function() { return 42; }"))
    check("HTTP link valide", _est_valeur_valide("https://example.com/foo"))
    check("Pub JS", not _est_valeur_valide("googletag.pubads().enableSingleRequest()"))

    # ──────────────────────────────────────────────────────────────
    print("\n=== TEST 9 : _fusionner_donnees ===")
    pdf = {"code": "111", "nom": "TOURNOI PDF", "club": "Club PDF", "debut": "01/01/2026", "fin": "", "surface": "Dur", "juge_arbitre": "", "lieu": ""}
    url = {"code": "111", "nom": "TOURNOI URL", "club": "", "debut": "", "fin": "10/01/2026", "surface": "", "juge_arbitre": "M. Dupont", "lieu": "Paris", "epreuves": [{"tarif_jeune": "10€"}]}
    merged = _fusionner_donnees(pdf, url)
    check("PDF prioritaire nom", merged["nom"] == "TOURNOI PDF", f"got: {merged['nom']!r}")
    check("URL comble fin", merged["fin"] == "10/01/2026", f"got: {merged['fin']!r}")
    check("URL comble JA", merged["juge_arbitre"] == "M. Dupont", f"got: {merged['juge_arbitre']!r}")
    check("Épreuves URL", len(merged.get("epreuves", [])) == 1)

    # ──────────────────────────────────────────────────────────────
    print("\n=== TEST 10 : Extraction depuis texte brut (fallback) ===")
    texte_brut = """
Épreuves du tournoi
SM
Simple Messieurs 11/12 ans
Âge : 11/12 ans
Tarif jeune : 9,50 €
Classement : NC à 30/5
Format : 2 sets

SD
Simple Dames 13/14 ans
Âge : 13/14 ans
"""
    epreuves = _extraire_epreuves_depuis_texte(texte_brut)
    check("Fallback texte: trouvé", len(epreuves) >= 1, f"got: {len(epreuves)}")
    if epreuves:
        ep = epreuves[0]
        check("Fallback texte: tarif", "9,50" in ep.get("tarif_jeune", ""), f"got: {ep.get('tarif_jeune')!r}")
        check("Fallback texte: classement", "30/5" in ep.get("classement", ""), f"got: {ep.get('classement')!r}")

    # ──────────────────────────────────────────────────────────────
    print("\n=== TEST 11 : _est_classement_valide ===")
    check("NC - N1 valide", _est_classement_valide("NC - N1"))
    check("NC - 30/5 valide", _est_classement_valide("NC - 30/5"))
    check("40 - 30/1 valide", _est_classement_valide("40 - 30/1"))
    check("30/5 - 30/1 valide", _est_classement_valide("30/5 - 30/1"))
    check("NC valide", _est_classement_valide("NC"))
    check("NC à 30/5 valide", _est_classement_valide("NC à 30/5"))
    check("Texte parasite rejeté", not _est_classement_valide(
        ", proximité et date d'inscription seront les critères pris en compte)"
    ))
    check("Phrase longue rejetée", not _est_classement_valide(
        "demandé (classement, proximité et date d'inscription seront les critères)"
    ))
    check("Vide rejeté", not _est_classement_valide(""))
    check("None rejeté", not _est_classement_valide(None))

    # ──────────────────────────────────────────────────────────────
    print("\n=== TEST 12 : Classement parasite (bug 196945) ===")
    # Reproduit la structure HTML réelle de tenup.fft.fr tournoi 196945 :
    # - Description contenant "Classement demandé (...)" en texte libre
    # - Vrai classement dans div.epreuve-detail-classement-detail
    HTML_CLASSEMENT_PARASITE = """
    <html><body>
    <div class="epreuve-card">
        <div class="badge">SM</div>
        <div>Simple Messieurs (TS)</div>
        <p>Classement demandé (classement, proximité et date d'inscription seront les critères pris en compte)</p>
        <div class="epreuve-detail-age-classement">
            <div>Âge : 11/12 ans</div>
            <div class="epreuve-detail-classement-detail">
                Classement : 30/5 - 30/1
            </div>
            <div class="epreuve-detail-format">Format : 2 - 2 sets à 6 jeux ; 3ème set = SJD à 10 points</div>
        </div>
        <div>Tarif jeune : 35,00 €</div>
    </div>
    </body></html>
    """
    info_parasite = parser_page_tournoi(HTML_CLASSEMENT_PARASITE, "196945")
    check("Épreuve trouvée", len(info_parasite["epreuves"]) >= 1,
          f"got: {len(info_parasite['epreuves'])}")
    if info_parasite["epreuves"]:
        ep = info_parasite["epreuves"][0]
        classement = ep.get("classement", "")
        check("Classement valide (pas de texte parasite)",
              "proximité" not in classement,
              f"got: {classement!r}")
        check("Classement = 30/5 - 30/1",
              "30/5" in classement and "30/1" in classement,
              f"got: {classement!r}")
        fmt = ep.get("format", "")
        check("Format via CSS selector",
              "2 sets" in fmt.lower() or "SJD" in fmt,
              f"got: {fmt!r}")

    # ──────────────────────────────────────────────────────────────
    print("\n=== TEST 13 : 'Retour aux résultats' filtré du nom d'épreuve ===")
    HTML_RETOUR_RESULTATS = """
    <html><body>
    <div class="epreuve-card">
        <a href="/tournois">Retour aux résultats</a>
        <div class="badge">SM</div>
        <div>Simple Messieurs 11/12 ans</div>
        <div>Âge : 11/12 ans</div>
        <div>Tarif jeune : 15,00 €</div>
        <div>Classement : NC - N1</div>
        <div>Format : 2 sets à 6 jeux</div>
    </div>
    </body></html>
    """
    info_retour = parser_page_tournoi(HTML_RETOUR_RESULTATS, "111111")
    check("Épreuve trouvée", len(info_retour["epreuves"]) >= 1)
    if info_retour["epreuves"]:
        ep = info_retour["epreuves"][0]
        nom = ep.get("nom_epreuve", "")
        check("Nom != 'Retour aux résultats'",
              "retour" not in nom.lower(),
              f"got: {nom!r}")
        check("Nom contient 'Simple Messieurs'",
              "Simple Messieurs" in nom,
              f"got: {nom!r}")

    # ──────────────────────────────────────────────────────────────
    print("\n=== TEST 14 : Fallback texte avec classement parasite ===")
    texte_parasite = """
SM
Simple Messieurs 11/12 ans
Âge : 11/12 ans
Tarif jeune : 35,00 €
Classement demandé (classement, proximité et date d'inscription)
Classement : 30/5 - 30/1
Format : 2 sets

SD
"""
    epreuves_p = _extraire_epreuves_depuis_texte(texte_parasite)
    check("Fallback: trouvé", len(epreuves_p) >= 1)
    if epreuves_p:
        ep = epreuves_p[0]
        cls = ep.get("classement", "")
        check("Fallback: pas de texte parasite",
              "proximité" not in cls,
              f"got: {cls!r}")

    # ──────────────────────────────────────────────────────────────
    print(f"\n{'='*50}")
    print(f"Résultat : {passed} OK, {failed} FAIL")
    if errors:
        print(f"Tests échoués : {', '.join(errors)}")
    return failed == 0


if __name__ == "__main__":
    success = run_tests()
    sys.exit(0 if success else 1)
