#!/usr/bin/env python3
"""
06-generer-mails.py : prépare les mails de prospection à partir de ta liste LOCALE.

Ce script NE ENVOIE RIEN. Il écrit un brouillon .eml par prospect (s'ouvre dans
Outlook, Thunderbird, Apple Mail) et un rapport. Tu relis, puis tu envoies à la main.

Tes coordonnées (e-mail, téléphone, SIRET) viennent d'un fichier config-local.json, pas du script.

Utilisation :
  1. Crée config-local.json hors du repo, par exemple dans ~/prospection-local/ :
       {"email_pro": "...", "telephone": "...", "siret": "...", "site": "ducoinweb.fr"}
  2. Garde ta liste en local : le tableau 08-suivi-prospects.xlsx (onglet Prospects)
     ou un CSV avec les mêmes en-têtes de colonnes.
  3. Lance :
       python3 06-generer-mails.py ~/prospection-local/08-suivi-prospects.xlsx \
           --config ~/prospection-local/config-local.json
     Options : --sortie DOSSIER, --date AAAA-MM-JJ (pour simuler un autre jour),
               --groupe NOM (ex. Restauration), --limite N (nombre maximum de mails).
     Le fichier 03-prospects-64-propre.xlsx (colonnes Nom, Email, État du site) est accepté tel quel.
     Premier contact : seulement les lignes dont « État du site » vaut « Aucun site connu »
     ET dont la vérification (04-verifier-sites.py, option --verif) dit « Aucun site apparent ».
  4. Ouvre le dossier de sortie, relis le rapport.txt, vérifie chaque brouillon,
     envoie, puis note les dates d'envoi dans le tableau de suivi (le script ne
     modifie jamais ton tableau).

Dépendance : openpyxl, seulement si ta liste est un .xlsx (pip install openpyxl).
"""

import argparse
import csv
import datetime as dt
import json
import re
import sys
from email.message import EmailMessage
from pathlib import Path

# ---------------------------------------------------------------------------
# MODÈLES : brouillon repris de 05-emails-prospection.docx, À VALIDER AVEC TOI.
# Les mots entre accolades sont remplacés par le script. Pour changer un texte,
# modifie seulement les chaînes ci-dessous.
# Ne modifie pas le paragraphe MENTION : il contient l'information exigée par le RGPD.
# ---------------------------------------------------------------------------
LIEN_RDV = "https://cal.meetergo.com/ducoinco/30-min-meeting-or-ducoinco"

SIGNATURE = (
    "Cordialement,\n"
    "{nom}\n"
    "{marque}\n"
    "{telephone} | {email_pro}\n"
    "SIRET {siret}"
)

MENTION = (
    "Vous recevez ce message parce que votre entreprise figure sur {source}. "
    "J'utilise vos coordonnées uniquement pour vous proposer mes services, et je les "
    "conserve 3 ans au plus à compter de la collecte ou de notre dernier échange. "
    "Pour ne plus recevoir de message de ma part, répondez « STOP » : je supprime vos "
    "coordonnées de ma liste de prospection et je ne vous recontacte pas. Vos droits et "
    "l'identité du responsable de traitement : section « Prospection commerciale » de la "
    "politique de confidentialité, sur mon site."
)

MODELES = {
    "A": {  # Premier contact (jour 0)
        "objet": "Un site pour {activite} à {ville}, en ligne en 5 jours",
        "corps": (
            "{salut}\n\n"
            "Je suis {nom}, je crée des sites vitrines pour des professionnels. En cherchant "
            "les entreprises de {ville}, je n'ai pas trouvé de site pour {entreprise}.\n\n"
            "Je propose un site de 5 pages au plus pour 900 €, en ligne 5 jours après réception "
            "de vos textes et photos. Le nom de domaine et l'hébergement sont à votre nom et je "
            "vous remets les accès. Il n'y a pas de contrat de maintenance. Le seul frais "
            "extérieur est le nom de domaine, environ 10 à 15 € par an.\n\n"
            "Si cela vous intéresse, vous pouvez réserver un appel gratuit de 30 minutes : "
            + LIEN_RDV
        ),
    },
    "B": {  # Première relance (5 jours ouvrés après le premier contact)
        "objet": "Re : un site pour {entreprise}",
        "corps": (
            "{salut}\n\n"
            "Je reviens vers vous au sujet de mon message du {date_premier_contact}.\n\n"
            "Deux formules, à prix fixe : Essentiel à 900 € (jusqu'à 5 pages) et Complet à "
            "1 600 € (jusqu'à 8 pages, galerie photo et référencement local renforcé). Les deux "
            "incluent le formulaire de contact, la prise de rendez-vous en ligne, le référencement "
            "de base et deux allers-retours de retouches.\n\n"
            "L'appel de 30 minutes est gratuit et débouche sur un devis écrit, sans engagement : "
            + LIEN_RDV
        ),
    },
    "C": {  # Deuxième relance (7 jours ouvrés après la première relance), dernier message
        "objet": "Dernier message : site de {entreprise}",
        "corps": (
            "{salut}\n\n"
            "Ce message est le dernier de ma part. Si un site n'est pas d'actualité pour "
            "{entreprise}, ignorez-le : je ne vous écrirai plus.\n\n"
            "Si le sujet vous intéresse plus tard, mon adresse reste {email_pro} et vous pouvez "
            "réserver un appel gratuit ici : " + LIEN_RDV
        ),
    },
}

NOM_PAR_DEFAUT = "Pierre-Antton Ducoin"
MARQUE_PAR_DEFAUT = "P-A Ducoin, web design et développement"

# Noms de colonnes acceptés en plus de ceux du tableau 04 (ex. 03-prospects-64-propre.xlsx)
ALIAS_COLONNES = {"Nom": "Nom de l'entreprise", "Email": "Email professionnel"}

# Valeur de « État du site » pour laquelle le modèle A (« je n'ai pas trouvé de site ») est exact.
# Les autres valeurs sont écartées du premier contact : vérifie à la main avant (règle 2 de 01).
ETAT_SITE_OK = "Aucun site connu"

# Libellé de la source dans la phrase d'information du mail
SOURCES_AFFICHAGE = {"OpenStreetMap (Overpass)": "OpenStreetMap, base de données cartographique ouverte"}

# Statuts pour lesquels on n'écrit plus (même logique que les formules du tableau 04)
STATUTS_STOP = {"Opposition", "Pas intéressé", "Client", "Rendez-vous pris", "Numéro erroné"}

# Adresses de messagerie grand public : à vérifier à la main (règle : adresses
# professionnelles uniquement, voir 01-conformite-prospection).
DOMAINES_GRAND_PUBLIC = {
    "gmail.com", "yahoo.fr", "yahoo.com", "hotmail.com", "hotmail.fr", "outlook.com",
    "outlook.fr", "live.fr", "live.com", "orange.fr", "wanadoo.fr", "free.fr", "sfr.fr",
    "laposte.net", "icloud.com", "me.com", "msn.com", "aol.com", "gmx.fr", "gmx.com",
    "proton.me", "protonmail.com", "bbox.fr", "neuf.fr",
}

EMAIL_RE = re.compile(r"^[^@\s]+@[^@\s]+\.[^@\s]+$")


# ---------------------------------------------------------------------------
# Lecture des données
# ---------------------------------------------------------------------------
def en_date(valeur):
    """Convertit une cellule (date Excel, JJ/MM/AAAA ou AAAA-MM-JJ) en date, sinon None."""
    if valeur is None or valeur == "":
        return None
    if isinstance(valeur, dt.datetime):
        return valeur.date()
    if isinstance(valeur, dt.date):
        return valeur
    texte = str(valeur).strip()
    for fmt in ("%d/%m/%Y", "%Y-%m-%d", "%d-%m-%Y", "%d.%m.%Y"):
        try:
            return dt.datetime.strptime(texte, fmt).date()
        except ValueError:
            pass
    return None


def jours_ouvres_apres(depart, n):
    """Ajoute n jours ouvrés (lundi à vendredi, sans jours fériés, comme WORKDAY du tableau)."""
    jour = depart
    while n > 0:
        jour += dt.timedelta(days=1)
        if jour.weekday() < 5:
            n -= 1
    return jour


def normaliser(ligne):
    """Ajoute les noms de colonnes standard quand le fichier utilise des alias."""
    for alias, standard in ALIAS_COLONNES.items():
        if alias in ligne and standard not in ligne:
            ligne[standard] = ligne[alias]
    return ligne


def texte(valeur):
    return "" if valeur is None else str(valeur).strip()


def lire_xlsx(chemin):
    try:
        import openpyxl
    except ImportError:
        sys.exit("openpyxl manquant : pip install openpyxl (ou donne un fichier .csv).")
    classeur = openpyxl.load_workbook(chemin, data_only=True)
    feuille = classeur["Prospects"] if "Prospects" in classeur.sheetnames else classeur.active
    lignes = list(feuille.iter_rows(values_only=True))
    entetes = [texte(c) for c in lignes[0]]
    prospects = [normaliser(dict(zip(entetes, ligne))) for ligne in lignes[1:]]
    opposes = set()
    if "Opposition" in classeur.sheetnames:
        for ligne in list(classeur["Opposition"].iter_rows(values_only=True))[1:]:
            if len(ligne) > 1 and ligne[1]:
                opposes.add(texte(ligne[1]).lower())
    return prospects, opposes


def lire_csv(chemin):
    with open(chemin, newline="", encoding="utf-8-sig") as f:
        echantillon = f.read(2048)
        f.seek(0)
        dialecte = csv.Sniffer().sniff(echantillon, delimiters=",;\t")
        return [normaliser(dict(l)) for l in csv.DictReader(f, dialect=dialecte)], set()


# ---------------------------------------------------------------------------
# Choix du modèle pour un prospect
# ---------------------------------------------------------------------------
def choisir_modele(p, aujourdhui):
    """Renvoie (modèle, raison_si_aucun)."""
    if texte(p.get("Statut")) in STATUTS_STOP:
        return None, "statut « %s »" % texte(p.get("Statut"))
    premier = en_date(p.get("Date du 1er contact"))
    relance1 = en_date(p.get("Relance 1 envoyée le"))
    relance2 = en_date(p.get("Relance 2 envoyée le"))
    if premier is None:
        etat = texte(p.get("État du site"))
        if etat and etat != ETAT_SITE_OK:
            return None, "État du site : « %s » (le modèle A suppose qu'il n'y a pas de site)" % etat
        return "A", None
    if texte(p.get("Canal du 1er contact")) not in ("", "Email"):
        return None, "1er contact par un autre canal que l'e-mail"
    if relance1 is None:
        if jours_ouvres_apres(premier, 5) <= aujourdhui:
            return "B", None
        return None, "relance 1 pas encore due (%s)" % jours_ouvres_apres(premier, 5).strftime("%d/%m/%Y")
    if relance2 is None:
        if jours_ouvres_apres(relance1, 7) <= aujourdhui:
            return "C", None
        return None, "relance 2 pas encore due (%s)" % jours_ouvres_apres(relance1, 7).strftime("%d/%m/%Y")
    return None, "3 messages déjà envoyés : arrêt"


def verifier(p, opposes):
    """Renvoie la raison d'écarter le prospect, ou None s'il est exploitable."""
    email = texte(p.get("Email professionnel")).lower()
    if not email:
        return "pas d'e-mail professionnel"
    if not EMAIL_RE.match(email):
        return "e-mail invalide : %s" % email
    if email in opposes or texte(p.get("Nom de l'entreprise")).lower() in opposes:
        return "présent dans l'onglet Opposition"
    if email.split("@")[1] in DOMAINES_GRAND_PUBLIC:
        return "adresse grand public (à vérifier à la main) : %s" % email
    if not texte(p.get("Source de la donnée")):
        return "source de la donnée vide (obligatoire)"
    for champ in ("Nom de l'entreprise", "Activité", "Ville"):
        if not texte(p.get(champ)):
            return "champ manquant : %s" % champ
    return None


# ---------------------------------------------------------------------------
# Écriture des brouillons
# ---------------------------------------------------------------------------
def nom_fichier(texte_libre):
    return re.sub(r"[^A-Za-z0-9]+", "-", texte_libre).strip("-").lower()[:40] or "prospect"


def construire(p, modele, cfg):
    premier = en_date(p.get("Date du 1er contact"))
    prenom = texte(p.get("Prénom du dirigeant"))
    valeurs = {
        "salut": "Bonjour %s," % prenom if prenom else "Bonjour,",
        "entreprise": texte(p.get("Nom de l'entreprise")),
        "activite": texte(p.get("Activité")).lower(),
        "ville": texte(p.get("Ville")),
        "source": SOURCES_AFFICHAGE.get(texte(p.get("Source de la donnée")), texte(p.get("Source de la donnée"))),
        "date_premier_contact": premier.strftime("%d/%m/%Y") if premier else "",
        "nom": cfg.get("nom", NOM_PAR_DEFAUT),
        "marque": cfg.get("marque", MARQUE_PAR_DEFAUT),
        "email_pro": cfg["email_pro"],
        "telephone": cfg["telephone"],
        "siret": cfg["siret"],
    }
    m = MODELES[modele]
    corps = m["corps"].format(**valeurs) + "\n\n" + SIGNATURE.format(**valeurs)
    corps += "\n\n" + MENTION.format(**valeurs)
    if cfg.get("site"):
        corps = corps.replace(valeurs["marque"], valeurs["marque"] + "\n" + cfg["site"], 1)
    msg = EmailMessage()
    msg["To"] = texte(p.get("Email professionnel"))
    msg["From"] = cfg["email_pro"]
    msg["Subject"] = m["objet"].format(**valeurs)
    msg["X-Unsent"] = "1"  # s'ouvre comme brouillon dans Outlook et Thunderbird
    msg.set_content(corps)
    return msg


def main():
    ap = argparse.ArgumentParser(description="Génère des brouillons de mails de prospection (aucun envoi).")
    ap.add_argument("liste", help="tableau .xlsx (onglet Prospects) ou .csv, en local")
    ap.add_argument("--config", required=True, help="config-local.json (hors repo)")
    ap.add_argument("--sortie", default=str(Path.home() / "prospection-local" / "mails"),
                    help="dossier de sortie")
    ap.add_argument("--date", help="AAAA-MM-JJ : simule un autre jour (par défaut aujourd'hui)")
    ap.add_argument("--verif", help="verification-sites.csv produit par 04-verifier-sites.py")
    ap.add_argument("--sans-verification", action="store_true",
                    help="autorise le modèle A sans vérification du site (déconseillé)")
    ap.add_argument("--groupe", help="ne traiter que ce groupe (colonne Groupe), ex. Restauration")
    ap.add_argument("--limite", type=int, help="nombre maximum de mails à générer")
    a = ap.parse_args()

    cfg = json.loads(Path(a.config).read_text(encoding="utf-8"))
    for cle in ("email_pro", "telephone", "siret"):
        if not cfg.get(cle):
            sys.exit("config-local.json : champ « %s » manquant." % cle)

    aujourdhui = en_date(a.date) if a.date else dt.date.today()
    if aujourdhui is None:
        sys.exit("--date doit être au format AAAA-MM-JJ.")

    sortie = Path(a.sortie).expanduser() / aujourdhui.isoformat()
    sortie.mkdir(parents=True, exist_ok=True)

    chemin = Path(a.liste).expanduser()
    prospects, opposes = lire_xlsx(chemin) if chemin.suffix.lower() == ".xlsx" else lire_csv(chemin)

    verdicts = {}
    if a.verif:
        with open(Path(a.verif).expanduser(), newline="", encoding="utf-8") as f:
            verdicts = {int(r["ligne"]): r for r in csv.DictReader(f)}

    generes, ecartes = [], []
    for i, p in enumerate(prospects, start=2):  # ligne 2 = 1er prospect, comme dans Excel
        if not texte(p.get("Nom de l'entreprise")):
            continue
        nom = texte(p.get("Nom de l'entreprise"))
        if a.groupe and texte(p.get("Groupe")).lower() != a.groupe.lower():
            continue
        modele, raison = choisir_modele(p, aujourdhui)
        if modele is None:
            ecartes.append((i, nom, raison))
            continue
        if modele == "A" and not a.sans_verification:
            v = verdicts.get(i)
            if v is None:
                ecartes.append((i, nom, "site non vérifié : lance 04-verifier-sites.py puis --verif"))
                continue
            if v["verdict"] != "Aucun site apparent":
                ecartes.append((i, nom, "vérification du site : %s %s" % (v["verdict"], v.get("url", ""))))
                continue
        raison = verifier(p, opposes)
        if raison:
            ecartes.append((i, nom, raison))
            continue
        if a.limite is not None and len(generes) >= a.limite:
            ecartes.append((i, nom, "limite atteinte (--limite)"))
            continue
        msg = construire(p, modele, cfg)
        fichier = sortie / ("%s-ligne%03d-%s.eml" % (modele, i, nom_fichier(nom)))
        fichier.write_bytes(bytes(msg))
        generes.append((i, nom, modele, msg["To"], msg["Subject"], fichier.name))

    rapport = ["Rapport du %s" % aujourdhui.strftime("%d/%m/%Y"),
               "Relis chaque brouillon. Puis note les dates d'envoi dans le tableau de suivi.", "",
               "%d brouillon(s) généré(s) :" % len(generes)]
    for i, nom, modele, a_qui, objet, f in generes:
        rapport.append("  ligne %d | modèle %s | %s | %s | %s" % (i, modele, nom, a_qui, f))
    rapport += ["", "%d ligne(s) écartée(s) :" % len(ecartes)]
    for i, nom, raison in ecartes:
        rapport.append("  ligne %d | %s | %s" % (i, nom, raison))
    (sortie / "rapport.txt").write_text("\n".join(rapport) + "\n", encoding="utf-8")

    print("\n".join(rapport[:4]))
    print("Dossier : %s" % sortie)


if __name__ == "__main__":
    main()
