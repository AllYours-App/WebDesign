#!/usr/bin/env python3
"""
16-verifier-sites.py : vérifie qu'une entreprise n'a PAS de site avant de lui écrire.

Pourquoi : la colonne « État du site » du fichier prospects vient d'OpenStreetMap, qui
rate beaucoup de sites existants. Ce script refait la vérification que tu fais à la main
(recherche du numéro, du nom, de l'e-mail) et écrit un verdict par entreprise.

Il tourne EN LOCAL, sur ta machine. Rien n'est écrit dans le repo : le script refuse un
dossier de sortie situé dans un dépôt Git.

Deux contrôles, dans cet ordre :
  1. Domaine de l'e-mail (gratuit, sans clé) : si l'adresse est en @entreprise.fr et que
     entreprise.fr répond comme un site, l'entreprise a un site.
  2. Recherche web (clé Brave Search, voir plus bas) : une recherche sur le téléphone,
     puis sur le nom et la ville si besoin. Un résultat hors annuaires et réseaux sociaux
     dont le titre ou l'extrait contient le nom ou le numéro = site probable.

Verdicts écrits dans verification-sites.csv :
  - « Aucun site apparent » : aucun des deux contrôles n'a trouvé de site. Seul verdict
    que 11-generer-mails.py accepte pour le premier contact.
  - « Site trouvé »         : url dans la colonne suivante. On n'écrit pas à cette entreprise.
  - « À vérifier »          : recherche en erreur ou sans résultat. Fais-le à la main.
Même après « Aucun site apparent », garde un coup d'œil humain sur un échantillon.

Mise en place :
  1. Crée une clé sur https://brave.com/search/api (le plan actuel offre 5 $ de crédit par
     mois, environ 1 000 requêtes, carte bancaire demandée pour la vérification : vérifie
     les conditions sur le site au moment de l'inscription).
  2. Ajoute dans ton config-local.json (hors repo) : "brave_api_key": "..."
  3. Lance :
       python3 16-verifier-sites.py ~/prospection-local/prospects-64-propre.xlsx \
           --config ~/prospection-local/config-local.json
     Options : --groupe NOM, --max-requetes N (défaut 900), --inclure-sans-email,
               --sans-recherche (seulement le contrôle du domaine, sans clé),
               --sortie DOSSIER.
  4. Relance 11-generer-mails.py avec  --verif <dossier de sortie>/verification-sites.csv

Les réponses de la recherche sont gardées dans cache-recherches.json (même dossier) :
relancer le script ne dépense pas de requêtes déjà faites.

Garde 11-generer-mails.py dans le même dossier que ce script (il en réutilise les lecteurs).

Dépendance : openpyxl pour un .xlsx (pip install openpyxl). Le reste est la bibliothèque
standard de Python.
"""

import argparse
import csv
import json
import re
import socket
import sys
import time
import unicodedata
import urllib.error
import urllib.parse
import urllib.request
from pathlib import Path

# Importe les lecteurs de 11-generer-mails.py (même dossier)
import importlib.util

_spec = importlib.util.spec_from_file_location("gen", Path(__file__).with_name("11-generer-mails.py"))
gen = importlib.util.module_from_spec(_spec)
_spec.loader.exec_module(gen)

VERDICT_OK = "Aucun site apparent"
VERDICT_SITE = "Site trouvé"
VERDICT_VERIF = "À vérifier"

# Annuaires, réseaux sociaux, agrégateurs : un résultat sur ces domaines n'est PAS un site de l'entreprise.
DOMAINES_IGNORES = {
    "facebook.com", "instagram.com", "linkedin.com", "twitter.com", "x.com", "tiktok.com", "youtube.com",
    "pinterest.com", "pinterest.fr", "google.com", "google.fr", "goo.gl", "maps.apple.com", "waze.com",
    "pagesjaunes.fr", "mappy.com", "118712.fr", "118218.fr", "cylex.fr", "kompass.com", "europages.fr",
    "societe.com", "pappers.fr", "infogreffe.fr", "verif.com", "manageo.fr", "societeinfo.com", "score3.fr",
    "annuaire-entreprises.data.gouv.fr", "data.gouv.fr", "entreprises.lefigaro.fr", "bodacc.fr", "pagesjaunes.com",
    "tripadvisor.fr", "tripadvisor.com", "thefork.fr", "thefork.com", "lafourchette.com", "yelp.fr", "yelp.com",
    "booking.com", "airbnb.fr", "airbnb.com", "expedia.fr", "hotels.com", "trivago.fr", "kayak.fr", "agoda.com",
    "abritel.fr", "gites-de-france.com", "petitfute.com", "restaurantguru.com", "foursquare.com", "mapquest.com",
    "openstreetmap.org", "wikipedia.org", "wikidata.org", "doctolib.fr", "leboncoin.fr", "ubereats.com",
    "deliveroo.fr", "annuairefrancais.fr", "journaldunet.com", "fr-fr.facebook.com", "mon-artisan.fr",
    "starofservice.com", "travaux.com", "batiweb.com", "ordre.pharmacien.fr", "ameli.fr", "santé.fr", "sante.fr",
    "pagesblanches.fr", "pages-blanches.fr", "justacote.com", "hoodspot.fr", "ouvert.fr", "horaires.lefigaro.fr",
    "linternaute.com", "tripadvisor.co.uk", "net-entreprises.fr", "dirigeant.societe.com", "fr.mappy.com",
}
MOTS_BANALS = {
    "sarl", "sas", "sasu", "eurl", "snc", "sci", "etablissements", "ets", "entreprise", "societe", "restaurant",
    "bar", "cafe", "hotel", "boulangerie", "coiffure", "coiffeur", "pharmacie", "garage", "agence", "immobiliere",
    "le", "la", "les", "du", "des", "de", "et", "chez", "au", "aux", "l", "d", "pays", "basque",
}
EN_TETES = {"User-Agent": "Mozilla/5.0 (compatible; verification-sites/1.0)"}


def sans_accents(s):
    return "".join(c for c in unicodedata.normalize("NFKD", s) if not unicodedata.combining(c)).lower()


def jetons_nom(nom):
    """Mots significatifs du nom (4 lettres et plus, hors mots banals)."""
    mots = re.findall(r"[a-z0-9]+", sans_accents(nom))
    return [m for m in mots if len(m) >= 4 and m not in MOTS_BANALS]


def chiffres(tel):
    return re.sub(r"\D", "", tel or "")


def premier_telephone(tel):
    morceau = re.split(r"[/;,]", tel or "")[0].strip()
    d = chiffres(morceau)
    if d.startswith("33") and len(d) == 11:
        d = "0" + d[2:]
    return d if len(d) == 10 else ""


def tel_formate(d):
    return " ".join(d[i:i + 2] for i in range(0, 10, 2))


def domaine_de(url):
    h = urllib.parse.urlparse(url).netloc.lower().split(":")[0]
    return h[4:] if h.startswith("www.") else h


def est_ignore(dom):
    return any(dom == d or dom.endswith("." + d) for d in DOMAINES_IGNORES)


# --------------------------------------------------------------------------
# Contrôle 1 : le domaine de l'e-mail héberge-t-il un site ?
# --------------------------------------------------------------------------
def domaine_repond(domaine):
    try:
        socket.gethostbyname(domaine)
    except OSError:
        return None
    for schema in ("https", "http"):
        try:
            req = urllib.request.Request("%s://%s/" % (schema, domaine), headers=EN_TETES)
            with urllib.request.urlopen(req, timeout=7) as r:
                if r.status < 400:
                    return r.geturl()
        except (urllib.error.URLError, OSError, ValueError):
            continue
    return None


# --------------------------------------------------------------------------
# Contrôle 2 : recherche web (Brave Search API)
# --------------------------------------------------------------------------
class Recherche:
    def __init__(self, cle, dossier, max_requetes, pause=1.1):
        self.cle, self.max, self.pause = cle, max_requetes, pause
        self.nb = 0
        self.fichier = dossier / "cache-recherches.json"
        self.cache = json.loads(self.fichier.read_text(encoding="utf-8")) if self.fichier.exists() else {}

    def requete(self, q):
        if q in self.cache:
            return self.cache[q]
        if self.nb >= self.max:
            raise RuntimeError("limite de requêtes atteinte (--max-requetes)")
        url = "https://api.search.brave.com/res/v1/web/search?" + urllib.parse.urlencode(
            {"q": q, "country": "fr", "search_lang": "fr", "count": 8})
        req = urllib.request.Request(url, headers={"Accept": "application/json", "X-Subscription-Token": self.cle})
        with urllib.request.urlopen(req, timeout=20) as r:
            data = json.loads(r.read().decode("utf-8"))
        res = [{"url": x.get("url", ""), "titre": x.get("title", ""), "extrait": x.get("description", "")}
               for x in data.get("web", {}).get("results", [])]
        self.nb += 1
        self.cache[q] = res
        self.fichier.write_text(json.dumps(self.cache, ensure_ascii=False), encoding="utf-8")
        time.sleep(self.pause)
        return res


def site_dans_resultats(resultats, nom, tel_chiffres, email_domaine=""):
    """Renvoie l'url d'un site probable de l'entreprise parmi les résultats, sinon None."""
    jetons = jetons_nom(nom)
    for r in resultats:
        dom = domaine_de(r["url"])
        if not dom or est_ignore(dom):
            continue
        texte = sans_accents(" ".join([r["titre"], r["extrait"], r["url"]]))
        if tel_chiffres and tel_chiffres in re.sub(r"\D", "", texte):
            return r["url"]
        if email_domaine and dom.endswith(email_domaine):
            return r["url"]
        if jetons:
            trouves = sum(1 for j in jetons if j in texte)
            if trouves >= max(1, (len(jetons) + 1) // 2):
                return r["url"]
    return None


def verifier(p, rech):
    nom = gen.texte(p.get("Nom de l'entreprise"))
    ville = gen.texte(p.get("Ville"))
    email = gen.texte(p.get("Email professionnel")).lower()
    tel = premier_telephone(gen.texte(p.get("Téléphone")))
    dom_mail = email.split("@")[1] if "@" in email else ""

    # 1. domaine de l'e-mail
    if dom_mail and dom_mail not in gen.DOMAINES_GRAND_PUBLIC:
        u = domaine_repond(dom_mail)
        if u:
            return VERDICT_SITE, u, "domaine de l'e-mail"

    if rech is None:
        return VERDICT_VERIF, "", "recherche non faite"

    # 2. recherche web : téléphone d'abord, puis nom + ville
    requetes = []
    if tel:
        requetes.append('"%s"' % tel_formate(tel))
    requetes.append("%s %s" % (nom, ville))
    erreurs, vus = 0, 0
    for q in requetes:
        try:
            res = rech.requete(q)
        except RuntimeError:
            raise
        except (urllib.error.URLError, OSError, ValueError, json.JSONDecodeError):
            erreurs += 1
            continue
        vus += len(res)
        u = site_dans_resultats(res, nom, tel, dom_mail if dom_mail not in gen.DOMAINES_GRAND_PUBLIC else "")
        if u:
            return VERDICT_SITE, u, "recherche : " + ("téléphone" if q.startswith('"') else "nom et ville")
    if erreurs == len(requetes) or vus == 0:
        return VERDICT_VERIF, "", "recherche en erreur ou sans résultat"
    return VERDICT_OK, "", "domaine et recherche"


def main():
    ap = argparse.ArgumentParser(description="Vérifie l'absence de site avant le premier contact.")
    ap.add_argument("liste", help="tableau .xlsx (onglet Prospects) ou .csv, en local")
    ap.add_argument("--config", required=True, help="config-local.json (hors repo), avec brave_api_key")
    ap.add_argument("--sortie", default=str(Path.home() / "prospection-local" / "verification"))
    ap.add_argument("--groupe")
    ap.add_argument("--max-requetes", type=int, default=900)
    ap.add_argument("--inclure-sans-email", action="store_true", help="contrôle aussi les lignes sans e-mail (liste d'appel)")
    ap.add_argument("--sans-recherche", action="store_true", help="contrôle du domaine seulement, sans clé")
    a = ap.parse_args()

    cfg = json.loads(Path(a.config).read_text(encoding="utf-8"))
    sortie = Path(a.sortie).expanduser()
    if gen.dans_un_depot_git(sortie):
        sys.exit("Refusé : %s est dans un dépôt Git. Choisis un dossier hors repo." % sortie)
    sortie.mkdir(parents=True, exist_ok=True)

    chemin = Path(a.liste).expanduser()
    prospects, _ = gen.lire_xlsx(chemin) if chemin.suffix.lower() == ".xlsx" else gen.lire_csv(chemin)

    rech = None
    if not a.sans_recherche:
        if not cfg.get("brave_api_key"):
            sys.exit("config-local.json : « brave_api_key » manquante (ou utilise --sans-recherche).")
        rech = Recherche(cfg["brave_api_key"], sortie, a.max_requetes)

    candidats = []
    for i, p in enumerate(prospects, start=2):
        if not gen.texte(p.get("Nom de l'entreprise")):
            continue
        if a.groupe and gen.texte(p.get("Groupe")).lower() != a.groupe.lower():
            continue
        etat = gen.texte(p.get("État du site"))
        if etat and etat != gen.ETAT_SITE_OK:
            continue  # site déjà connu : inutile de vérifier
        if not a.inclure_sans_email and gen.verifier(p, set()) is not None:
            continue  # même règles que le générateur : e-mail pro, source, ville
        candidats.append((i, p))

    print("%d entreprise(s) à vérifier." % len(candidats))
    chemin_csv = sortie / "verification-sites.csv"
    deja = {}
    if chemin_csv.exists():
        with open(chemin_csv, newline="", encoding="utf-8") as f:
            deja = {int(r["ligne"]): r for r in csv.DictReader(f)}
    arret = False
    for n, (i, p) in enumerate(candidats, 1):
        if i in deja and deja[i]["verdict"] != VERDICT_VERIF:
            continue
        try:
            verdict, url, methode = verifier(p, rech)
        except RuntimeError as e:
            print("Arrêt : %s" % e)
            arret = True
            break
        deja[i] = {"ligne": i, "nom": gen.texte(p.get("Nom de l'entreprise")), "verdict": verdict, "url": url, "methode": methode}
        if n % 20 == 0:
            print("  %d/%d" % (n, len(candidats)))
            with open(chemin_csv, "w", newline="", encoding="utf-8") as f:
                w = csv.DictWriter(f, fieldnames=["ligne", "nom", "verdict", "url", "methode"])
                w.writeheader()
                w.writerows(deja[k] for k in sorted(deja))
    with open(chemin_csv, "w", newline="", encoding="utf-8") as f:
        w = csv.DictWriter(f, fieldnames=["ligne", "nom", "verdict", "url", "methode"])
        w.writeheader()
        w.writerows(deja[k] for k in sorted(deja))
    comptes = {}
    for r in deja.values():
        comptes[r["verdict"]] = comptes.get(r["verdict"], 0) + 1
    print("Résultat :", comptes, "| requêtes de recherche :", rech.nb if rech else 0)
    print("Fichier :", chemin_csv)
    if arret:
        print("Relance plus tard : les lignes déjà traitées sont conservées.")


if __name__ == "__main__":
    main()
