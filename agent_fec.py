#!/usr/bin/env python3
"""
Agent factures -> FEC (importable dans ACD Compta Expert).

Usage :
  python agent_fec.py export.csv --siren 123456789 --exercice 2026
  python agent_fec.py export.xlsx facture1.pdf scan.jpg --siren 123456789 --type achat

Étapes : lecture -> mapping auto des colonnes -> format pivot -> contrôles -> FEC + rapport.
Le mapping est mémorisé dans mappings/<empreinte>.json (modifiable à la main).
PDF/images : Factur-X (XML intégré) -> texte du PDF (règles) -> OCR PaddleOCR (facultatif)
-> IA Claude (facultatif, payant, si ANTHROPIC_API_KEY est défini).
"""
import argparse, hashlib, io, json, os, re, shutil, sys, tempfile, unicodedata, zipfile
VERSION = "2026.11.06-l"
import xml.etree.ElementTree as ET
from collections import OrderedDict
from decimal import Decimal, ROUND_HALF_UP
from datetime import datetime
import pandas as pd

# ---------- Paramétrage comptable (à adapter au dossier) ----------
COMPTES = {
    "vente": {"journal": ("VE", "Ventes"), "tiers": "411000", "charge_produit": "706000",
              "libelle_cp": "Prestations de services", "sens_tiers": "D"},
    "achat": {"journal": ("AC", "Achats"), "tiers": "401000", "charge_produit": "607000",
              "libelle_cp": "Achats de marchandises", "sens_tiers": "C"},
}
TVA_COMPTES = {  # taux -> (TVA collectée 4457xxx, TVA déductible 4456xxx) : 7 caractères, xxx = taux x 10
    "20": ("4457200", "4456200"), "10": ("4457100", "4456100"),
    "5.5": ("4457055", "4456055"), "2.1": ("4457021", "4456021"),
}
TVA_INCONNUE = ("4457000", "4456000")
NATURES = {"706": "Prestations de services", "707": "Ventes de marchandises", "701": "Ventes de produits finis"}

def produits_par_taux(prefixe="706", taux=("20", "10", "5.5", "2.1", "0")):
    """Comptes de produits par taux de TVA, sur 7 caractères : préfixe (706/707/701) + taux x 100 sur 4 chiffres.
    Ex. 706 : 20 % -> 7062000 ; 10 % -> 7061000 ; 5,5 % -> 7060550 ; 2,1 % -> 7060210 ; 0 % -> 7060000."""
    out = {}
    for t in taux:
        v = float(t)
        out[str(t)] = (f"{prefixe}{int(round(v * 100)):04d}", f"{NATURES.get(prefixe, 'Produits')} {('%g' % v).replace('.', ',')} %")
    return out

PRODUITS_TVA = produits_par_taux("706")
TVA_LIB = "TVA"

# ---------- Mapping des colonnes ----------
SYNONYMES = {
    "date": ["date facture", "date de facture", "date de la facture", "date d emission", "date emission", "date piece",
             "date de piece", "date doc", "date document", "date comptable", "date de facturation", "date facturation",
             "invoice date", "issue date", "date of issue", "date creation", "date de creation", "created", "date"],
    "numero": ["numero facture", "numero de facture", "n facture", "n de facture", "no facture", "num facture", "n piece",
               "numero piece", "numero de piece", "n de piece", "reference facture", "ref facture", "invoice number",
               "invoice no", "invoice", "document", "numero", "reference", "ref", "num", "facture", "piece",
               "n de document", "numero document", "n document", "n avoir", "numero avoir", "id facture", "facture n",
               "facture no", "n", "no", "nr", "doc", "n doc", "num doc", "num piece", "id"],
    "tiers": ["client", "nom client", "nom du client", "raison sociale", "fournisseur", "nom fournisseur",
              "nom du fournisseur", "tiers", "nom tiers", "societe", "entreprise", "denomination", "customer",
              "customer name", "supplier", "vendor", "contact", "libelle tiers", "intitule tiers",
              "raison sociale client", "raison sociale fournisseur", "client raison sociale", "nom commercial",
              "enseigne", "acheteur", "destinataire", "emetteur", "beneficiaire", "intitule client", "libelle client",
              "intitule fournisseur", "nom", "company", "nom societe", "nom entreprise", "buyer", "seller"],
    "code_tiers": ["code client", "code tiers", "compte client", "compte tiers", "code fournisseur", "compte fournisseur",
                   "n client", "numero client", "compte auxiliaire", "auxiliaire", "customer id", "customer code",
                   "id client", "ref client"],
    "ht": ["total ht", "montant ht", "ht", "hors taxe", "hors taxes", "net ht", "total hors taxe", "total hors taxes",
           "montant hors taxe", "montant hors taxes", "base ht", "base", "base hors taxe", "sous total", "subtotal", "amount excl", "total excl tax",
           "net amount", "prix ht", "total net", "montant net ht", "untaxed amount"],
    "tva": ["total tva", "montant tva", "tva", "taxe", "taxes", "vat", "tax amount", "total taxes", "montant de la tva",
            "total vat", "montant taxe"],
    "ttc": ["total ttc", "montant ttc", "ttc", "total", "net a payer", "amount incl", "total amount", "montant total",
            "total toutes taxes comprises", "gross", "total due", "total a payer", "total general", "grand total",
            "amount due", "montant a payer"],
    "taux_tva": ["taux tva", "taux de tva", "taux", "% tva", "tva %", "vat rate", "tax rate", "tva pct", "code tva", "taux taxe", "tx", "tx tva", "tva taux"],
    "libelle": ["libelle", "designation", "description", "objet", "intitule", "produit", "article"],
    "compte": ["compte produit", "compte comptable", "compte", "compte de produit", "compte de charge", "code comptable", "account"],
}
OBLIGATOIRES = ["date", "numero", "tiers"]

def norm(s):
    s = unicodedata.normalize("NFKD", str(s)).encode("ascii", "ignore").decode().lower()
    return re.sub(r"[^a-z0-9% ]+", " ", s).strip()

# Mots qui, présents dans l'intitulé, interdisent un champ (ex. « N° client » n'est pas un n° de facture)
EXCLUS = {
    "date": ["echeance", "reglement", "paiement", "livraison", "naissance", "limite", "modification", "maj", "export",
             "envoi", "relance", "encaissement", "valeur", "debut", "fin", "periode", "expedition", "due"],
    "numero": ["client", "fournisseur", "tiers", "customer", "supplier", "siret", "siren", "intracom", "intracommunautaire",
               "telephone", "tel", "commande", "bon", "bl", "livraison", "devis", "ligne", "article", "produit", "compte",
               "lot", "serie", "cheque", "iban", "rib", "postal", "rue", "voie", "date", "montant", "total", "ht", "ttc",
               "tva", "type", "statut", "etat", "nombre", "quantite", "qte", "page", "order", "po", "contrat", "dossier",
               "code", "mode", "taux", "nom", "libelle", "designation", "affaire", "projet", "chantier", "immatriculation"],
    "tiers": ["code", "n", "no", "numero", "num", "id", "ref", "reference", "compte", "adresse", "rue", "ville", "cp",
              "postal", "pays", "email", "mail", "tel", "telephone", "siret", "siren", "tva", "civilite", "prenom",
              "date", "montant", "total", "iban", "type", "categorie", "groupe", "famille", "commercial", "vendeur"],
    "code_tiers": ["date", "montant", "total", "siret", "siren", "tva", "nom", "raison", "adresse"],
    "ht": ["unitaire", "pu", "remise", "escompte", "acompte", "deja", "reste", "quantite", "qte", "taux", "ttc", "tva"],
    "tva": ["intracom", "intracommunautaire", "numero", "n", "taux", "%", "regime", "exoneration", "mention", "code", "ht",
            "ttc", "pct"],
    "ttc": ["ht", "tva", "regle", "paye", "payee", "reste", "solde", "acompte", "deja", "encaisse", "quantite", "qte",
            "unitaire", "remise", "nombre", "taux", "lignes"],
    "taux_tva": ["remise", "escompte", "commission", "marge", "change", "horaire", "montant"],
    "libelle": ["tiers", "client", "fournisseur", "compte", "tva", "journal"],
    "compte": ["client", "fournisseur", "tiers", "auxiliaire", "bancaire", "iban", "libelle", "intitule"],
}
CLIENT_MOTS, FOURN_MOTS = r"\b(client|customer|acheteur|destinataire|buyer)s?\b", r"\b(fournisseur|supplier|vendor|emetteur|seller)s?\b"

def _profil(serie):
    """Statistiques simples sur le contenu d'une colonne (échantillon des 200 premières valeurs non vides)."""
    v = [str(x).strip() for x in serie.dropna().head(400) if str(x).strip() not in ("", "nan", "None")][:200]
    n = len(v) or 1
    lettres = lambda x: len(re.findall(r"[A-Za-zÀ-ÿ]", x))
    chiffres = lambda x: len(re.findall(r"\d", x))
    return dict(n=len(v),
                date=sum(date(x) is not None for x in v) / n,
                montant=sum(montant(x) is not None and not re.search(r"[A-Za-zÀ-ÿ]{2,}", x) for x in v) / n,
                decimal=sum(bool(re.fullmatch(r"-?\d+[.,]\d{1,2}", x.replace(" ", ""))) for x in v) / n,
                texte=sum(lettres(x) >= 2 and lettres(x) > chiffres(x) for x in v) / n,
                chiffre=sum(chiffres(x) >= 1 for x in v) / n,
                taux=sum((lambda m: m is not None and 0 <= m <= 30)(montant(x)) for x in v) / n,
                distinct=len(set(v)), mail=sum("@" in x for x in v) / n,
                code=sum(bool(re.fullmatch(r"[A-Za-z]{0,6}[-_/.]?\d[\w\-/.]*", x)) for x in v) / n)

def _score_entete(champ, n, type_=None):
    """Score de l'intitulé normalisé n pour un champ (None = impossible)."""
    mots = set(n.split())
    best = None
    for k, syn in enumerate(SYNONYMES[champ]):
        if n == syn:
            sc = 100 + 3 * len(syn.split()) - 0.2 * k
        elif re.search(rf"\b{re.escape(syn)}\b", n) and len(syn) > 2:
            sc = 55 + 3 * len(syn.split()) - 2 * (len(mots) - len(syn.split())) - 0.2 * k
        else:
            continue
        best = sc if best is None else max(best, sc)
    if best is None: return None
    if best < 100 and any(x in mots or (" " in x and x in n) for x in EXCLUS.get(champ, [])):
        return None                                             # intitulé partiel avec un mot exclu
    if champ in ("tiers", "code_tiers") and type_ in ("vente", "achat"):
        bon, mauvais = (CLIENT_MOTS, FOURN_MOTS) if type_ == "vente" else (FOURN_MOTS, CLIENT_MOTS)
        if re.search(mauvais, n): best -= 60                    # colonne de l'autre partie (ex. fournisseur en vente)
        elif re.search(bon, n): best += 15
    if champ == "tiers" and n in ("raison sociale", "societe", "entreprise", "denomination", "nom", "contact", "company"):
        best -= 8                                               # intitulé générique : passe après « nom client »
    return best

def _score_contenu(champ, p):
    """Bonus/malus selon le contenu ; None = contenu incompatible avec le champ."""
    if not p or p["n"] == 0: return 0
    if champ == "date":
        return None if p["date"] < 0.6 else 15 * p["date"]
    if champ in ("ht", "tva", "ttc"):
        return None if p["montant"] < 0.6 or p["date"] > 0.5 else 10 * p["montant"]
    if champ == "taux_tva":
        return None if p["taux"] < 0.6 or p["date"] > 0.5 else 10
    if champ == "numero":
        if p["chiffre"] < 0.5 or p["date"] > 0.6 or p["mail"] > 0.2: return None
        sc = 10 - 25 * p["decimal"]
        if p["n"] >= 3 and p["distinct"] == 1: sc -= 40         # même valeur partout : pas un n° de facture
        return sc
    if champ == "tiers":
        if p["texte"] < 0.5 or p["date"] > 0.5 or p["mail"] > 0.3 or p["code"] > 0.6: return None
        sc = 10 * p["texte"]
        if p["n"] >= 3 and p["distinct"] == 1: sc -= 45         # même nom partout : sans doute votre propre société
        return sc
    return 0

def mapping_regles(colonnes, df=None, type_=None):
    """Associe chaque champ à la meilleure colonne : intitulé (synonymes, mots exclus, client/fournisseur)
    + contrôle du contenu (dates, montants, textes) si le tableau est fourni. Affectation globale par score."""
    cols = [c for c in colonnes]
    profils = {c: _profil(df[c]) for c in cols} if df is not None else {}
    cand = []
    for champ in SYNONYMES:
        for c in cols:
            sc = _score_entete(champ, norm(c), type_)
            if sc is None: continue
            if df is not None:
                cc = _score_contenu(champ, profils.get(c))
                if cc is None: continue
                sc += cc
            cand.append((sc, champ, c))
    res, pris = {}, set()
    for sc, champ, c in sorted(cand, key=lambda x: -x[0]):
        if champ not in res and c not in pris and sc > 0:
            res[champ] = c; pris.add(c)
    if df is not None:
        # tiers de secours : 2e meilleure colonne de nom (ex. « Nom » quand « Raison sociale » est vide pour un particulier)
        sec = [(sc, c) for sc, ch, c in cand if ch == "tiers" and c not in pris and sc > 40
               and not (profils[c]["n"] >= 3 and profils[c]["distinct"] == 1)]
        if "tiers" in res and sec:
            res["tiers2"] = max(sec)[1]; pris.add(res["tiers2"])
        # champs obligatoires sans intitulé reconnu : on devine par le contenu
        libres = [c for c in cols if c not in pris]
        if "date" not in res:
            ok = [(profils[c]["date"], c) for c in libres if profils[c]["date"] >= 0.8 and not EXCLUS_ENT(norm(c), "date")]
            if ok: res["date"] = max(ok)[1]; libres.remove(res["date"])
        if "numero" not in res:
            ok = [(profils[c]["code"] + profils[c]["distinct"] / max(profils[c]["n"], 1), c) for c in libres
                  if profils[c]["code"] >= 0.8 and profils[c]["decimal"] < 0.2 and profils[c]["date"] < 0.2
                  and not EXCLUS_ENT(norm(c), "numero")]
            if ok: res["numero"] = max(ok)[1]; libres.remove(res["numero"])
        if "tiers" not in res:
            ok = [(profils[c]["texte"] + profils[c]["distinct"] / 1000, c) for c in libres
                  if profils[c]["texte"] >= 0.8 and profils[c]["distinct"] > 1 and profils[c]["mail"] < 0.2
                  and profils[c]["code"] < 0.3
                  and not EXCLUS_ENT(norm(c), "tiers")]
            if ok: res["tiers"] = max(ok)[1]; libres.remove(res["tiers"])
        if "tiers" not in res and res.get("libelle") and norm(res["libelle"]) in ("intitule", "libelle", "nom", "intitule compte"):
            pl = profils[res["libelle"]]                        # « Intitulé » seul : c'est le nom du tiers, pas un libellé
            if pl["texte"] >= 0.8 and pl["code"] < 0.3 and pl["distinct"] > 1:
                res["tiers"] = res.pop("libelle")
        # contrôle HT / TVA / TTC : si HT > TTC sur la plupart des lignes, les deux colonnes sont inversées
        if "ht" in res and "ttc" in res:
            a = [(montant(x), montant(y)) for x, y in zip(df[res["ht"]].head(100), df[res["ttc"]].head(100))]
            a = [(x, y) for x, y in a if x is not None and y is not None and x > 0]
            if a and sum(x > y for x, y in a) > 0.7 * len(a):
                res["ht"], res["ttc"] = res["ttc"], res["ht"]
    return res

def EXCLUS_ENT(n, champ):
    return any(x in set(n.split()) for x in EXCLUS.get(champ, []))

def cle_api():
    """Clé API valide ou chaîne vide (ignore les valeurs factices comme sk-ant-...)."""
    cle = os.getenv("ANTHROPIC_API_KEY", "").strip().strip("\"'“”‘’ ")
    return cle if len(cle) >= 40 and "..." not in cle else ""

def mapping_ia(df, deja):
    """Demande à Claude de deviner les champs manquants (optionnel)."""
    manquants = [c for c in SYNONYMES if c not in deja]
    if not manquants or not cle_api():
        return {}
    try:
        import anthropic
        extrait = df.head(5).to_csv(index=False)
        rep = anthropic.Anthropic(api_key=cle_api()).messages.create(
            model="claude-sonnet-4-6", max_tokens=500,
            messages=[{"role": "user", "content":
                f"Export d'un logiciel de facturation :\n{extrait}\n"
                f"Pour chaque champ parmi {manquants}, donne le nom exact de la colonne "
                f"correspondante ou null. Réponds uniquement en JSON {{champ: colonne}}."}])
        txt = re.sub(r"```json|```", "", rep.content[0].text).strip()
        data = json.loads(txt)
        if not isinstance(data, dict): return {}
        return {k: v for k, v in data.items() if v in df.columns}
    except Exception as e:
        print(f"[IA indisponible : {e}]", file=sys.stderr); return {}

CHAMPS_LIBELLES = {"date": "Date de facture", "numero": "N° de facture", "tiers": "Client / Fournisseur",
                   "code_tiers": "Code tiers (facultatif)", "ht": "Montant HT", "tva": "Montant TVA",
                   "ttc": "Montant TTC", "taux_tva": "Taux de TVA", "libelle": "Libellé (facultatif)",
                   "compte": "Compte comptable (facultatif)", "tiers2": "Nom du tiers de secours (si le 1er est vide)"}
MAPPING_V = 2   # version de la détection automatique : les anciens mappings automatiques sont recalculés

def empreinte(df):
    return hashlib.md5("|".join(map(str, df.columns)).encode()).hexdigest()[:10]

def sauver_mapping(df, m, dossier="mappings"):
    os.makedirs(dossier, exist_ok=True)
    m = {**{k: v for k, v in m.items() if not k.startswith("_")}, "_manuel": True}
    json.dump(m, open(os.path.join(dossier, f"{empreinte(df)}.json"), "w", encoding="utf-8"), indent=2, ensure_ascii=False)

def exporter_reglages(comptes, tva_comptes, produits_tva=None, tva_nd=None, vehicule=None, dossier="mappings"):
    """Tout ce qui se règle dans l'app, dans un dictionnaire sauvegardable (JSON)."""
    maps = {}
    if os.path.isdir(dossier):
        for f in os.listdir(dossier):
            if f.endswith(".json"): maps[f[:-5]] = json.load(open(os.path.join(dossier, f), encoding="utf-8"))
    return {"comptes": comptes, "tva_comptes": tva_comptes, "produits_tva": produits_tva, "tva_nd": tva_nd, "vehicule": vehicule, "mappings": maps}

def importer_reglages(d, dossier="mappings"):
    os.makedirs(dossier, exist_ok=True)
    for emp, m in (d.get("mappings") or {}).items():
        json.dump(m, open(os.path.join(dossier, f"{emp}.json"), "w", encoding="utf-8"), indent=2, ensure_ascii=False)
    return d.get("comptes"), d.get("tva_comptes"), d.get("produits_tva"), d.get("tva_nd"), d.get("vehicule")

def charger_mapping(df, dossier="mappings", type_=None):
    """Mapping mémorisé à la main (prioritaire) sinon détection automatique (recalculée si la détection a évolué)."""
    if all(c in df.columns for c in COLONNES_STD):              # tableau produit par la lecture PDF / XML : colonnes connues
        return dict(zip(["date", "numero", "tiers", "ht", "taux_tva", "tva", "ttc"], COLONNES_STD))
    chemin = os.path.join(dossier, f"{empreinte(df)}.json")
    if os.path.exists(chemin):
        try: m = json.load(open(chemin, encoding="utf-8"))
        except Exception: m = {}
        if m.get("_manuel") or (m.get("_v") == MAPPING_V and m.get("_type") == type_):
            print(f"Mapping mémorisé : {chemin}")
            return {k: v for k, v in m.items() if not k.startswith("_") and v in df.columns}
    m = mapping_regles(df.columns, df, type_)
    m.update(mapping_ia(df, m))
    os.makedirs(dossier, exist_ok=True)
    json.dump({**m, "_v": MAPPING_V, "_type": type_}, open(chemin, "w", encoding="utf-8"), indent=2, ensure_ascii=False)
    print(f"Mapping détecté et enregistré : {chemin}")
    return m

# ---------- Lecture & conversions ----------
COLONNES_STD = ["Date facture", "N° facture", "Client", "Total HT", "Taux TVA", "Total TVA", "Total TTC"]

def lire_ia(fichier, type_="vente"):
    """Extrait les factures d'un PDF/image via Claude (payant). Nécessite ANTHROPIC_API_KEY."""
    if not cle_api():
        raise ValueError("Lecture IA : aucune clé API valide (elle commence par sk-ant- et fait environ 100 caractères).")
    import anthropic, base64
    ext = os.path.splitext(fichier)[1].lower()
    media = {".pdf": "application/pdf", ".png": "image/png", ".jpg": "image/jpeg", ".jpeg": "image/jpeg"}[ext]
    b64 = base64.standard_b64encode(open(fichier, "rb").read()).decode()
    bloc = {"type": "document" if ext == ".pdf" else "image",
            "source": {"type": "base64", "media_type": media, "data": b64}}
    qui = "le CLIENT (destinataire)" if type_ == "vente" else "le FOURNISSEUR (émetteur)"
    consigne = (f"Extrais toutes les factures de ce document. Pour 'tiers', donne {qui}. "
                "Dates au format JJ/MM/AAAA. Montants négatifs pour un avoir. Une entrée de 'lignes' par taux de TVA. "
                "Réponds uniquement par un JSON, sans texte autour : "
                '[{"date":"","numero":"","tiers":"","total_ttc":0.00,'
                '"lignes":[{"taux_tva":20,"ht":0.00,"tva":0.00}]}]')
    cle = cle_api()
    try:
        rep = anthropic.Anthropic(api_key=cle).messages.create(
            model="claude-sonnet-4-6", max_tokens=4000,
            messages=[{"role": "user", "content": [bloc, {"type": "text", "text": consigne}]}])
    except anthropic.AuthenticationError:
        raise ValueError(f"Clé API refusée par Anthropic (clé lue : {len(cle)} caractères, "
                         f"commence par « {cle[:7]} », finit par « {cle[-4:]} »). "
                         "Une clé valide commence par sk-ant- et fait environ 100 caractères.")
    except anthropic.APIStatusError as e:
        raise ValueError(f"Erreur API Anthropic : {getattr(e, 'message', e)}")
    try:
        data = json.loads(re.sub(r"```json|```", "", rep.content[0].text).strip())
    except Exception:
        raise ValueError(f"Réponse IA illisible pour {os.path.basename(fichier)}")
    data = [data] if isinstance(data, dict) else data
    rows, anom = [], []
    for fac in data:
        somme, total = Decimal("0"), montant(fac.get("total_ttc"))
        for l in fac.get("lignes") or []:
            ht, tva = montant(l.get("ht")), montant(l.get("tva"))
            if ht is None: continue
            tva = tva if tva is not None else Decimal("0")
            somme += ht + tva
            rows.append({"Date facture": fac.get("date"), "N° facture": fac.get("numero"),
                         "Client": fac.get("tiers"), "Total HT": str(ht),
                         "Taux TVA": str(l.get("taux_tva", "")), "Total TVA": str(tva),
                         "Total TTC": str(ht + tva)})
        if total is not None and abs(somme - total) > Decimal("0.02"):
            anom.append((f"Lecture IA : total reconstitué ({somme}) <> TTC affiché ({total})", fac.get("numero", "")))
    df = pd.DataFrame(rows, columns=COLONNES_STD)
    df.attrs["anomalies"] = anom
    return df

# ---------- Lecture gratuite : XML (Factur-X/CII, UBL), texte du PDF, OCR (facultatif) ----------
STD_TAUX = [20.0, 10.0, 5.5, 2.1, 0.0]
MOIS = {"janvier": 1, "janv": 1, "jan": 1, "fevrier": 2, "fevr": 2, "fev": 2, "mars": 3, "mar": 3, "avril": 4,
        "avr": 4, "mai": 5, "juin": 6, "jun": 6, "juillet": 7, "juil": 7, "aout": 8, "septembre": 9, "sept": 9,
        "sep": 9, "octobre": 10, "oct": 10, "novembre": 11, "nov": 11, "decembre": 12, "dec": 12,
        "january": 1, "february": 2, "feb": 2, "march": 3, "april": 4, "may": 5, "june": 6, "july": 7, "jul": 7,
        "august": 8, "aug": 8, "september": 9, "october": 10, "november": 11, "december": 12}
NB = r"(-?(?:\d{1,3}(?:[ .,]\d{3})+|\d+)(?:[.,]\d{1,2})?)(?![\d]|[.,]\d|\s*%)"
TOKEN = r"([A-Z0-9][A-Z0-9\-/_.]*\d[A-Z0-9\-/_.]*)"
D_NUM = r"(?<![\d/.\-])(?:\d{4}-\d{2}-\d{2}|\d{1,2}[/.\-]\d{1,2}[/.\-]\d{2,4})(?!\d)"
D_FR = r"\d{1,2}(?:er)?\s+[A-Za-zÀ-ÿ]{3,9}\.?,?\s+\d{4}"
D_EN = r"[A-Za-zÀ-ÿ]{3,9}\.?\s+\d{1,2}(?:st|nd|rd|th)?,?\s+\d{4}"
D_ALL = re.compile("(" + D_NUM + ")|(" + D_FR + ")|(" + D_EN + ")", re.I)

# ----- XML : Factur-X / ZUGFeRD (CII) et UBL -----
def _loc(e): return e.tag.rsplit("}", 1)[-1]
def _desc(n, nom): return [e for e in n.iter() if _loc(e) == nom]
def _enf(n, nom): return [e for e in n if _loc(e) == nom]
def _txt(n, nom):
    r = _desc(n, nom)
    return (r[0].text or "").strip() if r else ""

def _parse_cii(racine, type_):
    doc = _desc(racine, "ExchangedDocument")[0]
    d = _txt(_desc(doc, "IssueDateTime")[0], "DateTimeString")
    d = f"{d[6:8]}/{d[4:6]}/{d[:4]}" if len(d) == 8 and d.isdigit() else (_conv_date(d) or d)
    agr = _desc(racine, "ApplicableHeaderTradeAgreement")[0]
    tiers = _txt(_desc(agr, "BuyerTradeParty" if type_ == "vente" else "SellerTradeParty")[0], "Name")
    reg = _desc(racine, "ApplicableHeaderTradeSettlement")[0]
    signe = -1 if _txt(doc, "TypeCode") == "381" else 1
    rows, somme = [], Decimal("0")
    for t in _enf(reg, "ApplicableTradeTax"):
        ht, tva = montant(_txt(t, "BasisAmount")), montant(_txt(t, "CalculatedAmount")) or Decimal("0")
        if ht is None: continue
        ht, tva = ht * signe, tva * signe
        somme += ht + tva
        rows.append({"Date facture": d, "N° facture": _txt(doc, "ID"), "Client": tiers, "Total HT": str(ht),
                     "Taux TVA": _txt(t, "RateApplicablePercent"), "Total TVA": str(tva), "Total TTC": str(ht + tva)})
    total = montant(_txt(_desc(racine, "SpecifiedTradeSettlementHeaderMonetarySummation")[0], "GrandTotalAmount"))
    return rows, somme, (total * signe if total is not None else None), _txt(doc, "ID")

def _parse_ubl(racine, type_):
    numero = (_enf(racine, "ID") or [None])[0]
    numero = (numero.text or "").strip() if numero is not None else ""
    d = _conv_date(_txt(racine, "IssueDate")) or _txt(racine, "IssueDate")
    p = _desc(racine, "AccountingCustomerParty" if type_ == "vente" else "AccountingSupplierParty")[0]
    tiers = _txt(p, "RegistrationName") or _txt(p, "Name")
    signe = -1 if _loc(racine) == "CreditNote" or _txt(racine, "InvoiceTypeCode") == "381" else 1
    rows, somme = [], Decimal("0")
    for tt in _enf(racine, "TaxTotal"):
        for st in _enf(tt, "TaxSubtotal"):
            ht, tva = montant(_txt(st, "TaxableAmount")), montant(_txt(st, "TaxAmount")) or Decimal("0")
            if ht is None: continue
            ht, tva = ht * signe, tva * signe
            somme += ht + tva
            rows.append({"Date facture": d, "N° facture": numero, "Client": tiers, "Total HT": str(ht),
                         "Taux TVA": _txt(st, "Percent"), "Total TVA": str(tva), "Total TTC": str(ht + tva)})
    lm = _enf(racine, "LegalMonetaryTotal")
    total = montant(_txt(lm[0], "TaxInclusiveAmount")) if lm else None
    if not rows and lm:                                   # facture sans TVA détaillée
        ht = montant(_txt(lm[0], "TaxExclusiveAmount"))
        if ht is not None:
            rows.append({"Date facture": d, "N° facture": numero, "Client": tiers, "Total HT": str(ht * signe),
                         "Taux TVA": "0", "Total TVA": "0", "Total TTC": str(ht * signe)})
            somme = ht * signe
    return rows, somme, (total * signe if total is not None else None), numero

def _parse_xml(contenu, type_):
    try:
        racine = ET.fromstring(contenu)
        nom = _loc(racine)
        if nom == "CrossIndustryInvoice": rows, somme, total, num = _parse_cii(racine, type_)
        elif nom in ("Invoice", "CreditNote"): rows, somme, total, num = _parse_ubl(racine, type_)
        else: return None
        if not rows: return None
        df = pd.DataFrame(rows, columns=COLONNES_STD)
        df.attrs["anomalies"] = []
        if total is not None and abs(somme - total) > Decimal("0.02"):
            df.attrs["anomalies"].append((f"XML : total reconstitué ({somme}) <> TTC déclaré ({total})", num))
        return df
    except Exception:
        return None

def lire_facturx(fichier, type_="vente"):
    """XML Factur-X / ZUGFeRD / UBL embarqué dans un PDF. None s'il n'y en a pas."""
    try:
        from pypdf import PdfReader
        for n, c in PdfReader(fichier).attachments.items():
            if n.lower().endswith(".xml"):
                df = _parse_xml(c[0], type_)
                if df is not None: return df
    except Exception:
        pass
    return None

def lire_xml(fichier, type_="vente"):
    df = _parse_xml(open(fichier, "rb").read(), type_)
    if df is None:
        raise ValueError("XML non reconnu (Factur-X/CII ou UBL attendu)")
    return df

# ----- Texte des PDF -----
def textes_pdf(fichier):
    """Plusieurs extractions du texte (chacune = liste de pages) : les mises en page varient."""
    out = []
    try:
        import pdfplumber
        with pdfplumber.open(fichier) as pdf:
            out.append([(p.extract_text() or "") for p in pdf.pages])
            out.append([(p.extract_text(layout=True) or "") for p in pdf.pages])
    except Exception:
        pass
    try:
        from pypdf import PdfReader
        out.append([(p.extract_text() or "") for p in PdfReader(fichier).pages])
    except Exception:
        pass
    return [v for v in out if "".join(v).strip()]

_OCR = None
def ocr_texte(fichier):
    """OCR via PaddleOCR (facultatif : pip install -r requirements-ocr.txt). None si indisponible."""
    global _OCR
    try:
        from paddleocr import PaddleOCR
    except ImportError:
        return None
    try:
        import tempfile
        pages = [fichier]
        if fichier.lower().endswith(".pdf"):
            import pdfplumber
            tmp, pages = tempfile.mkdtemp(), []
            with pdfplumber.open(fichier) as pdf:
                for i, p in enumerate(pdf.pages[:10]):
                    ch = os.path.join(tmp, f"p{i}.png"); p.to_image(resolution=200).save(ch); pages.append(ch)
        if _OCR is None:
            try: _OCR = PaddleOCR(lang="fr", use_textline_orientation=True)
            except Exception: _OCR = PaddleOCR(lang="fr", use_angle_cls=True)
        sortie = []
        for ch in pages:
            items = []
            if hasattr(_OCR, "predict"):
                for res in _OCR.predict(ch):
                    for t, b in zip(res["rec_texts"], res["rec_boxes"]):
                        items.append((float(b[1]), float(b[0]), t))
            else:
                for res in _OCR.ocr(ch, cls=True) or []:
                    for box, (t, _) in (res or []):
                        items.append((box[0][1], box[0][0], t))
            ligne, y0 = [], None
            for y, x, t in sorted(items):
                if y0 is not None and y - y0 > 15:
                    sortie.append("  ".join(t2 for _, t2 in sorted(ligne))); ligne = []
                if not ligne: y0 = y
                ligne.append((x, t))
            if ligne: sortie.append("  ".join(t2 for _, t2 in sorted(ligne)))
        return "\n".join(sortie)
    except Exception as e:
        print(f"[OCR indisponible : {e}]", file=sys.stderr); return None

# ----- Moteur de règles sur le texte -----
def _norm_txt(t):
    t = t.replace("\u00a0", " ").replace("\u202f", " ").replace("’", "'").replace("‘", "'")
    return re.sub(r"[ \t]{3,}", "  ", t)

def _num(texte, motifs, dernier=False):
    """Montant qui suit un libellé : sur la même ligne, ou seul sur la ligne suivante."""
    for m in motifs:
        vals = []
        for pat in (m + r"[^\d\n\-]{0,60}" + NB,
                    m + r"[^\n\d]{0,40}\n[ ]*[€$£]?[ ]*" + NB + r"[ ]*(?:€|eur|euros?|\$|£)?[ ]*(?:\n|$)"):
            vals = [montant(r.group(1)) for r in re.finditer(pat, texte, re.I | re.M)]
            vals = [v for v in vals if v is not None]
            if vals: break
        if vals: return vals[-1] if dernier else vals[0]
    return None

DEV = r"(?:\(?\s*(?:eur|€|euros?)\s*\)?\s*)?"
MOTIFS_HT = [r"total\s*" + DEV + r"(?:net\s*)?h\.?\s?t\.?(?![a-z])", r"total\s*hors\s*t(?:axes?|\.?v\.?a\.?)",
             r"montant\s*(?:total\s*)?(?:h\.?\s?t\.?(?![a-z])|hors\s*taxes?)",
             r"total\s*(?:net\s*)?(?:avant|hors)\s*(?:tva|taxes?)", r"net\s*h\.?\s?t\.?(?![a-z])",
             r"total\s*(?:excl(?:uding|\.)?\s*(?:vat|tax)|before\s*tax)", r"untaxed\s*amount", r"net\s*amount", r"total\s*net\s*amount",
             r"sous[\s-]?total(?:\s*h\.?\s?t\.?(?![a-z]))?", r"subtotal", r"montant\s*net", r"base\s*h\.?\s?t\.?(?![a-z])"]
MOTIFS_TTC = [r"total\s*" + DEV + r"(?:g[ée]n[ée]ral\s*)?t\.?t\.?c\.?", r"montant\s*(?:total\s*)?t\.?t\.?c\.?",
              r"total\s*(?:tva\s*(?:incluse|comprise)|toutes?\s*taxes?\s*comprises?)",
              r"total\s*(?:incl(?:uding|\.)?\s*(?:vat|tax))", r"grand\s*total", r"gross\s*total", r"total\s*gross", r"total\s*amount", r"total\s*due",
              r"amount\s*due", r"net\s*[àa]\s*payer", r"total\s*[àa]\s*(?:payer|r[ée]gler)", r"[àa]\s*r[ée]gler",
              r"montant\s*total", r"total\s*g[ée]n[ée]ral"]
MOTIFS_TVA = [r"total\s*t\.?v\.?a\.?", r"montant\s*(?:de\s*la\s*)?t\.?v\.?a\.?", r"total\s*(?:vat|tax(?:es)?)"]
MOTIF_TOTAL_SEUL = [r"^[ ]*total(?!\s*(?:h|t\.?v|hors|net|g|excl|incl|amount|due))"]

def _tva_ligne(lignes):
    """Ligne du type 'TVA 20 % : 80,00' : dernier montant à décimales (ignore n° de TVA, mentions d'exonération)."""
    for l in lignes:
        if re.match(r"\W*(?:t\.?v\.?a\.?|vat)\b", l, re.I) and \
           not re.search(r"non\s*applicable|n°|siren|siret|293|intracom|exon|incluse|comprise|\bFR\d|\bon\s*[€$£]?\s*\d|\bsur\s*[€$£]?\s*\d", l, re.I):
            v = [montant(x) for x in re.findall(NB, re.sub(r"\d{1,2}(?:[.,]\d{1,2})?\s*%", " ", l)) if re.search(r"[.,]\d{2}$", x)]
            v = [x for x in v if x is not None]
            if v: return v[-1]
    return None

def _totaux_tableau(lignes):
    """Tableau : en-tête 'HT / TVA / TTC' puis une ligne de montants."""
    for i, l in enumerate(lignes):
        if re.search(r"h\.?t\b", l, re.I) and re.search(r"t\.?t\.?c|t\.?v\.?a", l, re.I):
            for s in lignes[i + 1:i + 4]:
                v = [montant(x) for x in re.findall(NB, s)]
                v = [x for x in v if x is not None]
                if len(v) >= 3 and abs(v[-3] + v[-2] - v[-1]) <= Decimal("0.02"): return v[-3], v[-1]
                if len(v) == 2 and 0 < v[0] < v[1] <= v[0] * Decimal("1.25"): return v[0], v[1]
    return None, None

def _taux_lignes(lignes):
    """Lignes à un seul taux (ex. '20 %  145,83  29,17') -> (taux, base, tva), selon que la ligne
    contient base+TVA (A) ou base+TTC (B). Les montants sont reconnus par le calcul, pas par leur place."""
    A, B = [], []
    for li, l in enumerate(lignes):
        p = list(re.finditer(r"(\d{1,2}(?:[.,]\d{1,2})?)\s*%", l))
        if len(p) != 1: continue
        r = Decimal(p[0].group(1).replace(",", "."))
        nums = [montant(x) for x in re.findall(NB, l[:p[0].start()] + " " + l[p[0].end():])]
        nums = [abs(x) for x in nums if x is not None]
        pa = [(b, v) for i, b in enumerate(nums) for j, v in enumerate(nums)
              if i != j and b > v and abs(b * r / 100 - v) <= Decimal("0.02")]
        pb = [(b, t - b) for i, b in enumerate(nums) for j, t in enumerate(nums)
              if i != j and t > b and r > 0 and abs(b * (1 + r / 100) - t) <= Decimal("0.02")]
        if pa: b, v = max(pa); A.append((float(r), b, v))
        elif pb: b, v = max(pb); B.append((float(r), b, v))
        elif len(nums) == 1 and li + 1 < len(lignes):          # base sur la ligne du taux, TVA sur la ligne suivante
            suiv = [montant(x) for x in re.findall(NB, lignes[li + 1])]
            suiv = [abs(x) for x in suiv if x is not None]
            if len(suiv) == 1 and nums[0] > suiv[0] and abs(nums[0] * r / 100 - suiv[0]) <= Decimal("0.02"):
                A.append((float(r), nums[0], suiv[0]))
    return A, B

def _ventiler(candidats, ht, tva_tot):
    for rows in candidats:
        if not rows: continue
        g = OrderedDict()
        for r, b, v in rows:
            b0, v0 = g.get(r, (Decimal("0"), Decimal("0")))
            g[r] = (b0 + b, v0 + v)
        if abs(sum(b for b, _ in g.values()) - abs(ht)) <= Decimal("0.05") and \
           abs(sum(v for _, v in g.values()) - abs(tva_tot)) <= Decimal("0.05"):
            return [(r, b, v) for r, (b, v) in g.items()]
    return None

# ----- date -----
def _conv_date(d):
    d = d.strip()
    r = re.fullmatch(r"(\d{4})-(\d{2})-(\d{2})", d)
    if r: return f"{r.group(3)}/{r.group(2)}/{r.group(1)}"
    if re.fullmatch(D_NUM, d):
        r = re.fullmatch(r"(\d{1,2})([/.\-])(\d{1,2})[/.\-](\d{2,4})", d)
        if r and int(r.group(3)) > 12 and int(r.group(1)) <= 12:      # 03/15/2026 : format américain évident
            return f"{r.group(3)}/{r.group(1)}/{r.group(4)}"
        return d
    r = re.fullmatch(r"(\d{1,2})(?:er)?\s+([A-Za-zÀ-ÿ]{3,9})\.?,?\s+(\d{4})", d)
    if r: jour, mois, an = r.groups()
    else:
        r = re.fullmatch(r"([A-Za-zÀ-ÿ]{3,9})\.?\s+(\d{1,2})(?:st|nd|rd|th)?,?\s+(\d{4})", d)
        if not r: return ""
        mois, jour, an = r.groups()
    m = MOIS.get(norm(mois))
    return f"{int(jour):02d}/{m:02d}/{an}" if m else ""

LAB_DATE = [r"date\s*(?:de\s*la\s*|de\s*|d')?(?:facture|facturation|[ée]mission|[ée]dition|pi[eè]ce|document|comptable)",
            r"(?:invoice|issue|issued)\s*date", r"date\s*of\s*(?:invoice|issue)", r"[ée]mise?\s*le", r"factur[ée]e?\s*le",
            r"facture\s*du", r"fait\s*le", r"\bdate\b"]

def _date_texte(texte):
    t = _norm_txt(texte)
    ds = []
    for r in D_ALL.finditer(t):
        v = _conv_date(r.group(0))
        if v: ds.append((r.start(), r.end(), v))
    for pat in LAB_DATE:                                   # date derrière un libellé (même ligne ou ligne suivante)
        for m in re.finditer(pat, t, re.I):
            for d0, _, v in ds:
                if d0 >= m.end():
                    if not t[m.end():d0].strip(" :.-\t\n"): return v
                    break
    for d0, _, v in ds:                                    # sinon 1re date qui n'est pas une échéance
        if not re.search(r"[ée]ch[ée]ance|due|limite|paiement|r[èe]glement|livraison|validit|expir|avant le", t[max(0, d0 - 30):d0], re.I):
            return v
    return ds[0][2] if ds else ""

# ----- numéro -----
PATS_NUM = [
    (True, r"(?:facture|avoir|invoice|credit\s*note|note\s*de\s*cr[ée]dit|pi[eè]ce|document)\s*(?:d'avoir\s*)?"
           r"(?:n[°ºo]\.?|num[ée]ro|number|no\.?|nr\.?|#|r[ée]f(?:[ée]rence)?\.?)\s*:?\s*"),
    (True, r"(?:n[°ºo]\.?|num[ée]ro|number|no\.?|#|r[ée]f(?:[ée]rence)?\.?)\s*(?:de\s*|d')?(?:facture|invoice|pi[eè]ce|avoir)\s*:?\s*"),
    (False, r"(?:facture|avoir|invoice)\s*:?\s*"),
    (True, r"\bnum[ée]ro\s*:?\s*"),
    (True, r"\bn[°º]\s*:?\s*"),
]

PLAQUE = re.compile(r"^[A-Z]{2}[-\s]?\d{3}[-\s]?[A-Z]{2}$|^\d{1,4}[-\s]?[A-Z]{2,3}[-\s]?\d{2,3}$", re.I)
CTX_NON_FACTURE = (r"immat\w*|plaque|v[ée]hicule|chassis|châssis|\bvin\b|\bkm\b|kilom\w*|s[ée]rie|t[ée]l\w*|mobile|portable|"
                   r"fax|siren|siret|\btva\b|iban|\bbic\b|\brcs\b|\bnaf\b|\bape\b|client|compte|commande|contrat|dossier|"
                   r"devis|bon\s*de|livraison|colis|suivi|imei")
LIGNE_VEHICULE = r"immat\w*|plaque|v[ée]hicule|marque|mod[èe]le|kilom\w*|\bkm\b|chassis|châssis"

def _plaques(t):
    """Immatriculations citées dans le texte (à ne jamais prendre pour un numéro de facture)."""
    out = set()
    for pat in (r"(?:immat\w*|plaque\w*)\W{0,25}([A-Z0-9][A-Z0-9\-]{4,10})", r"(?:immat\w*|plaque\w*)\W{0,25}([A-Z]{2}[\s-]?\d{3}[\s-]?[A-Z]{2})"):
        for r in re.finditer(pat, t, re.I):
            out.add(re.sub(r"[^A-Z0-9]", "", r.group(1).upper()))
    return out

def _numero(texte):
    t = _norm_txt(texte)
    plaques = _plaques(t)
    for explicite, pat in PATS_NUM:
        for r in re.finditer(pat + TOKEN, t, re.I):
            c = r.group(1).strip(".-/_,")
            if not c or re.fullmatch(D_NUM, c) or re.fullmatch(r"(?:19|20)\d{2}", c) or re.fullmatch(r"\d{9}|\d{14}", c): continue
            if len(c) < (1 if explicite else 4): continue
            if PLAQUE.match(c) or re.sub(r"[^A-Z0-9]", "", c.upper()) in plaques: continue          # immatriculation
            if c.isdigit() and len(c) <= 4 and re.match(r"[\s-]*[A-Z]{2,3}[\s-]*\d{2,3}\b", t[r.end(1):r.end(1) + 8]): continue   # 1234 AB 56
            ls_, le_ = t.rfind("\n", 0, r.start(1)) + 1, t.find("\n", r.end(1))
            ligne = t[ls_:(le_ if le_ >= 0 else len(t))]
            ctx = t[ls_:r.start(1)]                                  # libellé sur la même ligne...
            if not ctx.strip():                                      # ...ou sur la ligne précédente (valeur seule)
                ctx = t[t.rfind("\n", 0, max(ls_ - 1, 0)) + 1:ls_]
            if re.search(r"factur|invoice|avoir", ctx + " " + ligne, re.I): pass
            elif re.search(CTX_NON_FACTURE, ctx, re.I) or re.search(LIGNE_VEHICULE, ligne, re.I): continue
            return c
    return ""

# ----- tiers -----
ETIQ = {"vente": r"(?:factur[ée]e?\s*[àa]|adresse\s*de\s*facturation|[àa]\s*l'attention\s*de|[àa]\s*destination\s*de|"
                 r"destinataire|client|customer|bill(?:ed)?\s*to|sold\s*to|adress[ée]e?\s*[àa]|nom\s*du\s*client)",
        "achat": r"(?:fournisseur|[ée]metteur|vendeur|vendu\s*par|prestataire|envoy[ée]e?\s*par|from|seller|supplier)"}
NON_NOM = (r"^(?:factur\w*|avoir|invoice|devis|date|num[ée]ro|n[°º]|total|page|adresse|adress[ée]\w*|siret|siren|tva|iban|bic|"
           r"t[ée]l|email|e-mail|r[ée]f|code|france|paiement|conditions|mode|ech[ée]ance|objet|d[ée]signation|"
           r"description|quantit[ée]|prix|montant|bon\s*de|commande|client|destinataire|livr\w*|bill|sold|ship|"
           r"doit|[àa]\s*l'attention|[àa]\s*destination)\b")
RUE = r"(?:rue|avenue|av|bd|boulevard|chemin|impasse|place|all[ée]e|route|cours|quai|zone|za|zi|lotissement|r[ée]sidence)"
LABELS_ACHETEUR = (r"(?:factur[ée]e?\s*[àa]|adress[ée]e?\s*[àa]|adresse\s*de\s*facturation|[àa]\s*l'attention|"
                   r"[àa]\s*destination|destinataire|client|customer|bill(?:ed)?\s*to|sold\s*to|livr[ée]e?\s*[àa]|"
                   r"nom\s*du\s*client|ship\s*to)")
LEGAL = (r"siret|siren|\brcs\b|capital|n[°º]\s*(?:de\s*)?tva|tva\s*intra\w*|identifiant\s*tva|\bape\b|\bnaf\b|"
         r"immatricul|vat\s*(?:id|no|number)")
SEP_LEGAL = (r"\s*[|·•]\s*|\s+[-–—]\s+|,\s|\s+(?:au\s+)?capital|\s+siret|\s+siren|\s+rcs|\s+n[°º]|\s+tva|\s+ape\b|"
             r"\s+naf\b|\s+immatricul|\s+soci[ée]t[ée]\s+[àa]")

TITRE_MOTS = {"facture", "factures", "avoir", "invoice", "devis", "proforma", "pro", "forma", "acompte", "acomptes", "situation",
              "recapitulatif", "releve", "decompte", "honoraires", "honoraire", "note", "bon", "recu", "quittance", "ticket",
              "commande", "duplicata", "original", "copie", "reglement", "document", "piece", "detail", "detaillee", "definitive",
              "provisoire", "rectificative", "complementaire", "finale", "intermediaire", "mensuelle", "trimestrielle",
              "annuelle", "credit", "debit", "livraison", "echeancier", "attestation", "synthese"}
TITRE_LIANTS = {"de", "du", "des", "la", "le", "les", "d", "l", "n", "no", "nr", "numero", "num", "et", "vente", "achat",
                "service", "services", "prestation", "prestations", "travaux", "frais", "a", "au", "aux", "en", "sur", "pour",
                "periode", "ref", "reference", "date"}

def _est_titre(s):
    """Vrai si le texte n'est qu'un titre de document (Facture, Note d'honoraires, Facture de situation n°2...)."""
    jetons = [t for t in norm(s).split() if not re.fullmatch(r"n?[\d%.,/\-]+\w{0,2}", t)]   # "n°2" -> "n2" est un numéro
    return all(t in TITRE_MOTS or t in TITRE_LIANTS for t in jetons)

def _nom_valide(s):
    s = s.strip(" :-|")
    if len(re.findall(r"[A-Za-zÀ-ÿ]", s)) < 3 or "@" in s or re.search(r"\d{8,}", s.replace(" ", "")): return False
    if _est_titre(s) or re.match(NON_NOM, s, re.I) or re.match(r"^\d{5}\b", s): return False
    if re.match(r"^n[°º]|^no\.?\s*\d", s, re.I): return False                         # "N° FA-2026-12" n'est pas un nom
    if sum(c.isdigit() for c in s) > sum(c.isalpha() for c in s): return False         # plus de chiffres que de lettres : une référence
    if re.match(rf"^(?:\d{{1,4}}\s*(?:bis|ter)?[\s,]+)?{RUE}\b", s, re.I): return False
    return True

def _nettoyer_nom(s):
    s = s.strip(" :-|")
    if _est_titre(s): return ""                                   # le titre du document n'est pas un tiers
    s = re.sub(r"^(?:facture|avoir|invoice)\s+", "", s, flags=re.I)   # "Facture Dupont SARL" (titre collé au nom)
    s = re.split(r"\s{2,}|\s+(?:date|n[°º]|num[ée]ro|r[ée]f|page|t[ée]l|email|e-mail|siret|siren|tva)\b", s, 1, flags=re.I)[0].strip(" :-|")
    return "" if _est_titre(s) else s

FORMES = {"sarl", "sas", "sasu", "eurl", "sa", "sci", "snc", "scop", "selarl", "earl", "gaec", "ets", "etablissements",
          "societe", "entreprise", "cie", "and", "les", "des", "du", "la", "le", "de", "et"}

def _meme_societe(a, b):
    ja = {t for t in norm(a).split() if len(t) >= 3 and t not in FORMES}
    jb = {t for t in norm(b).split() if len(t) >= 3 and t not in FORMES}
    return bool(ja and jb) and (ja <= jb or jb <= ja)

def _emetteur_legal(lignes):
    """Nom de l'émetteur d'après les mentions légales (SIRET, RCS, capital, n° TVA...), en pied de page ou en en-tête."""
    for i, l in enumerate(lignes):
        if re.search(LEGAL, l, re.I) and not re.match(rf"\s*{LABELS_ACHETEUR}\b", l, re.I):
            for cand in (re.split(SEP_LEGAL, l, 1, flags=re.I)[0], lignes[i - 1] if i else ""):
                c = _nettoyer_nom(cand)
                if c and _nom_valide(c): return c
    return ""

def _tiers_texte(lignes, type_):
    for i, l in enumerate(lignes):                                  # 1. libellé explicite
        m = re.match(rf"\s*{ETIQ[type_]}\b\s*:?\s*(.*)$", l, re.I)
        if not m: continue
        for cand in [m.group(1)] + lignes[i + 1:i + 5]:
            c = _nettoyer_nom(cand)
            if c and _nom_valide(c): return c
    em = _emetteur_legal(lignes)                                    # 2. mentions légales de l'émetteur
    if type_ == "achat":
        if em: return em
        skip = 0                                                    # 3. sinon 1re ligne hors bloc client
        for l in lignes[:15]:
            if skip: skip -= 1; continue
            if re.match(rf"\s*{LABELS_ACHETEUR}\b", l, re.I): skip = 4; continue
            c = _nettoyer_nom(l)
            if c and _nom_valide(c): return c
    elif em:                                                        # client : 1re ligne qui n'est pas l'émetteur
        for l in lignes[:15]:
            c = _nettoyer_nom(l)
            if c and _nom_valide(c) and not _meme_societe(c, em): return c
    return ""

def lire_texte(texte, type_, nom):
    """Extraction par règles à partir du texte d'une facture (une facture par document)."""
    t = _norm_txt(texte)
    lignes = [l.strip() for l in t.splitlines() if l.strip()]
    anom, numero = [], _numero(t)
    trouve = bool(numero)
    if not numero:
        numero = nom; anom.append(("N° de facture non trouvé (nom du fichier utilisé)", nom))
    ht, ttc, tva_exp = _num(t, MOTIFS_HT), _num(t, MOTIFS_TTC), _num(t, MOTIFS_TVA)
    if tva_exp is None: tva_exp = _tva_ligne(lignes)
    if ttc is None: ttc = _num(t, MOTIF_TOTAL_SEUL, dernier=True)
    if ht is None or ttc is None:
        h2, t2 = _totaux_tableau(lignes)
        ht = ht if ht is not None else h2; ttc = ttc if ttc is not None else t2
    if ht is None and ttc is not None and tva_exp is not None: ht = ttc - tva_exp
    if ttc is None and ht is not None and tva_exp is not None: ttc = ht + tva_exp
    if re.search(r"tva\s*non\s*applicable|art(?:icle|\.)?\s*293\s*b|franchise\s*en\s*base|exon[ée]r[ée]e?\s*de\s*tva|vat\s*exempt", t, re.I):
        if ht is None and ttc is not None: ht = ttc
        if ttc is None and ht is not None: ttc = ht
    if ht is None or ttc is None or ht == 0:
        return None
    if tva_exp is not None and abs(ht + tva_exp - ttc) > Decimal("0.02"):
        anom.append((f"HT ({ht}) + TVA ({tva_exp}) <> TTC ({ttc})", numero))
    tva_tot = ttc - ht
    A, B = _taux_lignes(lignes)
    vent = _ventiler([list(OrderedDict.fromkeys(A)), A, B], ht, tva_tot)
    if vent is None:
        if tva_tot == 0: vent = [(0.0, ht, Decimal("0"))]
        else:
            calc = float(tva_tot / ht * 100)
            proche = min(STD_TAUX, key=lambda x: abs(x - calc))
            if abs(proche - calc) <= 0.3: taux = proche
            else:
                taux = round(calc, 1)
                anom.append((f"Taux de TVA calculé {taux} % : facture multi-taux ou erreur, à vérifier", numero))
            vent = [(taux, ht, tva_tot)]
    else:
        diff = ttc - sum(b + v for _, b, v in vent)
        if diff and abs(diff) <= Decimal("0.05"):
            r_, b_, v_ = vent[-1]; vent[-1] = (r_, b_, v_ + diff)
    avoir = re.search(r"^\s*(?:facture\s*d')?avoir\b(?!\s*lieu)|credit\s*note|note\s*de\s*cr[ée]dit|facture\s*d'avoir",
                      "\n".join(lignes[:25]), re.I | re.M)
    sg = -1 if (avoir and ht > 0) else 1
    if type_ == "achat" and re.search(r"cadeaux?\b|coffrets?\s*cadeau|carte\s*cadeau|bon\s*cadeau", t, re.I):
        anom.append((msg_cadeau(), numero))
    if re.search(r"autoliquidation|reverse\s*charge", t, re.I):
        anom.append(("Autoliquidation de TVA : écritures de TVA à traiter manuellement", numero))
    if re.search(r"\b(?:USD|GBP|CHF)\b|[$£]", t) and not re.search(r"€|\bEUR\b|euros?", t, re.I):
        anom.append(("Devise autre que l'euro : montants à convertir", numero))
    tiers = _tiers_texte(lignes, type_)
    if not tiers:
        tiers = "A IDENTIFIER"; anom.append(("Tiers non identifié", numero))
    d = _date_texte(t)
    df = pd.DataFrame([{"Date facture": d, "N° facture": numero, "Client": tiers, "Total HT": str(sg * b),
                        "Taux TVA": taux_str(r), "Total TVA": str(sg * v), "Total TTC": str(sg * (b + v))}
                       for r, b, v in vent], columns=COLONNES_STD)
    df.attrs["anomalies"], df.attrs["numero_trouve"] = anom, trouve
    return df

def _lire_pages(pages, type_, nom):
    """Un PDF peut contenir plusieurs factures (une par page) : on les sépare si les numéros diffèrent."""
    if len(pages) > 1:
        res = []
        for i, p in enumerate(pages, 1):
            if len(p.strip()) < 40: continue
            d = lire_texte(p, type_, f"{nom}-p{i}")
            if d is not None and d.attrs.get("numero_trouve"):
                d["Page"] = i
                res.append(d)
        if len(res) >= 2 and len({d.iloc[0]["N° facture"] for d in res}) >= 2:
            df = pd.concat(res, ignore_index=True)
            df.attrs["anomalies"] = [a for d in res for a in d.attrs["anomalies"]]
            return df
    return lire_texte("\n".join(pages), type_, nom)

def lire_pdf(fichier, type_="vente"):
    """Cascade : XML intégré -> texte du PDF (3 extractions) -> OCR (si installé) -> IA (si clé API)."""
    ext = os.path.splitext(fichier)[1].lower()
    nom = os.path.splitext(os.path.basename(fichier))[0]
    variantes = []
    if ext == ".pdf":
        df = lire_facturx(fichier, type_)
        if df is not None: return df
        variantes = textes_pdf(fichier)
    if not any(len("".join(v).strip()) >= 40 for v in variantes):
        t = ocr_texte(fichier)
        variantes = [[t]] if t else []
    best = None
    for pages in variantes:
        cand = _lire_pages(pages, type_, nom)
        if cand is not None and (best is None or len(cand.attrs["anomalies"]) < len(best.attrs["anomalies"])):
            best = cand
            if not cand.attrs["anomalies"]: break
    if best is not None: return best
    texte = "\n".join(variantes[0]) if variantes else ""
    apercu = " | ".join(l.strip() for l in texte.splitlines() if l.strip())[:500]
    apercu = f" Texte lu : {apercu}" if apercu else ""
    if cle_api():
        try:
            return lire_ia(fichier, type_)
        except ValueError as e:
            raise ValueError(f"Règles : montants HT/TTC introuvables ; IA : {e}{apercu}")
    raise ValueError("Lecture automatique impossible" + (" (scan : installez PaddleOCR ou renseignez une clé API)"
                     if len(texte.strip()) < 40 else " (montants HT/TTC introuvables)") + "." + apercu)

def lire(fichier, type_="vente"):
    if fichier.lower().endswith(".xml"):
        return lire_xml(fichier, type_)
    if fichier.lower().endswith((".pdf", ".png", ".jpg", ".jpeg")):
        return lire_pdf(fichier, type_)
    if fichier.lower().endswith((".xlsx", ".xls", ".xlsm")):
        brut = pd.read_excel(fichier, dtype=str, header=None)
        i = _entete_index(brut.fillna("").values.tolist())
        df = brut.iloc[i + 1:].reset_index(drop=True)
        df.columns = [str(c) if pd.notna(c) else f"col{k}" for k, c in enumerate(brut.iloc[i])]
        return df.dropna(how="all")
    for enc in ("utf-8-sig", "cp1252", "latin-1"):
        try:
            texte = open(fichier, encoding=enc).read()
            break
        except UnicodeDecodeError:
            continue
    lignes = texte.splitlines()
    delim = max(";\t,|", key=lambda d: sum(l.count(d) for l in lignes[:20]))
    i = _entete_index([l.split(delim) for l in lignes[:30]])
    return pd.read_csv(io.StringIO("\n".join(lignes[i:])), sep=delim, dtype=str, engine="python").dropna(how="all")

def _entete_index(lignes_cellules):
    """Index de la 1re ligne contenant au moins 3 intitulés de colonnes connus."""
    connus = {x for syns in SYNONYMES.values() for x in syns}
    for i, cells in enumerate(lignes_cellules[:30]):
        if sum(norm(c) in connus for c in cells) >= 3: return i
    return 0

def montant(v):
    if v is None or (isinstance(v, float) and pd.isna(v)) or str(v).strip() in ("", "nan"):
        return None
    s = re.sub(r"[^\d,.\-]", "", str(v).replace("\u202f", "").replace(" ", ""))
    if s.startswith("(") : s = "-" + s
    if re.fullmatch(r"-?\d{1,3}([.,]\d{3})+", s): s = re.sub(r"[.,]", "", s)
    if "," in s and "." in s:
        s = s.replace(",", "") if s.rfind(".") > s.rfind(",") else s.replace(".", "").replace(",", ".")
    else:
        s = s.replace(",", ".")
    try: return Decimal(s).quantize(Decimal("0.01"), ROUND_HALF_UP)
    except Exception: return None

def date(v):
    s = str(v).strip()[:10]
    for f in ("%d/%m/%Y", "%Y-%m-%d", "%d-%m-%Y", "%d.%m.%Y", "%d/%m/%y", "%d.%m.%y", "%d-%m-%y", "%Y/%m/%d"):
        try: return datetime.strptime(s, f)
        except ValueError: pass
    return None

def taux_str(t):
    return str(t).rstrip("0").rstrip(".") if "." in str(t) else str(t)

# ---------- Format pivot ----------
def colonnes_par_taux(df):
    """Colonnes du type 'HT 20%', 'Base HT 5,5', 'TVA 10 %' -> {taux: {'ht': col, 'tva': col}} (au moins un taux avec HT)."""
    res = {}
    for c in df.columns:
        h = str(c).lower()
        r = re.search(r"(?<![\d.,])(\d{1,2}(?:[.,]\d{1,2})?)\s*%", h) or \
            re.search(r"(?:ht|base|tva|taxe|vat)\D{0,10}?(?<![\d.,])(\d{1,2}(?:[.,]\d{1,2})?)(?![\d.,])", h)
        if not r or re.search(r"\bttc\b", h): continue
        t = Decimal(r.group(1).replace(",", "."))
        nh = norm(h)
        est_tva, est_ht = re.search(r"\b(tva|taxe|vat)\b", nh), re.search(r"\b(ht|hors|base|net)\b", nh)
        kind = "tva" if est_tva and not est_ht else "ht" if est_ht and not est_tva else None
        if kind and t <= 30: res.setdefault(taux_str(t), {})[kind] = c
    return {t: v for t, v in res.items() if "ht" in v}

def vers_pivot(df, m, anomalies, src=""):
    for c in OBLIGATOIRES:
        if c not in m:
            raise ValueError(f"{src} : colonne « {CHAMPS_LIBELLES[c]} » non trouvée. Choisissez-la dans "
                             "Réglages → Colonnes des CSV / Excel, puis relancez.")
    lignes, wide = [], colonnes_par_taux(df)
    if len(wide) < 2: wide = {}
    for i, r in df.iterrows():
        n = f"{src}:{i + 2}"
        d = date(r[m["date"]])
        if d is None:
            anomalies.append((n, "Date illisible", r[m["date"]])); continue
        g0 = lambda k: (str(r[m[k]]).strip() if k in m and m[k] in r.index and pd.notna(r[m[k]]) and str(r[m[k]]).strip() not in ("nan", "None") else "")
        g = lambda k: (g0(k) or g0("tiers2")) if k == "tiers" else g0(k)
        page = str(int(float(r["Page"]))) if "Page" in df.columns and pd.notna(r["Page"]) else ""
        if wide:                                              # une colonne HT / TVA par taux : une ligne par taux
            parts = []
            for t, cols in wide.items():
                h = montant(r[cols["ht"]]) if "ht" in cols else None
                v = montant(r[cols["tva"]]) if "tva" in cols else None
                if not h and not v: continue
                h = h or Decimal("0")
                parts.append((t, h, v if v is not None else (h * Decimal(t) / 100).quantize(Decimal("0.01"), ROUND_HALF_UP)))
            if parts:
                ttc_row = montant(r[m["ttc"]]) if "ttc" in m else None
                tot = sum(h + v for _, h, v in parts)
                if ttc_row is not None and abs(tot - ttc_row) > Decimal("0.02"):
                    anomalies.append((n, f"Colonnes par taux : HT+TVA ({tot}) <> TTC ({ttc_row})", g("numero")))
                for t, h, v in parts:
                    lignes.append(dict(n=n, date=d, numero=g("numero"), tiers=g("tiers"), code=g("code_tiers"),
                                       libelle=g("libelle"), compte=g("compte"), page=page, ht=h, tva=v, ttc=h + v, taux=t))
                continue
        ht = montant(r[m["ht"]]) if "ht" in m else None
        tva = montant(r[m["tva"]]) if "tva" in m else None
        ttc = montant(r[m["ttc"]]) if "ttc" in m else None
        taux = montant(r[m["taux_tva"]]) if "taux_tva" in m else None
        if ht is None and ttc is not None and tva is not None: ht = ttc - tva
        if ht is None and ttc is not None and taux is not None:
            ht = (ttc / (1 + taux / 100)).quantize(Decimal("0.01"), ROUND_HALF_UP)
        if ht is None:
            anomalies.append((n, "Montant HT introuvable", "")); continue
        if tva is None:
            tva = (ht * taux / 100).quantize(Decimal("0.01"), ROUND_HALF_UP) if taux is not None \
                  else (ttc - ht if ttc is not None else Decimal("0"))
        if taux is None:
            taux = (tva / ht * 100).quantize(Decimal("0.1")) if ht else Decimal("0")
        lignes.append(dict(n=n, date=d, numero=g("numero"), tiers=g("tiers"), code=g("code_tiers"),
                           libelle=g("libelle"), compte=g("compte"), page=page, ht=ht, tva=tva, ttc=ttc,
                           taux=taux_str(taux)))
    return lignes

# ---------- Table d'imputation (fournisseur/client -> compte) ----------
# Modèle de départ INDICATIF (plan comptable général) : à valider avec le plan de chaque dossier.
_IMP = {
    ("606100", "Fournitures non stockables (énergie, eau)"): [
        "edf", "engie", "totalenergies", "total energies", "direct energie", "enedis", "grdf", "eni", "plenitude",
        "ekwateur", "ilek", "mint energie", "veolia", "suez", "saur"],
    ("606130", "Carburants"): ["esso", "shell", "bp", "avia", "q8", "dkv"],
    ("606300", "Fournitures d'entretien et petit équipement"): [
        "leroy merlin", "castorama", "brico depot", "bricomarche", "ikea", "amazon", "amazon business", "fnac", "darty",
        "boulanger", "cdiscount", "ldlc", "materiel net", "back market", "backmarket", "dell", "lenovo", "apple", "hp",
        "carrefour", "auchan", "lidl", "leclerc", "intermarche", "gifi"],
    ("606400", "Fournitures administratives"): [
        "bureau vallee", "staples", "office depot", "lyreco", "manutan", "viking", "majuscule", "rajapack"],
    ("613200", "Locations immobilières"): ["foncia", "nexity", "citya", "regus", "wework"],
    ("613500", "Locations mobilières et abonnements logiciels"): [
        "hertz", "avis", "europcar", "ada", "rent a car", "leasys", "ayvens", "ald automotive", "leaseplan", "loxam",
        "kiloutou", "boels", "elis",
        "google workspace", "google cloud", "google", "microsoft", "adobe", "ovh", "ovhcloud", "gandi", "ionos", "o2switch",
        "scaleway", "digitalocean", "amazon web services", "aws", "slack", "zoom", "notion", "canva", "dropbox", "github",
        "gitlab", "shopify", "wix", "squarespace", "mailchimp", "brevo", "sendinblue", "hubspot", "salesforce", "openai",
        "anthropic", "atlassian", "figma", "airtable", "zapier", "docusign", "yousign", "pennylane", "tiime", "indy",
        "freebe", "henrri", "axonaut", "sellsy", "zervant", "quickbooks", "xero", "cegid", "sage", "ebp", "acd"],
    ("615000", "Entretien et réparations"): [
        "norauto", "speedy", "midas", "feu vert", "euromaster", "point s", "vulco", "carglass", "onet"],
    ("616000", "Primes d'assurance"): [
        "axa", "maif", "macif", "matmut", "allianz", "groupama", "generali", "hiscox", "april", "swisslife", "mma", "gan",
        "aviva", "ag2r", "malakoff humanis", "harmonie mutuelle", "covea", "pacifica", "direct assurance", "euler hermes"],
    ("618100", "Documentation générale"): [
        "lefebvre dalloz", "dalloz", "editions legislatives", "wolters kluwer", "les echos", "le monde", "le figaro"],
    ("618400", "Versements à des organismes de formation"): ["cegos", "ifocop", "openclassrooms", "udemy", "coursera", "linkedin learning"],
    ("622600", "Honoraires"): [
        "expert comptable", "expertise comptable", "commissaire aux comptes", "avocat", "avocats", "notaire",
        "legalstart", "legalplace", "captain contrat"],
    ("622700", "Frais d'actes et de contentieux"): ["huissier", "infogreffe", "greffe", "inpi"],
    ("623000", "Publicité, annonces et recrutement"): [
        "google ads", "facebook", "meta platforms", "instagram", "linkedin", "pages jaunes", "solocal", "vistaprint",
        "tiktok", "pinterest", "indeed", "welcome to the jungle", "apec", "pole emploi", "france travail"],
    ("624100", "Transports (colis et expéditions)"): [
        "colissimo", "chronopost", "dhl", "ups", "fedex", "tnt", "dpd", "gls", "mondial relay", "relais colis", "boxtal",
        "sendcloud", "packlink", "geodis", "schenker", "heppner"],
    ("625100", "Voyages, déplacements, péages et parkings"): [
        "vinci autoroutes", "sanef", "aprr", "escota", "cofiroute", "sapn", "indigo", "saemes", "onepark",
        "zenpark", "q park", "effia"],
    ("625110", "Voyages : transport de personnes (train, avion, taxi)"): [
        "sncf", "ouigo", "air france", "easyjet", "ryanair", "transavia", "blablacar", "uber", "bolt", "g7", "ratp",
        "navigo", "flixbus", "eurostar"],
    ("625610", "Missions : hébergement (hôtel)"): ["booking", "airbnb", "ibis", "accor", "novotel", "mercure", "campanile", "kyriad", "hotel"],
    ("625700", "Réceptions"): ["uber eats", "deliveroo", "just eat", "mcdonald", "restaurant", "traiteur"],
    ("626000", "Frais postaux et télécommunications"): [
        "orange", "sfr", "bouygues", "free", "free mobile", "la poste", "ringover", "aircall", "onoff", "wifirst", "lebara",
        "sosh", "red by sfr", "numericable", "nordnet"],
    ("627000", "Services bancaires et assimilés"): [
        "qonto", "shine", "bnp paribas", "societe generale", "credit agricole", "caisse d epargne", "lcl", "credit mutuel",
        "cic", "banque populaire", "la banque postale", "boursorama", "revolut", "stripe", "paypal", "sumup", "zettle",
        "alma", "mollie", "adyen", "gocardless", "hello bank", "fortuneo", "bred", "hsbc", "n26"],
    ("628100", "Cotisations"): ["cci", "chambre de commerce", "chambre des metiers", "capeb", "ffb", "medef", "cpme"],
    ("628800", "Autres services extérieurs (sécurité, surveillance)"): ["securitas", "verisure", "prosegur"],
}
IMPUT_DEFAUT = sorted([(m, c, lib, "achat") for (c, lib), ms in _IMP.items() for m in ms], key=lambda r: (r[1], r[0]))
COLS_IMPUT = ["motif", "compte", "libelle", "type"]

def charger_imputations(chemin="imputations.csv"):
    """Table motif;compte;libelle;type (type = achat, vente ou vide pour les deux)."""
    if not os.path.exists(chemin):
        pd.DataFrame(IMPUT_DEFAUT, columns=COLS_IMPUT).to_csv(chemin, index=False, sep=";")
    df = pd.read_csv(chemin, sep=";", dtype=str).fillna("")
    for c in COLS_IMPUT:
        if c not in df.columns: df[c] = ""
    return df[COLS_IMPUT]

ANCIENS_COMPTES = {"606100": ["esso", "shell", "bp", "avia", "q8", "dkv"],
                   "625100": ["sncf", "ouigo", "air france", "easyjet", "ryanair", "transavia", "blablacar", "uber", "bolt", "g7",
                              "ratp", "navigo", "flixbus", "eurostar", "booking", "airbnb", "ibis", "accor", "novotel",
                              "mercure", "campanile", "kyriad"]}

def migrer_modele(df):
    """Met à jour les anciennes lignes du modèle (carburant, transport de personnes, hôtels) vers leurs nouveaux comptes,
    sans toucher aux lignes que l'utilisateur a déjà modifiées. Renvoie (table, nombre de lignes migrées)."""
    cible = {m: (c, l) for m, c, l, _ in IMPUT_DEFAUT}
    df, n = df.copy(), 0
    for i, r in df.iterrows():
        m = norm(_s(r["motif"]))
        if m in ANCIENS_COMPTES.get(_s(r["compte"]), []) and m in cible and cible[m][0] != _s(r["compte"]):
            df.at[i, "compte"], df.at[i, "libelle"] = cible[m]
            n += 1
    return df, n

def fusionner_modele(df):
    """Ajoute à la table les règles du modèle dont le motif n'y figure pas encore (ne modifie rien d'existant)."""
    connus = {norm(m) for m in df["motif"]}
    neuf = [r for r in IMPUT_DEFAUT if norm(r[0]) not in connus]
    return pd.concat([df, pd.DataFrame(neuf, columns=COLS_IMPUT)], ignore_index=True), len(neuf)

def _s(v):
    return "" if v is None or (isinstance(v, float) and pd.isna(v)) else str(v).strip()

def preparer_regles(regles):
    out = [(norm(_s(r["motif"])), _s(r["compte"]), _s(r["libelle"]), norm(_s(r.get("type", "")))) for _, r in regles.iterrows()]
    return sorted([r for r in out if r[0] and r[1]], key=lambda r: -len(r[0]))  # motif le plus précis d'abord

def trouver_regle(tiers, regles, type_=None):
    n = norm(tiers)
    for motif, compte, lib, typ in regles:
        if typ in ("", "tous", "nan") or type_ is None or typ == type_:
            if re.search(rf"(?<![a-z0-9]){re.escape(motif)}(?![a-z0-9])", n):
                return compte, lib
    return None

# ---------- TVA non déductible (achats) ----------
# Règles de DÉPART, INDICATIVES, à valider selon le dossier : par COMPTE de charge (début du numéro de compte),
# % de TVA non déductible, incorporé à la charge. Le carburant est géré à part (véhicule + carburant, voir plus bas).
COLS_ND = ["compte", "pct_non_deductible", "commentaire"]
TVA_ND_DEFAUT = [
    ("625110", 100, "Transport de personnes (train, avion, taxi, VTC) : TVA non déductible (art. 206 IV-2 annexe II CGI)"),
    ("625610", 100, "Hébergement des dirigeants et salariés : TVA non déductible ; déductible pour l'hébergement d'un tiers (client)"),
]

def tva_nd_defaut_df():
    return pd.DataFrame(TVA_ND_DEFAUT, columns=COLS_ND)

# --- Carburant : % de TVA DÉDUCTIBLE selon le véhicule (VP = véhicule de tourisme, VU = utilitaire) et l'énergie ---
VEHICULES = {"VP": "VP · véhicule de tourisme", "VU": "VU · véhicule utilitaire"}
CARBURANTS = {   # énergie -> (déductible VP, déductible VU, remarque)
    "Gazole": (80, 100, ""),
    "Essence (SP95, SP98, E10)": (80, 100, ""),
    "Superéthanol E85": (80, 100, ""),
    "GPL carburant (liquéfié)": (100, 100, "Sources divergentes en VP (80 % ou 100 %) : à valider."),
    "GPL gazeux": (50, 100, "Limité à 50 % en VP selon la plupart des sources : à valider."),
    "GNV (gaz naturel)": (100, 100, ""),
    "Électricité (recharge)": (100, 100, "Hybrides : appliquer l'énergie réellement achetée (essence/gazole ou électricité)."),
}
COMPTE_CARBURANT_DEFAUT = "606130"

def pct_deductible_carburant(vehicule, carburant):
    vp, vu, _ = CARBURANTS.get(carburant, (80, 100, ""))
    return vu if vehicule == "VU" else vp

def regle_carburant(vehicule="VP", carburant="Gazole", comptes=COMPTE_CARBURANT_DEFAUT):
    """Règles de TVA non déductible du carburant, pour chaque compte carburant (séparés par des virgules)."""
    if carburant not in CARBURANTS: return []                  # None / inconnu : pas de règle carburant
    ded = pct_deductible_carburant(vehicule, carburant)
    rem = CARBURANTS.get(carburant, (0, 0, ""))[2]
    com = f"Carburant : {vehicule} · {carburant} · {ded} % déductible" + (f" ({rem})" if rem else "")
    return [(re.sub(r"\W", "", c), Decimal(100 - ded), com)
            for c in str(comptes).replace(";", ",").split(",") if re.sub(r"\W", "", c)] if ded < 100 else []

def preparer_nd(df, vehicule="VP", carburant="Gazole", comptes_carburant=COMPTE_CARBURANT_DEFAUT):
    """Règles par compte (tableau modifiable) + règle carburant. Les lignes sans compte sont ignorées."""
    out = []
    for _, r in df.iterrows():
        p, cpt = montant(_s(r.get("pct_non_deductible"))), re.sub(r"\W", "", _s(r.get("compte")))
        if p is not None and 0 < p <= 100 and cpt:
            out.append((cpt, p, _s(r.get("commentaire"))))
    return out + regle_carburant(vehicule, carburant, comptes_carburant)

def trouver_nd(compte, regles):
    """Règle la plus précise = compte le plus long qui commence le compte de charge (à égalité, la 1re du tableau)."""
    best = None
    for cpt, p, com in regles:
        if compte.startswith(cpt) and (best is None or len(cpt) > len(best[0])): best = (cpt, p, com)
    return (best[1], best[2]) if best else None

SEUIL_CADEAU = 73   # € TTC par bénéficiaire et par an (art. 23 N, annexe IV du CGI) : à revérifier chaque année

def msg_cadeau():
    return (f"Cadeau : TVA déductible seulement si la valeur est ≤ {SEUIL_CADEAU} € TTC par bénéficiaire et par an "
            "(seuil à vérifier pour l'année) ; au-delà, TVA non déductible. Relevé n° 2067 si le total des cadeaux dépasse 3 000 €.")

INFOS = ("TVA non déductible appliquée", "Ventilation TVA lue sur la pièce")      # messages informatifs : non surlignés

def tagger_alertes(ecritures, anomalies):
    """Ajoute à chaque écriture le texte des alertes/erreurs qui la concernent (colonne 'alerte')."""
    msgs = {}
    for ref, msg, detail in anomalies:
        if msg.startswith(INFOS): continue
        f = str(ref).split(":")[0]
        for e in ecritures:
            if (detail == e["piece"] and f == e["fichier"]) or (detail and detail == e["tiers"]) or ref == e["ref"]:
                lst = msgs.setdefault(e["num"], [])
                if msg not in lst: lst.append(msg)
    for e in ecritures:
        e["alerte"] = " | ".join(msgs.get(e["num"], []))

# ---------- Contrôles + écritures ----------
FORMES_DEBUT = {"sarl", "sas", "sasu", "eurl", "sa", "sci", "snc", "scop", "selarl", "earl", "gaec", "ets", "etablissements",
                "ste", "societe", "mr", "mme", "monsieur", "madame", "m", "le", "la", "les", "l"}

def code_tiers(nom, type_):
    """Compte de tiers sur 7 caractères : C + 6 caractères du nom (clients) ou F + 6 (fournisseurs)."""
    jetons = norm(nom).split()
    while len(jetons) > 1 and jetons[0] in FORMES_DEBUT: jetons.pop(0)       # SARL, SCI, M., Le... ignorés en tête
    base = re.sub(r"[^A-Z0-9]", "", "".join(jetons).upper())[:6] or "DIVERS"
    return ("C" if type_ == "vente" else "F") + base

def code_tiers_unique(nom, type_, codes, anomalies, ref):
    """Évite que deux tiers différents reçoivent le même compte (suffixe 1, 2, ... à la place du dernier caractère)."""
    base, cle, cand, k = code_tiers(nom, type_), norm(nom), None, 0
    cand = base
    while codes.get(cand, cle) != cle and k < 99:
        k += 1
        cand = base[:7 - len(str(k))] + str(k)
    if cand != base:
        anomalies.append((ref, f"Compte tiers {base} déjà pris par un autre tiers : {cand} attribué", nom))
    codes[cand] = cle
    return cand

def construire_ecritures(lignes, type_, anomalies, regles=None, comptes=None, tva_comptes=None, produits_tva=None,
                         tva_nd=None, tiers_au_compte=True):
    cfg = (comptes or COMPTES)[type_]
    tvac = tva_comptes or TVA_COMPTES
    regles, signales, codes, tronques, nd_regles = regles or [], set(), {}, set(), tva_nd or []
    prod = produits_tva or PRODUITS_TVA
    pref = next(iter(prod.values()))[0][:3] if prod else "706"
    factures = OrderedDict()
    for l in lignes:
        factures.setdefault((l["numero"], l["tiers"]), []).append(l)
    vus, ecritures, num = set(), [], 0
    idx = 0 if type_ == "vente" else 1
    for (numero, tiers), ls in factures.items():
        if not numero: anomalies.append((ls[0]["n"], "N° de facture vide", tiers))
        cle = (numero, tiers, sum(x["ttc"] or 0 for x in ls))
        if numero and cle in vus: anomalies.append((ls[0]["n"], "Doublon possible", numero))
        vus.add(cle)
        d = ls[0]["date"]
        src, page = ls[0]["n"].rsplit(":", 1)[0], ls[0].get("page", "")
        if ls[0].get("pj"): src, page = ls[0]["pj"], ls[0].get("pj_page", "")
        ht_tot = sum(x["ht"] for x in ls); tva_par_taux = OrderedDict()
        for x in ls: tva_par_taux[x["taux"]] = tva_par_taux.get(x["taux"], Decimal("0")) + x["tva"]
        tva_tot = sum(tva_par_taux.values()); ttc_calc = ht_tot + tva_tot
        ttcs = [x["ttc"] for x in ls]
        ttc_src = sum(ttcs) if all(v is not None for v in ttcs) else None
        if ttc_src is not None and abs(ttc_src - ttc_calc) > Decimal("0.02") \
           and not (len(ls) > 1 and abs(ttcs[0] - ttc_calc) <= Decimal("0.02")):
            anomalies.append((ls[0]["n"], f"HT+TVA ({ttc_calc}) <> TTC ({ttc_src})", numero))
        for t in tva_par_taux:
            if t not in tvac and t not in ("0", "0.0"):
                anomalies.append((ls[0]["n"], f"Taux de TVA inhabituel : {t}%", numero))
        avoir = ttc_calc < 0
        num += 1
        lib = f"{numero} {tiers}".strip()[:100]
        code = code_tiers_unique(tiers, type_, codes, anomalies, ls[0]["n"])
        def ecr(compte, libc, aux_n, aux_l, montant_, sens):
            if montant_ == 0: return
            if len(compte) > 7:                                  # comptes limités à 7 caractères
                if compte not in tronques:
                    tronques.add(compte); anomalies.append((ls[0]["n"], f"Compte {compte} tronqué à 7 caractères : {compte[:7]}", numero))
                compte = compte[:7]
            sens = sens if not avoir else ("C" if sens == "D" else "D")
            ecritures.append(dict(num=num, src=src, page=page, tiers=tiers, fichier=ls[0]["n"].rsplit(":", 1)[0], ref=ls[0]["n"], date=d, compte=compte, libc=libc, aux=aux_n, auxl=aux_l,
                                  piece=numero, lib=lib, debit=abs(montant_) if sens == "D" else Decimal("0"),
                                  credit=abs(montant_) if sens == "C" else Decimal("0")))
        st = cfg["sens_tiers"]; sc = "C" if st == "D" else "D"
        if tiers_au_compte:                                      # compte de tiers directement au niveau du compte
            ecr(code, tiers, "", "", ttc_calc, st)
        else:                                                    # collectif + compte auxiliaire
            ecr(cfg["tiers"], "Clients" if type_ == "vente" else "Fournisseurs", code, tiers, ttc_calc, st)
        charges, deduc = OrderedDict(), OrderedDict()           # une ligne par compte de charge/produit, une par taux de TVA
        for x in ls:
            cpt, libc = x["compte"], cfg["libelle_cp"]
            if not cpt:
                reg = trouver_regle(tiers, regles, type_)
                if reg: cpt, libc = reg[0], (reg[1] or libc)
                elif type_ == "vente":                           # compte 7 par taux de TVA
                    cpt, libc = prod.get(x["taux"]) or produits_par_taux(pref, [x["taux"]])[x["taux"]]
                else:
                    cpt = cfg["charge_produit"]
                    if type_ == "achat" and tiers not in signales:
                        signales.add(tiers)
                        anomalies.append((x["n"], f"Aucune règle d'imputation : compte par défaut {cpt}", tiers))
            if type_ == "achat" and (cpt.startswith("6234") or re.search(r"cadeau|coffret|\bgift\b|bon d achat", norm(tiers + " " + x.get("libelle", "")))) \
               and not any(a[1] == msg_cadeau() and a[2] == numero for a in anomalies):
                anomalies.append((x["n"], msg_cadeau(), numero))
            nd = Decimal("0")
            if type_ == "achat" and nd_regles:                   # TVA non déductible : incorporée à la charge
                rg = trouver_nd(cpt, nd_regles)
                if rg:
                    nd = (x["tva"] * rg[0] / 100).quantize(Decimal("0.01"), ROUND_HALF_UP)
                    if (tiers, "nd") not in signales:
                        signales.add((tiers, "nd"))
                        anomalies.append((x["n"], f"TVA non déductible appliquée : {rg[0]} % ({rg[1] or 'règle'})", tiers))
            charges[(cpt, libc)] = charges.get((cpt, libc), Decimal("0")) + x["ht"] + nd
            deduc[x["taux"]] = deduc.get(x["taux"], Decimal("0")) + x["tva"] - nd
        for (cpt, libc), v in charges.items():
            ecr(cpt, libc, "", "", v, sc)
        for t, v in deduc.items():
            ecr(tvac.get(t, TVA_INCONNUE)[idx], TVA_LIB, "", "", v, sc)
    return ecritures

# ---------- Export FEC ----------
FEC_COLS = ["JournalCode","JournalLib","EcritureNum","EcritureDate","CompteNum","CompteLib","CompAuxNum",
            "CompAuxLib","PieceRef","PieceDate","EcritureLib","Debit","Credit","EcritureLet","DateLet",
            "ValidDate","Montantdevise","Idevise"]

def ecrire_fec(ecritures, type_, siren, exercice, sortie, comptes=None, separateur="tab"):
    jc, jl = (comptes or COMPTES)[type_]["journal"]
    fmt = lambda x: f"{x:.2f}".replace(".", ",")
    d_tot = sum(e["debit"] for e in ecritures); c_tot = sum(e["credit"] for e in ecritures)
    if d_tot != c_tot: raise ValueError(f"FEC déséquilibré : D={d_tot} C={c_tot}")
    fin = max(e["date"] for e in ecritures).strftime("%Y%m%d")
    nom = os.path.join(sortie, f"{siren}FEC{exercice}1231.txt") if exercice else os.path.join(sortie, f"{siren}FEC{fin}.txt")
    with open(nom, "w", encoding="utf-8", newline="") as f:
        sep = "|" if separateur == "pipe" else "\t"
        f.write(sep.join(FEC_COLS) + "\r\n")
        for e in ecritures:
            dd = e["date"].strftime("%Y%m%d")
            f.write(sep.join([jc, jl, str(e["num"]), dd, e["compte"], e["libc"], e["aux"], e["auxl"],
                               e["piece"], dd, e["lib"], fmt(e["debit"]), fmt(e["credit"]),
                               "", "", dd, "", ""]) + "\r\n")
    return nom, d_tot

# ---------- ZIP et rapprochement factures <-> pièces justificatives ----------
EXT_OK = {".csv", ".xlsx", ".xls", ".xlsm", ".pdf", ".png", ".jpg", ".jpeg", ".xml"}
EXT_TAB = (".csv", ".xlsx", ".xls", ".xlsm")
EXT_DOC = (".pdf", ".png", ".jpg", ".jpeg")

def extraire_zip(chemin_zip, dest, max_fichiers=2000, max_octets=500 * 1024 * 1024):
    """Extrait à plat les fichiers utiles d'un zip (sans dossiers, sans fichiers cachés), noms rendus uniques."""
    os.makedirs(dest, exist_ok=True)
    sortie, total = [], 0
    with zipfile.ZipFile(chemin_zip) as z:
        for info in z.infolist():
            nom = os.path.basename(info.filename.replace("\\", "/"))
            if info.is_dir() or not nom or nom.startswith((".", "~$")) or "__MACOSX" in info.filename: continue
            base, ext = os.path.splitext(nom)
            if ext.lower() not in EXT_OK: continue
            total += info.file_size
            if len(sortie) >= max_fichiers or total > max_octets:
                raise ValueError("Archive trop volumineuse (plus de 2000 fichiers ou 500 Mo).")
            cible, k = os.path.join(dest, nom), 2
            while os.path.exists(cible):
                cible = os.path.join(dest, f"{base}_{k}{ext}"); k += 1
            with z.open(info) as src, open(cible, "wb") as out:
                shutil.copyfileobj(src, out)
            sortie.append(cible)
    return sortie

def classer(chemins):
    """(fichiers comptables CSV/Excel, pièces PDF/images, XML de factures)"""
    low = lambda p: p.lower()
    return ([p for p in chemins if low(p).endswith(EXT_TAB)], [p for p in chemins if low(p).endswith(EXT_DOC)],
            [p for p in chemins if low(p).endswith(".xml")])

_IDX = {}
def _indexer_piece(chemin):
    cle = (chemin, os.path.getmtime(chemin))
    if cle not in _IDX:
        pages = []
        if chemin.lower().endswith(".pdf"):
            v = textes_pdf(chemin)
            pages = v[0] if v else []
        texte = _norm_txt("\n".join(pages))
        if chemin.lower().endswith(".pdf"):                      # Factur-X : le n° et les montants sont dans le XML intégré
            try:
                from pypdf import PdfReader
                for n, c in PdfReader(chemin).attachments.items():
                    if n.lower().endswith(".xml"):
                        texte += "\n" + re.sub(r"<[^>]+>", " ", c[0].decode("utf-8", errors="replace"))
            except Exception:
                pass
        _IDX[cle] = {"chemin": chemin, "nom": os.path.basename(chemin), "pages": pages, "multi": len(pages) > 1,
                     "texte": texte,
                     "montants": {abs(m) for m in (montant(x) for x in re.findall(NB, texte)) if m is not None},
                     "dates": {_conv_date(r.group(0)) for r in D_ALL.finditer(texte)} - {""},
                     "jetons": set(norm(texte + " " + os.path.basename(chemin)).split())}
    return _IDX[cle]

def _re_numero(numero):
    runs = re.findall(r"[A-Za-z0-9]+", numero or "")
    if not runs: return None
    return re.compile(r"(?<![A-Za-z0-9])" + r"[\s\-_/.]*".join(re.escape(r) for r in runs) + r"(?![A-Za-z0-9])", re.I)

def _score_piece(f, info):
    s, met = 0.0, []
    rx, fort = f["rx"], len(re.sub(r"[^A-Za-z0-9]", "", f["numero"])) >= 4
    if rx and rx.search(info["nom"]): s += 5 if fort else 2; met.append("n° dans le nom du fichier")
    if rx and rx.search(info["texte"]): s += 4 if fort else 1.5; met.append("n° dans la facture")
    if f["montants"] & info["montants"]: s += 2; met.append("montant")
    jt = {t for t in norm(f["tiers"]).split() if len(t) >= 3 and t not in FORMES}
    if jt and jt <= info["jetons"]: s += 1.5; met.append("tiers")
    if f["date"] in info["dates"]: s += 1; met.append("date")
    return s, met

def _ventilation_piece(info, page, ht, tva):
    """Ventilation de TVA par taux lue dans la pièce (XML Factur-X, sinon texte du PDF), si elle retombe sur les totaux
    de l'écriture. Renvoie [(taux, base HT, TVA)] avec au moins 2 taux, ou None."""
    if info["chemin"].lower().endswith(".pdf") and not info["multi"]:
        df = lire_facturx(info["chemin"], "achat")
        if df is not None and len(df) >= 2:
            vent = [(float(r["Taux TVA"] or 0), abs(montant(r["Total HT"]) or Decimal("0")), abs(montant(r["Total TVA"]) or Decimal("0")))
                    for _, r in df.iterrows()]
            if abs(sum(b for _, b, _ in vent) - abs(ht)) <= Decimal("0.05") and abs(sum(v for _, _, v in vent) - abs(tva)) <= Decimal("0.05"):
                return vent
    texte = info["pages"][int(page) - 1] if page and info["pages"] else info["texte"]
    lignes = [l.strip() for l in _norm_txt(texte).splitlines() if l.strip()]
    A, B = _taux_lignes(lignes)
    vent = _ventiler([list(OrderedDict.fromkeys(A)), A, B], ht, tva)
    return vent if vent and len(vent) >= 2 else None

def rapprocher(lignes, pieces, manuels=None, ventiler=True):
    """Associe chaque facture du fichier comptable à sa pièce (n° de facture dans le nom ou le texte du PDF,
    sinon montant + tiers + date). Renseigne l['pj'] / l['pj_page'] et renvoie le détail du rapprochement."""
    infos = [_indexer_piece(p) for p in pieces]
    par_nom = {i["nom"]: i for i in infos}
    facts = OrderedDict()
    for l in lignes:
        cle = f"{l['n'].rsplit(':', 1)[0]}|{l['numero']}"
        f = facts.setdefault(cle, {"cle": cle, "numero": l["numero"], "tiers": l["tiers"], "date": l["date"].strftime("%d/%m/%Y"),
                                   "ht": Decimal("0"), "tva": Decimal("0"), "ttcs": [], "rx": _re_numero(l["numero"])})
        f["ht"] += l["ht"]; f["tva"] += l["tva"]; f["ttcs"].append(l["ttc"])
    for f in facts.values():
        f["ttc"] = f["ht"] + f["tva"]
        f["montants"] = {abs(f["ttc"])} | {abs(x) for x in f["ttcs"] if x is not None}
    cands = {}
    for k, f in facts.items():
        cands[k] = sorted(((sc, i["nom"], met) for i in infos for sc, met in [_score_piece(f, i)] if sc >= 3), key=lambda x: -x[0])
    res, utilises = {}, set()
    for k in sorted(cands, key=lambda k: -(cands[k][0][0] if cands[k] else 0)):
        liste = cands[k]
        if not liste: continue
        if len(liste) > 1 and liste[0][0] == liste[1][0] and liste[0][1] != liste[1][1]:
            res[k] = ("", 0, ["ambigu : plusieurs pièces possibles"]); continue
        for sc, nom, met in liste:
            if par_nom[nom]["multi"] or nom not in utilises:
                res[k] = (nom, sc, met); utilises.add(nom); break
    for k, nom in (manuels or {}).items():                      # choix manuels de l'utilisateur
        if k in facts:
            res[k] = (nom, 99, ["choix manuel"]) if nom else ("", 0, ["aucune (choix manuel)"])
    pages = {}
    for k, (nom, sc, met) in res.items():                       # PDF multi-factures : page de chaque facture
        if nom and par_nom[nom]["multi"] and sum(1 for v in res.values() if v[0] == nom) >= 2:
            rx = facts[k]["rx"]
            pages[k] = next((str(i) for i, p in enumerate(par_nom[nom]["pages"], 1) if rx and rx.search(p)), "")
    cle_de = lambda l: f"{l['n'].rsplit(':', 1)[0]}|{l['numero']}"
    ventile = {}
    if ventiler:                                              # multi-taux : ventilation relue sur la pièce rapprochée
        for k, (nom, sc, met) in res.items():
            if not nom or k not in facts: continue
            inv = [l for l in lignes if cle_de(l) == k]
            if not inv or len({l["taux"] for l in inv}) != 1 or len({l["compte"] for l in inv}) != 1: continue
            vent = _ventilation_piece(par_nom[nom], pages.get(k, ""), facts[k]["ht"], facts[k]["tva"])
            if not vent: continue
            sg = -1 if facts[k]["ht"] < 0 else 1
            neuves = [dict(inv[0], ht=sg * b, tva=sg * v, ttc=sg * (b + v), taux=taux_str(r)) for r, b, v in vent]
            pos = lignes.index(inv[0])
            lignes[:] = [l for l in lignes if cle_de(l) != k]
            lignes[pos:pos] = neuves
            ventile[k] = " + ".join(f"{taux_str(r)} %" for r, _, _ in vent)
    for l in lignes:
        k = cle_de(l)
        if res.get(k, ("",))[0]:
            l["pj"], l["pj_page"] = res[k][0], pages.get(k, "")
    return [{"cle": k, "numero": f["numero"], "tiers": f["tiers"], "ttc": str(f["ttc"]),
             "piece": res.get(k, ("",))[0], "page": pages.get(k, ""), "ventilation": ventile.get(k, ""),
             "methode": ", ".join(res[k][2]) if k in res else "aucune pièce trouvée"} for k, f in facts.items()]

def creer_archive(ecritures, fec, sources, type_, comptes=None):
    """Zip : FEC + dossier pieces_jointes (1 fichier par facture, nommé JOURNAL_NUMECRITURE_REFPIECE) + index."""
    jc = (comptes or COMPTES)[type_]["journal"][0]
    factures = OrderedDict()
    for e in ecritures: factures.setdefault(e["num"], e)
    tmp, index, pieces = tempfile.mkdtemp(), [], []
    for num, e in factures.items():
        src = sources.get(e.get("src", ""))
        ext = os.path.splitext(src)[1].lower() if src else ""
        if not src or ext not in (".pdf", ".png", ".jpg", ".jpeg", ".xml"): continue
        ref = re.sub(r"[^A-Za-z0-9._-]+", "_", e["piece"] or f"ecr{num}")[:40]
        nom = f"{jc}_{num:05d}_{ref}{ext}"
        chemin = os.path.join(tmp, nom)
        if ext == ".pdf" and e.get("page"):
            from pypdf import PdfReader, PdfWriter
            w = PdfWriter(); w.add_page(PdfReader(src).pages[int(e["page"]) - 1])
            with open(chemin, "wb") as fh: w.write(fh)
        else:
            shutil.copyfile(src, chemin)
        pieces.append((nom, chemin))
        index.append([jc, num, e["piece"], e["date"].strftime("%Y%m%d"), e["aux"] or e["compte"],
                      e["src"] + (f" p.{e['page']}" if e.get("page") else ""), nom])
    if not pieces: return None
    idx = os.path.join(tmp, "index_pieces.csv")
    pd.DataFrame(index, columns=["JournalCode", "EcritureNum", "PieceRef", "EcritureDate", "CompteTiers",
                                 "Fichier source", "Pièce jointe"]).to_csv(idx, index=False, sep=";")
    lisez = os.path.join(tmp, "LISEZMOI.txt")
    open(lisez, "w", encoding="utf-8").write(
        "Contenu : le FEC, le dossier pieces_jointes et index_pieces.csv.\n"
        "Chaque pièce porte le nom JOURNAL_NUMEROECRITURE_REFERENCE : le numéro d'écriture est celui de la colonne "
        "EcritureNum du FEC. index_pieces.csv fait le lien écriture <-> pièce.\n"
        "Le FEC lui-même ne peut pas contenir de pièce jointe : le rattachement dans le logiciel comptable "
        "dépend de ses fonctions d'import.\n")
    zip_path = os.path.splitext(fec)[0] + "_avec_pieces.zip"
    with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as z:
        z.write(fec, os.path.basename(fec))
        for nom, chemin in pieces: z.write(chemin, f"pieces_jointes/{nom}")
        z.write(idx, "index_pieces.csv"); z.write(lisez, "LISEZMOI.txt")
    return zip_path

def valider_ecritures(ecr):
    """Contrôles d'une liste d'écritures : renvoie [(n° d'écriture, message, bloquant)]."""
    pb, soldes = [], OrderedDict()
    for e in ecr:
        if not e["compte"]: pb.append((e["num"], "Compte vide", True))
        elif len(e["compte"]) > 7: pb.append((e["num"], f"Compte {e['compte']} de plus de 7 caractères", False))
        if e["debit"] < 0 or e["credit"] < 0: pb.append((e["num"], "Montant négatif", True))
        if e["debit"] and e["credit"]: pb.append((e["num"], "Débit et crédit sur la même ligne", False))
        soldes[e["num"]] = soldes.get(e["num"], Decimal("0")) + e["debit"] - e["credit"]
    pb += [(n, f"Écriture déséquilibrée : écart de {v}", True) for n, v in soldes.items() if v != 0]
    return pb

def ecritures_depuis_df(df, originaux):
    """Reconstruit les écritures à partir du tableau modifié par l'utilisateur."""
    par_num = {}
    for e in originaux: par_num.setdefault(e["num"], e)
    out = []
    for i, (_, r) in enumerate(df.iterrows(), 1):
        deb, cre = montant(_s(r["debit"])) or Decimal("0"), montant(_s(r["credit"])) or Decimal("0")
        if not _s(r["compte"]) and not deb and not cre: continue                 # ligne vide : ignorée
        try: num = int(float(r["num"]))
        except (TypeError, ValueError):
            raise ValueError(f"ligne {i} du tableau : indiquez le n° d'écriture auquel elle appartient")
        o = par_num.get(num, {})
        try: d = pd.to_datetime(r["date"]).to_pydatetime()
        except Exception: d = o.get("date") or datetime.now()
        if pd.isna(d): d = o.get("date") or datetime.now()
        out.append(dict(num=num, src=o.get("src", ""), page=o.get("page", ""), tiers=o.get("tiers", ""),
                        fichier=o.get("fichier", ""), ref=o.get("ref", ""), date=d, compte=_s(r["compte"]), libc=_s(r["libc"]),
                        aux=_s(r["aux"]), auxl=_s(r.get("auxl", "")), piece=_s(r["piece"]) or o.get("piece", ""),
                        lib=_s(r["lib"]), debit=deb, credit=cre, alerte=_s(r.get("alerte", ""))))
    return sorted(out, key=lambda e: e["num"])

def regenerer(df, originaux, type_, siren, exercice, sortie, comptes=None, separateur="tab", sources=None):
    """FEC (et zip des pièces) recalculés à partir des écritures modifiées à la main."""
    ecr = ecritures_depuis_df(df, originaux)
    bloquants = [f"écriture {n} : {m}" for n, m, b in valider_ecritures(ecr) if b]
    if bloquants:
        raise ValueError("Modifications refusées — " + " ; ".join(bloquants[:5]))
    fec, total = ecrire_fec(ecr, type_, siren, exercice, sortie, comptes, separateur)
    return dict(ecritures=ecr, fec=fec, zip=creer_archive(ecr, fec, sources or {}, type_, comptes), total=total)

def traiter(fichiers, type_="vente", siren="000000000", exercice="", sortie=".", regles=None,
            mappings=None, comptes=None, tva_comptes=None, separateur="tab", tiers_corriges=None, pieces=None, rapprochements=None, produits_tva=None, tva_nd=None, tiers_au_compte=True, ventilations=None,
            vehicule="VP", carburant="Gazole", comptes_carburant=COMPTE_CARBURANT_DEFAUT, ventiler_pieces=True):
    anomalies, lignes = [], []
    for f in fichiers:
        nom = os.path.basename(f)
        try:
            df = lire(f, type_)
        except ValueError as e:
            anomalies.append((nom, str(e), "")); continue
        anomalies += [(nom, a, b) for a, b in df.attrs.get("anomalies", [])]
        if df.empty:
            anomalies.append((nom, "Aucune facture extraite", "")); continue
        lignes += vers_pivot(df, (mappings or {}).get(nom) or charger_mapping(df, type_=type_), anomalies, nom)
    if not lignes:
        raise ValueError("Aucune écriture exploitable : " + " ; ".join(f"{a[0]} - {a[1]}" for a in anomalies[:3]))
    for l in lignes:                                         # noms de tiers corrigés dans l'app
        cle = f"{l['n'].rsplit(':', 1)[0]}|{l['numero']}"
        if (tiers_corriges or {}).get(cle):
            l["tiers"], l["code"] = tiers_corriges[cle], ""
    cle_de = lambda l: f"{l['n'].rsplit(':', 1)[0]}|{l['numero']}"
    for cle, rows in (ventilations or {}).items():            # ventilation de TVA saisie à la main (multi-taux)
        inv = [l for l in lignes if cle_de(l) == cle]
        if not inv or not rows: continue
        modele, ancien, neuves = inv[0], sum(l["ht"] + l["tva"] for l in inv), []
        for rw in rows:
            h, v, t = montant(str(rw.get("ht"))), montant(str(rw.get("tva"))), montant(str(rw.get("taux")))
            if h is None or t is None: continue
            neuves.append(dict(modele, ht=h, tva=v if v is not None else (h * t / 100).quantize(Decimal("0.01"), ROUND_HALF_UP),
                               taux=taux_str(t)))
        if not neuves: continue
        for l in neuves: l["ttc"] = l["ht"] + l["tva"]
        nouveau = sum(l["ttc"] for l in neuves)
        if abs(nouveau - ancien) > Decimal("0.05"):
            anomalies.append((cle.split("|")[0], f"Ventilation TVA : total saisi {nouveau} <> total facture {ancien}", modele["numero"]))
        pos = lignes.index(modele)
        lignes[:] = [l for l in lignes if cle_de(l) != cle]
        lignes[pos:pos] = neuves
        anomalies[:] = [a for a in anomalies if not (f"{a[0]}|{a[2]}" == cle and a[1].startswith("Taux de TVA calculé"))]
    multi = []
    for a in anomalies:                                       # factures dont les taux n'ont pas pu être ventilés
        if a[1].startswith("Taux de TVA calculé") and not any(m["cle"] == f"{a[0]}|{a[2]}" for m in multi):
            inv = [l for l in lignes if cle_de(l) == f"{a[0]}|{a[2]}"]
            if inv:
                multi.append({"cle": f"{a[0]}|{a[2]}", "numero": a[2], "tiers": inv[0]["tiers"],
                              "ht": str(sum(l["ht"] for l in inv)), "tva": str(sum(l["tva"] for l in inv))})
    rapp = rapprocher(lignes, pieces, rapprochements, ventiler_pieces) if pieces else []
    for rr in rapp:
        if rr.get("ventilation"):
            anomalies.append((rr["cle"].split("|")[0], f"Ventilation TVA lue sur la pièce : {rr['ventilation']}", rr["numero"]))
        if not rr["piece"]:
            anomalies.append((rr["cle"].split("|")[0], "Pièce justificative non trouvée (" + rr["methode"] + ")", rr["numero"]))
    factures = list(OrderedDict(((f"{l['n'].rsplit(':', 1)[0]}|{l['numero']}"),
                                 {"cle": f"{l['n'].rsplit(':', 1)[0]}|{l['numero']}", "fichier": l["n"].rsplit(":", 1)[0],
                                  "numero": l["numero"], "tiers": l["tiers"]}) for l in lignes).values())
    ecr = construire_ecritures(lignes, type_, anomalies,
                               preparer_regles(regles if regles is not None else charger_imputations()),
                               comptes, tva_comptes, produits_tva,
                               preparer_nd(tva_nd if tva_nd is not None else tva_nd_defaut_df(), vehicule, carburant, comptes_carburant),
                               tiers_au_compte)
    tagger_alertes(ecr, anomalies)
    fec, total = ecrire_fec(ecr, type_, siren, exercice, sortie, comptes, separateur)
    sources = {os.path.basename(f): f for f in list(fichiers) + list(pieces or [])}
    archive = creer_archive(ecr, fec, sources, type_, comptes)
    rapport = None
    if anomalies:
        rapport = os.path.join(sortie, "rapport_anomalies.csv")
        pd.DataFrame(anomalies, columns=["ligne_source", "anomalie", "detail"]).to_csv(rapport, index=False, sep=";")
    sans = sorted({a[2] for a in anomalies if a[1].startswith("Aucune règle d'imputation")})
    return dict(sources=sources, multi_taux=multi, rapprochements=rapp, factures=factures, fec=fec, zip=archive, ecritures=ecr, anomalies=anomalies, rapport=rapport, total=total, sans_regle=sans)

def main():
    p = argparse.ArgumentParser()
    p.add_argument("fichiers", nargs="+"); p.add_argument("--siren", default="000000000")
    p.add_argument("--exercice", default=""); p.add_argument("--type", choices=["vente", "achat"], default="vente")
    p.add_argument("--sortie", default=".")
    p.add_argument("--imputations", default="imputations.csv")
    p.add_argument("--separateur", choices=["tab", "pipe"], default="tab")
    a = p.parse_args()
    tous = []
    for f in a.fichiers:
        tous += extraire_zip(f, tempfile.mkdtemp()) if f.lower().endswith(".zip") else [f]
    tab, docs, xmls = classer(tous)
    fichiers, pieces = (tab + xmls, docs) if tab and docs else (tous, None)
    try:
        r = traiter(fichiers, a.type, a.siren, a.exercice, a.sortie, charger_imputations(a.imputations),
                    separateur=a.separateur, pieces=pieces)
    except ValueError as e:
        sys.exit(str(e))
    print(f"FEC généré : {r['fec']}  ({len(r['ecritures'])} lignes, débit = crédit = {r['total']})")
    if r["zip"]: print(f"FEC + pièces jointes : {r['zip']}")
    if r["sans_regle"]: print("Sans règle d'imputation (à ajouter dans imputations.csv) : " + ", ".join(r["sans_regle"]))
    print(f"{len(r['anomalies'])} anomalie(s) : {r['rapport']}" if r["anomalies"] else "Aucune anomalie.")

if __name__ == "__main__":
    main()
