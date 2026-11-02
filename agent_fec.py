#!/usr/bin/env python3
"""
Agent factures -> FEC (importable dans ACD Compta Expert).

Usage :
  python agent_fec.py export.csv --siren 123456789 --exercice 2026
  python agent_fec.py export.xlsx facture1.pdf scan.jpg --siren 123456789 --type achat

Étapes : lecture -> mapping auto des colonnes -> format pivot -> contrôles -> FEC + rapport.
Le mapping est mémorisé dans mappings/<empreinte>.json (modifiable à la main).
Si ANTHROPIC_API_KEY est défini et le package `anthropic` installé,
l'IA est consultée pour les colonnes que les règles n'ont pas reconnues.
"""
import argparse, hashlib, json, os, re, sys, unicodedata
from collections import OrderedDict
from decimal import Decimal, ROUND_HALF_UP
from datetime import datetime
import pandas as pd

# ---------- Paramétrage comptable (à adapter au dossier) ----------
COMPTES = {
    "vente": {"journal": ("VE", "Ventes"), "tiers": "411", "charge_produit": "706000",
              "libelle_cp": "Prestations de services", "sens_tiers": "D"},
    "achat": {"journal": ("AC", "Achats"), "tiers": "401", "charge_produit": "607000",
              "libelle_cp": "Achats de marchandises", "sens_tiers": "C"},
}
TVA_COMPTES = {  # taux -> (compte ventes, compte achats)
    "20": ("44571200", "44566200"), "10": ("44571100", "44566100"),
    "5.5": ("44571055", "44566055"), "2.1": ("44571021", "44566021"),
}
TVA_LIB = "TVA"

# ---------- Mapping des colonnes ----------
SYNONYMES = {
    "date": ["date facture", "date de facture", "date", "date emission", "date piece", "invoice date", "issue date"],
    "numero": ["numero facture", "n facture", "n piece", "numero", "reference", "ref", "invoice number", "invoice no", "num"],
    "tiers": ["client", "fournisseur", "tiers", "raison sociale", "nom client", "customer", "supplier", "societe"],
    "code_tiers": ["code client", "code tiers", "compte client", "code fournisseur", "compte tiers", "customer code"],
    "ht": ["total ht", "montant ht", "ht", "hors taxe", "net ht", "subtotal", "amount excl"],
    "tva": ["total tva", "montant tva", "tva", "taxe", "vat", "tax amount"],
    "ttc": ["total ttc", "montant ttc", "ttc", "total", "net a payer", "amount incl", "total amount"],
    "taux_tva": ["taux tva", "taux de tva", "taux", "% tva", "vat rate", "tax rate"],
    "libelle": ["libelle", "designation", "description", "objet"],
    "compte": ["compte produit", "compte comptable", "compte"],
}
OBLIGATOIRES = ["date", "numero", "tiers"]

def norm(s):
    s = unicodedata.normalize("NFKD", str(s)).encode("ascii", "ignore").decode().lower()
    return re.sub(r"[^a-z0-9% ]+", " ", s).strip()

def mapping_regles(colonnes):
    res, pris = {}, set()
    cols = {c: norm(c) for c in colonnes}
    for champ, syns in SYNONYMES.items():           # 1) correspondance exacte
        for c, n in cols.items():
            if c not in pris and n in syns:
                res[champ] = c; pris.add(c); break
    for champ, syns in SYNONYMES.items():           # 2) correspondance partielle
        if champ in res: continue
        for s in syns:
            for c, n in cols.items():
                if c not in pris and re.search(rf"\b{re.escape(s)}\b", n):
                    res[champ] = c; pris.add(c); break
            if champ in res: break
    return res

def mapping_ia(df, deja):
    """Demande à Claude de deviner les champs manquants (optionnel)."""
    manquants = [c for c in SYNONYMES if c not in deja]
    if not manquants or not os.getenv("ANTHROPIC_API_KEY"):
        return {}
    try:
        import anthropic
        extrait = df.head(5).to_csv(index=False)
        rep = anthropic.Anthropic().messages.create(
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

def charger_mapping(df, dossier="mappings"):
    empreinte = hashlib.md5("|".join(map(str, df.columns)).encode()).hexdigest()[:10]
    chemin = os.path.join(dossier, f"{empreinte}.json")
    if os.path.exists(chemin):
        print(f"Mapping mémorisé : {chemin}")
        return json.load(open(chemin, encoding="utf-8"))
    m = mapping_regles(df.columns)
    m.update(mapping_ia(df, m))
    os.makedirs(dossier, exist_ok=True)
    json.dump(m, open(chemin, "w", encoding="utf-8"), indent=2, ensure_ascii=False)
    print(f"Mapping détecté et enregistré : {chemin}")
    return m

# ---------- Lecture & conversions ----------
COLONNES_STD = ["Date facture", "N° facture", "Client", "Total HT", "Taux TVA", "Total TVA", "Total TTC"]

def lire_pdf(fichier, type_="vente"):
    """Extrait les factures d'un PDF/image via Claude. Nécessite ANTHROPIC_API_KEY."""
    if not os.getenv("ANTHROPIC_API_KEY"):
        raise ValueError("Lecture PDF/image : définissez la variable ANTHROPIC_API_KEY.")
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
    rep = anthropic.Anthropic().messages.create(
        model="claude-sonnet-4-6", max_tokens=4000,
        messages=[{"role": "user", "content": [bloc, {"type": "text", "text": consigne}]}])
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

def lire(fichier, type_="vente"):
    if fichier.lower().endswith((".pdf", ".png", ".jpg", ".jpeg")):
        return lire_pdf(fichier, type_)
    if fichier.lower().endswith((".xlsx", ".xls", ".xlsm")):
        return pd.read_excel(fichier, dtype=str)
    for enc in ("utf-8-sig", "cp1252", "latin-1"):
        try:
            return pd.read_csv(fichier, sep=None, engine="python", dtype=str, encoding=enc)
        except UnicodeDecodeError:
            continue
    raise ValueError("Encodage illisible")

def montant(v):
    if v is None or (isinstance(v, float) and pd.isna(v)) or str(v).strip() in ("", "nan"):
        return None
    s = re.sub(r"[^\d,.\-]", "", str(v).replace("\u202f", "").replace(" ", ""))
    if s.startswith("(") : s = "-" + s
    if "," in s and "." in s:
        s = s.replace(",", "") if s.rfind(".") > s.rfind(",") else s.replace(".", "").replace(",", ".")
    else:
        s = s.replace(",", ".")
    try: return Decimal(s).quantize(Decimal("0.01"), ROUND_HALF_UP)
    except Exception: return None

def date(v):
    s = str(v).strip()[:10]
    for f in ("%d/%m/%Y", "%Y-%m-%d", "%d-%m-%Y", "%d.%m.%Y", "%d/%m/%y", "%Y/%m/%d"):
        try: return datetime.strptime(s, f)
        except ValueError: pass
    return None

def taux_str(t):
    return str(t).rstrip("0").rstrip(".") if "." in str(t) else str(t)

# ---------- Format pivot ----------
def vers_pivot(df, m, anomalies, src=""):
    for c in OBLIGATOIRES:
        if c not in m:
            raise ValueError(f"Colonne obligatoire non détectée : '{c}'. Complétez le fichier mappings/*.json")
    lignes = []
    for i, r in df.iterrows():
        n = f"{src}:{i + 2}"
        d = date(r[m["date"]])
        if d is None:
            anomalies.append((n, "Date illisible", r[m["date"]])); continue
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
        g = lambda k: (str(r[m[k]]).strip() if k in m and pd.notna(r[m[k]]) else "")
        lignes.append(dict(n=n, date=d, numero=g("numero"), tiers=g("tiers"), code=g("code_tiers"),
                           libelle=g("libelle"), compte=g("compte"), ht=ht, tva=tva, ttc=ttc,
                           taux=taux_str(taux)))
    return lignes

# ---------- Contrôles + écritures ----------
def aux_code(l, prefixe):
    if l["code"]: return l["code"].upper()[:17]
    base = re.sub(r"[^A-Z0-9]", "", norm(l["tiers"]).upper())[:8]
    return ("F" if prefixe == "401" else "C") + base if base else "DIVERS"

def construire_ecritures(lignes, type_, anomalies):
    cfg = COMPTES[type_]
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
        ht_tot = sum(x["ht"] for x in ls); tva_par_taux = OrderedDict()
        for x in ls: tva_par_taux[x["taux"]] = tva_par_taux.get(x["taux"], Decimal("0")) + x["tva"]
        tva_tot = sum(tva_par_taux.values()); ttc_calc = ht_tot + tva_tot
        ttc_src = sum(x["ttc"] for x in ls) if all(x["ttc"] is not None for x in ls) else None
        if ttc_src is not None and abs(ttc_src - ttc_calc) > Decimal("0.02"):
            anomalies.append((ls[0]["n"], f"HT+TVA ({ttc_calc}) <> TTC ({ttc_src})", numero))
        for t in tva_par_taux:
            if t not in TVA_COMPTES and t not in ("0", "0.0"):
                anomalies.append((ls[0]["n"], f"Taux de TVA inhabituel : {t}%", numero))
        avoir = ttc_calc < 0
        num += 1
        lib = f"{numero} {tiers}".strip()[:100]
        aux = aux_code(ls[0], cfg["tiers"])
        def ecr(compte, libc, aux_n, aux_l, montant_, sens):
            if montant_ == 0: return
            sens = sens if not avoir else ("C" if sens == "D" else "D")
            ecritures.append(dict(num=num, date=d, compte=compte, libc=libc, aux=aux_n, auxl=aux_l,
                                  piece=numero, lib=lib, debit=abs(montant_) if sens == "D" else Decimal("0"),
                                  credit=abs(montant_) if sens == "C" else Decimal("0")))
        st = cfg["sens_tiers"]; sc = "C" if st == "D" else "D"
        ecr(cfg["tiers"], "Clients" if type_ == "vente" else "Fournisseurs", aux, tiers, ttc_calc, st)
        for x in ls:
            ecr(x["compte"] or cfg["charge_produit"], cfg["libelle_cp"], "", "", x["ht"], sc)
        for t, v in tva_par_taux.items():
            cpt = TVA_COMPTES.get(t, ("44571000", "44566000"))[idx]
            ecr(cpt, TVA_LIB, "", "", v, sc)
    return ecritures

# ---------- Export FEC ----------
FEC_COLS = ["JournalCode","JournalLib","EcritureNum","EcritureDate","CompteNum","CompteLib","CompAuxNum",
            "CompAuxLib","PieceRef","PieceDate","EcritureLib","Debit","Credit","EcritureLet","DateLet",
            "ValidDate","Montantdevise","Idevise"]

def ecrire_fec(ecritures, type_, siren, exercice, sortie):
    jc, jl = COMPTES[type_]["journal"]
    fmt = lambda x: f"{x:.2f}".replace(".", ",")
    d_tot = sum(e["debit"] for e in ecritures); c_tot = sum(e["credit"] for e in ecritures)
    if d_tot != c_tot: raise ValueError(f"FEC déséquilibré : D={d_tot} C={c_tot}")
    fin = max(e["date"] for e in ecritures).strftime("%Y%m%d")
    nom = os.path.join(sortie, f"{siren}FEC{exercice}1231.txt") if exercice else os.path.join(sortie, f"{siren}FEC{fin}.txt")
    with open(nom, "w", encoding="utf-8", newline="") as f:
        f.write("\t".join(FEC_COLS) + "\r\n")
        for e in ecritures:
            dd = e["date"].strftime("%Y%m%d")
            f.write("\t".join([jc, jl, str(e["num"]), dd, e["compte"], e["libc"], e["aux"], e["auxl"],
                               e["piece"], dd, e["lib"], fmt(e["debit"]), fmt(e["credit"]),
                               "", "", dd, "", ""]) + "\r\n")
    return nom, d_tot

def traiter(fichiers, type_="vente", siren="000000000", exercice="", sortie="."):
    anomalies, lignes = [], []
    for f in fichiers:
        nom = os.path.basename(f)
        df = lire(f, type_)
        anomalies += [(nom, a, b) for a, b in df.attrs.get("anomalies", [])]
        if df.empty:
            anomalies.append((nom, "Aucune facture extraite", "")); continue
        lignes += vers_pivot(df, charger_mapping(df), anomalies, nom)
    if not lignes:
        raise ValueError("Aucune écriture exploitable (voir le rapport d'anomalies).")
    ecr = construire_ecritures(lignes, type_, anomalies)
    fec, total = ecrire_fec(ecr, type_, siren, exercice, sortie)
    rapport = None
    if anomalies:
        rapport = os.path.join(sortie, "rapport_anomalies.csv")
        pd.DataFrame(anomalies, columns=["ligne_source", "anomalie", "detail"]).to_csv(rapport, index=False, sep=";")
    return dict(fec=fec, ecritures=ecr, anomalies=anomalies, rapport=rapport, total=total)

def main():
    p = argparse.ArgumentParser()
    p.add_argument("fichiers", nargs="+"); p.add_argument("--siren", default="000000000")
    p.add_argument("--exercice", default=""); p.add_argument("--type", choices=["vente", "achat"], default="vente")
    p.add_argument("--sortie", default=".")
    a = p.parse_args()
    try:
        r = traiter(a.fichiers, a.type, a.siren, a.exercice, a.sortie)
    except ValueError as e:
        sys.exit(str(e))
    print(f"FEC généré : {r['fec']}  ({len(r['ecritures'])} lignes, débit = crédit = {r['total']})")
    print(f"{len(r['anomalies'])} anomalie(s) : {r['rapport']}" if r["anomalies"] else "Aucune anomalie.")

if __name__ == "__main__":
    main()
