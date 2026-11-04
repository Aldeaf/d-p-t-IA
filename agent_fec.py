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

# ---------- Lecture gratuite : Factur-X, texte du PDF, OCR (optionnel) ----------
STD_TAUX = [20.0, 10.0, 5.5, 2.1, 0.0]
MOIS = {"janvier": 1, "fevrier": 2, "mars": 3, "avril": 4, "mai": 5, "juin": 6, "juillet": 7,
        "aout": 8, "septembre": 9, "octobre": 10, "novembre": 11, "decembre": 12}
NB = r"(-?(?:\d{1,3}(?:[ \u00a0.,]\d{3})+|\d+)(?:[.,]\d{1,2})?)"
TOKEN = r"([A-Z0-9][A-Z0-9\-/_.]*\d[A-Z0-9\-/_.]*)"

def _loc(e):
    return e.tag.rsplit("}", 1)[-1]

def lire_facturx(fichier, type_="vente"):
    """Lit le XML Factur-X / ZUGFeRD embarqué dans un PDF. None s'il n'y en a pas."""
    try:
        from pypdf import PdfReader
        pieces = PdfReader(fichier).attachments
        xml = next((c[0] for n, c in pieces.items() if n.lower().endswith(".xml")), None)
        if xml is None: return None
        import xml.etree.ElementTree as ET
        racine = ET.fromstring(xml)
        desc = lambda n, nom: [e for e in n.iter() if _loc(e) == nom]
        val = lambda n, nom: ((desc(n, nom) or [None])[0].text or "").strip() if desc(n, nom) else ""
        doc = desc(racine, "ExchangedDocument")[0]
        d = val(desc(doc, "IssueDateTime")[0], "DateTimeString")
        d = f"{d[6:8]}/{d[4:6]}/{d[:4]}" if len(d) == 8 and d.isdigit() else d
        agr = desc(racine, "ApplicableHeaderTradeAgreement")[0]
        tiers = val(desc(agr, "BuyerTradeParty" if type_ == "vente" else "SellerTradeParty")[0], "Name")
        reg = desc(racine, "ApplicableHeaderTradeSettlement")[0]
        taxes = [e for e in reg if _loc(e) == "ApplicableTradeTax"]
        signe = -1 if val(doc, "TypeCode") == "381" else 1
        rows, somme = [], Decimal("0")
        for t in taxes:
            ht, tva = montant(val(t, "BasisAmount")), montant(val(t, "CalculatedAmount")) or Decimal("0")
            if ht is None: continue
            ht, tva = ht * signe, tva * signe
            somme += ht + tva
            rows.append({"Date facture": d, "N° facture": val(doc, "ID"), "Client": tiers, "Total HT": str(ht),
                         "Taux TVA": val(t, "RateApplicablePercent"), "Total TVA": str(tva), "Total TTC": str(ht + tva)})
        if not rows: return None
        df = pd.DataFrame(rows, columns=COLONNES_STD)
        total = montant(val(desc(racine, "SpecifiedTradeSettlementHeaderMonetarySummation")[0], "GrandTotalAmount"))
        df.attrs["anomalies"] = []
        if total is not None and abs(somme - total * signe) > Decimal("0.02"):
            df.attrs["anomalies"].append((f"Factur-X : total reconstitué ({somme}) <> TTC déclaré ({total})", val(doc, "ID")))
        return df
    except Exception:
        return None

def texte_pdf(fichier):
    try:
        import pdfplumber
        with pdfplumber.open(fichier) as pdf:
            return "\n".join((p.extract_text() or "") for p in pdf.pages)
    except Exception:
        return ""

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

def _num(texte, motifs):
    for m in motifs:
        r = re.search(m + r"[^\d\n\-]{0,40}" + NB, texte, re.I)
        if r and montant(r.group(1)) is not None:
            return montant(r.group(1))
    return None

def _totaux_tableau(lignes):
    """Cas du tableau : en-tête 'HT TVA TTC' puis une ligne de 3 montants."""
    for i, l in enumerate(lignes):
        if re.search(r"h\.?t\b", l, re.I) and re.search(r"t\.?t\.?c", l, re.I):
            for s in lignes[i + 1:i + 3]:
                v = [montant(x) for x in re.findall(NB, s)]
                v = [x for x in v if x is not None]
                if len(v) >= 3 and abs(v[-3] + v[-2] - v[-1]) <= Decimal("0.02"):
                    return v[-3], v[-1]
    return None, None

def _date_texte(texte):
    r = (re.search(r"date(?:\s*(?:de|d')\s*(?:facture|facturation|[ée]mission))?\s*:?\s*(\d{1,2}[/.\-]\d{1,2}[/.\-]\d{2,4})", texte, re.I)
         or re.search(r"\b(\d{1,2}[/.\-]\d{1,2}[/.\-]\d{4})\b", texte))
    if r: return r.group(1)
    r = re.search(r"(\d{1,2})(?:er)?\s+(janvier|f[ée]vrier|mars|avril|mai|juin|juillet|ao[uû]t|septembre|octobre|novembre|d[ée]cembre)\s+(\d{4})", texte, re.I)
    return f"{int(r.group(1)):02d}/{MOIS[norm(r.group(2))]:02d}/{r.group(3)}" if r else ""

def _tiers_texte(lignes, type_):
    etiq = (r"(?:factur[ée]e?\s*[àa]|client|destinataire|bill(?:ed)?\s*to|adress[ée]e?\s*[àa])" if type_ == "vente"
            else r"(?:fournisseur|[ée]metteur|vendeur|from)")
    for i, l in enumerate(lignes):
        m = re.match(rf"\s*{etiq}\s*:?\s*(.*)$", l, re.I)
        if m:
            if len(m.group(1).strip()) > 2: return m.group(1).strip()
            for s in lignes[i + 1:i + 3]:
                if s.strip(): return s.strip()
    if type_ == "achat":
        for l in lignes:
            if re.search(r"[A-Za-z]{3}", l) and not re.match(r"\s*(facture|avoir|invoice)", l, re.I): return l.strip()
    return ""

def lire_texte(texte, type_, nom):
    """Extraction par règles à partir du texte d'une facture (une facture par document)."""
    lignes = [l.strip() for l in texte.splitlines() if l.strip()]
    anom, numero = [], ""
    for pat in (r"(?:facture|avoir|invoice|credit note)\s*(?:n[°ºo]\.?|num[ée]ro|no\.?|#)?\s*:?\s*" + TOKEN,
                r"\bn[°º]\s*(?:de\s*facture)?\s*:?\s*" + TOKEN):
        for r in re.finditer(pat, texte, re.I):
            c = r.group(1).strip(".-/_")
            if not re.fullmatch(r"\d{1,2}[/.\-]\d{1,2}[/.\-]\d{2,4}", c):
                numero = c; break
        if numero: break
    if not numero:
        numero = nom; anom.append(("N° de facture non trouvé (nom du fichier utilisé)", nom))
    ht = _num(texte, [r"total\s*h\.?\s?t\.?", r"montant\s*h\.?\s?t\.?", r"sous[\s-]?total(?:\s*h\.?\s?t\.?)?", r"net\s*h\.?\s?t\.?"])
    ttc = _num(texte, [r"total\s*t\.?t\.?c\.?", r"net\s*[àa]\s*payer", r"montant\s*t\.?t\.?c", r"total\s*[àa]\s*payer", r"total\s*due", r"amount\s*due"])
    tva = _num(texte, [r"total\s*t\.?v\.?a\.?", r"montant\s*t\.?v\.?a\.?"])
    if ht is None or ttc is None:
        h2, t2 = _totaux_tableau(lignes)
        ht = ht if ht is not None else h2; ttc = ttc if ttc is not None else t2
    if ht is None and ttc is not None and tva is not None: ht = ttc - tva
    if ttc is None and ht is not None and tva is not None: ttc = ht + tva
    if ht is None or ttc is None or ht == 0:
        return None
    if tva is not None and abs(ht + tva - ttc) > Decimal("0.02"):
        anom.append((f"HT ({ht}) + TVA ({tva}) <> TTC ({ttc})", numero))
    tva = ttc - ht
    if re.search(r"\bavoir\b|credit note", " ".join(lignes[:15]), re.I) and ht > 0:
        ht, tva, ttc = -ht, -tva, -ttc
    calc = float(tva / ht * 100)
    proche = min(STD_TAUX, key=lambda t: abs(t - calc))
    if abs(proche - calc) <= 0.3: taux = proche
    else:
        taux = round(calc, 1)
        anom.append((f"Taux de TVA calculé {taux} % : facture multi-taux ou erreur, à vérifier", numero))
    tiers = _tiers_texte(lignes, type_)
    if not tiers:
        tiers = "A IDENTIFIER"; anom.append(("Tiers non identifié", numero))
    df = pd.DataFrame([{"Date facture": _date_texte(texte), "N° facture": numero, "Client": tiers,
                        "Total HT": str(ht), "Taux TVA": str(taux), "Total TVA": str(tva), "Total TTC": str(ttc)}],
                      columns=COLONNES_STD)
    df.attrs["anomalies"] = anom
    return df

def lire_pdf(fichier, type_="vente"):
    """Cascade : Factur-X -> texte du PDF -> OCR (si installé) -> IA (si clé API)."""
    ext = os.path.splitext(fichier)[1].lower()
    nom = os.path.splitext(os.path.basename(fichier))[0]
    texte = ""
    if ext == ".pdf":
        df = lire_facturx(fichier, type_)
        if df is not None: return df
        texte = texte_pdf(fichier)
    if len(texte.strip()) < 40:
        texte = ocr_texte(fichier) or ""
    df = lire_texte(texte, type_, nom) if len(texte.strip()) >= 40 else None
    if df is not None: return df
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

# ---------- Table d'imputation (fournisseur/client -> compte) ----------
IMPUT_DEFAUT = [
    ("edf", "606100", "Énergie - électricité"), ("engie", "606100", "Énergie - gaz"),
    ("orange", "626000", "Frais postaux et télécommunications"), ("sfr", "626000", "Frais postaux et télécommunications"),
    ("bouygues", "626000", "Frais postaux et télécommunications"), ("free", "626000", "Frais postaux et télécommunications"),
    ("la poste", "626000", "Frais postaux et télécommunications"),
    ("axa", "616000", "Primes d'assurance"), ("maif", "616000", "Primes d'assurance"),
    ("allianz", "616000", "Primes d'assurance"), ("groupama", "616000", "Primes d'assurance"),
    ("sncf", "625100", "Voyages et déplacements"), ("air france", "625100", "Voyages et déplacements"),
]

def charger_imputations(chemin="imputations.csv"):
    """Table motif;compte;libelle. Un modèle de départ est créé si le fichier n'existe pas."""
    if not os.path.exists(chemin):
        pd.DataFrame(IMPUT_DEFAUT, columns=["motif", "compte", "libelle"]).to_csv(chemin, index=False, sep=";")
    df = pd.read_csv(chemin, sep=";", dtype=str).fillna("")
    for c in ("motif", "compte", "libelle"):
        if c not in df.columns: df[c] = ""
    return df[["motif", "compte", "libelle"]]

def _s(v):
    return "" if v is None or (isinstance(v, float) and pd.isna(v)) else str(v).strip()

def preparer_regles(regles):
    out = [(norm(_s(r["motif"])), _s(r["compte"]), _s(r["libelle"])) for _, r in regles.iterrows()]
    return sorted([r for r in out if r[0] and r[1]], key=lambda r: -len(r[0]))  # motif le plus précis d'abord

def trouver_regle(tiers, regles):
    n = norm(tiers)
    for motif, compte, lib in regles:
        if re.search(rf"(?<![a-z0-9]){re.escape(motif)}(?![a-z0-9])", n):
            return compte, lib
    return None

# ---------- Contrôles + écritures ----------
def aux_code(l, prefixe):
    if l["code"]: return l["code"].upper()[:17]
    base = re.sub(r"[^A-Z0-9]", "", norm(l["tiers"]).upper())[:8]
    return ("F" if prefixe == "401" else "C") + base if base else "DIVERS"

def construire_ecritures(lignes, type_, anomalies, regles=None):
    cfg = COMPTES[type_]
    regles, signales = regles or [], set()
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
            cpt, libc = x["compte"], cfg["libelle_cp"]
            if not cpt:
                reg = trouver_regle(tiers, regles)
                if reg: cpt, libc = reg[0], (reg[1] or libc)
                else:
                    cpt = cfg["charge_produit"]
                    if type_ == "achat" and tiers not in signales:
                        signales.add(tiers)
                        anomalies.append((x["n"], f"Aucune règle d'imputation : compte par défaut {cpt}", tiers))
            ecr(cpt, libc, "", "", x["ht"], sc)
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

def traiter(fichiers, type_="vente", siren="000000000", exercice="", sortie=".", regles=None):
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
        lignes += vers_pivot(df, charger_mapping(df), anomalies, nom)
    if not lignes:
        raise ValueError("Aucune écriture exploitable : " + " ; ".join(f"{a[0]} - {a[1]}" for a in anomalies[:3]))
    ecr = construire_ecritures(lignes, type_, anomalies,
                               preparer_regles(regles if regles is not None else charger_imputations()))
    fec, total = ecrire_fec(ecr, type_, siren, exercice, sortie)
    rapport = None
    if anomalies:
        rapport = os.path.join(sortie, "rapport_anomalies.csv")
        pd.DataFrame(anomalies, columns=["ligne_source", "anomalie", "detail"]).to_csv(rapport, index=False, sep=";")
    sans = sorted({a[2] for a in anomalies if a[1].startswith("Aucune règle d'imputation")})
    return dict(fec=fec, ecritures=ecr, anomalies=anomalies, rapport=rapport, total=total, sans_regle=sans)

def main():
    p = argparse.ArgumentParser()
    p.add_argument("fichiers", nargs="+"); p.add_argument("--siren", default="000000000")
    p.add_argument("--exercice", default=""); p.add_argument("--type", choices=["vente", "achat"], default="vente")
    p.add_argument("--sortie", default=".")
    p.add_argument("--imputations", default="imputations.csv")
    a = p.parse_args()
    try:
        r = traiter(a.fichiers, a.type, a.siren, a.exercice, a.sortie, charger_imputations(a.imputations))
    except ValueError as e:
        sys.exit(str(e))
    print(f"FEC généré : {r['fec']}  ({len(r['ecritures'])} lignes, débit = crédit = {r['total']})")
    if r["sans_regle"]: print("Sans règle d'imputation (à ajouter dans imputations.csv) : " + ", ".join(r["sans_regle"]))
    print(f"{len(r['anomalies'])} anomalie(s) : {r['rapport']}" if r["anomalies"] else "Aucune anomalie.")

if __name__ == "__main__":
    main()
