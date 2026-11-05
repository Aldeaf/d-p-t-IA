"""Interface : streamlit run app.py"""
import json, os, tempfile
import pandas as pd
import streamlit as st
import agent_fec as ag

st.set_page_config(page_title="Factures → FEC", page_icon="🧾")
st.title("🧾 Factures → FEC pour ACD")
ss = st.session_state

# Vérifie que agent_fec.py est bien la dernière version (sinon message clair au lieu d'une erreur masquée)
_REQUIS = ["VERSION", "CHAMPS_LIBELLES", "importer_reglages", "exporter_reglages", "fusionner_modele", "sauver_mapping", "creer_archive", "OBLIGATOIRES", "TVA_COMPTES"]
_manque = [n for n in _REQUIS if not hasattr(ag, n)]
if _manque:
    st.error("agent_fec.py n'est pas à jour sur GitHub (il manque : " + ", ".join(_manque) +
             "). Remplacez-le par la dernière version, puis faites Manage app → Reboot app.")
    st.stop()
rerun = getattr(st, "rerun", None) or st.experimental_rerun   # compatibilité anciennes versions de Streamlit

# ---------- état initial (avant tout widget) ----------
def tva_df(d):
    return pd.DataFrame([(k, v[0], v[1]) for k, v in d.items()], columns=["taux", "compte_ventes", "compte_achats"])

def appliquer_reglages(c, t):
    for ty, base in (c or {}).items():
        ss[f"cpt_{ty}_tiers"], ss[f"cpt_{ty}_cp"], ss[f"cpt_{ty}_lib"] = base["tiers"], base["charge_produit"], base["libelle_cp"]
        ss[f"cpt_{ty}_jcode"], ss[f"cpt_{ty}_jlib"] = base["journal"][0], base["journal"][1]
    if t:
        ss.tva_base = tva_df(t)
        ss.pop("ed_tva", None)

if "init" not in ss:
    ss.init = True
    ss.tmp = tempfile.mkdtemp()
    ss.imput_base = ag.charger_imputations()
    ss.tva_base = tva_df(ag.TVA_COMPTES)
    if os.path.exists("reglages.json"):                      # réglages sauvegardés dans GitHub
        appliquer_reglages(*ag.importer_reglages(json.load(open("reglages.json", encoding="utf-8"))))
if "reglages_a_charger" in ss:                               # import demandé dans l'onglet Réglages
    appliquer_reglages(*ag.importer_reglages(ss.pop("reglages_a_charger")))

def comptes_ui():
    c = {}
    for t, b in ag.COMPTES.items():
        c[t] = {"journal": (ss.get(f"cpt_{t}_jcode", b["journal"][0]), ss.get(f"cpt_{t}_jlib", b["journal"][1])),
                "tiers": ss.get(f"cpt_{t}_tiers", b["tiers"]), "charge_produit": ss.get(f"cpt_{t}_cp", b["charge_produit"]),
                "libelle_cp": ss.get(f"cpt_{t}_lib", b["libelle_cp"]), "sens_tiers": b["sens_tiers"]}
    tva = {}
    for _, r in ss.get("tva_courant", ss.tva_base).iterrows():
        taux = str(r["taux"]).strip().replace(",", ".")
        if taux and taux != "nan":
            tva[taux] = (str(r["compte_ventes"]).strip(), str(r["compte_achats"]).strip())
    return c, tva

def mapping_choisi(nom):
    m = {ch: ss.get(f"map_{nom}_{ch}") for ch in ag.CHAMPS_LIBELLES}
    return {k: v for k, v in m.items() if v and v != "(aucune)"}

def afficher_piece(e):
    """Affiche la facture d'origine d'une écriture (PDF, image, XML ou ligne du CSV/Excel)."""
    chemin = os.path.join(ss.tmp, e["src"])
    page = f", page {e['page']}" if e["page"] else ""
    st.markdown(f"**👁️ Facture {e['piece']}** · {e['src']}{page}")
    ext = os.path.splitext(chemin)[1].lower()
    try:
        if ext == ".pdf":
            import pdfplumber
            with pdfplumber.open(chemin) as pdf:
                pages = [pdf.pages[int(e["page"]) - 1]] if e["page"] else pdf.pages[:4]
                for p in pages:
                    st.image(p.to_image(resolution=110).original, use_container_width=True)
        elif ext in (".png", ".jpg", ".jpeg"):
            st.image(chemin, use_container_width=True)
        elif ext == ".xml":
            st.code(open(chemin, encoding="utf-8", errors="replace").read()[:4000], language="xml")
        else:
            df = ag.lire(chemin)
            st.dataframe(df[df.astype(str).apply(lambda l: str(e["piece"]) in l.values, axis=1)], use_container_width=True)
    except Exception as ex:
        st.warning(f"Aperçu impossible : {ex}")

# ---------- page ----------
LIBELLES = {"vente": "🧑‍💼 Clients (factures de vente)", "achat": "🏭 Fournisseurs (factures d'achat)"}
type_ = st.radio("Je traite des…", list(LIBELLES), format_func=LIBELLES.get, horizontal=True)
col1, col2 = st.columns(2)
siren = col1.text_input("SIREN du dossier", "000000000")
exercice = col2.text_input("Exercice (AAAA)", "2026")
with st.sidebar:
    st.caption(f"Version {ag.VERSION}")
    cle = st.text_input("Clé API Anthropic (facultative, secours IA)", type="password")
    if cle:
        os.environ["ANTHROPIC_API_KEY"] = cle

tab1, tab2, tab3 = st.tabs(["📄 Traitement", "📒 Imputations", "⚙️ Réglages"])

with tab1:
    st.caption("CSV/Excel : règles · XML Factur-X/UBL · PDF texte : règles · Scans : OCR si installé, sinon IA (clé facultative)")
    fichiers = st.file_uploader("Déposez vos exports ou factures", accept_multiple_files=True,
                                type=["csv", "xlsx", "xls", "pdf", "xml", "png", "jpg", "jpeg"])
    chemins = {}
    for f in fichiers or []:
        p = os.path.join(ss.tmp, f.name)
        open(p, "wb").write(f.getbuffer())
        chemins[f.name] = p
    tabulaires = [n for n in chemins if n.lower().endswith((".csv", ".xlsx", ".xls", ".xlsm"))]

    def lancer():
        c, t = comptes_ui()
        maps = {n: mapping_choisi(n) for n in tabulaires if mapping_choisi(n)}
        try:
            with st.spinner("Analyse en cours…"):
                ss.res = ag.traiter(list(chemins.values()), type_, siren, exercice, ss.tmp,
                                    regles=ss.get("imput_courant", ss.imput_base), mappings=maps, comptes=c, tva_comptes=t,
                                    separateur="pipe" if ss.get("sep_fec") == "Pipe |" else "tab")
            ss.err = None
        except ValueError as e:
            ss.res, ss.err = None, str(e)

    if chemins and st.button(f"Générer le FEC ({'clients' if type_ == 'vente' else 'fournisseurs'})", type="primary"):
        lancer()

    if ss.get("err"):
        st.error(ss.err)
    r = ss.get("res")
    if r:
        c1, c2, c3 = st.columns(3)
        c1.metric("Lignes d'écriture", len(r["ecritures"]))
        c2.metric("Total débit = crédit", f"{r['total']:,.2f} €".replace(",", " "))
        c3.metric("Anomalies", len(r["anomalies"]))
        st.download_button("⬇️ Télécharger le FEC", open(r["fec"], "rb").read(),
                           file_name=os.path.basename(r["fec"]), mime="text/plain", type="primary")
        if r.get("zip"):
            st.download_button("📎 FEC + pièces jointes (zip)", open(r["zip"], "rb").read(),
                               file_name=os.path.basename(r["zip"]), mime="application/zip")
        if r["sans_regle"]:
            st.subheader("Comptes à renseigner")
            st.caption("Ces tiers n'ont pas de règle : le compte par défaut a été utilisé. "
                       "Indiquez le compte (ex. 606100) pour le mémoriser et régénérer le FEC.")
            nouv = st.data_editor(pd.DataFrame({"motif": r["sans_regle"], "compte": "", "libelle": ""}),
                                  hide_index=True, use_container_width=True, key="nouv_regles")
            if st.button("Enregistrer ces règles et régénérer"):
                ok = nouv[nouv["compte"].astype(str).str.strip() != ""].copy()
                if not ok.empty:
                    ok["type"] = "achat"
                    ss.imput_base = pd.concat([ss.get("imput_courant", ss.imput_base), ok], ignore_index=True)
                    ss.pop("ed_imput", None)
                    ss.imput_courant = ss.imput_base
                    lancer()
                    rerun()
        rest = [a for a in r["anomalies"] if not a[1].startswith("Aucune règle d'imputation")]
        if rest:
            st.warning("Points à vérifier avant import dans ACD :")
            st.dataframe(pd.DataFrame(rest, columns=["Fichier:ligne", "Anomalie", "Détail"]), use_container_width=True)
            st.download_button("Télécharger le rapport", open(r["rapport"], "rb").read(), file_name="rapport_anomalies.csv")
        st.subheader("Écritures")
        st.caption("Touchez une ligne pour afficher la facture correspondante.")
        ecr = pd.DataFrame(r["ecritures"])
        sel = None
        try:
            ev = st.dataframe(ecr[["num", "date", "compte", "aux", "piece", "lib", "debit", "credit"]],
                              on_select="rerun", selection_mode="single-row", hide_index=True,
                              use_container_width=True, key="tbl_ecr")
            if ev.selection.rows:
                sel = ecr.iloc[ev.selection.rows[0]]
        except TypeError:                                   # anciennes versions de Streamlit
            uniques = ecr.drop_duplicates("num")
            choix = st.selectbox("Facture à afficher", ["(aucune)"] + [f"{p.num} - {p.piece}" for p in uniques.itertuples()])
            if choix != "(aucune)":
                sel = uniques[uniques["num"] == int(choix.split(" - ")[0])].iloc[0]
        if sel is not None:
            afficher_piece(sel)

with tab2:
    st.caption("Si le nom du tiers contient le motif, le compte est utilisé. Colonne type : achat, vente, ou vide pour les deux. "
               "Les règles du modèle sont indicatives : vérifiez-les avec le plan comptable du dossier.")
    if st.button("➕ Ajouter les fournisseurs courants du modèle"):
        base, n = ag.fusionner_modele(ss.get("imput_courant", ss.imput_base))
        ss.imput_base = base
        ss.pop("ed_imput", None)
        st.success(f"{n} règle(s) ajoutée(s).")
    up = st.file_uploader("Importer ma table (CSV)", type="csv", key="imp_up")
    if up is not None and ss.get("imp_nom") != (up.name, up.size):
        tmp_csv = os.path.join(ss.tmp, "imputations_import.csv")
        open(tmp_csv, "wb").write(up.getbuffer())
        ss.imput_base, ss.imp_nom = ag.charger_imputations(tmp_csv), (up.name, up.size)
        ss.pop("ed_imput", None)
    ss.imput_courant = st.data_editor(ss.imput_base, num_rows="dynamic", use_container_width=True,
                                      hide_index=True, key="ed_imput")
    st.download_button("⬇️ Télécharger la table", ss.imput_courant.to_csv(index=False, sep=";").encode("utf-8"),
                       file_name="imputations.csv", mime="text/csv")
    st.caption("Pour la conserver : remplacez imputations.csv dans GitHub par ce fichier.")

with tab3:
    with st.expander("📑 Colonnes des CSV / Excel (mapping)", expanded=bool(tabulaires)):
        if not tabulaires:
            st.info("Déposez un CSV ou Excel dans l'onglet Traitement pour régler ses colonnes. "
                    "Les PDF et XML n'ont pas de colonnes à régler.")
        for nom in tabulaires:
            st.markdown(f"**{nom}**")
            try:
                df = ag.lire(chemins[nom])
            except Exception as e:
                st.error(f"Lecture impossible : {e}"); continue
            detecte = ag.charger_mapping(df)
            opts = ["(aucune)"] + [str(c) for c in df.columns]
            for champ, lib in ag.CHAMPS_LIBELLES.items():
                cur = detecte.get(champ)
                st.selectbox(lib + (" *" if champ in ag.OBLIGATOIRES else ""), opts,
                             index=opts.index(cur) if cur in opts else 0, key=f"map_{nom}_{champ}")
            manque = [ag.CHAMPS_LIBELLES[c] for c in ag.OBLIGATOIRES if c not in mapping_choisi(nom)]
            if manque:
                st.warning("Colonne(s) obligatoire(s) non choisie(s) : " + ", ".join(manque))
            st.dataframe(df.head(3), use_container_width=True)
            if st.button("💾 Mémoriser pour les fichiers de même format", key=f"save_{nom}"):
                ag.sauver_mapping(df, mapping_choisi(nom))
                st.success("Mapping mémorisé.")
    with st.expander("🧮 Comptes et journaux"):
        for t, lib in (("vente", "Clients (ventes)"), ("achat", "Fournisseurs (achats)")):
            b = ag.COMPTES[t]
            st.markdown(f"**{lib}**")
            a1, a2 = st.columns(2)
            a1.text_input("Compte de tiers", b["tiers"], key=f"cpt_{t}_tiers")
            a2.text_input("Compte produit/charge par défaut", b["charge_produit"], key=f"cpt_{t}_cp")
            a3, a4 = st.columns(2)
            a3.text_input("Code journal", b["journal"][0], key=f"cpt_{t}_jcode")
            a4.text_input("Libellé du journal", b["journal"][1], key=f"cpt_{t}_jlib")
            st.text_input("Libellé du compte par défaut", b["libelle_cp"], key=f"cpt_{t}_lib")
        st.markdown("**Comptes de TVA par taux**")
        ss.tva_courant = st.data_editor(ss.tva_base, num_rows="dynamic", hide_index=True,
                                        use_container_width=True, key="ed_tva")
    with st.expander("📄 Format du FEC"):
        st.radio("Séparateur des colonnes", ["Tabulation", "Pipe |"], horizontal=True, key="sep_fec")
        st.caption("Les deux sont admis pour un FEC. Choisissez celui que votre import ACD accepte (certaines sources "
                   "indiquent que le FEC pour ACD utilise le séparateur pipe : à confirmer auprès d'ACD).")
    with st.expander("💾 Sauvegarde des réglages"):
        c, t = comptes_ui()
        st.download_button("⬇️ Télécharger mes réglages (reglages.json)",
                           json.dumps(ag.exporter_reglages(c, t), ensure_ascii=False, indent=2).encode("utf-8"),
                           file_name="reglages.json", mime="application/json")
        st.caption("Pour les garder : ajoutez ce fichier à la racine de GitHub, il sera rechargé au démarrage.")
        rg = st.file_uploader("Restaurer des réglages (reglages.json)", type="json", key="reg_up")
        if rg is not None and ss.get("reg_nom") != (rg.name, rg.size):
            ss.reg_nom = (rg.name, rg.size)
            ss.reglages_a_charger = json.loads(rg.getvalue().decode("utf-8"))
            rerun()
