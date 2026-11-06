"""Interface : streamlit run app.py"""
import json, os, tempfile
import pandas as pd
import streamlit as st
import agent_fec as ag

st.set_page_config(page_title="Agent Fec", page_icon="🧾")
st.title("🧾 Agent Fec")
ss = st.session_state

# Vérifie que agent_fec.py est bien la dernière version (sinon message clair au lieu d'une erreur masquée)
_REQUIS = ["VERSION", "CHAMPS_LIBELLES", "importer_reglages", "exporter_reglages", "fusionner_modele", "sauver_mapping", "creer_archive", "extraire_zip", "rapprocher", "produits_par_taux", "preparer_nd", "colonnes_par_taux", "tva_nd_defaut_df", "OBLIGATOIRES", "TVA_COMPTES", "regenerer", "migrer_modele", "VEHICULES", "CARBURANTS", "pct_deductible_carburant", "valider_ecritures", "COLS_ND"]
_manque = [n for n in _REQUIS if not hasattr(ag, n)]
if _manque:
    st.error("agent_fec.py n'est pas à jour sur GitHub (il manque : " + ", ".join(_manque) +
             "). Remplacez-le par la dernière version, puis faites Manage app → Reboot app.")
    st.stop()
rerun = getattr(st, "rerun", None) or st.experimental_rerun   # compatibilité anciennes versions de Streamlit

# ---------- état initial (avant tout widget) ----------
def tva_df(d):
    return pd.DataFrame([(k, v[0], v[1]) for k, v in d.items()], columns=["taux", "compte_ventes", "compte_achats"])

def produits_df(d):
    return pd.DataFrame([(k, v[0], v[1]) for k, v in d.items()], columns=["taux", "compte", "libelle"])

AUCUN = "(aucune règle carburant)"

def appliquer_reglages(c, t, p=None, nd=None, veh=None):
    for ty, base in (c or {}).items():
        ss[f"cpt_{ty}_tiers"], ss[f"cpt_{ty}_cp"], ss[f"cpt_{ty}_lib"] = base["tiers"], base["charge_produit"], base["libelle_cp"]
        ss[f"cpt_{ty}_jcode"], ss[f"cpt_{ty}_jlib"] = base["journal"][0], base["journal"][1]
    if t:
        ss.tva_base = tva_df(t)
        ss.pop("ed_tva", None)
    if p:
        ss.prod_base = produits_df(p)
        ss.pop("ed_prod", None)
    if nd:
        ss.nd_base = pd.DataFrame(nd, columns=ag.COLS_ND)
        ss.pop("ed_nd", None)
    if veh:
        if veh.get("vehicule") in ag.VEHICULES: ss.veh_type = veh["vehicule"]
        if veh.get("carburant") in ag.CARBURANTS or veh.get("carburant") == AUCUN: ss.veh_carb = veh["carburant"]
        if veh.get("comptes"): ss.veh_cpt = veh["comptes"]

if "init" not in ss:
    ss.init = True
    ss.tmp = tempfile.mkdtemp()
    ss.imput_base = ag.charger_imputations()
    ss.tva_base = tva_df(ag.TVA_COMPTES)
    ss.prod_base = produits_df(ag.PRODUITS_TVA)
    ss.nd_base = ag.tva_nd_defaut_df()
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
    prod = {}
    for _, r in ss.get("prod_courant", ss.prod_base).iterrows():
        taux = ag.taux_str(str(r["taux"]).strip().replace(",", "."))
        if taux and taux != "nan" and str(r["compte"]).strip():
            prod[taux] = (str(r["compte"]).strip(), str(r["libelle"]).strip())
    return c, tva, prod

def nd_ui():
    """Règles de TVA non déductible à appliquer (tableau vide si désactivé)."""
    if not ss.get("nd_actif", True):
        return pd.DataFrame(columns=ag.COLS_ND)
    return ss.get("nd_courant", ss.nd_base)

def tiers_au_compte_ui():
    return str(ss.get("mode_tiers", "Au niveau du compte")).startswith("Au niveau")

def mapping_choisi(nom):
    m = {ch: ss.get(f"map_{nom}_{ch}") for ch in ag.CHAMPS_LIBELLES}
    return {k: v for k, v in m.items() if v and v != "(aucune)"}

def afficher_piece(e):
    """Affiche la facture d'origine d'une écriture (PDF, image, XML ou ligne du CSV/Excel)."""
    chemin = ss.get("index_chemins", {}).get(e["src"]) or os.path.join(ss.tmp, e["src"])
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
    st.caption(f"Agent Fec · version {ag.VERSION}")
    cle = st.text_input("Clé API Anthropic (facultative, secours IA)", type="password")
    if cle:
        os.environ["ANTHROPIC_API_KEY"] = cle

tab1, tab2, tab3 = st.tabs(["📄 Traitement", "📒 Imputations", "⚙️ Réglages"])

with tab1:
    st.caption("CSV/Excel : règles · XML Factur-X/UBL · PDF texte : règles · Scans : OCR si installé, sinon IA (clé facultative)")
    fichiers = st.file_uploader("Déposez vos exports ou factures", accept_multiple_files=True,
                                type=["zip", "csv", "xlsx", "xls", "pdf", "xml", "png", "jpg", "jpeg"])
    ss.setdefault("zips", {})
    chemins = {}
    for f in fichiers or []:
        p = os.path.join(ss.tmp, f.name)
        open(p, "wb").write(f.getbuffer())
        if f.name.lower().endswith(".zip"):                  # zip : extraction unique, mise en cache
            cle = (f.name, f.size)
            if cle not in ss.zips:
                try:
                    ss.zips[cle] = ag.extraire_zip(p, os.path.join(ss.tmp, f"zip{len(ss.zips)}"))
                except Exception as ex:
                    st.error(f"{f.name} : {ex}"); ss.zips[cle] = []
            for q in ss.zips[cle]:
                chemins[os.path.basename(q)] = q
        else:
            chemins[f.name] = p
    ss.index_chemins = chemins
    tabulaires, docs, xmls = [[os.path.basename(p) for p in L] for L in ag.classer(list(chemins.values()))]
    pieces_noms = []
    if tabulaires and docs:
        if st.checkbox("📎 Rapprocher automatiquement les PDF / images du fichier Excel (pièces justificatives)",
                       value=True, key="use_pieces"):
            pieces_noms = docs
    if pieces_noms:
        st.checkbox("🧮 Analyser les pièces pour proposer la TVA multi-taux (une ligne par taux)", value=True, key="ventiler_pieces")
    a_traiter = tabulaires + xmls + ([] if pieces_noms else docs)
    if tabulaires and docs:
        st.caption(f"{len(tabulaires)} fichier(s) comptable(s) · {len(docs)} PDF/image(s) · "
                   + ("les PDF servent de pièces justificatives." if pieces_noms else "les PDF sont lus comme des factures."))

    def lancer():
        c, t, p = comptes_ui()
        maps = {n: mapping_choisi(n) for n in tabulaires if mapping_choisi(n)}
        try:
            with st.spinner("Analyse en cours…"):
                ss.res = ag.traiter([chemins[n] for n in a_traiter], type_, siren, exercice, ss.tmp,
                                    regles=ss.get("imput_courant", ss.imput_base), mappings=maps, comptes=c, tva_comptes=t, produits_tva=p,
                                    separateur="pipe" if ss.get("sep_fec") == "Pipe |" else "tab",
                                    tiers_corriges=ss.get("tiers_corriges"),
                                    tva_nd=nd_ui(), tiers_au_compte=tiers_au_compte_ui(), ventilations=ss.get("ventilations"),
                                    pieces=[chemins[n] for n in pieces_noms], rapprochements=ss.get("rapp_manuels"),
                                    vehicule=ss.get("veh_type", "VP"),
                                    carburant=None if (ss.get("veh_carb", "Gazole") == AUCUN or not ss.get("nd_actif", True)) else ss.get("veh_carb", "Gazole"),
                                    comptes_carburant=ss.get("veh_cpt", ag.COMPTE_CARBURANT_DEFAUT),
                                    ventiler_pieces=ss.get("ventiler_pieces", True))
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
        if r.get("multi_taux"):
            st.warning(f"{len(r['multi_taux'])} facture(s) avec plusieurs taux de TVA non détectés : à ventiler.")
            with st.expander("🧮 Ventiler la TVA (multi-taux)", expanded=True):
                st.caption("Une ligne par taux : ajoutez des lignes (+) pour 2 ou 3 taux. Le total doit égaler la facture.")
                numeros = [m["numero"] for m in r["multi_taux"]]
                init = pd.DataFrame([{"facture": m["numero"], "taux": "", "ht": m["ht"], "tva": m["tva"]} for m in r["multi_taux"]])
                ev = st.data_editor(init, num_rows="dynamic", hide_index=True, use_container_width=True,
                                    column_config={"facture": st.column_config.SelectboxColumn("facture", options=numeros)},
                                    key=f"ed_vent_{ss.get('n_vent', 0)}")
                if st.button("Appliquer la ventilation et régénérer"):
                    vent = dict(ss.get("ventilations") or {})
                    par_num = {m["numero"]: m["cle"] for m in r["multi_taux"]}
                    for num_f, grp in ev.groupby("facture"):
                        rows = [{"taux": g["taux"], "ht": g["ht"], "tva": g["tva"]} for _, g in grp.iterrows()
                                if str(g["taux"]).strip() not in ("", "nan", "None")]
                        if rows and num_f in par_num:
                            vent[par_num[num_f]] = rows
                    ss.ventilations = vent
                    ss.n_vent = ss.get("n_vent", 0) + 1
                    lancer()
                    rerun()
        if r.get("rapprochements"):
            rap = pd.DataFrame(r["rapprochements"])
            nb = int((rap["piece"] != "").sum())
            st.info(f"📎 {nb} facture(s) sur {len(rap)} rapprochée(s) d'une pièce justificative.")
            with st.expander("🔗 Rapprochement factures ↔ pièces", expanded=nb < len(rap)):
                if "ventilation" not in rap: rap["ventilation"] = ""
                rap["ventilation"] = rap["ventilation"].apply(lambda v: ", ".join(map(str, v)) if isinstance(v, (list, tuple)) else (str(v) if v else ""))
                aff = rap[["numero", "tiers", "ttc", "piece", "ventilation", "methode"]].copy()
                aff["piece"] = aff["piece"].replace("", "(aucune)")
                ed2 = st.data_editor(aff, hide_index=True, use_container_width=True,
                                     disabled=["numero", "tiers", "ttc", "ventilation", "methode"],
                                     column_config={"piece": st.column_config.SelectboxColumn("pièce", options=["(aucune)"] + pieces_noms)},
                                     key=f"ed_rapp_{ss.get('n_rapp', 0)}")
                if st.button("Appliquer le rapprochement et régénérer"):
                    man = dict(ss.get("rapp_manuels") or {})
                    for i in rap.index:
                        if ed2.loc[i, "piece"] != aff.loc[i, "piece"]:
                            man[rap.loc[i, "cle"]] = "" if ed2.loc[i, "piece"] == "(aucune)" else ed2.loc[i, "piece"]
                    ss.rapp_manuels = man
                    ss.n_rapp = ss.get("n_rapp", 0) + 1
                    lancer()
                    rerun()
        with st.expander("✏️ Vérifier ou corriger les noms de tiers"):
            fac = pd.DataFrame(r["factures"])
            ed = st.data_editor(fac[["fichier", "numero", "tiers"]], hide_index=True, use_container_width=True,
                                disabled=["fichier", "numero"], key=f"ed_tiers_{ss.get('n_tiers', 0)}")
            if st.button("Appliquer les corrections et régénérer"):
                corr = dict(ss.get("tiers_corriges") or {})
                for i in fac.index:
                    nouveau = str(ed.loc[i, "tiers"]).strip()
                    if nouveau and nouveau != str(fac.loc[i, "tiers"]).strip():
                        corr[fac.loc[i, "cle"]] = nouveau
                ss.tiers_corriges = corr
                ss.n_tiers = ss.get("n_tiers", 0) + 1
                lancer()
                rerun()
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
        ecr = pd.DataFrame(r["ecritures"])
        if "alerte" not in ecr: ecr["alerte"] = ""
        ecr["alerte"] = ecr["alerte"].fillna("")
        nb_al = int((ecr.drop_duplicates("num")["alerte"] != "").sum()) if len(ecr) else 0
        if nb_al:
            st.caption(f"🟡 {nb_al} écriture(s) en alerte (surlignées en jaune) : voir la colonne « alerte ».")
        st.caption("Touchez une ligne pour afficher la facture correspondante.")
        vue = ecr[["num", "date", "compte", "aux", "piece", "lib", "debit", "credit", "alerte"]].copy()
        vue["date"] = pd.to_datetime(vue["date"]).dt.strftime("%d/%m/%Y")
        for k in ("debit", "credit"):
            vue[k] = vue[k].astype(float)
        jaune = vue["alerte"] != ""
        # une alerte colore toutes les lignes de l'écriture concernée
        nums_al = set(vue.loc[jaune, "num"])
        def _jaune(row):
            return ["background-color: #FFF9C4; color: #000"] * len(row) if row["num"] in nums_al else [""] * len(row)
        sel = None
        try:
            ev = st.dataframe(vue.style.apply(_jaune, axis=1).format({"debit": "{:.2f}", "credit": "{:.2f}"}),
                              on_select="rerun", selection_mode="single-row", hide_index=True,
                              use_container_width=True, key="tbl_ecr")
            if ev.selection.rows:
                sel = ecr.iloc[ev.selection.rows[0]]
        except Exception:                                   # anciennes versions de Streamlit
            st.dataframe(vue.style.apply(_jaune, axis=1), hide_index=True, use_container_width=True)
            uniques = ecr.drop_duplicates("num")
            choix = st.selectbox("Facture à afficher", ["(aucune)"] + [f"{p.num} - {p.piece}" for p in uniques.itertuples()])
            if choix != "(aucune)":
                sel = uniques[uniques["num"] == int(choix.split(" - ")[0])].iloc[0]
        if sel is not None:
            afficher_piece(sel)

        with st.expander("✏️ Modifier les écritures", expanded=False):
            st.caption("Modifiez comptes, libellés, montants ou dates ; ajoutez une ligne (+) en indiquant le n° d'écriture. "
                       "Le débit doit égaler le crédit pour chaque écriture.")
            cols_ed = ["num", "date", "compte", "libc", "aux", "auxl", "piece", "lib", "debit", "credit", "alerte"]
            base = ecr[cols_ed].copy()
            base["date"] = pd.to_datetime(base["date"]).dt.strftime("%Y-%m-%d")
            for k in ("debit", "credit"):
                base[k] = base[k].astype(float)
            ed_e = st.data_editor(base, num_rows="dynamic", hide_index=True, use_container_width=True,
                                  disabled=["alerte"], key=f"ed_ecr_{ss.get('n_ecr', 0)}")
            if st.button("Valider mes modifications et régénérer le FEC"):
                try:
                    nv = ag.regenerer(ed_e, r["ecritures"], type_, siren, exercice, ss.tmp, comptes=comptes_ui()[0],
                                      separateur="pipe" if ss.get("sep_fec") == "Pipe |" else "tab",
                                      sources={os.path.basename(q): q for q in ss.get("index_chemins", {}).values()})
                    pb = [(n, m) for n, m, b in ag.valider_ecritures(nv["ecritures"]) if not b]
                    for e in nv["ecritures"]:
                        e["alerte"] = " ; ".join(m for n, m in pb if n == e["num"]) or e.get("alerte", "")
                    ss.res = {**r, **nv}
                    ss.n_ecr = ss.get("n_ecr", 0) + 1
                    rerun()
                except ValueError as ex:
                    st.error(str(ex))


with tab2:
    st.caption("Si le nom du tiers contient le motif, le compte est utilisé. Colonne type : achat, vente, ou vide pour les deux. "
               "Les règles du modèle sont indicatives : vérifiez-les avec le plan comptable du dossier.")
    if st.button("➕ Ajouter les fournisseurs courants du modèle"):
        base, n = ag.fusionner_modele(ss.get("imput_courant", ss.imput_base))
        ss.imput_base = base
        ss.pop("ed_imput", None)
        st.success(f"{n} règle(s) ajoutée(s).")
    if st.button("🔄 Mettre à jour les comptes carburant / transport / hôtels du modèle"):
        base, n = ag.migrer_modele(ss.get("imput_courant", ss.imput_base))
        ss.imput_base = base
        ss.pop("ed_imput", None)
        st.success(f"{n} règle(s) migrée(s) (les lignes que vous avez modifiées sont conservées).")
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
            detecte = ag.charger_mapping(df, type_=type_)
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
        st.radio("Compte de tiers", ["Au niveau du compte (C/F + 6 caractères)", "Collectif + compte auxiliaire"],
                 key="mode_tiers")
        st.caption("Comptes limités à 7 caractères. Compte de tiers : C + 6 premiers caractères du nom (clients), "
                   "F + 6 (fournisseurs), généré automatiquement.")
        st.markdown("**Comptes de produits par taux de TVA (ventes)**")
        n1, n2 = st.columns([3, 2])
        nature = n1.radio("Nature des ventes", list(ag.NATURES), format_func=lambda k: f"{k} · {ag.NATURES[k]}",
                          key="nature_ventes")
        if n2.button("Proposer les comptes"):
            ss.prod_base = produits_df(ag.produits_par_taux(nature))
            ss.pop("ed_prod", None)
        ss.prod_courant = st.data_editor(ss.prod_base, num_rows="dynamic", hide_index=True,
                                         use_container_width=True, key="ed_prod")
        st.markdown("**Comptes de TVA par taux**")
        ss.tva_courant = st.data_editor(ss.tva_base, num_rows="dynamic", hide_index=True,
                                        use_container_width=True, key="ed_tva")
    with st.expander("🚫 TVA non déductible (achats)"):
        st.checkbox("Appliquer les règles de TVA non déductible", value=True, key="nd_actif")
        st.markdown("**⛽ Véhicule et carburant**")
        v1, v2 = st.columns(2)
        v1.selectbox("Type de véhicule", list(ag.VEHICULES), format_func=ag.VEHICULES.get, key="veh_type")
        v2.selectbox("Carburant / énergie", [AUCUN] + list(ag.CARBURANTS), index=1, key="veh_carb")
        st.text_input("Compte(s) de charge carburant (séparés par une virgule)", ag.COMPTE_CARBURANT_DEFAUT, key="veh_cpt")
        carb = ss.get("veh_carb", "Gazole")
        if carb in ag.CARBURANTS:
            ded = ag.pct_deductible_carburant(ss.get("veh_type", "VP"), carb)
            st.info(f"TVA déductible : **{ded} %** · non déductible : **{100 - ded} %** (ajoutée à la charge).")
            rem = ag.CARBURANTS[carb][2]
            if rem:
                st.warning(rem)
        else:
            st.caption("Aucune règle carburant : la TVA du carburant est déduite en totalité.")
        st.markdown("**📋 Autres règles, par compte**")
        st.caption("Règles de départ INDICATIVES, à valider selon le dossier. Compte : début du compte de charge "
                   "(ex. 6256 couvre 625610) ; % : part de TVA non déductible, ajoutée à la charge.")
        if st.button("Remettre les règles types"):
            ss.nd_base = ag.tva_nd_defaut_df()
            ss.pop("ed_nd", None)
        ss.nd_courant = st.data_editor(ss.nd_base, num_rows="dynamic", hide_index=True,
                                       use_container_width=True, key="ed_nd")
    with st.expander("📄 Format du FEC"):
        st.radio("Séparateur des colonnes", ["Tabulation", "Pipe |"], horizontal=True, key="sep_fec")
        st.caption("Les deux sont admis pour un FEC. Choisissez celui que votre import ACD accepte (certaines sources "
                   "indiquent que le FEC pour ACD utilise le séparateur pipe : à confirmer auprès d'ACD).")
    with st.expander("💾 Sauvegarde des réglages"):
        c, t, p = comptes_ui()
        st.download_button("⬇️ Télécharger mes réglages (reglages.json)",
                           json.dumps(ag.exporter_reglages(c, t, p, ss.get("nd_courant", ss.nd_base).to_dict("records"),
                                                   {"vehicule": ss.get("veh_type", "VP"), "carburant": ss.get("veh_carb", "Gazole"),
                                                    "comptes": ss.get("veh_cpt", ag.COMPTE_CARBURANT_DEFAUT)}), ensure_ascii=False, indent=2).encode("utf-8"),
                           file_name="reglages.json", mime="application/json")
        st.caption("Pour les garder : ajoutez ce fichier à la racine de GitHub, il sera rechargé au démarrage.")
        rg = st.file_uploader("Restaurer des réglages (reglages.json)", type="json", key="reg_up")
        if rg is not None and ss.get("reg_nom") != (rg.name, rg.size):
            ss.reg_nom = (rg.name, rg.size)
            ss.reglages_a_charger = json.loads(rg.getvalue().decode("utf-8"))
            rerun()
