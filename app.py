"""Interface : streamlit run app.py"""
import io, os, tempfile
import pandas as pd
import streamlit as st
import agent_fec as ag

st.set_page_config(page_title="Factures → FEC", page_icon="🧾")
st.title("🧾 Factures → FEC pour ACD")
st.caption("CSV/Excel : règles · PDF Factur-X : XML intégré · PDF texte : règles · Scans : OCR si installé, sinon IA (clé facultative)")
ss = st.session_state

LIBELLES = {"vente": "🧑‍💼 Clients (factures de vente)", "achat": "🏭 Fournisseurs (factures d'achat)"}
type_ = st.radio("Je traite des…", list(LIBELLES), format_func=LIBELLES.get, horizontal=True)
col1, col2 = st.columns(2)
siren = col1.text_input("SIREN du dossier", "000000000")
exercice = col2.text_input("Exercice (AAAA)", "2026")

with st.sidebar:
    cle = st.text_input("Clé API Anthropic (facultative, secours IA)", type="password")
    if cle:
        os.environ["ANTHROPIC_API_KEY"] = cle

if "imput" not in ss:
    ss.imput = ag.charger_imputations()

with st.expander("📒 Table d'imputation (compte par fournisseur)"):
    st.caption("Si le nom du tiers contient le motif, le compte est utilisé. Pour la conserver : téléchargez-la, "
               "puis remplacez imputations.csv dans GitHub (ou réimportez-la ici à chaque session).")
    up = st.file_uploader("Importer ma table (CSV)", type="csv", key="imp_up")
    if up is not None and ss.get("imp_nom") != (up.name, up.size):
        tmp_csv = os.path.join(tempfile.mkdtemp(), "imputations.csv")
        open(tmp_csv, "wb").write(up.getbuffer())
        ss.imput, ss.imp_nom = ag.charger_imputations(tmp_csv), (up.name, up.size)
    ss.imput = st.data_editor(ss.imput, num_rows="dynamic", use_container_width=True, hide_index=True, key="ed_imput")
    st.download_button("⬇️ Télécharger la table", ss.imput.to_csv(index=False, sep=";").encode("utf-8"),
                       file_name="imputations.csv", mime="text/csv")

def lancer():
    try:
        with st.spinner("Analyse en cours…"):
            ss.res = ag.traiter(ss.chemins, type_, siren, exercice, ss.tmp, regles=ss.imput)
        ss.err = None
    except ValueError as e:
        ss.res, ss.err = None, str(e)

fichiers = st.file_uploader("Déposez vos exports ou factures", accept_multiple_files=True,
                            type=["csv", "xlsx", "xls", "pdf", "xml", "png", "jpg", "jpeg"])

if fichiers and st.button(f"Générer le FEC ({'clients' if type_ == 'vente' else 'fournisseurs'})", type="primary"):
    ss.tmp, ss.chemins = tempfile.mkdtemp(), []
    for f in fichiers:
        p = os.path.join(ss.tmp, f.name)
        open(p, "wb").write(f.getbuffer())
        ss.chemins.append(p)
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

    if r["sans_regle"]:
        st.subheader("Comptes à renseigner")
        st.caption("Ces tiers n'ont pas de règle : le compte par défaut a été utilisé. "
                   "Indiquez le compte (ex. 606100) pour le mémoriser et régénérer le FEC.")
        nouv = st.data_editor(pd.DataFrame({"motif": r["sans_regle"], "compte": "", "libelle": ""}),
                              hide_index=True, use_container_width=True, key="nouv_regles")
        if st.button("Enregistrer ces règles et régénérer"):
            ok = nouv[nouv["compte"].astype(str).str.strip() != ""]
            if not ok.empty:
                ss.imput = pd.concat([ss.imput, ok], ignore_index=True)
                lancer()
                st.rerun()

    rest = [a for a in r["anomalies"] if not a[1].startswith("Aucune règle d'imputation")]
    if rest:
        st.warning("Points à vérifier avant import dans ACD :")
        st.dataframe(pd.DataFrame(rest, columns=["Fichier:ligne", "Anomalie", "Détail"]), use_container_width=True)
        st.download_button("Télécharger le rapport", open(r["rapport"], "rb").read(), file_name="rapport_anomalies.csv")
    st.subheader("Aperçu des écritures")
    st.dataframe(pd.DataFrame(r["ecritures"])[["num", "date", "compte", "aux", "piece", "lib", "debit", "credit"]],
                 use_container_width=True)
