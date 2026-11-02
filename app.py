"""Interface : streamlit run app.py"""
import os, tempfile
import pandas as pd
import streamlit as st
import agent_fec as ag

st.set_page_config(page_title="Factures → FEC", page_icon="🧾")
st.title("🧾 Factures → FEC pour ACD")

with st.sidebar:
    type_ = st.radio("Type de factures", ["vente", "achat"], format_func=str.capitalize)
    siren = st.text_input("SIREN du dossier", "000000000")
    exercice = st.text_input("Exercice (AAAA)", "2026")
    cle = st.text_input("Clé API Anthropic (PDF/images)", type="password")
    if cle:
        os.environ["ANTHROPIC_API_KEY"] = cle

fichiers = st.file_uploader("Déposez vos exports ou factures", accept_multiple_files=True,
                            type=["csv", "xlsx", "xls", "pdf", "png", "jpg", "jpeg"])

if fichiers and st.button("Générer le FEC", type="primary"):
    tmp = tempfile.mkdtemp()
    chemins = []
    for f in fichiers:
        p = os.path.join(tmp, f.name)
        open(p, "wb").write(f.getbuffer())
        chemins.append(p)
    try:
        with st.spinner("Analyse en cours…"):
            r = ag.traiter(chemins, type_, siren, exercice, tmp)
    except ValueError as e:
        st.error(str(e)); st.stop()

    c1, c2, c3 = st.columns(3)
    c1.metric("Lignes d'écriture", len(r["ecritures"]))
    c2.metric("Total débit = crédit", f"{r['total']:,.2f} €".replace(",", " "))
    c3.metric("Anomalies", len(r["anomalies"]))

    st.download_button("⬇️ Télécharger le FEC", open(r["fec"], "rb").read(),
                       file_name=os.path.basename(r["fec"]), mime="text/plain", type="primary")
    if r["anomalies"]:
        st.warning("Points à vérifier avant import dans ACD :")
        st.dataframe(pd.DataFrame(r["anomalies"], columns=["Fichier:ligne", "Anomalie", "Détail"]),
                     use_container_width=True)
        st.download_button("Télécharger le rapport", open(r["rapport"], "rb").read(),
                           file_name="rapport_anomalies.csv")
    st.subheader("Aperçu des écritures")
    st.dataframe(pd.DataFrame(r["ecritures"])[["num", "date", "compte", "aux", "piece", "lib", "debit", "credit"]],
                 use_container_width=True)
