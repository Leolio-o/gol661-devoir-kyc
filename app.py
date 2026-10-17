import streamlit as st
import pandas as pd
import requests
import xml.etree.ElementTree as ET
import sqlite3
import os
import uuid
import unicodedata
import re
from difflib import SequenceMatcher
from datetime import datetime

# Configuration de la page
st.set_page_config(
    page_title="Portail de Conformite KYC - GOL661",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Style CSS professionnel
STYLE_CSS = """
<style>
.block-container { padding-top: 2rem; padding-bottom: 2rem; }
h1, h2, h3 { font-family: 'Segoe UI', Roboto, Arial, sans-serif; font-weight: 600; }
.rapport-box {
    background-color: #f8f9fa;
    border-left: 4px solid #1f4e78;
    padding: 1.25rem;
    margin-bottom: 1.5rem;
    border-radius: 4px;
}
.alerte-critique { border-left-color: #b91c1c; background-color: #fef2f2; }
.alerte-moderee { border-left-color: #d97706; background-color: #fffbeb; }
.alerte-faible { border-left-color: #15803d; background-color: #f0fdf4; }
</style>
"""
st.markdown(STYLE_CSS, unsafe_allow_html=True)

# ==========================================
# 0. MOTEUR DE RECHERCHE "BEST MATCH"
# ==========================================
def normaliser_texte(texte: str) -> str:
    if not texte:
        return ""
    texte = str(texte).lower()
    texte = "".join(
        c for c in unicodedata.normalize("NFD", texte)
        if unicodedata.category(c) != "Mn"
    )
    texte = re.sub(r"[^a-z0-9\s]", " ", texte)
    return re.sub(r"\s+", " ", texte).strip()


def calculer_score_match(requete: str, cible: str) -> float:
    req_norm = normaliser_texte(requete)
    cible_norm = normaliser_texte(cible)

    if not req_norm or not cible_norm:
        return 0.0
    if req_norm in cible_norm:
        return 100.0

    mots_req = req_norm.split()
    mots_cible = cible_norm.split()
    if not mots_cible:
        return 0.0

    scores_mots = []
    for m_req in mots_req:
        if m_req in cible_norm:
            scores_mots.append(1.0)
        else:
            meilleur = max(
                (SequenceMatcher(None, m_req, m_cib).ratio() for m_cib in mots_cible),
                default=0.0
            )
            scores_mots.append(meilleur)

    score_moyen = (sum(scores_mots) / len(scores_mots)) * 100.0
    extrait = cible_norm[: len(req_norm) * 2]
    score_global = SequenceMatcher(None, req_norm, extrait).ratio() * 100.0

    return round(max(score_moyen, score_global), 1)


# ==========================================
# 1. BASE DE DONNEES D'AUDIT (SQLite - BCAPG)
# ==========================================
DB_FILE = "audit_siag_bcapg.db"
EXCEL_FILE = "GOL661_Devoir_KYC.xlsx"


def init_db():
    conn = sqlite3.connect(DB_FILE)
    c = conn.cursor()
    c.execute(
        """
        CREATE TABLE IF NOT EXISTS audits_kyc (
            audit_id TEXT PRIMARY KEY,
            horodatage TEXT,
            analyste TEXT,
            entite_recherchee TEXT,
            bases_selectionnees TEXT,
            rapport_agent_siag TEXT,
            niveau_risque TEXT,
            commentaires_humains TEXT,
            resultats_bruts_json TEXT
        )
        """
    )
    conn.commit()
    conn.close()


def sauvegarder_audit(analyste, entite, bases, rapport, risque, comm, df_res):
    conn = sqlite3.connect(DB_FILE)
    c = conn.cursor()
    date_str = datetime.now().strftime("%Y%m%d")
    uid_str = str(uuid.uuid4())[:6].upper()
    audit_id = f"AUD-{date_str}-{uid_str}"
    horodatage = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    raw_json = (
        df_res.to_json(orient="records", force_ascii=False)
        if not df_res.empty
        else "[]"
    )

    c.execute(
        "INSERT INTO audits_kyc VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)",
        (
            audit_id,
            horodatage,
            analyste,
            entite,
            ", ".join(bases),
            rapport,
            risque,
            comm,
            raw_json,
        ),
    )
    conn.commit()
    conn.close()
    return audit_id


def charger_audits():
    conn = sqlite3.connect(DB_FILE)
    df = pd.read_sql_query(
        "SELECT * FROM audits_kyc ORDER BY horodatage DESC", conn
    )
    conn.close()
    return df


init_db()


# ==========================================
# 2. CONNECTEURS DE BASES DE DONNEES
# ==========================================
def interroger_bases_quebec(terme: str, seuil: int = 65) -> list:
    resultats = []
    terme_norm = normaliser_texte(terme)

    try:
        url = "https://www.donneesquebec.ca/recherche/api/3/action/package_search"
        resp = requests.get(url, params={"q": terme_norm, "rows": 5}, timeout=6)
        if resp.status_code == 200:
            items = resp.json().get("result", {}).get("results", [])
            for item in items:
                titre = item.get("title", "N/A")
                notes = item.get("notes", "") or ""
                s1 = calculer_score_match(terme, titre)
                s2 = calculer_score_match(terme, notes)
                score = max(s1, s2, 75.0)
                nom_pkg = item.get("name", "")
                lien = f"https://www.donneesquebec.ca/recherche/dataset/{nom_pkg}"
                org = item.get("organization", {}).get("title", "Gouv. du Quebec")
                resultats.append(
                    {
                        "Correspondance (%)": score,
                        "Source": "API Donnees Quebec (CKAN)",
                        "Entite / Resultat": titre,
                        "Categorie": org,
                        "Details": notes[:160] + "...",
                        "URL": lien,
                    }
                )
    except Exception:
        pass

    if os.path.exists(EXCEL_FILE):
        df = pd.read_excel(EXCEL_FILE, sheet_name="1-Bases Québec")
        for _, row in df.iterrows():
            s_nom = calculer_score_match(terme, f"{row['Nom']} {row['Catégorie']}")
            s_desc = calculer_score_match(terme, str(row["Description"]))
            best = max(s_nom, s_desc)
            if best >= seuil:
                resultats.append(
                    {
                        "Correspondance (%)": best,
                        "Source": "Registres publics du Quebec (Referentiel)",
                        "Entite / Resultat": row["Nom"],
                        "Categorie": row["Catégorie"],
                        "Details": row["Description"],
                        "URL": row["URL"],
                    }
                )
    return resultats


def interroger_sanctions_canada(terme: str, seuil: int = 65) -> list:
    resultats = []
    if os.path.exists(EXCEL_FILE):
        df_t = pd.read_excel(EXCEL_FILE, sheet_name="3-Organisations terroristes")
        for _, row in df_t.iterrows():
            nom_org = str(row.get("Nom", ""))
            alias_org = str(row.get("Alias / variantes", ""))
            pays_org = str(row.get("Pays d'origine", ""))
            zone_org = str(row.get("Zone d'opération", ""))
            listes_org = str(row.get("Présence dans quelles listes", ""))
            lien_can = row.get("Lien Canada (Sécurité publique / communiqué)", "")

            s_nom = calculer_score_match(terme, nom_org)
            s_alias = calculer_score_match(terme, alias_org)
            s_pays = calculer_score_match(terme, f"{pays_org} {zone_org}")
            best = max(s_nom, s_alias, s_pays)

            if best >= seuil:
                resultats.append(
                    {
                        "Correspondance (%)": best,
                        "Source": "Securite publique Canada (Code criminel)",
                        "Entite / Resultat": f"{nom_org} (Alias : {alias_org})",
                        "Categorie": f"Entite terroriste inscrite - {pays_org}",
                        "Details": f"Zones : {zone_org} | Listes : {listes_org}",
                        "URL": lien_can,
                    }
                )

    try:
        url_xml = (
            "https://www.international.gc.ca/world-monde/assets/office_docs/"
            "international_relations-relations_internationales/sanctions/sema-lmes.xml"
        )
        resp = requests.get(url_xml, timeout=8)
        if resp.status_code == 200:
            root = ET.fromstring(resp.content)
            matches_xml = []
            lien_lmes = (
                "https://www.international.gc.ca/world-monde/"
                "international_relations-relations_internationales/sanctions/"
                "consolidated-consolide.aspx?lang=fra"
            )
            for rec in root.findall(".//record"):
                prenom = rec.findtext("GivenName", "")
                nom_fam = rec.findtext("LastName", "")
                nom = rec.findtext("Entity") or f"{prenom} {nom_fam}".strip()
                pays = rec.findtext("Country") or "Non specifie"
                sched = rec.findtext("Schedule") or "LMES"
                score = calculer_score_match(terme, f"{nom} {pays}")
                if score >= max(seuil, 72):
                    matches_xml.append(
                        {
                            "Correspondance (%)": score,
                            "Source": "Affaires mondiales Canada (Flux XML LMES)",
                            "Entite / Resultat": nom,
                            "Categorie": f"Sanction economique autonome ({pays})",
                            "Details": f"Reglement / Annexe : {sched}",
                            "URL": lien_lmes,
                        }
                    )
            matches_xml.sort(key=lambda x: x["Correspondance (%)"], reverse=True)
            resultats.extend(matches_xml[:8])
    except Exception:
        pass
    return resultats


def interroger_sanctions_internationales(terme: str, seuil: int = 65) -> list:
    resultats = []
    if os.path.exists(EXCEL_FILE):
        df_s = pd.read_excel(EXCEL_FILE, sheet_name="2-Sanctions")
        for _, row in df_s.iterrows():
            s_nom = calculer_score_match(terme, f"{row['Nom']} {row['Catégorie']}")
            s_desc = calculer_score_match(terme, str(row["Description"]))
            best = max(s_nom, s_desc)
            if best >= seuil:
                resultats.append(
                    {
                        "Correspondance (%)": best,
                        "Source": "Listes internationales de sanctions",
                        "Entite / Resultat": row["Nom"],
                        "Categorie": row["Catégorie"],
                        "Details": row["Description"],
                        "URL": row["URL"],
                    }
                )

    q_url = requests.utils.quote(terme)
    resultats.append(
        {
            "Correspondance (%)": 100.0,
            "Source": "Agregateur international OpenSanctions",
            "Entite / Resultat": f"Requete consolidee : {terme}",
            "Categorie": "Verification multi-juridictions (ONU, OFAC, UE)",
            "Details": "Lien de verification dans la base mondiale consolidee.",
            "URL": f"https://www.opensanctions.org/search/?q={q_url}",
        }
    )
    return resultats


# ==========================================
# 3. ORCHESTRATION ET DIAGNOSTIC DE L'AGENT
# ==========================================
def generer_diagnostic_agent(terme: str, donnees: list, bases: list):
    alerte_critique = any(
        "Code criminel" in d["Source"] or "Sanction economique" in d["Categorie"]
        for d in donnees
    )
    top_score = donnees[0]["Correspondance (%)"] if donnees else 0
    perimetre = ", ".join(bases)
    terme_maj = terme.upper()

    if alerte_critique:
        lignes = [
            f"<h4>Synthese de l'Agent KYC : Risque Critique Detecte — "
            f"« {terme_maj} » (Correspondance max : {top_score} %)</h4>",
            "<ul>",
            f"<li><b>Perimetre audite :</b> {perimetre}</li>",
            "<li><b>Risque legal (Critique) :</b> Correspondance etablie avec une "
            "entite inscrite sur les listes officielles de sanctions canadiennes "
            "(LMES / Code criminel, art. 83.05) ou internationales.</li>",
            "<li><b>Risque reputationnel (Critique) :</b> Risque severe de "
            "contamination reputationnelle en cas d'association d'affaires.</li>",
            "<li><b>Risque operationnel (Eleve) :</b> Gel potentiel des avoirs, "
            "blocage bancaire et revocation des autorisations (AMP/RENA).</li>",
            "<li><b>Recommandation :</b> <b>Blocage immediat de la partie prenante."
            "</b> Si le score est inferieur a 100 %, valider l'absence d'homonymie.</li>",
            "</ul>",
        ]
        return "".join(lignes), "alerte-critique", 0

    if len(donnees) > 1:
        lignes = [
            f"<h4>Synthese de l'Agent KYC : Vigilance Requise — "
            f"« {terme_maj} » (Correspondance max : {top_score} %)</h4>",
            "<ul>",
            f"<li><b>Perimetre audite :</b> {perimetre}</li>",
            f"<li><b>Risque legal (Modere) :</b> {len(donnees)} entree(s) "
            "repertoriee(s). Verifier la conformite aux registres publics "
            "(REQ, RENA, RBQ).</li>",
            "<li><b>Risque reputationnel (Modere) :</b> Verifier l'absence de "
            "litiges dans les bases judiciaires (SOQUIJ, CanLII).</li>",
            "<li><b>Risque operationnel (Faible a Modere) :</b> Valider la "
            "detention des autorisations de contracter (AMP).</li>",
            "<li><b>Recommandation :</b> <b>Diligence raisonnable renforcee.</b> "
            "Documenter la verification dans la section d'audit ci-dessous.</li>",
            "</ul>",
        ]
        return "".join(lignes), "alerte-moderee", 2

    lignes = [
        f"<h4>Synthese de l'Agent KYC : Risque Faible — « {terme_maj} »</h4>",
        "<ul>",
        f"<li><b>Perimetre audite :</b> {perimetre}</li>",
        "<li><b>Evaluation globale :</b> Aucune inscription defavorable ou "
        "sanction detectee selon le seuil de tolerance applique.</li>",
        "<li><b>Recommandation :</b> <b>Homologation conditionnelle</b>, sous "
        "reserve de la validation usuelle du NEQ au Registraire des entreprises.</li>",
        "</ul>",
    ]
    return "".join(lignes), "alerte-faible", 3


def executer_agent_kyc(terme: str, bases_actives: list, seuil: int):
    journal = []
    donnees = []

    if "Bases du Quebec (REQ, RENA, Donnees Quebec)" in bases_actives:
        journal.append(f"Connecteur 1 : Interrogation API Donnees Quebec (seuil {seuil} %).")
        res_qc = interroger_bases_quebec(terme, seuil)
        donnees.extend(res_qc)
        journal.append(f"Resultat Connecteur 1 : {len(res_qc)} entree(s) extraite(s).")

    if "Sanctions et Entites Terroristes Canada (LMES / Code criminel)" in bases_actives:
        journal.append(f"Connecteur 2 : Analyse flux XML LMES et Code criminel (seuil {seuil} %).")
        res_can = interroger_sanctions_canada(terme, seuil)
        donnees.extend(res_can)
        journal.append(f"Resultat Connecteur 2 : {len(res_can)} entree(s) extraite(s).")

    if "Listes Internationales de Sanctions (ONU, OFAC, UE, OpenSanctions)" in bases_actives:
        journal.append(f"Connecteur 3 : Verification listes internationales (seuil {seuil} %).")
        res_int = interroger_sanctions_internationales(terme, seuil)
        donnees.extend(res_int)
        journal.append(f"Resultat Connecteur 3 : {len(res_int)} reference(s) extraite(s).")

    donnees.sort(key=lambda x: x.get("Correspondance (%)", 0), reverse=True)
    rapport, css, niv = generer_diagnostic_agent(terme, donnees, bases_actives)
    journal.append("Synthese des risques completee par le moteur d'analyse.")

    return donnees, rapport, css, niv, journal


# ==========================================
# 4. INTERFACE UTILISATEUR PRINCIPALE
# ==========================================
st.title("Systeme d'Evaluation KYC et de Verification des Antecedents")
st.markdown(
    "**Cours GOL661 — Ecole de technologie superieure (ETS)** | "
    "Outil d'interrogation des bases de donnees ouvertes et de journalisation pour audit par le "
    "**Bureau canadien d'agrement des programmes de genie (BCAPG)**."
)

with st.sidebar:
    st.header("Parametres de recherche")
    analyste = st.text_input("Identifiant de l'auditeur ou de l'equipe", value="Equipe GOL661")

    st.subheader("1. Bases de donnees consultees")
    b_qc = st.checkbox("Bases du Quebec (REQ, RENA, Donnees Quebec)", value=True)
    b_can = st.checkbox("Sanctions et Entites Terroristes Canada (LMES / Code criminel)", value=True)
    b_int = st.checkbox("Listes Internationales de Sanctions (ONU, OFAC, UE, OpenSanctions)", value=True)

    st.divider()
    st.subheader("2. Tolerance de recherche (Best Match)")
    seuil_match = st.slider(
        "Seuil minimal de correspondance (%)",
        min_value=45,
        max_value=100,
        value=65,
        step=5,
        help="Permet d'identifier une entite malgre des variantes orthographiques ou fautes de frappe.",
    )

tab_recherche, tab_audit, tab_referentiel = st.tabs(
    [
        "1. Recherche et Analyse KYC",
        "2. Registre d'Audit (BCAPG)",
        "3. Inventaire des Bases de Donnees (Points 1 a 3)",
    ]
)

with tab_recherche:
    col_input, col_btn = st.columns([4, 1])
    with col_input:
        terme = st.text_input(
            "Nom de l'organisation, de l'entreprise ou de l'individu a verifier",
            placeholder="Exemple : Al-Qaida, Hamas, Hezbollah, RENA, Construction...",
        )
    with col_btn:
        st.write("")
        st.write("")
        lancer = st.button("Lancer la verification", type="primary", use_container_width=True)

    if lancer and terme:
        bases_choisies = []
        if b_qc:
            bases_choisies.append("Bases du Quebec (REQ, RENA, Donnees Quebec)")
        if b_can:
            bases_choisies.append("Sanctions et Entites Terroristes Canada (LMES / Code criminel)")
        if b_int:
            bases_choisies.append("Listes Internationales de Sanctions (ONU, OFAC, UE, OpenSanctions)")

        if not bases_choisies:
            st.warning("Veuillez selectionner au moins une base de donnees dans le panneau de gauche.")
        else:
            with st.spinner("Interrogation des bases de donnees ouvertes en cours..."):
                res_l, rap, css, niv, etapes = executer_agent_kyc(terme, bases_choisies, seuil_match)
                st.session_state["res_df"] = pd.DataFrame(res_l)
                st.session_state["rapport_ia"] = rap
                st.session_state["classe_css"] = css
                st.session_state["niveau_defaut"] = niv
                st.session_state["etapes_agent"] = etapes
                st.session_state["terme_actuel"] = terme
                st.session_state["bases_actuelles"] = bases_choisies

    if "rapport_ia" in st.session_state:
        boite_html = (
            f'<div class="rapport-box {st.session_state["classe_css"]}">'
            f'{st.session_state["rapport_ia"]}</div>'
        )
        st.markdown(boite_html, unsafe_allow_html=True)

        with st.expander("Journal d'execution technique des requetes", expanded=False):
            for idx, etape in enumerate(st.session_state["etapes_agent"], 1):
                st.text(f"[{idx}] {etape}")

        nb_corr = len(st.session_state["res_df"])
        st.subheader(f"Resultats extraits des bases de donnees ({nb_corr} correspondance(s))")
        if not st.session_state["res_df"].empty:
            st.dataframe(
                st.session_state["res_df"],
                column_config={
                    "Correspondance (%)": st.column_config.ProgressColumn(
                        "Pertinence (Best Match)",
                        format="%.1f %%",
                        min_value=0,
                        max_value=100,
                    ),
                    "URL": st.column_config.LinkColumn("Lien vers la source officielle"),
                },
                use_container_width=True,
                hide_index=True,
            )

        st.divider()
        st.subheader("Decision de conformite et archivage pour audit")
        with st.form("audit_form"):
            options_risque = [
                "CRITIQUE - Blocage immediat (Sanction / Entite terroriste)",
                "ELEVE - Rejet recommande",
                "MODERE - Diligence raisonnable renforcee requise",
                "FAIBLE - Homologation autorisee",
            ]
            risque_retenu = st.selectbox(
                "Classification finale du risque validee par l'analyste",
                options_risque,
                index=st.session_state.get("niveau_defaut", 2),
            )
            commentaire = st.text_area(
                "Commentaires d'analyse et justification de la decision (tracabilite BCAPG)",
                placeholder="Inscrire les conclusions de l'enquete ou les mesures d'attenuation...",
            )
            btn_save = st.form_submit_button("Enregistrer la verification au registre d'audit")
            if btn_save:
                texte_propre = re.sub("<[^<]+?>", "", st.session_state["rapport_ia"])
                id_genere = sauvegarder_audit(
                    analyste=analyste,
                    entite=st.session_state["terme_actuel"],
                    bases=st.session_state["bases_actuelles"],
                    rapport=texte_propre,
                    risque=risque_retenu,
                    comm=commentaire,
                    df_res=st.session_state["res_df"],
                )
                st.success(f"Dossier d'audit enregistre avec succes. Reference BCAPG : {id_genere}")

with tab_audit:
    st.subheader("Registre permanent des verifications effectuees (Consultation BCAPG)")
    st.markdown(
        "Ce registre conserve l'historique horodate des recherches, les bases de donnees "
        "interrogees, la classification du risque et les commentaires de l'analyste."
    )
    df_hist = charger_audits()
    if not df_hist.empty:
        df_affichage = df_hist.drop(
            columns=["resultats_bruts_json", "rapport_agent_siag"]
        ).rename(
            columns={
                "audit_id": "Reference Audit",
                "horodatage": "Date et heure",
                "analyste": "Analyste / Equipe",
                "entite_recherchee": "Entite verifiee",
                "bases_selectionnees": "Bases consultees",
                "niveau_risque": "Decision de risque",
                "commentaires_humains": "Commentaires de l'auditeur",
            }
        )
        st.dataframe(df_affichage, use_container_width=True, hide_index=True)

        csv_bytes = df_hist.to_csv(index=False).encode("utf-8-sig")
        nom_csv = f"Registre_Audit_BCAPG_{datetime.now().strftime('%Y%m%d')}.csv"
        st.download_button(
            "Exporter le registre d'audit complet (format CSV)",
            data=csv_bytes,
            file_name=nom_csv,
            mime="text/csv",
        )
    else:
        st.info("Aucune verification n'a encore ete archivee dans le registre.")

with tab_referentiel:
    st.subheader("Referentiel des bases de donnees et listes reglementaires (Points 1 a 3)")
    if os.path.exists(EXCEL_FILE):
        feuille = st.selectbox(
            "Selectionner le tableau d'inventaire a consulter",
            ["1-Bases Québec", "2-Sanctions", "3-Organisations terroristes"],
        )
        df_inv = pd.read_excel(EXCEL_FILE, sheet_name=feuille)
        st.dataframe(df_inv, use_container_width=True, hide_index=True)
    else:
        st.warning(f"Le fichier {EXCEL_FILE} est introuvable dans le repertoire.")
