import streamlit as st
import pandas as pd
import requests
import xml.etree.ElementTree as ET
import sqlite3
import json
import os
import uuid
import unicodedata
import re
from difflib import SequenceMatcher
from datetime import datetime

# Configuration de la page (sans émojis, style corporatif)
st.set_page_config(
    page_title="Portail de Conformité KYC - GOL661",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Style CSS professionnel
st.markdown("""
    <style>
    .block-container {
        padding-top: 2rem;
        padding-bottom: 2rem;
    }
    h1, h2, h3 {
        font-family: 'Segoe UI', Roboto, Helvetica, Arial, sans-serif;
        font-weight: 600;
    }
    .rapport-box {
        background-color: #f8f9fa;
        border-left: 4px solid #1f4e78;
        padding: 1.25rem;
        margin-bottom: 1.5rem;
        border-radius: 4px;
    }
    .alerte-critique {
        border-left-color: #b91c1c;
        background-color: #fef2f2;
    }
    .alerte-moderee {
        border-left-color: #d97706;
        background-color: #fffbeb;
    }
    .alerte-faible {
        border-left-color: #15803d;
        background-color: #f0fdf4;
    }
    </style>
""", unsafe_allow_html=True)

# ==========================================
# 0. MOTEUR DE RECHERCHE "BEST MATCH" (FUZZY)
# ==========================================
def normaliser_texte(texte: str) -> str:
    """Normalise la chaîne : minuscules, suppression des accents et de la ponctuation."""
    if not texte:
        return ""
    texte = str(texte).lower()
    texte = ''.join(
        c for c in unicodedata.normalize('NFD', texte)
        if unicodedata.category(c) != 'Mn'
    )
    texte = re.sub(r'[^a-z0-9\s]', ' ', texte)
    return re.sub(r'\s+', ' ', texte).strip()

def calculer_score_match(requete: str, cible: str) -> float:
    """Calcule un indice de similarité (0 à 100 %) insensible à la casse, aux accents et aux fautes de frappe."""
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
            meilleur_mot = max(
                (SequenceMatcher(None, m_req, m_cib).ratio() for m_cib in mots_cible),
                default=0.0
            )
            scores_mots.append(meilleur_mot)

    score_moyen = (sum(scores_mots) / len(scores_mots)) * 100.0
    score_global = SequenceMatcher(None, req_norm, cible_norm[:len(req_norm)*2]).ratio() * 100.0
    
    return round(max(score_moyen, score_global), 1)

# ==========================================
# 1. BASE DE DONNÉES D'AUDIT (SQLite - BCAPG)
# ==========================================
DB_FILE = "audit_siag_bcapg.db"

def init_db():
    conn = sqlite3.connect(DB_FILE)
    c = conn.cursor()
    c.execute('''
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
    ''')
    conn.commit()
    conn.close()

def sauvegarder_audit(analyste, entite, bases, rapport_ia, risque, commentaire, df_res):
    conn = sqlite3.connect(DB_FILE)
    c = conn.cursor()
    audit_id = f"AUD-{datetime.now().strftime('%Y%m%d')}-{str(uuid.uuid4())[:6].upper()}"
    horodatage = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    raw_json = df_res.to_json(orient="records", force_ascii=False) if not df_res.empty else "[]"
    
    c.execute('''
        INSERT INTO audits_kyc VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
    ''', (audit_id, horodatage, analyste, entite, ", ".join(bases), rapport_ia, risque, commentaire, raw_json))
    conn.commit()
    conn.close()
    return audit_id

def charger_audits():
    conn = sqlite3.connect(DB_FILE)
    df = pd.read_sql_query("SELECT * FROM audits_kyc ORDER BY horodatage DESC", conn)
    conn.close()
    return df

init_db()

# ==========================================
# 2. CONNECTEURS DE BASES DE DONNÉES
# ==========================================
def interroger_bases_quebec(terme: str, seuil: int = 65) -> list:
    resultats = []
    terme_norm = normaliser_texte(terme)
    
    try:
        url = "https://www.donneesquebec.ca/recherche/api/3/action/package_search"
        resp = requests.get(url, params={"q": terme_norm, "rows": 5}, timeout=6)
        if resp.status_code == 200:
            for item in resp.json().get("result", {}).get("results", []):
                titre = item.get("title", "N/A")
                notes = item.get("notes", "") or ""
                score = max(calculer_score_match(terme, titre), calculer_score_match(terme, notes), 75.0)
                resultats.append({
                    "Correspondance (%)": score,
                    "Source": "API Données Québec (CKAN)",
                    "Entité / Résultat": titre,
                    "Catégorie": item.get("organization", {}).get("title", "Gouvernement du Québec"),
                    "Détails": notes[:160] + "...",
                    "URL": f"https://www.donneesquebec.ca/recherche/dataset/{item.get('name', '')}"
                })
    except Exception:
        pass

    if os.path.exists("GOL661_Devoir_KYC.xlsx"):
        df = pd.read_excel("GOL661_Devoir_KYC.xlsx", sheet_name="1-Bases Québec")
        for _, row in df.iterrows():
            score_nom = calculer_score_match(terme, f"{row['Nom']} {row['Catégorie']}")
            score_desc = calculer_score_match(terme, str(row['Description']))
            meilleur_score = max(score_nom, score_desc)
            if meilleur_score >= seuil:
                resultats.append({
                    "Correspondance (%)": meilleur_score,
                    "Source": "Registres publics du Québec (Référentiel)",
                    "Entité / Résultat": row["Nom"],
                    "Catégorie": row["Catégorie"],
                    "Détails": row["Description"],
                    "URL": row["URL"]
                })
    return resultats

def interroger_sanctions_canada(terme: str, seuil: int = 65) -> list:
    resultats = []
    if os.path.exists("GOL661_Devoir_KYC.xlsx"):
        df_t = pd.read_excel("GOL661_Devoir_KYC.xlsx", sheet_name="3-Organisations terroristes")
        for _, row in df_t.iterrows():
            score_nom = calculer_score_match(terme, str(row['Nom']))
            score_alias = calculer_score_match(terme, str(row['Alias / variantes']))
            score_pays = calculer_score_match(terme, f"{row['Pays d\'origine']} {row['Zone d\'opération']}")
            meilleur_score = max(score_nom, score_alias, score_pays)
            
            if meilleur_score >= seuil:
                resultats.append({
                    "Correspondance (%)": meilleur_score,
                    "Source": "Sécurité publique Canada (Code criminel)",
                    "Entité / Résultat": f"{row['Nom']} (Alias : {row['Alias / variantes']})",
                    "Catégorie": f"Entité terroriste inscrite - {row['Pays d\'origine']}",
                    "Détails": f"Zones d'opération : {row['Zone d\'opération']} | Listes : {row['Présence dans quelles listes']}",
                    "URL": row["Lien Canada (Sécurité publique / communiqué)"]
                })

    try:
        url_xml = "https://www.international.gc.ca/world-monde/assets/office_docs/international_relations-relations_internationales/sanctions/sema-lmes.xml"
        resp = requests.get(url_xml, timeout=8)
        if resp.status_code == 200:
            root = ET.fromstring(resp.content)
            matches_xml = []
            for rec in root.findall(".//record"):
                nom = rec.findtext("Entity") or f"{rec.findtext('GivenName', '')} {rec.findtext('LastName', '')}".strip()
                pays = rec.findtext("Country") or "Non spécifié"
                cible = f"{nom} {pays}"
                score = calculer_score_match(terme, cible)
                if score >= max(seuil, 72):
                    matches_xml.append({
                        "Correspondance (%)": score,
                        "Source": "Affaires mondiales Canada (Flux XML LMES)",
                        "Entité / Résultat": nom,
                        "Catégorie": f"Sanction économique autonome ({pays})",
                        "Détails": f"Règlement / Annexe : {rec.findtext('Schedule') or 'LMES'}",
                        "URL": "https://www.international.gc.ca/world-monde/international_relations-relations_internationales/sanctions/consolidated-consolide.aspx?lang=fra"
                    })
            matches_xml = sorted(matches_xml, key=lambda x: x["Correspondance (%)"], reverse=True)[:8]
            resultats.extend(matches_xml)
    except Exception:
        pass
    return resultats

def interroger_sanctions_internationales(terme: str, seuil: int = 65) -> list:
    resultats = []
    if os.path.exists("GOL661_Devoir_KYC.xlsx"):
        df_s = pd.read_excel("GOL661_Devoir_KYC.xlsx", sheet_name="2-Sanctions")
        for _, row in df_s.iterrows():
            score_nom = calculer_score_match(terme, f"{row['Nom']} {row['Catégorie']}")
            score_desc = calculer_score_match(terme, str(row['Description']))
            meilleur_score = max(score_nom, score_desc)
            if meilleur_score >= seuil:
                resultats.append({
                    "Correspondance (%)": meilleur_score,
                    "Source": "Listes internationales de sanctions (Référentiel)",
                    "Entité / Résultat": row["Nom"],
                    "Catégorie": row["Catégorie"],
                    "Détails": row["Description"],
                    "URL": row["URL"]
                })

    resultats.append({
        "Correspondance (%)": 100.0,
        "Source": "Agrégateur international OpenSanctions",
        "Entité / Résultat": f"Requête consolidée : {terme}",
        "Catégorie": "Vérification multi-juridictions (ONU, OFAC, UE, Interpol)",
        "Détails": "Lien de vérification directe dans la base internationale consolidée.",
        "URL": f"https://www.opensanctions.org/search/?q={requests.utils.quote(terme)}"
    })
    return resultats

# ==========================================
# 3. ORCHESTRATION ET ANALYSE DE L'AGENT SIAG
# ==========================================
def executer_agent_kyc(terme: str, bases_actives: list, seuil: int):
    journal_etapes = []
    donnees_collectees = []

    if "Bases du Québec (REQ, RENA, Données Québec)" in bases_actives:
        journal_etapes.append(f"Exécution du connecteur 1 : Interrogation de l'API Données Québec (CKAN) et des registres québécois (seuil de correspondance fixé à {seuil} %).")
        res_qc = interroger_bases_quebec(terme, seuil)
        donnees_collectees.extend(res_qc)
        journal_etapes.append(f"Résultat du connecteur 1 : {len(res_qc)} enregistrement(s) pertinent(s) extrait(s).")

    if "Sanctions et Entités Terroristes Canada (LMES / Code criminel)" in bases_actives:
        journal_etapes.append(f"Exécution du connecteur 2 : Analyse du flux XML d'Affaires mondiales Canada (LMES) et de la liste du Code criminel (seuil fixé à {seuil} %).")
        res_can = interroger_sanctions_canada(terme, seuil)
        donnees_collectees.extend(res_can)
        journal_etapes.append(f"Résultat du connecteur 2 : {len(res_can)} enregistrement(s) pertinent(s) extrait(s).")

    if "Listes Internationales de Sanctions (ONU, OFAC, UE, OpenSanctions)" in bases_actives:
        journal_etapes.append(f"Exécution du connecteur 3 : Interrogation des référentiels internationaux de sanctions (seuil fixé à {seuil} %).")
        res_int = interroger_sanctions_internationales(terme, seuil)
        donnees_collectees.extend(res_int)
        journal_etapes.append(f"Résultat du connecteur 3 : {len(res_int)} référence(s) internationale(s) extraite(s).")

    donnees_collectees = sorted(donnees_collectees, key=lambda x: x.get("Correspondance (%)", 0), reverse=True)
    rapport_ia, classe_css, niveau_defaut = generer_diagnostic_agent(terme, donnees_collectees, bases_actives)
    journal_etapes.append("Synthèse des risques (légal, réputationnel, opérationnel) complétée par le moteur d'analyse.")

    return donnees_collectees, rapport_ia, classe_css, niveau_defaut, journal_etapes

def generer_diagnostic_agent(terme: str, donnees: list, bases: list):
    alerte_critique = any(
        "Code criminel" in d["Source"] or "Sanction économique" in d["Catégorie"]
        for d in donnees
    )
    meilleur_match = donnees[0]["Correspondance (%)"] if donnees else 0
    
    if alerte_critique:
        texte = (
            f"<h4>Synthèse de l'Agent KYC : Risque Critique Détecté — « {terme.upper()} » (Correspondance maximale : {meilleur_
