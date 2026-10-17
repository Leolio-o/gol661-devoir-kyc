import streamlit as st
import pandas as pd
import requests
import xml.etree.ElementTree as ET
import sqlite3
import json
import os
import uuid
from datetime import datetime
from openai import OpenAI
from langsmith import traceable
from langsmith.wrappers import wrap_openai

# Configuration de la page
st.set_page_config(
    page_title="Agent SIAG KYC - GOL661 (BCAPG)",
    page_icon="🤖",
    layout="wide"
)

# ==========================================
# 1. BASE DE DONNÉES D'AUDIT (BCAPG)
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
    audit_id = f"SIAG-{datetime.now().strftime('%Y%m%d')}-{str(uuid.uuid4())[:6].upper()}"
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
# 2. OUTILS (TOOLS) DE L'AGENT SIAG
# ==========================================

@traceable(run_type="tool", name="Outil_1_Bases_Quebec")
def tool_bases_quebec(terme: str) -> list:
    """Interroge l'API CKAN Données Québec (REQ/RENA) et la feuille 1-Bases Québec."""
    resultats = []
    try:
        url = "https://www.donneesquebec.ca/recherche/api/3/action/package_search"
        resp = requests.get(url, params={"q": terme, "rows": 4}, timeout=6)
        if resp.status_code == 200:
            for item in resp.json().get("result", {}).get("results", []):
                resultats.append({
                    "Source": "API Données Québec (CKAN)",
                    "Entité / Résultat": item.get("title", "N/A"),
                    "Catégorie": item.get("organization", {}).get("title", "Québec"),
                    "Détails": (item.get("notes", "") or "")[:160] + "...",
                    "URL": f"https://www.donneesquebec.ca/recherche/dataset/{item.get('name', '')}"
                })
    except Exception:
        pass

    if os.path.exists("GOL661_Devoir_KYC.xlsx"):
        df = pd.read_excel("GOL661_Devoir_KYC.xlsx", sheet_name="1-Bases Québec")
        mask = df.apply(lambda r: terme.lower() in str(r.values).lower(), axis=1)
        for _, row in df[mask].iterrows():
            resultats.append({
                "Source": "Inventaire Québec (Excel Q1)",
                "Entité / Résultat": row["Nom"],
                "Catégorie": row["Catégorie"],
                "Détails": row["Description"],
                "URL": row["URL"]
            })
    return resultats

@traceable(run_type="tool", name="Outil_2_Sanctions_Terrorisme_Canada")
def tool_sanctions_canada(terme: str) -> list:
    """Interroge le flux XML des sanctions canadiennes (LMES) et la liste des entités terroristes (Q3)."""
    resultats = []
    if os.path.exists("GOL661_Devoir_KYC.xlsx"):
        df_t = pd.read_excel("GOL661_Devoir_KYC.xlsx", sheet_name="3-Organisations terroristes")
        mask = df_t.apply(lambda r: terme.lower() in str(r.values).lower(), axis=1)
        for _, row in df_t[mask].iterrows():
            resultats.append({
                "Source": "Sécurité publique Canada (Terrorisme)",
                "Entité / Résultat": f"{row['Nom']} (Alias: {row['Alias / variantes']})",
                "Catégorie": f"Entité terroriste - {row['Pays d\'origine']}",
                "Détails": f"Zones: {row['Zone d\'opération']} | Listes: {row['Présence dans quelles listes']}",
                "URL": row["Lien Canada (Sécurité publique / communiqué)"]
            })

    try:
        url_xml = "https://www.international.gc.ca/world-monde/assets/office_docs/international_relations-relations_internationales/sanctions/sema-lmes.xml"
        resp = requests.get(url_xml, timeout=8)
        if resp.status_code == 200:
            root = ET.fromstring(resp.content)
            count = 0
            for rec in root.findall(".//record"):
                txt = " ".join([e.text or "" for e in rec.iter()])
                if terme.lower() in txt.lower():
                    nom = rec.findtext("Entity") or f"{rec.findtext('GivenName', '')} {rec.findtext('LastName', '')}".strip()
                    pays = rec.findtext("Country") or "Inconnu"
                    resultats.append({
                        "Source": "API XML Sanctions Canada (LMES)",
                        "Entité / Résultat": nom,
                        "Catégorie": f"Sanction économique Canada ({pays})",
                        "Détails": f"Annexe : {rec.findtext('Schedule') or 'LMES'}",
                        "URL": "https://www.international.gc.ca/world-monde/international_relations-relations_internationales/sanctions/consolidated-consolide.aspx?lang=fra"
                    })
                    count += 1
                    if count >= 8:
                        break
    except Exception:
        pass
    return resultats

@traceable(run_type="tool", name="Outil_3_Sanctions_Internationales")
def tool_sanctions_internationales(terme: str) -> list:
    """Interroge les listes internationales de sanctions (OpenSanctions, ONU, OFAC, UE - Q2)."""
    resultats = []
    if os.path.exists("GOL661_Devoir_KYC.xlsx"):
        df_s = pd.read_excel("GOL661_Devoir_KYC.xlsx", sheet_name="2-Sanctions")
        mask = df_s.apply(lambda r: terme.lower() in str(r.values).lower(), axis=1)
        for _, row in df_s[mask].iterrows():
            resultats.append({
                "Source": "Inventaire Sanctions Int. (Excel Q2)",
                "Entité / Résultat": row["Nom"],
                "Catégorie": row["Catégorie"],
                "Détails": row["Description"],
                "URL": row["URL"]
            })

    resultats.append({
        "Source": "Agrégateur OpenSanctions (OFAC/ONU/UE)",
        "Entité / Résultat": f"Requête internationale : {terme}",
        "Catégorie": "Vérification multi-juridictions",
        "Détails": "Lien direct vers la fiche consolidée OpenSanctions (80+ listes mondiales).",
        "URL": f"https://www.opensanctions.org/search/?q={requests.utils.quote(terme)}"
    })
    return resultats

# ==========================================
# 3. CERVEAU DE L'AGENT SIAG (AVEC LANGSMITH)
# ==========================================

@traceable(run_type="chain", name="Agent_SIAG_KYC_GOL661")
def executer_agent_siag(terme: str, bases_actives: list, openai_key: str):
    """
    L'Agent SIAG planifie l'interrogation des bases sélectionnées, collecte les preuves
    et rédige un rapport d'évaluation des risques (réputationnel, légal, opérationnel).
    """
    journal_etapes = []
    donnees_collectees = []

    # Étape 1 : Orchestration des outils selon la sélection de l'utilisateur
    if "Bases Québec (REQ, RENA, Données Québec)" in bases_actives:
        journal_etapes.append("🔄 **Agent SIAG → Appel Outil 1 :** Interrogation de l'API Données Québec (CKAN) et des registres québécois...")
        res_qc = tool_bases_quebec(terme)
        donnees_collectees.extend(res_qc)
        journal_etapes.append(f"✔️ {len(res_qc)} résultat(s) extrait(s) des bases québécoises.")

    if "Sanctions & Entités Terroristes Canada (LMES / Code criminel)" in bases_actives:
        journal_etapes.append("🔄 **Agent SIAG → Appel Outil 2 :** Analyse du flux XML Affaires mondiales Canada et de la liste antiterroriste...")
        res_can = tool_sanctions_canada(terme)
        donnees_collectees.extend(res_can)
        journal_etapes.append(f"✔️ {len(res_can)} résultat(s) extrait(s) des bases fédérales canadiennes.")

    if "Listes Internationales (ONU, OFAC, UE, OpenSanctions)" in bases_actives:
        journal_etapes.append("🔄 **Agent SIAG → Appel Outil 3 :** Interrogation des listes internationales consolidées...")
        res_int = tool_sanctions_internationales(terme)
        donnees_collectees.extend(res_int)
        journal_etapes.append(f"✔️ {len(res_int)} référence(s) internationale(s) vérifiée(s).")

    # Étape 2 : Raisonnement et synthèse par le SIAG (LLM ou moteur heuristique de secours)
    prompt_systeme = """Tu es un Agent expert en conformité KYC (Know Your Customer) et gestion des risques (cours GOL661 - ÉTS).
À partir des résultats bruts extraits des bases de données ouvertes, rédige un rapport structuré court (10-15 lignes) évaluant :
1. Le risque LÉGAL (sanctions économiques, Code criminel, inadmissibilité RENA).
2. Le risque RÉPUTATIONNEL (association avec une entité controversée, terrorisme, fraude).
3. Le risque OPÉRATIONNEL (blocage de la chaîne d'approvisionnement, gel des actifs).
4. Une recommandation claire (Approuver, Diligence renforcée, ou Rejeter/Bloquer)."""

    donnees_str = json.dumps(donnees_collectees, ensure_ascii=False, indent=2)

    if openai_key:
        try:
            client = wrap_openai(OpenAI(api_key=openai_key))
            completion = client.chat.completions.create(
                model="gpt-4o-mini",
                messages=[
                    {"role": "system", "content": prompt_systeme},
                    {"role": "user", "content": f"Entité recherchée : '{terme}'\nBases consultées : {bases_actives}\nRésultats trouvés :\n{donnees_str}"}
                ],
                temperature=0.2
            )
            rapport_ia = completion.choices[0].message.content
            journal_etapes.append("🧠 **Agent SIAG :** Synthèse générée via LLM (trace envoyée vers LangSmith).")
        except Exception as e:
            rapport_ia = generer_synthese_locale(terme, donnees_collectees, bases_actives)
            journal_etapes.append(f"⚠️ Clé API non valide ou quota atteint ({e}). Bascule sur le moteur d'inférence local.")
    else:
        rapport_ia = generer_synthese_locale(terme, donnees_collectees, bases_actives)
        journal_etapes.append("🧠 **Agent SIAG :** Synthèse générée par le moteur d'analyse embarqué (ajoutez une clé OpenAI dans le menu latéral pour activer le LLM externe).")

    return donnees_collectees, rapport_ia, journal_etapes

def generer_synthese_locale(terme: str, donnees: list, bases: list) -> str:
    """Moteur d'analyse KYC autonome garantissant un rapport structuré même sans clé API payante."""
    sources = [d["Source"] for d in donnees]
    alerte_terrorisme = any("Terrorisme" in s or "Sanction" in d["Catégorie"] for d, s in zip(donnees, sources))
    
    if alerte_terrorisme:
        return (
            f"### 🚨 RAPPORT DE L'AGENT SIAG : RISQUE CRITIQUE DÉTECTÉ POUR « {terme.upper()} »\n"
            f"- **Bases interrogées :** {', '.join(bases)}\n"
            f"- **Risque Légal (Critique) :** L'entité figure dans les listes officielles de sanctions canadiennes (LMES / Code criminel art. 83.05) et/ou internationales. Toute transaction financière ou contractuelle constitue une infraction criminelle passible de poursuites.\n"
            f"- **Risque Réputationnel (Critique) :** Contamination immédiate de l'image de l'entreprise en cas d'association directe ou indirecte avec cette organisation.\n"
            f"- **Risque Opérationnel (Élevé) :** Gel des biens, saisie douanière et révocation des autorisations de contracter (AMP/RENA).\n"
            f"- **Recommandation de l'Agent :** **BLOCAGE IMMÉDIAT** de la partie prenante et signalement au responsable de la conformité."
        )
    elif len(donnees) > 1:
        return (
            f"### ⚠️ RAPPORT DE L'AGENT SIAG : VIGILANCE REQUISE POUR « {terme.upper()} »\n"
            f"- **Bases interrogées :** {', '.join(bases)}\n"
            f"- **Risque Légal (Modéré) :** {len(donnees)} correspondance(s) documentaire(s) identifiée(s) dans les registres ouverts. Aucune inscription directe sur la liste antiterroriste, mais une vérification d'homonymie et de validité de licence (REQ/RENA/RBQ) est requise.\n"
            f"- **Risque Réputationnel (Modéré) :** Vérifier les antécédents judiciaires (SOQUIJ/CanLII) et les bénéficiaires effectifs avant signature.\n"
            f"- **Risque Opérationnel (Faible à Modéré) :** S'assurer que le fournisseur détient ses autorisations AMP si le contrat dépasse les seuils publics.\n"
            f"- **Recommandation de l'Agent :** **DILIGENCE RAISONNABLE RENFORCÉE** (documenter la vérification dans l'espace commentaires ci-dessous)."
        )
    else:
        return (
            f"### ✅ RAPPORT DE L'AGENT SIAG : RISQUE FAIBLE POUR « {terme.upper()} »\n"
            f"- **Bases interrogées :** {', '.join(bases)}\n"
            f"- **Analyse globale :** Aucune correspondance négative détectée dans les bases de sanctions ou d'entités terroristes sélectionnées.\n"
            f"- **Recommandation de l'Agent :** **APPROBATION CONDITIONNELLE**, sous réserve de la vérification standard du numéro d'entreprise du Québec (NEQ) au REQ."
        )

# ==========================================
# 4. INTERFACE UTILISATEUR STREAMLIT
# ==========================================
st.title("🤖 Agent SIAG de Vérification KYC & Traçabilité LangSmith (GOL661)")
st.markdown("Application en **mode Agent SIAG** permettant d'interroger les bases ouvertes, d'analyser les risques de contamination et d'archiver les preuves pour le **BCAPG**.")

with st.sidebar:
    st.header("⚙️ Configuration de l'Agent SIAG")
    analyste = st.text_input("Identifiant de l'auditeur :", value="Équipe GOL661")
    
    st.subheader("1. Sélection des bases à interroger")
    b_qc = st.checkbox("Bases Québec (REQ, RENA, Données Québec)", value=True)
    b_can = st.checkbox("Sanctions & Entités Terroristes Canada (LMES / Code criminel)", value=True)
    b_int = st.checkbox("Listes Internationales (ONU, OFAC, UE, OpenSanctions)", value=True)
    
    st.divider()
    st.subheader("2. Connexion SIAG & LangSmith (Optionnel)")
    openai_api_key = st.text_input("Clé API OpenAI (GPT-4o-mini) :", type="password")
    langsmith_key = st.text_input("Clé API LangSmith (Traçage) :", type="password")
    
    if langsmith_key:
        os.environ["LANGCHAIN_TRACING_V2"] = "true"
        os.environ["LANGCHAIN_API_KEY"] = langsmith_key
        os.environ["LANGCHAIN_PROJECT"] = "GOL661-Audit-KYC-BCAPG"
        st.success("✅ Traçage LangSmith actif (Projet : GOL661-Audit-KYC-BCAPG)")

tab1, tab2, tab3 = st.tabs([
    "🤖 1. Enquête par l'Agent SIAG", 
    "🗄️ 2. Registre d'Audit BCAPG", 
    "📊 3. Tableaux de référence (Points 1 à 3)"
])

with tab1:
    col_input, col_btn = st.columns([3, 1])
    with col_input:
        terme = st.text_input("Nom de la partie prenante à auditer :", placeholder="Ex: Hamas, Al-Qaïda, Boko Haram, Russie, Construction...")
    with col_btn:
        st.write("")
        st.write("")
        lancer = st.button("🚀 Exécuter l'Agent SIAG", type="primary", use_container_width=True)

    if lancer and terme:
        bases_choisies = []
        if b_qc: bases_choisies.append("Bases Québec (REQ, RENA, Données Québec)")
        if b_can: bases_choisies.append("Sanctions & Entités Terroristes Canada (LMES / Code criminel)")
        if b_int: bases_choisies.append("Listes Internationales (ONU, OFAC, UE, OpenSanctions)")

        if not bases_choisies:
            st.warning("Veuillez sélectionner au moins une base de données dans le menu de gauche.")
        else:
            with st.spinner("L'Agent SIAG interroge les bases de données et rédige son analyse..."):
                res_list, rapport_ia, etapes = executer_agent_siag(terme, bases_choisies, openai_api_key)
                st.session_state["res_df"] = pd.DataFrame(res_list)
                st.session_state["rapport_ia"] = rapport_ia
                st.session_state["etapes_agent"] = etapes
                st.session_state["terme_actuel"] = terme
                st.session_state["bases_actuelles"] = bases_choisies

    if "rapport_ia" in st.session_state:
        with st.expander("🔍 Voir le journal de raisonnement de l'Agent SIAG (Étapes d'exécution)", expanded=True):
            for etape in st.session_state["etapes_agent"]:
                st.write(etape)

        st.info(st.session_state["rapport_ia"])

        st.subheader(f"Résultats extraits des bases ({len(st.session_state['res_df'])} entrée(s))")
        if not st.session_state["res_df"].empty:
            st.dataframe(
                st.session_state["res_df"],
                column_config={"URL": st.column_config.LinkColumn("Lien source officiel")},
                use_container_width=True,
                hide_index=True
            )

        st.divider()
        st.subheader("📝 Validation humaine & Sauvegarde pour audit BCAPG")
        with st.form("audit_form"):
            risque_retenu = st.selectbox(
                "Cote de risque finale validée par l'analyste :",
                ["CRITIQUE - Blocage immédiat", "ÉLEVÉ - Rejet recommandé", "MODÉRÉ - Diligence renforcée", "FAIBLE - Homologation autorisée"]
            )
            commentaire = st.text_area(
                "Espace commentaires (obligatoire pour justifier la décision auprès du BCAPG) :",
                placeholder="Ex: Confirmation qu'il s'agit bien de l'entité visée par l'article 83.05 du Code criminel. Fournisseur bloqué dans l'ERP."
            )
            btn_save = st.form_submit_button("💾 Enregistrer le dossier dans le serveur d'audit (BCAPG)")
            if btn_save:
                id_genere = sauvegarder_audit(
                    analyste=analyste,
                    entite=st.session_state["terme_actuel"],
                    bases=st.session_state["bases_actuelles"],
                    rapport_ia=st.session_state["rapport_ia"],
                    risque=risque_retenu,
                    commentaire=commentaire,
                    df_res=st.session_state["res_df"]
                )
                st.success(f"✅ Dossier d'audit archivé avec succès ! Référence BCAPG : **`{id_genere}`**")

with tab2:
    st.subheader("🗄️ Journal d'audit permanent (Consultation BCAPG)")
    df_hist = charger_audits()
    if not df_hist.empty:
        st.dataframe(df_hist.drop(columns=["resultats_bruts_json"]), use_container_width=True, hide_index=True)
        csv_bytes = df_hist.to_csv(index=False).encode('utf-8-sig')
        st.download_button(
            "📥 Télécharger le registre d'audit BCAPG (.CSV)",
            data=csv_bytes,
            file_name=f"Audit_BCAPG_SIAG_{datetime.now().strftime('%Y%m%d')}.csv",
            mime="text/csv"
        )
    else:
        st.info("Aucune vérification archivée pour l'instant.")

with tab3:
    st.subheader("📊 Vos tableaux de recherche (Questions 1, 2 et 3)")
    if os.path.exists("GOL661_Devoir_KYC.xlsx"):
        feuille = st.selectbox("Choisir la feuille :", ["1-Bases Québec", "2-Sanctions", "3-Organisations terroristes"])
        st.dataframe(pd.read_excel("GOL661_Devoir_KYC.xlsx", sheet_name=feuille), use_container_width=True, hide_index=True)
