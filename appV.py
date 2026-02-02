import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import hashlib
from io import BytesIO
import datetime as dt
from dateutil.relativedelta import relativedelta
import warnings
import numpy as np
warnings.filterwarnings('ignore')

# Fonction pour formater les dates en français
def format_date_fr(date):
    if pd.isna(date) or date is None:
        return ""
    if isinstance(date, str):
        date = pd.to_datetime(date)
    months = {
        1: 'janvier', 2: 'février', 3: 'mars', 4: 'avril',
        5: 'mai', 6: 'juin', 7: 'juillet', 8: 'août',
        9: 'septembre', 10: 'octobre', 11: 'novembre', 12: 'décembre'
    }
    return f"{date.day} {months[date.month]} {date.year}"

# Fonction pour hasher le fichier pour le cache
@st.cache_data
def get_file_hash(uploaded_file):
    return hashlib.md5(uploaded_file.read()).hexdigest()

# Charger et nettoyer les données (FIX pour Quantité : forcer en float)
@st.cache_data
def load_and_clean_data(file_hash, file_bytes):
    try:
        xls = pd.ExcelFile(BytesIO(file_bytes))
        data = {}
        for sheet in xls.sheet_names:
            df = pd.read_excel(BytesIO(file_bytes), sheet_name=sheet)
            # Nettoyage : Supprimer lignes vides
            df = df.dropna(how='all')  # Supprimer lignes entièrement vides
            
            # Détecter et convertir dates Excel (seulement sur colonnes int64)
            int_df = df.select_dtypes(include=['int64'])
            if not int_df.empty:
                unique_counts = int_df.nunique()  # nunique SEULEMENT sur les int64 (taille correcte)
                date_mask = unique_counts < len(df)  # Masque booléen de la bonne taille
                date_cols = int_df.columns[date_mask].tolist()  # Colonnes potentielles dates
                for col in date_cols:
                    if not (sheet == "Achats" and col == "Quantité") and "Année" not in col:
                        df[col] = pd.to_datetime(df[col], unit='D', origin='1899-12-30', errors='coerce')
            else:
                date_cols = []  # Pas de colonnes à convertir
            
            # Nettoyage supplémentaire : Remplacer NaN par 0 dans colonnes numériques
            numeric_cols = df.select_dtypes(include=[np.number]).columns
            df[numeric_cols] = df[numeric_cols].fillna(0)
            
            # FIX SPÉCIFIQUE : Forcer "Quantité" en float dans Achats (éviter confusion date/nombre)
            if sheet == "Achats" and "Quantité" in df.columns:
                # Forcer Quantité comme numérique AVANT toute détection de date
                df["Quantité"] = pd.to_numeric(df["Quantité"], errors="coerce")
                df["Quantité"] = df["Quantité"].fillna(0)
                # Valeurs négatives ou absurdes -> valeur absolue
                df["Quantité"] = df["Quantité"].abs()
            
            data[sheet] = df
        return data
    except Exception as e:
        st.error(f"Erreur lors du chargement : {e}. Vérifiez le format Excel.")
        return {}

# Fonction utilitaire pour pré-formater colonnes avec espaces (pour tableaux)
def pre_format_columns(df, money_cols, quantity_cols):
    df_formatted = df.copy()
    for col in money_cols:
        if col in df.columns:
            df_formatted[col] = df_formatted[col].apply(lambda x: f"{x:,.0f}".replace(",", " ") + " Ar")
    for col in quantity_cols:
        if col in df.columns:
            if col == "Litres":
                df_formatted[col] = df_formatted[col].apply(lambda x: f"{x:.1f} L")
            elif col == "Kilométrage":
                df_formatted[col] = df_formatted[col].apply(lambda x: f"{x:,.0f}".replace(",", " ") + " km")
            elif col == "Km_Parcourus":
                df_formatted[col] = df_formatted[col].apply(lambda x: f"{x:,.0f}".replace(",", " ") + " km")
            elif col == "Quantité":
                df_formatted[col] = df_formatted[col].apply(lambda x: f"{x:.1f}")
            else:
                df_formatted[col] = df_formatted[col].apply(lambda x: f"{x:.0f}")
    
    # Formatage des colonnes dates
    date_cols = df_formatted.select_dtypes(include=['datetime64[ns]']).columns
    for col in date_cols:
        df_formatted[col] = df_formatted[col].apply(format_date_fr)
    
    return df_formatted

# Fonction utilitaire pour formater les colonnes monétaires avec "Ar" (espace comme séparateur)
def format_money_columns(df, money_cols):
    config = {}
    for col in money_cols:
        if col in df.columns:
            config[col] = st.column_config.NumberColumn(label=col, format="%.0f Ar")
    return config

# Fonction utilitaire pour formater les colonnes litres avec "L"
def format_liters_columns(df, liter_cols):
    config = {}
    for col in liter_cols:
        if col in df.columns:
            config[col] = st.column_config.NumberColumn(label=col, format="%.1f L")
    return config

# Configuration de la page
st.set_page_config(page_title="Suivi Véhicules OMNIS ", layout="wide", initial_sidebar_state="expanded")
st.title("🚗📊 Suivi des Véhicules OMNIS ")

# Sidebar pour filtres globaux
st.sidebar.header("🔧 Filtres Globaux")
uploaded_file = st.sidebar.file_uploader("📁 Charger le fichier Excel", type=["xlsx"])

if uploaded_file:
    with st.spinner("Chargement des données..."):
        file_hash = get_file_hash(uploaded_file)
        data = load_and_clean_data(file_hash, uploaded_file.getvalue())
    
    if not data:
        st.error("Impossible de charger les données. Vérifiez le fichier.")
        st.stop()
    
    st.sidebar.success("✅ Données chargées")
    
    # Stats de chargement (bonus)
    with st.sidebar.expander("📈 Stats Chargement"):
        for sheet, df in data.items():
            st.write(f"{sheet}: {len(df)} lignes")

    # Récupérer les DataFrames avec gestion d'erreurs (AJOUT "Carburant")
    required_sheets = ["Parc_Véhicules", "Entretien", "Réparations Internes", "Prestation externe", 
                       "Suivi_Kilométrage", "Garage", "Fournisseurs", "Achats", "Assurance", "Visite_Technique", "Carburant"]
    dfs = {}
    for sheet in required_sheets:
        if sheet in data:
            dfs[sheet] = data[sheet]
        else:
            st.error(f"Feuille '{sheet}' manquante. Utilisez l'Excel généré pour tester.")
            st.stop()

    df_vehicules = dfs["Parc_Véhicules"]
    directions = sorted(df_vehicules["Direction"].dropna().unique())
    selected_directions = st.sidebar.multiselect("🏢 Directions", options=directions, default=directions)
    date_start, date_end = st.sidebar.date_input("📅 Période", value=(dt.date(2024, 1, 1), dt.date(2026, 1, 18)))

    df_vehicules_filtered = df_vehicules[df_vehicules["Direction"].isin(selected_directions)]
    if df_vehicules_filtered.empty:
        st.warning("Aucune direction sélectionnée valide.")
        st.stop()

    # Sélection véhicule
    selected_vehicle = st.selectbox("🚗 Véhicule", options=df_vehicules_filtered["Immatriculation"].unique())

    # Infos véhicule filtrées
    vehicule_info = df_vehicules_filtered[df_vehicules_filtered["Immatriculation"] == selected_vehicle].iloc[0]
    df_vehicle_specific = {}
    for k, v in dfs.items():
        if "Immatriculation" in v.columns:
            df_vehicle_specific[k] = v[v["Immatriculation"] == selected_vehicle]

    # Dashboard Global en haut (KPIs en 2 lignes, unité Km ajoutée, format espace, SUPPRIMÉ deltas)
    # Ligne 1 : 4 KPIs
    col1, col2, col3, col4 = st.columns(4)
    with col1:
        df_km = df_vehicle_specific.get("Suivi_Kilométrage", pd.DataFrame())
        dernier_km = df_km.sort_values("Date", ascending=False)["Kilométrage"].iloc[0] if not df_km.empty else 0
        st.metric("📏 Kilométrage", f"{int(dernier_km):,}".replace(",", " ") + " km")
    with col2:
        total_entretien = df_vehicle_specific.get("Entretien", pd.DataFrame())["Coût_Total"].sum()
        st.metric("🛠 Coût Entretien", f"{total_entretien:,.0f}".replace(",", " ") + " Ar")
    with col3:
        total_reparations = (df_vehicle_specific.get("Réparations Internes", pd.DataFrame())["Coût_Total"].sum() + 
                             df_vehicle_specific.get("Prestation externe", pd.DataFrame())["Coût_Total"].sum())
        st.metric("🔧 Coût Réparations", f"{total_reparations:,.0f}".replace(",", " ") + " Ar")
    with col4:
        total_achats = df_vehicle_specific.get("Achats", pd.DataFrame())["Prix_Total"].sum()
        st.metric("🛒 Achats", f"{total_achats:,.0f}".replace(",", " ") + " Ar")
    
    # Ligne 2 : 3 KPIs
    col5, col6, col7 = st.columns(3)
    with col5:
        cout_total_veh = total_entretien + total_reparations + total_achats
        st.metric("💰 Coût Total Véhicule", f"{cout_total_veh:,.0f}".replace(",", " ") + " Ar")
    with col6:
        df_carbu = df_vehicle_specific.get("Carburant", pd.DataFrame())
        total_litres = df_carbu["Litres"].sum()
        st.metric("⛽ Total Litres", f"{total_litres:,.1f}".replace(",", " ") + " L")
    with col7:
        total_carbu_ar = df_carbu["Total_Ar"].sum()
        st.metric("⛽ Coût Carburant", f"{total_carbu_ar:,.0f}".replace(",", " ") + " Ar")

    # Alertes (ex. : assurances expirées)
    today = pd.to_datetime(dt.date.today())  # Convertir en datetime64[ns] pour compatibilité pandas
    df_ass = df_vehicle_specific.get("Assurance", pd.DataFrame())
    if not df_ass.empty and "Date_Fin" in df_ass.columns:
        # Normaliser les dates pour ignorer l'heure
        df_ass["Date_Fin"] = pd.to_datetime(df_ass["Date_Fin"]).dt.normalize()
        ass_exp = df_ass[df_ass["Date_Fin"] < today]
        if not ass_exp.empty:
            st.error(f"⚠️ {len(ass_exp)} assurance(s) expirée(s) pour {selected_vehicle} !")

    df_vt = df_vehicle_specific.get("Visite_Technique", pd.DataFrame())
    if not df_vt.empty and "Etat" in df_vt.columns:
        vt_exp = df_vt[df_vt["Etat"] == "Expiré"]
        if not vt_exp.empty:
            st.warning(f"🔍 {len(vt_exp)} visite(s) technique(s) à renouveler.")

    # Onglets améliorés (AJOUT onglet "⛽ Carburant")
    tab1, tab2, tab3, tab4, tab5, tab6, tab7 = st.tabs([
        "📋 Fiche Véhicule", "🛠 Entretien & Réparations",
        "📈 Kilométrage & Performances", "📋 Assurance & Visites",
        "🛒 Achats & Fournisseurs", "⛽ Carburant", "📊 Dashboard Global & Export"
    ])

    with tab1:
        st.subheader(f"📌 Détails : {selected_vehicle}")
        # Appliquer formatage pour les dates et potentiellement monétaires
        money_cols_veh = ["Prix_Achat"] if "Prix_Achat" in vehicule_info.index else []
        quantity_cols_veh = []
        veh_df_formatted = pre_format_columns(vehicule_info.to_frame().T, money_cols_veh, quantity_cols_veh)
        st.dataframe(veh_df_formatted, use_container_width=True)

    with tab2:
        # Entretien
        st.subheader("🛠 Entretien")
        df_e = df_vehicle_specific.get("Entretien", pd.DataFrame())
        if df_e.empty:
            st.info("Aucun entretien.")
        else:
            df_e_formatted = pre_format_columns(df_e, ["Coût_Total"], [])
            st.dataframe(df_e_formatted, use_container_width=True)
            if 'Type_Entretien' in df_e.columns and 'Coût_Total' in df_e.columns:
                fig = px.pie(df_e, names='Type_Entretien', values='Coût_Total', title='Répartition Coûts Entretien (Ar)')
                fig.update_traces(textinfo='label+percent+value', texttemplate='%{label}<br>%{percent}<br>%{value} Ar')
                st.plotly_chart(fig, use_container_width=True)

        # Réparations Internes
        st.subheader("🔧 Réparations Internes")
        df_ri = df_vehicle_specific.get("Réparations Internes", pd.DataFrame())
        if df_ri.empty:
            st.info("Aucune réparation interne.")
        else:
            df_ri_formatted = pre_format_columns(df_ri, ["Coût_Total"], [])
            st.dataframe(df_ri_formatted, use_container_width=True)
            if 'Date d_entrée à Andraharo' in df_ri.columns and 'Coût_Total' in df_ri.columns:
                fig_ri = px.bar(df_ri, x='Date d_entrée à Andraharo', y='Coût_Total', color='Panne', 
                                title='Évolution Coûts Réparations Internes (Ar)')
                fig_ri.update_yaxes(title_text="Coût (Ar)")
                st.plotly_chart(fig_ri, use_container_width=True)

        # Prestations Externes
        st.subheader("🌐 Prestations Externes")
        df_pe = df_vehicle_specific.get("Prestation externe", pd.DataFrame())
        if df_pe.empty:
            st.info("Aucune prestation externe.")
        else:
            df_pe_formatted = pre_format_columns(df_pe, ["Coût_Total"], [])
            st.dataframe(df_pe_formatted, use_container_width=True)
            if 'Type de Prestation' in df_pe.columns and 'Coût_Total' in df_pe.columns:
                fig_pe = px.pie(df_pe, names='Type de Prestation', values='Coût_Total', title='Répartition Prestations (Ar)')
                fig_pe.update_traces(textinfo='label+percent+value', texttemplate='%{label}<br>%{percent}<br>%{value} Ar')
                st.plotly_chart(fig_pe, use_container_width=True)

    with tab3:
     st.subheader("📈 Suivi Kilométrage")  # CHANGÉ EN BAR CHART
     df_km = df_vehicle_specific.get("Suivi_Kilométrage", pd.DataFrame())
     if df_km.empty:
        st.info("Pas de données kilométriques.")
     else:
        # TRI ET CALCUL KM PARCOCUS (nouveau)
        df_km = df_km.sort_values("Date").reset_index(drop=True)  # Trier par date pour diff correcte
        df_km['Km_Parcourus'] = df_km['Kilométrage'].diff().fillna(0)  # Diff km + 0 pour 1ère ligne
        
        # Tableau formaté (avec espaces pour milliers)
        df_km_formatted = pre_format_columns(df_km, [], ["Kilométrage", "Km_Parcourus"])  # Ajoute Km_Parcourus
        st.dataframe(df_km_formatted, use_container_width=True)
        
        if 'Date' in df_km.columns and 'Km_Parcourus' in df_km.columns:
            # Bar chart avec km parcourus
            fig_km = px.bar(df_km, x='Date', y='Km_Parcourus', title='Évolution Km Parcourus (Bar Chart)')
            fig_km.update_yaxes(title_text="Km Parcourus entre Dates")
            st.plotly_chart(fig_km, use_container_width=True)

    with tab4:  # SÉQUENTIEL (Haut/Bas) au lieu de côte à côte
        st.subheader("📋 Assurances")
        df_ass_display = df_vehicle_specific.get("Assurance", pd.DataFrame())
        if not df_ass_display.empty:
            df_ass_formatted = pre_format_columns(df_ass_display, ["Montant"], [])
            st.dataframe(df_ass_formatted, use_container_width=True)
        else:
            st.dataframe(df_ass_display, use_container_width=True)
        
        st.subheader("🔍 Visites Techniques")
        df_vt_display = df_vehicle_specific.get("Visite_Technique", pd.DataFrame())
        df_vt_display_formatted = pre_format_columns(df_vt_display, [], [])
        st.dataframe(df_vt_display_formatted, use_container_width=True)  # Pas de monétaire ici

    with tab5:  # SÉQUENTIEL (Haut/Bas) au lieu de côte à côte
        st.subheader("🛒 Achats")
        df_ach = df_vehicle_specific.get("Achats", pd.DataFrame())
        if df_ach.empty:
            st.info("Aucun achat.")
        else:
            df_ach_formatted = pre_format_columns(df_ach, ["Prix_Unitaire", "Prix_Total"], ["Quantité"])
            st.dataframe(df_ach_formatted, use_container_width=True)
            if 'Nom_du_fournisseur' in df_ach.columns and 'Prix_Total' in df_ach.columns:
                fig_ach = px.pie(df_ach, names='Nom_du_fournisseur', values='Prix_Total', title='Achats par Fournisseur (Ar)')
                fig_ach.update_traces(textinfo='label+percent+value', texttemplate='%{label}<br>%{percent}<br>%{value} Ar')
                st.plotly_chart(fig_ach, use_container_width=True)
        
        st.subheader("📇 Fournisseurs")
        st.dataframe(dfs["Fournisseurs"], use_container_width=True)

    with tab6:  # ONGLET CARBURANT (SUPPRIMÉ PIE)
        st.subheader("⛽ Consommation Carburant")
        df_carbu = df_vehicle_specific.get("Carburant", pd.DataFrame())
        if df_carbu.empty:
            st.info("Aucune donnée carburant.")
        else:
            # Tableau avec formats
            df_carbu_formatted = pre_format_columns(df_carbu, ["Prix_Litre", "Total_Ar"], ["Litres"])
            st.dataframe(df_carbu_formatted, use_container_width=True)
            
            # Graphique Litres par date (bar) - UNIQUEMENT
            if 'Date' in df_carbu.columns and 'Litres' in df_carbu.columns:
                fig_litres = px.bar(df_carbu, x='Date', y='Litres', color='Type_Carburant', title='Évolution Consommation (L)')
                fig_litres.update_yaxes(title_text="Litres (L)")
                st.plotly_chart(fig_litres, use_container_width=True)

    with tab7:
        st.subheader("📊 Dashboard Global")
        # KPIs globaux
        total_veh = len(df_vehicules_filtered)
        total_coût = (dfs.get("Entretien", pd.DataFrame())["Coût_Total"].sum() + 
                      dfs.get("Réparations Internes", pd.DataFrame())["Coût_Total"].sum() + 
                      dfs.get("Prestation externe", pd.DataFrame())["Coût_Total"].sum() + 
                      dfs.get("Achats", pd.DataFrame())["Prix_Total"].sum() + 
                      dfs.get("Carburant", pd.DataFrame())["Total_Ar"].sum())  # AJOUT Carburant
        col_g1, col_g2, col_g3 = st.columns(3)
        col_g1.metric("🚗 Nb Véhicules", total_veh)
        col_g2.metric("💰 Coût Total Global", f"{total_coût:,.0f}".replace(",", " ") + " Ar")
        col_g3.metric("⏱ Dernière MAJ", format_date_fr(today))

        # Graphique global : Coûts par direction
        df_coûts_dir = df_vehicules.merge(dfs.get("Entretien", pd.DataFrame()), on="Immatriculation", how="left")
        df_coûts_dir["Coût_Total"] = df_coûts_dir["Coût_Total"].fillna(0)
        money_cols_global = ["Coût_Total"]
        config_global = format_money_columns(df_coûts_dir, money_cols_global)
        fig_global = px.bar(df_coûts_dir.groupby("Direction")["Coût_Total"].sum().reset_index(), 
                            x="Direction", y="Coût_Total", title="Coûts par Direction (Ar)")
        fig_global.update_yaxes(title_text="Coût (Ar)")
        st.plotly_chart(fig_global, use_container_width=True)
        
        # AJOUT : Répartition Carburant par Type (pie globale)
        df_carbu_global = dfs.get("Carburant", pd.DataFrame())
        if not df_carbu_global.empty and 'Type_Carburant' in df_carbu_global.columns and 'Total_Ar' in df_carbu_global.columns:
            fig_carbu_type = px.pie(df_carbu_global, names='Type_Carburant', values='Total_Ar', title='Répartition Coûts Carburant par Type (Ar)')
            fig_carbu_type.update_traces(textinfo='label+percent+value', texttemplate='%{label}<br>%{percent}<br>%{value} Ar')
            st.plotly_chart(fig_carbu_type, use_container_width=True)

        # Export Rapport
        st.subheader("📥 Générer Rapport")
        resume_data = {
            "Immatriculation": selected_vehicle,
            "Direction": vehicule_info["Direction"],
            "Kilométrage Actuel": f"{int(dernier_km):,}".replace(",", " ") + " km",
            "Coût Entretien": f"{total_entretien:,.0f}".replace(",", " ") + " Ar",
            "Coût Réparations": f"{total_reparations:,.0f}".replace(",", " ") + " Ar",
            "Coût Achats": f"{total_achats:,.0f}".replace(",", " ") + " Ar",
            "Coût Total Mécanique": f"{cout_total_veh:,.0f}".replace(",", " ") + " Ar",
            "Total Litres Carburant": f"{total_litres:,.1f}".replace(",", " ") + " L",
            "Coût Carburant": f"{total_carbu_ar:,.0f}".replace(",", " ") + " Ar",
            "Coût Total Global": f"{cout_total_veh + total_carbu_ar:,.0f}".replace(",", " ") + " Ar",
            "Date Rapport": format_date_fr(today.date())
        }
        df_resume = pd.DataFrame([resume_data])
        st.dataframe(df_resume, use_container_width=True)

        # Export Excel amélioré (avec format Ar et L, espaces pour milliers)
        buffer = BytesIO()
        with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
            workbook = writer.book
            header_format = workbook.add_format({'bold': True, 'text_wrap': True, 'valign': 'top', 'fg_color': '#D7E4BC'})
            money_format = workbook.add_format({'num_format': '# ##0 "Ar"'})  # Espaces au lieu de virgules
            liter_format = workbook.add_format({'num_format': '#,##0.0 "L"'})  # Format avec "L"
            
            df_resume.to_excel(writer, sheet_name="Résumé", index=False)
            worksheet = writer.sheets["Résumé"]
            for col_num, value in enumerate(df_resume.columns.values):
                worksheet.write(0, col_num, value, header_format)
            # Appliquer formats aux colonnes
            for col in ["Coût Entretien", "Coût Réparations", "Coût Achats", "Coût Total Mécanique", "Coût Carburant", "Coût Total Global"]:
                col_idx = list(df_resume.columns).index(col) + 1
                worksheet.set_column(col_idx, col_idx, None, money_format)
            for col in ["Total Litres Carburant"]:
                col_idx = list(df_resume.columns).index(col) + 1
                worksheet.set_column(col_idx, col_idx, None, liter_format)

            # Pour les autres feuilles
            sheets_money = {
                "Entretien": ["Coût_Total"],
                "Réparations Internes": ["Coût_Total"],
                "Prestation externe": ["Coût_Total"],  # Corrigé "Prestations Externes"
                "Assurance": ["Montant"],
                "Achats": ["Prix_Unitaire", "Prix_Total"],
                "Carburant": ["Prix_Litre", "Total_Ar"],
                "Suivi_Kilométrage": []  # Renommé pour cohérence
            }
            sheets_liter = {
                "Carburant": ["Litres"]
            }
            for sheet_name, df_sheet_orig in [("Entretien", df_vehicle_specific.get("Entretien", pd.DataFrame())),
                                         ("Réparations Internes", df_vehicle_specific.get("Réparations Internes", pd.DataFrame())),
                                         ("Prestation externe", df_vehicle_specific.get("Prestation externe", pd.DataFrame())),
                                         ("Assurance", df_vehicle_specific.get("Assurance", pd.DataFrame())),
                                         ("Visite_Technique", df_vehicle_specific.get("Visite_Technique", pd.DataFrame())),
                                         ("Achats", df_vehicle_specific.get("Achats", pd.DataFrame())),
                                         ("Carburant", df_vehicle_specific.get("Carburant", pd.DataFrame())),
                                         ("Suivi_Kilométrage", df_vehicle_specific.get("Suivi_Kilométrage", pd.DataFrame())),
                                         ("Parc_Véhicules", df_vehicules)]:
                df_sheet = df_sheet_orig.copy()
                # Formater les dates en texte français pour l'export
                date_cols = df_sheet.select_dtypes(include=['datetime64[ns]']).columns
                for col in date_cols:
                    df_sheet[col] = df_sheet[col].apply(format_date_fr)
                df_sheet.to_excel(writer, sheet_name=sheet_name, index=False)
                ws = writer.sheets[sheet_name]
                if sheet_name in sheets_money:
                    for col in sheets_money[sheet_name]:
                        if col in df_sheet.columns:
                            col_idx = list(df_sheet.columns).index(col) + 1
                            ws.set_column(col_idx, col_idx, None, money_format)
                if sheet_name in sheets_liter:
                    for col in sheets_liter[sheet_name]:
                        if col in df_sheet.columns:
                            col_idx = list(df_sheet.columns).index(col) + 1
                            ws.set_column(col_idx, col_idx, None, liter_format)

        buffer.seek(0)
        st.download_button(
            label="📥 Télécharger Rapport Excel",
            data=buffer,
            file_name=f"Rapport_{selected_vehicle}_{today.date().strftime('%Y%m%d')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
else:
    st.info("👆 Veuillez charger un fichier Excel pour commencer.")
# Footer fixe avec nom du créateur
st.markdown(
    """
    <style>
    .footer {
        position: fixed;
        bottom: 0;
        left: 0;
        width: 100%;
        background-color: #f0f2f6;
        border-top: 1px solid #d6d9dc;
        text-align: center;
        padding: 10px;
        font-size: 16px;
        z-index: 1000;
        color: #666;
    }
    </style>
    <div class="footer">
       <i style='color:red; font-weight:bold;'>Créé par RANAIVOSOA Tojoarimanana Hiratriniala / Tél : +261 33 51 880 19</i>
    </div>
    """,
    unsafe_allow_html=True
)