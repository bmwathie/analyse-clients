import pandas as pd
import streamlit as st
import plotly.express as px
import plotly.graph_objects as go
import plotly.io as pio
from datetime import datetime, timedelta
import numpy as np
import io
import locale

# Configuration de base de Plotly - optimisation de performance
pio.templates.default = "plotly_white"
pio.renderers.default = "browser"

# Cache pour améliorer la performance
@st.cache_data(ttl=3600)
def load_and_process_data(file):
    return pd.read_excel(file)

# Définir la locale pour le format français des dates
try:
    locale.setlocale(locale.LC_TIME, 'fr_FR.UTF-8')
except:
    try:
        locale.setlocale(locale.LC_TIME, 'fr_FR')
    except:
        pass  # Si aucune locale française n'est disponible, on continue sans

# Configuration de la page Streamlit
st.set_page_config(
    page_title="Analyse des Clients",
    page_icon="📊",
    layout="wide"
)

# Titre principal
st.title("📊 Analyse de l'Activité des Clients")

# Sidebar pour télécharger le fichier
st.sidebar.header("Import des Données")
uploaded_file = st.sidebar.file_uploader("Téléchargez votre fichier Excel", type=["xlsx", "xls"])

# Options de débogage
st.sidebar.markdown("### Options")
debug_mode = st.sidebar.checkbox("Activer le mode débogage", value=False)

# Fonction utilitaire pour convertir les types numpy en types Python standard
def convert_numpy_types(value):
    """Convertit les types numpy en types Python standard."""
    if isinstance(value, (np.int64, np.int32, np.int16, np.int8)):
        return int(value)
    elif isinstance(value, (np.float64, np.float32, np.float16)):
        return float(value)
    elif isinstance(value, np.bool_):
        return bool(value)
    elif isinstance(value, np.ndarray):
        return value.tolist()
    return value

# Fonction pour traiter les données
@st.cache_data(ttl=3600)
def process_data(df):
    # Créer une copie du DataFrame pour éviter les avertissements SettingWithCopyWarning
    df = df.copy()
    
    # Convertir la colonne Date en datetime en tenant compte du format français (JJ/MM/AAAA)
    df['Date'] = pd.to_datetime(df['Date'], format='%d/%m/%Y', errors='coerce')
    
    # Gérer les valeurs manquantes dans la colonne Date
    if df['Date'].isna().any():
        st.warning(f"Attention: {df['Date'].isna().sum()} dates n'ont pas pu être converties. Vérifiez le format des dates.")
    
    # Créer une colonne pour l'année
    df['Année'] = df['Date'].dt.year
    
    # Créer une colonne pour le mois
    df['Mois'] = df['Date'].dt.month
    
    # Créer une colonne pour le jour de la semaine (0=lundi, 6=dimanche)
    df['jour_semaine'] = df['Date'].dt.dayofweek
    
    # Marquer les commandes du weekend (samedi=5, dimanche=6)
    df['est_weekend'] = df['jour_semaine'].isin([5, 6])
    
    # S'assurer que toutes les colonnes de type objet sont des chaînes de caractères
    object_columns = df.select_dtypes(include=['object']).columns
    for col in object_columns:
        df[col] = df[col].astype(str)
    
    # Identifier l'année actuelle pour les calculs d'activité
    current_year = datetime.now().year
    
    # Identifier les clients uniques
    unique_clients = df['Id Client'].unique()
    
    # Analyser l'activité des clients par année
    client_activity = {}
    for client in unique_clients:
        client_data = df[df['Id Client'] == client]
        client_years = client_data['Année'].unique()
        
        # Calculer les métriques du client
        nb_commandes = len(client_data)
        total_achats = client_data['Total'].sum()
        
        # Calculer l'ancienneté en mois
        premiere_date = client_data['Date'].min()
        derniere_date = client_data['Date'].max()
        anciennete_mois = (datetime.now() - premiere_date).days // 30
        
        # Calculer la fréquence des commandes par semaine et par mois
        if nb_commandes > 1:
            duree_semaines = max(1, (derniere_date - premiere_date).days / 7)
            frequence_semaine = nb_commandes / duree_semaines
            # Fréquence par mois (en weekends)
            frequence_mois = (nb_commandes / (duree_semaines / 4))
        else:
            frequence_semaine = 1 if nb_commandes == 1 else 0
            frequence_mois = 1 if nb_commandes == 1 else 0
        
        # Compter les commandes du weekend
        nb_commandes_weekend = len(client_data[client_data['est_weekend']])
        
        # Check si actif en 2023, 2024 et 2025
        actif_2023 = 2023 in client_years
        actif_2024 = 2024 in client_years
        actif_2025 = 2025 in client_years
        
        client_activity[client] = {
            'premiere_annee': min(client_years) if len(client_years) > 0 else None,
            'derniere_annee': max(client_years) if len(client_years) > 0 else None,
            'premiere_commande': premiere_date,
            'derniere_commande': derniere_date,
            'anciennete_mois': anciennete_mois,
            'années_actif': sorted(client_years),
            'actif_2023': actif_2023,
            'actif_2024': actif_2024,
            'actif_2025': actif_2025,
            'nb_commandes': nb_commandes,
            'nb_commandes_weekend': nb_commandes_weekend,
            'total_achats': total_achats,
            'frequence_semaine': frequence_semaine,
            'frequence_mois': frequence_mois
        }
    
    # Créer un DataFrame pour l'analyse des clients
    client_df = pd.DataFrame.from_dict(client_activity, orient='index')
    client_df.reset_index(inplace=True)
    client_df.rename(columns={'index': 'Id Client'}, inplace=True)
    
    # Fusionner avec les informations client
    # S'assurer que toutes les colonnes existent ou créer des colonnes vides par défaut
    for col in ['Prenom et nom', 'adresse email', 'Numero', 'Adresse']:
        if col not in df.columns:
            df[col] = ''

    # Assurez-vous que la colonne 'Id Client' est du même type dans les deux DataFrames
    client_df['Id Client'] = client_df['Id Client'].astype(str)
    df['Id Client'] = df['Id Client'].astype(str)
    
    client_info = df[['Id Client', 'Prenom et nom', 'adresse email', 'Numero', 'Adresse']].drop_duplicates('Id Client')
    client_analysis = pd.merge(client_df, client_info, on='Id Client', how='left')
    
    # Identifier les catégories selon les nouveaux critères
    client_analysis['Catégorie'] = 'Client Inactif'
    
    # Année actuelle pour référence
    current_year = datetime.now().year
    
    # Afficher des informations de débogage si activé
    if debug_mode:
        st.sidebar.write(f"### Année actuelle: {current_year}")
        unique_years = sorted(client_analysis['premiere_annee'].dropna().unique())
        st.sidebar.write(f"Années premières commandes: {unique_years}")
        unique_last_years = sorted(client_analysis['derniere_annee'].dropna().unique())
        st.sidebar.write(f"Années dernières commandes: {unique_last_years}")
    
    # Identifier les clients avec des dates futures (après 2025)
    if debug_mode:
        future_clients = client_analysis[(client_analysis['premiere_annee'] > 2025) | 
                                       (client_analysis['derniere_annee'] > 2025)]
        if not future_clients.empty:
            st.sidebar.write(f"Clients avec dates futures (>{current_year}): {len(future_clients)}")
            st.sidebar.dataframe(future_clients[['Id Client', 'premiere_annee', 'derniere_annee']], use_container_width=True)
    
    # 1. Actif depuis 2023 (première commande en 2023 et dernière en 2024 ou 2025)
    mask1 = ((client_analysis['premiere_annee'] == 2023) & (client_analysis['derniere_annee'] == 2025))
    client_analysis.loc[mask1, 'Catégorie'] = 'Actif depuis 2023'
    
    # 2. Actif depuis 2024 (première commande en 2024 et dernière en 2025)
    mask2 = (client_analysis['premiere_annee'] == 2024) & (client_analysis['derniere_annee'] == 2025)
    client_analysis.loc[mask2, 'Catégorie'] = 'Actif depuis 2024'
    
    # 3. Actif depuis 2025 (première et dernière commande en 2025)
    mask3 = (client_analysis['premiere_annee'] == 2025) & (client_analysis['derniere_annee'] == 2025)
    client_analysis.loc[mask3, 'Catégorie'] = 'Actif depuis 2025'
    
    # 4. Clients de 2023 uniquement (première et dernière en 2023)
    mask4 = (client_analysis['premiere_annee'] == 2023) & (client_analysis['derniere_annee'] == 2023)
    client_analysis.loc[mask4, 'Catégorie'] = 'Clients de 2023'
    
    # 5. Clients de 2024 uniquement (première et dernière en 2024)
    mask5 = (client_analysis['premiere_annee'] == 2024) & (client_analysis['derniere_annee'] == 2024)
    client_analysis.loc[mask5, 'Catégorie'] = 'Clients de 2024'
    
    # Débogage: afficher les clients encore marqués comme 'Client Inactif'
    if debug_mode:
        autres = client_analysis[client_analysis['Catégorie'] == 'Client Inactif']
        if not autres.empty:
            st.sidebar.write(f"### Clients classés comme 'Client Inactif': {len(autres)}")
            st.sidebar.dataframe(autres[['Id Client', 'premiere_annee', 'derniere_annee', 'Prenom et nom']], use_container_width=True)
    
    # Calcul du ratio de commandes weekend par client
    total_clients = len(client_analysis)
    total_commandes = df.shape[0]
    commandes_weekend = df[df['est_weekend']].shape[0]
    
    # Nombre moyen de commandes par week-end vs base client
    weekends_depuis_2023 = (datetime.now() - pd.Timestamp('2023-01-01')).days / 7 * 2
    ratio_commandes_we = commandes_weekend / weekends_depuis_2023 / total_clients if total_clients > 0 else 0
    
    return df, client_analysis, ratio_commandes_we

# Afficher les données si un fichier est téléchargé
if uploaded_file is not None:
    try:
        # Lire le fichier Excel
        df = load_and_process_data(uploaded_file)
        
        # Afficher les premières lignes du dataframe
        st.subheader("Aperçu des données")
        st.dataframe(df.head(), use_container_width=True)
        
        # Standardiser les noms de colonnes pour gérer les variations ou espaces
        df.columns = [col.strip() for col in df.columns]
        
        # Faire correspondre les colonnes aux noms attendus
        column_mappings = {
            'DateId Client': 'Date',  # Cas où les colonnes seraient fusionnées ou mal formatées
            'Id Client': 'Id Client',
            'Moyen de paiement': 'Moyen de Paiement',
            'email': 'adresse email',
            'Nom': 'Prenom et nom',
            'Telephone': 'Numero',
            'N° et Rue': 'Adresse',  # Mapper les anciennes colonnes d'adresse vers 'Adresse'
            'N° et rue': 'Adresse'    # Mapper les anciennes colonnes d'adresse vers 'Adresse'
        }
        
        # Appliquer les mappings si nécessaire
        for old_col, new_col in column_mappings.items():
            if old_col in df.columns and new_col not in df.columns:
                df.rename(columns={old_col: new_col}, inplace=True)
        
        # Vérifier les colonnes nécessaires
        required_columns = ['Date', 'Id Client', 'Total', 'adresse email', 'Prenom et nom']
        missing_columns = [col for col in required_columns if col not in df.columns]
    
        if missing_columns:
            st.error(f"Colonnes manquantes dans le fichier: {', '.join(missing_columns)}")
        else:
            # Si 'Adresse' n'existe pas, créer à partir des anciennes colonnes si disponibles
            if 'Adresse' not in df.columns:
                address_components = []
                
                # Vérifier et ajouter chaque composant d'adresse s'il existe
                if 'N° et rue' in df.columns:
                    address_components.append('N° et rue')
                if 'Code postal' in df.columns:
                    address_components.append('Code postal')
                if 'Ville' in df.columns:
                    address_components.append('Ville')
                
                if address_components:
                    # Créer une colonne Adresse en combinant les composants disponibles
                    df['Adresse'] = df[address_components].apply(
                        lambda row: ', '.join([str(row[c]) for c in address_components if pd.notna(row[c]) and str(row[c]).strip() != '']), 
                        axis=1
                    )
                else:
                    # Si aucun composant d'adresse n'est disponible, créer une colonne vide
                    df['Adresse'] = ''
            
            # Traiter les données
            processed_df, client_analysis, ratio_commandes_we = process_data(df)
            
            # Afficher les résultats de l'analyse
            st.subheader("Analyse des Clients")
            
            # Créer des onglets pour différentes vues
            tab1, tab2, tab3 = st.tabs(["📊 Vue d'ensemble", "🔍 Détails des Clients", "📥 Exporter les Résultats"])
            
            with tab1:
                # Préparer les comptages par catégorie pour l'affichage
                order = ['Actif depuis 2023', 'Actif depuis 2024', 'Actif depuis 2025', 'Clients de 2023', 'Clients de 2024', 'Client Inactif']
                category_counts = client_analysis['Catégorie'].value_counts().reset_index()
                category_counts.columns = ['Catégorie', 'Nombre de Clients']
                
                # Ajouter une colonne pour l'ordre d'affichage
                category_map = {cat: i for i, cat in enumerate(order)}
                category_counts['order'] = category_counts['Catégorie'].map(lambda x: category_map.get(x, 999))
                category_counts = category_counts.sort_values('order').drop('order', axis=1)
                
                # Créer une mise en page pour les métriques (1 ligne avec toutes les catégories)
                st.subheader("Nombre de clients par catégorie")
                cols = st.columns(len(order))
                
                # Afficher le nombre total de clients en premier
                total_col = st.columns(1)[0]
                total_col.metric("Nombre Total de Clients", len(client_analysis))
                
                # Afficher chaque catégorie dans une colonne
                for i, cat in enumerate(order):
                    if i < len(cols):
                        cat_count = category_counts[category_counts['Catégorie'] == cat]['Nombre de Clients'].values
                        cols[i].metric(cat, cat_count[0] if len(cat_count) > 0 else 0)
                
                # Graphique de répartition des catégories avec nombre de clients dans la légende
                # Créer une colonne supplémentaire pour l'étiquette de la légende
                category_counts['Étiquette'] = category_counts.apply(
                    lambda x: f"{x['Catégorie']} ({x['Nombre de Clients']} clients)", axis=1
                )
                
                fig = px.pie(
                    category_counts, 
                    values='Nombre de Clients', 
                    names='Étiquette',
                    title='Répartition des Clients par Catégorie',
                    color='Catégorie',
                    color_discrete_sequence=px.colors.qualitative.Set3
                )
                st.plotly_chart(fig, use_container_width=True)
                
                # Graphique de l'évolution du nombre de clients par année
                client_years = processed_df.groupby(['Année', 'Id Client']).size().reset_index()
                clients_per_year = client_years.groupby('Année').size().reset_index()
                clients_per_year.columns = ['Année', 'Nombre de Clients']
                
                fig = px.bar(
                    clients_per_year,
                    x='Année',
                    y='Nombre de Clients',
                    title="Évolution du Nombre de Clients par Année",
                    color_discrete_sequence=['#2E86C1']
                )
                st.plotly_chart(fig, use_container_width=True)
            
            with tab2:
                # Détails des clients avec filtre par catégorie
                st.subheader("Liste détaillée des clients")
                
                # Système de filtrage avancé
                st.write("### Filtres avancés")
                
                # Disposition en colonnes pour les filtres
                col1, col2, col3 = st.columns(3)
                
                # Filtre 1 : Par catégorie
                with col1:
                    category_filter = st.selectbox(
                        "Par catégorie",
                        ['Tous'] + order
                    )
                
                # Filtre 2 : Par pouvoir d'achat
                with col2:
                    min_total = int(client_analysis['total_achats'].min()) if not client_analysis.empty else 0
                    max_total = int(client_analysis['total_achats'].max()) if not client_analysis.empty else 1000
                    
                    spending_threshold = st.slider(
                        "Total des achats min (€)",
                        min_value=min_total,
                        max_value=max_total,
                        value=min_total,
                        step=50
                    )
                
                # Filtre 3 : Par nombre de commandes
                with col3:
                    min_orders = int(client_analysis['nb_commandes'].min()) if not client_analysis.empty else 0
                    max_orders = int(client_analysis['nb_commandes'].max()) if not client_analysis.empty else 50
                    
                    min_orders_filter = st.slider(
                        "Nombre de commandes min",
                        min_value=min_orders,
                        max_value=max_orders,
                        value=min_orders,
                        step=1
                    )
                
                # Filtres avancés supplémentaires (repliables)
                with st.expander("Plus de filtres"):
                    col1, col2 = st.columns(2)
                    
                    # Filtre par date de dernière commande
                    with col1:
                        # Trouver les dates min et max pour le slider
                        if not client_analysis.empty:
                            min_date = client_analysis['derniere_commande'].min().date()
                            max_date = client_analysis['derniere_commande'].max().date()
                        else:
                            min_date = datetime.now().date() - timedelta(days=365)
                            max_date = datetime.now().date()
                            
                        last_order_date = st.date_input(
                            "Dernière commande après le",
                            value=min_date,
                            min_value=min_date,
                            max_value=max_date
                        )
                    
                    # Filtre par fréquence de commande
                    with col2:
                        min_freq = float(client_analysis['frequence_mois'].min()) if not client_analysis.empty else 0
                        max_freq = float(client_analysis['frequence_mois'].max()) if not client_analysis.empty else 10
                        
                        min_frequency = st.slider(
                            "Fréquence min (cmd/mois)",
                            min_value=min_freq,
                            max_value=max_freq,
                            value=min_freq,
                            step=0.5
                        )
                
                # Appliquer tous les filtres
                filtered_clients = client_analysis.copy()
                
                # Filtrer par catégorie si nécessaire
                if category_filter != 'Tous':
                    filtered_clients = filtered_clients[filtered_clients['Catégorie'] == category_filter]
                
                # Filtrer par montant total des achats
                filtered_clients = filtered_clients[filtered_clients['total_achats'] >= spending_threshold]
                
                # Filtrer par nombre de commandes
                filtered_clients = filtered_clients[filtered_clients['nb_commandes'] >= min_orders_filter]
                
                # Filtrer par date de dernière commande
                filtered_clients = filtered_clients[filtered_clients['derniere_commande'].dt.date >= last_order_date]
                
                # Filtrer par fréquence
                filtered_clients = filtered_clients[filtered_clients['frequence_mois'] >= min_frequency]
                
                # Afficher le compteur de résultats
                st.write(f"**{len(filtered_clients)} clients** correspondent aux critères sélectionnés")
                
                # Afficher le tableau filtré avec les métriques demandées
                columns_to_display = [
                    'Id Client', 'Prenom et nom', 'adresse email', 'Catégorie',
                    'premiere_commande', 'derniere_commande', 'anciennete_mois',
                    'nb_commandes', 'total_achats', 'frequence_mois'
                ]
                
                # Formater les données pour l'affichage
                # Créer une copie pour éviter les erreurs de mise à jour SettingWithCopyWarning
                display_df = filtered_clients[columns_to_display].copy()
                
                # Formater chaque colonne individuellement
                display_df['total_achats'] = display_df['total_achats'].round(2).astype(str) + ' €'
                display_df['frequence_mois'] = display_df['frequence_mois'].round(2).astype(str) + ' cmd/mois'
                
                # Convertir les dates en chaînes au format français
                try:
                    display_df['premiere_commande'] = pd.to_datetime(display_df['premiere_commande']).dt.strftime('%d/%m/%Y')
                except:
                    display_df['premiere_commande'] = display_df['premiere_commande'].astype(str)
                    
                try:
                    display_df['derniere_commande'] = pd.to_datetime(display_df['derniere_commande']).dt.strftime('%d/%m/%Y')
                except:
                    display_df['derniere_commande'] = display_df['derniere_commande'].astype(str)
                
                # Convertir ancienneté_mois en chaîne avec unité
                display_df['anciennete_mois'] = display_df['anciennete_mois'].astype(str) + ' mois'
                
                st.dataframe(
                    display_df,
                    hide_index=True,
                    column_config={
                        'nb_commandes': "Nombre de commandes",
                        'total_achats': "Total des achats",
                        'frequence_mois': "Fréquence (cmd/mois)",
                        'premiere_commande': "Date première commande",
                        'derniere_commande': "Date dernière commande",
                        'anciennete_mois': "Ancienneté"
                    },
                    use_container_width=True
                )
                
                # Exporter les données filtrées
                export_columns = ['Id Client', 'Prenom et nom', 'adresse email', 'Catégorie', 'Numero',
                                 'Adresse', 'premiere_commande', 'derniere_commande', 'anciennete_mois',
                                 'nb_commandes', 'total_achats', 'frequence_mois']
                
                export_data = filtered_clients[export_columns].copy()
                
                # Formater l'export
                try:
                    export_data['premiere_commande'] = pd.to_datetime(export_data['premiere_commande']).dt.strftime('%d/%m/%Y')
                except:
                    export_data['premiere_commande'] = export_data['premiere_commande'].astype(str)
                    
                try:
                    export_data['derniere_commande'] = pd.to_datetime(export_data['derniere_commande']).dt.strftime('%d/%m/%Y')
                except:
                    export_data['derniere_commande'] = export_data['derniere_commande'].astype(str)
                
                # Options d'exportation multiples
                st.write("### Exporter les résultats filtrés")
                
                export_col1, export_col2, export_col3 = st.columns(3)
                
                date_suffix = datetime.now().strftime('%Y%m%d_%H%M')
                file_prefix = f"clients_{'tous' if category_filter == 'Tous' else category_filter.lower().replace(' ', '_')}_{date_suffix}"
                
                with export_col1:
                    # Export Excel
                    if not export_data.empty:
                        try:
                            excel_buffer = io.BytesIO()
                            
                            # Export simple sans formatage avancé pour éviter les erreurs
                            with pd.ExcelWriter(excel_buffer, engine='xlsxwriter') as writer:
                                export_data.to_excel(writer, sheet_name='Clients', index=False)
                                
                                # Formatage minimal
                                workbook = writer.book
                                worksheet = writer.sheets['Clients']
                                
                                # Ajuster les largeurs de colonnes
                                for i, col in enumerate(export_data.columns):
                                    # Calculer la largeur basée sur la longueur des valeurs
                                    col_width = max(
                                        len(str(col)) + 2,  # Largeur de l'en-tête + marge
                                        export_data[col].astype(str).str.len().max() + 2  # Largeur max des données + marge
                                    )
                                    worksheet.set_column(i, i, col_width)
                            
                            excel_buffer.seek(0)
                            
                            st.download_button(
                                label="📥 Télécharger en Excel",
                                data=excel_buffer,
                                file_name=f"{file_prefix}.xlsx",
                                mime="application/vnd.ms-excel"
                            )
                            
                            st.success("L'export Excel a été préparé avec succès. Cliquez sur le bouton pour télécharger.")
                            
                        except Exception as e:
                            st.error(f"Erreur lors de la création du fichier Excel: {str(e)}")
                            st.info("Essayez plutôt les formats CSV ou JSON comme alternatives.")
                    else:
                        st.warning("Aucune donnée à exporter.")
                
                with export_col2:
                    # Export CSV
                    if not export_data.empty:
                        csv_buffer = io.StringIO()
                        export_data.to_csv(csv_buffer, sep=';', encoding='utf-8-sig', index=False)
                        csv_buffer.seek(0)
                        
                        st.download_button(
                            label="📄 CSV",
                            data=csv_buffer.getvalue(),
                            file_name=f"{file_prefix}.csv",
                            mime="text/csv"
                        )
                    else:
                        st.warning("Aucune donnée à exporter.")
                
                with export_col3:
                    # Export JSON
                    if not export_data.empty:
                        json_data = export_data.to_json(orient='records', date_format='iso', force_ascii=False)
                        
                        st.download_button(
                            label="🔄 JSON",
                            data=json_data,
                            file_name=f"{file_prefix}.json",
                            mime="application/json"
                        )
                    else:
                        st.warning("Aucune donnée à exporter.")
                
                # Détails d'un client spécifique
                st.subheader("Détails d'un client spécifique")
                
                # Assurer que nous avons des clients à afficher
                if not filtered_clients.empty:
                    selected_client = st.selectbox(
                        "Sélectionnez un client",
                        filtered_clients['Id Client'].tolist()
                    )
                    
                    if selected_client:
                        client_details = client_analysis[client_analysis['Id Client'] == selected_client].iloc[0]
                        client_purchases = processed_df[processed_df['Id Client'] == selected_client]
                        
                        col1, col2 = st.columns(2)
                        
                        with col1:
                            st.write("**Informations générales**")
                            st.write(f"Nom: {client_details['Prenom et nom']}")
                            st.write(f"Email: {client_details['adresse email']}")
                            st.write(f"Téléphone: {client_details['Numero']}")
                            st.write(f"Adresse: {client_details['Adresse']}")
                            st.write(f"Catégorie: {client_details['Catégorie']}")
                            st.write(f"Ancienneté: {client_details['anciennete_mois']} mois")
                            st.write(f"Nombre de commandes: {client_details['nb_commandes']}")
                            st.write(f"Total des achats: {client_details['total_achats']:.2f} €")
                            
                            # S'assurer que les dates sont affichées correctement
                            try:
                                st.write(f"Première commande: {client_details['premiere_commande'].strftime('%d/%m/%Y')}")
                            except:
                                st.write(f"Première commande: {client_details['premiere_commande']}")
                                
                            try:
                                st.write(f"Dernière commande: {client_details['derniere_commande'].strftime('%d/%m/%Y')}")
                            except:
                                st.write(f"Dernière commande: {client_details['derniere_commande']}")
                        
                        with col2:
                            st.write("**Historique des commandes**")
                            if not client_purchases.empty:
                                # Graphique d'historique des achats
                                try:
                                    grouped_purchases = client_purchases.groupby(pd.Grouper(key='Date', freq='ME')).agg({'Total': 'sum'}).reset_index()
                                    
                                    fig = px.line(
                                        grouped_purchases,
                                        x='Date',
                                        y='Total',
                                        title=f"Historique des achats de {client_details['Prenom et nom']}",
                                        markers=True
                                    )
                                    st.plotly_chart(fig, use_container_width=True)
                                except Exception as e:
                                    st.error(f"Impossible de générer le graphique d'historique: {str(e)}")
                                    st.write(f"Nombre de commandes: {len(client_purchases)}")
                            else:
                                st.write("Aucun historique d'achat disponible pour ce client.")
                else:
                    st.info("Aucun client ne correspond à ce filtre.")
            
            with tab3:
                # Exportation des résultats
                st.subheader("Exporter les résultats de l'analyse")
                
                # Préparer les données pour l'exportation
                export_data = client_analysis[[
                    'Id Client', 'Prenom et nom', 'adresse email', 'Numero', 'Adresse', 
                    'Catégorie', 'premiere_commande', 'derniere_commande', 'anciennete_mois', 
                    'nb_commandes', 'total_achats', 'frequence_mois'
                ]].copy()
                
                # Système de filtrage pour l'export
                col1, col2 = st.columns(2)
                
                with col1:
                    # Option pour filtrer les données à exporter
                    export_option = st.radio(
                        "Que souhaitez-vous exporter ?",
                        ['Tous les clients'] + order
                    )
                    
                    if export_option != 'Tous les clients':
                        export_data = export_data[export_data['Catégorie'] == export_option]
                
                with col2:
                    # Filtres supplémentaires
                    include_metrics = st.checkbox("Inclure les métriques d'analyse", value=True)
                    include_contact = st.checkbox("Inclure les coordonnées complètes", value=True)
                    
                    # Si coordonnées désactivées, retirer les colonnes correspondantes
                    if not include_contact:
                        export_data = export_data.drop(columns=['adresse email', 'Numero', 'Adresse'])
                    
                    # Si métriques désactivées, retirer les colonnes correspondantes
                    if not include_metrics:
                        export_data = export_data.drop(columns=['anciennete_mois', 'frequence_mois'])
                
                # Formater les dates pour l'export
                try:
                    export_data['premiere_commande'] = pd.to_datetime(export_data['premiere_commande']).dt.strftime('%d/%m/%Y')
                except:
                    export_data['premiere_commande'] = export_data['premiere_commande'].astype(str)
                    
                try:
                    export_data['derniere_commande'] = pd.to_datetime(export_data['derniere_commande']).dt.strftime('%d/%m/%Y')
                except:
                    export_data['derniere_commande'] = export_data['derniere_commande'].astype(str)
                
                # Prévisualisation des données à exporter
                st.write("### Aperçu des données à exporter")
                st.dataframe(export_data.head(5), use_container_width=True)
                
                # Information sur le nombre de lignes
                st.info(f"Le fichier exporté contiendra {len(export_data)} lignes de données.")
                
                # Formats d'exportation disponibles
                st.write("### Formats d'exportation disponibles")
                export_tab1, export_tab2, export_tab3, export_tab4 = st.tabs(["Excel", "CSV", "JSON", "Rapport PDF"])
                
                # Préfixe de nom de fichier basé sur la date et la sélection
                date_suffix = datetime.now().strftime('%Y%m%d_%H%M')
                file_prefix = f"analyse_clients_{export_option.lower().replace(' ', '_')}_{date_suffix}"
                
                with export_tab1:
                    # Excel avec options avancées
                    st.write("#### Export Excel")
                    
                    # Bouton pour Excel
                    if not export_data.empty:
                        try:
                            excel_buffer = io.BytesIO()
                            
                            # Export simple sans formatage avancé pour éviter les erreurs
                            with pd.ExcelWriter(excel_buffer, engine='xlsxwriter') as writer:
                                export_data.to_excel(writer, sheet_name='Clients', index=False)
                                
                                # Formatage minimal
                                workbook = writer.book
                                worksheet = writer.sheets['Clients']
                                
                                # Ajuster les largeurs de colonnes
                                for i, col in enumerate(export_data.columns):
                                    # Calculer la largeur basée sur la longueur des valeurs
                                    col_width = max(
                                        len(str(col)) + 2,  # Largeur de l'en-tête + marge
                                        export_data[col].astype(str).str.len().max() + 2  # Largeur max des données + marge
                                    )
                                    worksheet.set_column(i, i, col_width)
                            
                            excel_buffer.seek(0)
                            
                            st.download_button(
                                label="📥 Télécharger en Excel",
                                data=excel_buffer,
                                file_name=f"{file_prefix}.xlsx",
                                mime="application/vnd.ms-excel"
                            )
                            
                            st.success("L'export Excel a été préparé avec succès. Cliquez sur le bouton pour télécharger.")
                            
                        except Exception as e:
                            st.error(f"Erreur lors de la création du fichier Excel: {str(e)}")
                            st.info("Essayez plutôt les formats CSV ou JSON comme alternatives.")
                    else:
                        st.warning("Aucune donnée à exporter.")
                
                with export_tab2:
                    # Export CSV avec options
                    st.write("#### Export CSV")
                    
                    separator = st.selectbox(
                        "Séparateur", 
                        options=[";", ",", "|", "Tab"],
                        index=0
                    )
                    
                    if separator == "Tab":
                        separator = "\t"
                    
                    encoding = st.selectbox(
                        "Encodage",
                        options=["utf-8-sig", "utf-8", "latin1"],
                        index=0,
                        help="utf-8-sig est recommandé pour Excel"
                    )
                    
                    if not export_data.empty:
                        csv_buffer = io.StringIO()
                        export_data.to_csv(csv_buffer, sep=separator, encoding=encoding, index=False)
                        csv_buffer.seek(0)
                        
                        st.download_button(
                            label="📥 Télécharger en CSV",
                            data=csv_buffer.getvalue(),
                            file_name=f"{file_prefix}.csv",
                            mime="text/csv",
                            help="Télécharge un fichier CSV compatible avec Excel et autres logiciels"
                        )
                    else:
                        st.warning("Aucune donnée à exporter.")
                
                with export_tab3:
                    # Export JSON avec options
                    st.write("#### Export JSON")
                    
                    orient_option = st.selectbox(
                        "Format JSON",
                        options=[
                            "records (liste d'objets)",
                            "index (dictionnaire clé-valeur)",
                            "columns (format colonnes)"
                        ],
                        index=0
                    )
                    
                    orient_map = {
                        "records (liste d'objets)": "records",
                        "index (dictionnaire clé-valeur)": "index",
                        "columns (format colonnes)": "columns"
                    }
                    
                    orient = orient_map[orient_option]
                    
                    if not export_data.empty:
                        json_data = export_data.to_json(orient=orient, date_format='iso', force_ascii=False)
                        
                        st.download_button(
                            label="📥 Télécharger en JSON",
                            data=json_data,
                            file_name=f"{file_prefix}.json",
                            mime="application/json",
                            help="Télécharge un fichier JSON pour intégration technique"
                        )
                    else:
                        st.warning("Aucune donnée à exporter.")
                
                with export_tab4:
                    st.write("#### Rapport PDF")
                    st.info("""
                    Pour générer un rapport PDF complet:
                    
                    1. Exportez d'abord les données en Excel
                    2. Utilisez l'option d'impression vers PDF dans Excel
                    3. Vous pouvez également capturer les graphiques individuellement
                    """)
                    
                    # Générer un aperçu du rapport
                    st.write("#### Aperçu du rapport")
                    
                    # Créer des visualisations pour le rapport
                    fig_col1, fig_col2 = st.columns(2)
                    
                    with fig_col1:
                        # Créer un graphique pour l'exportation avec une configuration optimisée
                        fig = px.pie(
                            category_counts, 
                            values='Nombre de Clients', 
                            names='Étiquette',
                            title='Répartition des Clients par Catégorie',
                            color='Catégorie',
                            color_discrete_sequence=px.colors.qualitative.Set3,
                            width=600,
                            height=400
                        )
                        
                        # Afficher le graphique
                        st.plotly_chart(fig)
                    
                    with fig_col2:
                        # Graphique de répartition des commandes par catégorie
                        orders_by_cat = pd.DataFrame()
                        for cat in order:
                            cat_data = client_analysis[client_analysis['Catégorie'] == cat]
                            if not cat_data.empty:
                                orders_by_cat.loc[cat, 'Commandes'] = cat_data['nb_commandes'].sum()
                                orders_by_cat.loc[cat, 'CA Total'] = cat_data['total_achats'].sum()
                        
                        # Créer un bar chart pour les commandes
                        fig2 = px.bar(
                            orders_by_cat,
                            x=orders_by_cat.index,
                            y='Commandes',
                            title="Nombre de Commandes par Catégorie(je sais que c'est impertinent et que ca ne veut rien dire mais matay 🤣)",
                            color=orders_by_cat.index,
                            color_discrete_sequence=px.colors.qualitative.Set3,
                            width=600,
                            height=400
                        )
                        
                        st.plotly_chart(fig2)
                    
                    st.info("""
                    Pour télécharger les graphiques :
                    1. Survolez le graphique
                    2. Cliquez sur l'icône appareil photo 📸 dans la barre d'outils
                    3. Choisissez "Download plot as PNG"
                    """)
                
                # Option pour exporter des graphiques
                st.write("### Graphiques additionnels pour l'analyse")
                
                try:
                    # Créer un graphique pour l'exportation avec une configuration optimisée
                    fig3 = px.bar(
                        category_counts,
                        x='Catégorie',
                        y='Nombre de Clients',
                        title='Nombre de Clients par Catégorie',
                        color='Catégorie',
                        color_discrete_sequence=px.colors.qualitative.Set3,
                        width=1200,
                        height=500
                    )
                    
                    # Afficher le graphique
                    st.plotly_chart(fig3, use_container_width=True)
                    
                    # Graphique d'évolution mensuelle des commandes
                    if 'Date' in processed_df.columns:
                        monthly_orders = processed_df.groupby(pd.Grouper(key='Date', freq='M')).size().reset_index()
                        monthly_orders.columns = ['Mois', 'Nombre de Commandes']
                        
                        fig4 = px.line(
                            monthly_orders,
                            x='Mois',
                            y='Nombre de Commandes',
                            title="Évolution mensuelle du nombre de commandes",
                            markers=True,
                            width=1200,
                            height=500
                        )
                        
                        st.plotly_chart(fig4, use_container_width=True)
                except Exception as e:
                    st.error(f"Erreur lors de la création des graphiques : {str(e)}")
                
    except Exception as e:
        st.error(f"Une erreur s'est produite lors du traitement du fichier: {e}")
else:
    # Afficher des instructions si aucun fichier n'est téléchargé
    st.info("Veuillez télécharger votre fichier Excel pour commencer l'analyse.")
    
    # Exemple de format attendu
    st.subheader("Format de données attendu")
    example_data = {
        'Date': ['01/10/2023', '29/10/2023', '12/11/2023', '16/12/2023', '11/02/2024'],
        'Id Client': [1001, 1001, 1001, 1001, 1001],
        'Moyen de Paiement': ['Espèce', 'Espèce', 'Espèce', 'Espèce', 'Espèce'],
        'Total': [0, 0, 0, 0, 0],
        'adresse email': ['', '', '', '', 'mariamafallwathie@gmail.com'],
        'Prenom et nom': ['Mme wathie', 'Mme Wathie', 'Fall', 'Fall', 'mme wathie'],
        'Numero': ['624748439', '624847439', '624847439', '624847439', '624748439'],
        'Adresse': ['5rue Jules Massenet, 78330, Fontenay Le Fleury', '5rue, 78330, Fontenay Le Fleury', '5 Rue Jules Mas, 78330, Fontenay', '5rue jules masse, 78330, Fontenay Le Fleury', '5rue, 78330, Fontenay Le Fleury'],
    }
    for key in example_data:
        example_data[key] = [str(val) if pd.notna(val) else '' for val in example_data[key]]
    example_df = pd.DataFrame(example_data)
    st.dataframe(example_df)
    
    # Instructions d'utilisation
    st.subheader("Instructions d'utilisation")
    st.markdown("""
    1. Téléchargez votre fichier Excel en utilisant le bouton dans la barre latérale gauche
    2. L'application analysera automatiquement vos données et identifiera:
       - Actif depuis 2023 (première commande en 2023 et dernière en 2025)
       - Actif depuis 2024 (première commande en 2024 et dernière en 2025)
       - Actif depuis 2025 (première commande en 2025 et dernière en 2025)
       - Clients de 2023 uniquement
       - Clients de 2024 uniquement
    3. Vous pourrez visualiser les résultats sous forme de graphiques et tableaux
    4. Explorez les détails de chaque client en les sélectionnant dans la liste
    5. Exportez les résultats au format Excel ou les graphiques au format image
    
    **Note importante:** L'application traite les dates au format français (JJ/MM/AAAA). Assurez-vous que vos dates sont dans ce format.
    """)

# Pied de page
st.markdown("---")
st.markdown("BMW") 
