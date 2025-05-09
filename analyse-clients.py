import pandas as pd
import streamlit as st
import plotly.express as px
import plotly.graph_objects as go
import plotly.io as pio
from datetime import datetime
import numpy as np
import io
import locale

# Configuration de base de Plotly
pio.templates.default = "plotly_white"
pio.renderers.default = "browser"

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

# Fonction pour traiter les données
def process_data(df):
    # Convertir la colonne Date en datetime en tenant compte du format français (JJ/MM/AAAA)
    df['Date'] = pd.to_datetime(df['Date'], format='%d/%m/%Y', errors='coerce')
    
    # Gérer les valeurs manquantes dans la colonne Date
    if df['Date'].isna().any():
        st.warning(f"Attention: {df['Date'].isna().sum()} dates n'ont pas pu être converties. Vérifiez le format des dates.")
    
    # Créer une colonne pour l'année
    df['Année'] = df['Date'].dt.year
    
    # Créer une colonne pour le mois
    df['Mois'] = df['Date'].dt.month
    
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
        client_years = df[df['Id Client'] == client]['Année'].unique()
        client_data = df[df['Id Client'] == client]
        
        # Calculer les métriques du client
        nb_commandes = len(client_data)
        total_achats = client_data['Total'].sum()
        
        # Calculer la fréquence des commandes par semaine
        if nb_commandes > 1:
            date_min = client_data['Date'].min()
            date_max = client_data['Date'].max()
            duree_semaines = max(1, (date_max - date_min).days / 7)
            frequence_semaine = nb_commandes / duree_semaines
        else:
            frequence_semaine = 1 if nb_commandes == 1 else 0
        
        client_activity[client] = {
            'première_année': min(client_years) if len(client_years) > 0 else None,
            'dernière_année': max(client_years) if len(client_years) > 0 else None,
            'années_actif': sorted(client_years),
            'actif_2024': 2024 in client_years,
            'actif_2025': 2025 in client_years,
            'nb_commandes': nb_commandes,
            'total_achats': total_achats,
            'frequence_semaine': frequence_semaine,
            'derniere_commande': client_data['Date'].max()
        }
    
    # Créer un DataFrame pour l'analyse des clients
    client_df = pd.DataFrame.from_dict(client_activity, orient='index')
    client_df.reset_index(inplace=True)
    client_df.rename(columns={'index': 'Id Client'}, inplace=True)
    
    # Fusionner avec les informations client
    client_info = df[['Id Client', 'Prenom et nom', 'adresse email', 'Numero', 'N° et rue', 'Code postal', 'Ville']].drop_duplicates('Id Client')
    client_analysis = pd.merge(client_df, client_info, on='Id Client', how='left')
    
    # Identifier les catégories demandées
    client_analysis['Catégorie'] = 'Autre'
    
    # Clients depuis 2024 ou 2023 et toujours actifs (commandes en 2025, mais on ne compte que les achats à partir de 2024)
    mask_2023_2024 = ((client_analysis['première_année'] == 2023) | (client_analysis['première_année'] == 2024)) & (client_analysis['actif_2025'] == True)
    client_analysis.loc[mask_2023_2024, 'Catégorie'] = 'Clients depuis 2023/2024 et toujours actifs'
    
    # Clients depuis 2024 et plus là (première et dernière commande en 2024)
    client_analysis.loc[(client_analysis['première_année'] == 2024) & 
                       (client_analysis['dernière_année'] == 2024), 'Catégorie'] = 'Clients depuis 2024 et plus là'

    # Nouveaux clients 2025 : première et dernière commande en 2025
    client_analysis.loc[(client_analysis['première_année'] == 2025) & 
                       (client_analysis['dernière_année'] == 2025), 'Catégorie'] = 'Nouveaux clients 2025'

    # Pour les clients "Clients depuis 2023/2024 et toujours actifs", recalculer les métriques en ne prenant que les achats à partir de 2024
    ids_2023_2024 = client_analysis.loc[mask_2023_2024, 'Id Client']
    for cid in ids_2023_2024:
        client_data = df[(df['Id Client'] == cid) & (df['Année'] >= 2024)]
        nb_commandes = len(client_data)
        total_achats = client_data['Total'].sum()
        if nb_commandes > 1:
            date_min = client_data['Date'].min()
            date_max = client_data['Date'].max()
            duree_semaines = max(1, (date_max - date_min).days / 7)
            frequence_semaine = nb_commandes / duree_semaines
        else:
            frequence_semaine = 1 if nb_commandes == 1 else 0
        derniere_commande = client_data['Date'].max() if nb_commandes > 0 else None
        client_analysis.loc[client_analysis['Id Client'] == cid, 'nb_commandes'] = nb_commandes
        client_analysis.loc[client_analysis['Id Client'] == cid, 'total_achats'] = total_achats
        client_analysis.loc[client_analysis['Id Client'] == cid, 'frequence_semaine'] = frequence_semaine
        client_analysis.loc[client_analysis['Id Client'] == cid, 'derniere_commande'] = derniere_commande
    
    return df, client_analysis

# Afficher les données si un fichier est téléchargé
if uploaded_file is not None:
    try:
        # Lire le fichier Excel
        df = pd.read_excel(uploaded_file)
        
        # Afficher les premières lignes du dataframe
        st.subheader("Aperçu des données")
        st.dataframe(df.head())
        
        # Standardiser les noms de colonnes pour gérer les variations ou espaces
        df.columns = [col.strip() for col in df.columns]
        
        # Faire correspondre les colonnes aux noms attendus
        column_mappings = {
            'DateId Client': 'Date',  # Cas où les colonnes seraient fusionnées ou mal formatées
            'Id Client': 'Id Client',
            'Moyen de paiement': 'Moyen de Paiement',
            'email': 'adresse email',
            'Nom': 'Prenom et nom',
            'Telephone': 'Numero'
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
            # Traiter les données
            processed_df, client_analysis = process_data(df)
            
            # Afficher les résultats de l'analyse
            st.subheader("Analyse des Clients")
            
            # Créer des onglets pour différentes vues
            tab1, tab2, tab3 = st.tabs(["📊 Vue d'ensemble", "🔍 Détails des Clients", "📥 Exporter les Résultats"])
            
            with tab1:
                # Vue d'ensemble avec des chiffres clés
                col1, col2, col3 = st.columns(3)
                
                # Nombre total de clients
                with col1:
                    st.metric("Nombre Total de Clients", len(client_analysis))
                
                # Clients depuis 2023/2024 et toujours actifs
                with col2:
                    active_2024_count = len(client_analysis[client_analysis['Catégorie'] == 'Clients depuis 2023/2024 et toujours actifs'])
                    st.metric("Clients depuis 2023/2024 et toujours actifs", active_2024_count)
                
                # Nouveaux clients 2025
                with col3:
                    new_2025_count = len(client_analysis[client_analysis['Catégorie'] == 'Nouveaux clients 2025'])
                    st.metric("Nouveaux clients 2025", new_2025_count)
                
                # Graphique de répartition des catégories
                category_counts = client_analysis['Catégorie'].value_counts().reset_index()
                category_counts.columns = ['Catégorie', 'Nombre de Clients']
                
                fig = px.pie(
                    category_counts, 
                    values='Nombre de Clients', 
                    names='Catégorie',
                    title='Répartition des Clients par Catégorie',
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
                
                # Filtrer par catégorie
                category_filter = st.selectbox(
                    "Filtrer par catégorie",
                    ['Tous'] + list(client_analysis['Catégorie'].unique())
                )
                
                if category_filter == 'Tous':
                    filtered_clients = client_analysis
                else:
                    filtered_clients = client_analysis[client_analysis['Catégorie'] == category_filter]
                
                # Afficher le tableau filtré avec les métriques de poids
                columns_to_display = [
                    'Id Client', 'Prenom et nom', 'adresse email', 'Catégorie',
                    'nb_commandes', 'total_achats', 'frequence_semaine', 'derniere_commande'
                ]
                
                # Formater les données pour l'affichage
                display_df = filtered_clients[columns_to_display].copy()
                display_df['total_achats'] = display_df['total_achats'].round(2).apply(lambda x: f"{x:,.2f} €")
                display_df['frequence_semaine'] = display_df['frequence_semaine'].round(2).apply(lambda x: f"{x:.2f} / semaine")
                display_df['derniere_commande'] = display_df['derniere_commande'].dt.strftime('%d/%m/%Y')
                
                st.dataframe(
                    display_df,
                    hide_index=True,
                    column_config={
                        'nb_commandes': "Nombre de commandes",
                        'total_achats': "Total des achats",
                        'frequence_semaine': "Fréquence (commandes/semaine)",
                        'derniere_commande': "Dernière commande"
                    }
                )
                
                # Détails d'un client spécifique
                st.subheader("Détails d'un client spécifique")
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
                        st.write(f"Adresse: {client_details['N° et rue']}, {client_details['Code postal']} {client_details['Ville']}")
                        st.write(f"Catégorie: {client_details['Catégorie']}")
                        st.write(f"Première année: {client_details['première_année']}")
                        st.write(f"Dernière année: {client_details['dernière_année']}")
                    
                    with col2:
                        st.write("**Historique des commandes**")
                        if not client_purchases.empty:
                            # Graphique d'historique des achats
                            fig = px.line(
                                client_purchases.groupby(pd.Grouper(key='Date', freq='ME')).agg({'Total': 'sum'}).reset_index(),
                                x='Date',
                                y='Total',
                                title=f"Historique des achats de {client_details['Prenom et nom']}",
                                markers=True
                            )
                            st.plotly_chart(fig, use_container_width=True)
                        else:
                            st.write("Aucun historique d'achat disponible pour ce client.")
            
            with tab3:
                # Exportation des résultats
                st.subheader("Exporter les résultats de l'analyse")
                
                # Préparer les données pour l'exportation
                export_data = client_analysis[['Id Client', 'Prenom et nom', 'adresse email', 'Numero', 'Catégorie', 'première_année', 'dernière_année']]
                
                # Option pour filtrer les données à exporter
                export_option = st.radio(
                    "Que souhaitez-vous exporter ?",
                    ['Tous les clients', 'Clients depuis 2023/2024 et toujours actifs', 'Clients depuis 2024 et plus là', 'Nouveaux clients 2025']
                )
                
                if export_option != 'Tous les clients':
                    export_data = export_data[export_data['Catégorie'] == export_option]
                
                # Bouton pour déclencher l'exportation Excel
                buffer = io.BytesIO()
                
                with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
                    export_data.to_excel(writer, sheet_name='Analyse Clients', index=False)
                
                buffer.seek(0)
                
                st.download_button(
                    label="📥 Télécharger l'analyse en Excel",
                    data=buffer,
                    file_name=f"analyse_clients_{datetime.now().strftime('%Y%m%d')}.xlsx",
                    mime="application/vnd.ms-excel"
                )
                
                # Option pour exporter des graphiques
                st.write("**Exporter des graphiques**")
                
                try:
                    # Créer un graphique pour l'exportation avec une configuration optimisée
                    fig = px.pie(
                        category_counts, 
                        values='Nombre de Clients', 
                        names='Catégorie',
                        title='Répartition des Clients par Catégorie',
                        color_discrete_sequence=px.colors.qualitative.Set3,
                        width=1200,
                        height=800
                    )
                    
                    # Afficher le graphique avec les outils d'export
                    st.plotly_chart(fig, use_container_width=True)
                    
                    st.info("""
                    Pour télécharger le graphique :
                    1. Survolez le graphique
                    2. Cliquez sur l'icône appareil photo 📸 dans la barre d'outils
                    3. Choisissez "Download plot as PNG"
                    """)
                    
                except Exception as e:
                    st.error(f"Erreur lors de la création du graphique : {str(e)}")
                
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
        'N° et rue': ['5rue Jule Massenet', '5rue', '5 Rue Jules Mas', '5rue jules masse', '5rue'],
        'Code postal': ['', '78330', '78330', '78330', '78330'],
        'Ville': ['', 'Fontenay Le Fleury', 'Fontenay', 'Fontenay Le Fleury', 'Fontenay Le Fleury']
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
       - Les clients qui sont avec vous depuis 2024 et toujours actifs
       - Les clients arrivés en 2025 et plus là (inactifs)
    3. Vous pourrez visualiser les résultats sous forme de graphiques et tableaux
    4. Explorez les détails de chaque client en les sélectionnant dans la liste
    5. Exportez les résultats au format Excel ou les graphiques au format image
    
    **Note importante:** L'application traite les dates au format français (JJ/MM/AAAA). Assurez-vous que vos dates sont dans ce format.
    """)

# Pied de page
st.markdown("---")
st.markdown("BMW")