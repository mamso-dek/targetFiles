import os
import pandas as pd
from tqdm import tqdm


# %% Fonction pour récupérer tous les fichiers Excel
def get_excel_files(directory):
    excel_files = [os.path.join(directory, file) for file in os.listdir(directory) if file.endswith('.xlsx')]
    return excel_files


# %% Fonction pour reformater les données
def reformatData(file):
    # Lire les données depuis la feuille spécifique
    df = pd.read_excel(file, sheet_name='Statistiques Globales PE')

    departement = df.iloc[1, 2]  # Département
    commune = df.iloc[2, 2]  # Commune
    periode = df.iloc[4, 2]  # Période

    # Extraire l'année et le mois de la période
    annee = periode.split("AU")[0].split("/")[-1].strip()
    mois = periode.split("AU")[0].split("/")[1].strip()

    # Extraire les en-têtes et le corps des données
    headers = df.iloc[10, :].values
    body = df.iloc[11:20, :].copy()

    # Renommer les colonnes en fonction des en-têtes
    rename_cols = {old: new for old, new in zip(df.columns, headers)}
    body.rename(columns=rename_cols, inplace=True)

    # Ajouter des colonnes pour le département, la commune, l'année et le mois
    body.insert(loc=0, column='Departement', value=departement)
    body.insert(loc=1, column='Commune', value=commune)
    body.insert(loc=2, column='Année', value=annee)  # Insertion de la colonne 'Année'
    body.insert(loc=3, column='Mois', value=mois)  # Insertion de la colonne 'Mois'

    return body


# %% Fonction pour assembler les fichiers et trier les données
def assembler(files):
    # Fichier final
    final_df = pd.DataFrame()

    # Fusionner les fichiers
    for file in tqdm(files):
        df_ = reformatData(file)
        final_df = df_ if final_df.empty else pd.concat([final_df, df_], ignore_index=True)

    # Convertir les colonnes Année et Mois en entiers pour un tri correct
    final_df['Année'] = final_df['Année'].astype(int)
    final_df['Mois'] = final_df['Mois'].astype(int)

    # Trier les données par Année, Mois, Département et Commune
    final_df = final_df.sort_values(by=['Année', 'Mois', 'Departement', 'Commune'], ascending=True)

    # Sauvegarder le fichier final trié
    final_df.to_excel('final_df.xlsx', index=False)
    print("Fin de la fusion et du tri ...")

data_folder = 'data'
excel_files = get_excel_files(data_folder)

# %% Fusionner et trier les fichiers Excel
assembler(excel_files)

# %%
