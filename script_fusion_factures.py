import os
import glob
import pandas as pd
import json
from collections import defaultdict
import shutil  #déplacer des fichiers


#######################################################################################
#fusionne automatiquement des fichiers Excel de factures 
#basés sur un motif de nom de fichier spécifique, 
#puis met à jour et sauvegarde l'état de ces 
#fichiers pour éviter les traitements répétitifs.
#######################################################################################
#charger le fichers json                                                              #
#######################################################################################
def load_json(file_name):
    try:
        with open(file_name, 'r') as f:
            return json.load(f)
    except FileNotFoundError:
        return {}
#############################################
##auvegarde des données dans un fichier JSON.
##############################################
def save_json(file_name, data):
    with open(file_name, 'w') as f:
        json.dump(data, f)
########################################################################################
# Extraction du Numéro à la Fin du Nom de Fichier (Avant l'Extension)                  #
########################################################################################
#
# Objectif :
# ----------
# Identifier et extraire le numéro situé à la fin du nom d'un fichier, juste avant 
# l'extension. Ce numéro est crucial car, dans le contexte de fichiers fusionnés, 
# il représente souvent un identifiant unique (ID) qui doit être le dernier ID mentionné.
#
# Importance :
# ------------
# L'extraction de cet ID permet de gérer et d'identifier correctement les fichiers,
# en assumant que le numéro le plus récent (et donc le dernier dans l'ordre) est 
# représentatif de la version ou de l'instance la plus à jour du fichier concerné.
#
# Exemple :
# ---------
# Nom de fichier : "rapport_financier_14015585.xlsx"
# ID extrait : 14015585
########################################################################################
        
def extract_file_number(filename):
    # Extrait le numéro à la fin du nom de fichier (avant l'extension)
    base_name = os.path.basename(filename)
    try:
        return int(base_name.split('_')[-1].split('.')[0])
    except ValueError:
        return 0
# Chemins et noms de fichiers
folder_path = os.getcwd()
pattern = "Invoice Header Report by Supplier Group EU_*.xlsx"
state_file = 'files_state.json'
fused_files_file = 'fused_files.json'

######################################################################################## 
# Dossier de sauvegarde
backup_folder = os.path.join(folder_path, 'backup') 
########################################################################################
# verfier si  sauvegarde existe ,créé si c'est pas le cas 
########################################################################################
if not os.path.exists(backup_folder):
    os.makedirs(backup_folder)

########################################################################################
#état précédent et les fichiers déjà fusionnés
########################################################################################
previous_state = load_json(state_file)
fused_files = load_json(fused_files_file)


########################################################################################
# Trouver les fichiers actuels
########################################################################################
files_found = [file for file in glob.glob(os.path.join(folder_path, pattern))]

########################################################################################
# Mettre à jour l'état actuel et sauvegarder
########################################################################################
current_state = {os.path.basename(file): True for file in files_found}
save_json(state_file, current_state)

########################################################################################
# Identifier les fichiers nouveaux ou manquants
new_files = [file for file in files_found if os.path.basename(file) not in fused_files]
missing_files = [file for file in previous_state if file not in current_state]

########################################################################################
# Signaler les fichiers manquants
########################################################################################
if missing_files:
    print("Fichier(s) manquant(s) depuis la dernière exécution :")
    for file in missing_files:
        print(file)
####################################################################################
# Gestion des Données dans les Fichiers Excel liés aux Transactions Commerciales   #
####################################################################################
#
# Structure des données :
# -----------------------
# - Colonne "Invoice_Number" : Numéro de facture pour chaque transaction.
# - Colonne "Buyer_TN" (Transaction Number de l'acheteur) : Identifiant de l'acheteur.
#
# Note Importante :
# -----------------
# Un même "Invoice_Number" peut être lié à un unique "Buyer_TN", mais correspondre à 
# plusieurs "Transaction_Number". Cela indique la possibilité de réaliser plusieurs 
# transactions sous une seule facture pour un acheteur donné.
#
# Exemple de Structure de Données :
#  ------------------------------------------------
# | Invoice_Number | Buyer_TN | Transaction_Number |
# |----------------|----------|--------------------|
# |     12345      |   TN001  |        TR001       |
# |     12345      |   TN001  |        TR002       |
# |     12346      |   TN002  |        TR003       |
#  ------------------------------------------------
#
# Ce tableau montre que la facture 12345, liée à l'acheteur TN001, couvre deux transactions
# distinctes (TR001 et TR002). Cela illustre la gestion flexible des transactions commerciales.
transaction_file_mapping = defaultdict(list)
for file in new_files:
    df = pd.read_excel(file, sheet_name=0)
    for index, row in df.iterrows():
        transaction_no = row['TRANSACTION NO']
        transaction_file_mapping[transaction_no].append((file, extract_file_number(file), index))
#####################################################################
# Sélectionner l'entrée avec le numéro de fichier le plus élevé
#####################################################################
selected_entries = []
for transaction_no, files in transaction_file_mapping.items():
    
    selected_file = max(files, key=lambda x: x[1])
    selected_entries.append((selected_file[0], selected_file[2]))
#####################################################################
# Fusionner les données sélectionnées
#####################################################################
data_frames = []
for file, index in selected_entries:
    df = pd.read_excel(file, sheet_name=0)
    data_frames.append(df.loc[[index]])

#####################################################################
# Fusionner uniquement les nouveaux fichiers
#####################################################################
data_frames = []
for file in new_files:
    df = pd.read_excel(file, sheet_name=0)
    data_frames.append(df)
    fused_files[os.path.basename(file)] = True  # Marquer comme fusionné

if data_frames:
    all_data = pd.concat(data_frames, ignore_index=True)
    output_file = "Fusion_Invoices.xlsx"
    all_data.to_excel(output_file, index=False)
    print(f"Les nouveaux fichiers ont été fusionnés dans {output_file}.")
else:
    print("Aucun nouveau fichier à fusionner.")

# Sauvegarder les fichiers fusionnés
save_json(fused_files_file, fused_files)

for file, index in selected_entries:
    base_file_name = os.path.basename(file)
    destination = os.path.join(backup_folder, base_file_name)
    try:
        shutil.move(file, destination)
        print(f"Le fichier {file} a été déplacé vers {destination}.")
    except Exception as e:
        print(f"Erreur lors du déplacement du fichier {file}: {e}")