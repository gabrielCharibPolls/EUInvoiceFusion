import os
import glob
import pandas as pd
import json


########################################################
#fusionne automatiquement des fichiers Excel de factures 
#basés sur un motif de nom de fichier spécifique, 
#puis met à jour et sauvegarde l'état de ces 
#fichiers pour éviter les traitements répétitifs.
######################################################



###########################################
#charger le fichers json
###########################################
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

# Chemins et noms de fichiers
folder_path = os.getcwd()
pattern = "Invoice Header Report by Supplier Group EU_*.xlsx"
state_file = 'files_state.json'
fused_files_file = 'fused_files.json'

#############################################
#état précédent et les fichiers déjà fusionnés
#############################################
previous_state = load_json(state_file)
fused_files = load_json(fused_files_file)

#############################################
# Trouver les fichiers actuels
#############################################
files_found = [file for file in glob.glob(os.path.join(folder_path, pattern))]

# Mettre à jour l'état actuel et sauvegarder
current_state = {os.path.basename(file): True for file in files_found}
save_json(state_file, current_state)

# Identifier les fichiers nouveaux ou manquants
new_files = [file for file in files_found if os.path.basename(file) not in fused_files]
missing_files = [file for file in previous_state if file not in current_state]
#############################################
# Signaler les fichiers manquants
#############################################
if missing_files:
    print("Fichier(s) manquant(s) depuis la dernière exécution :")
    for file in missing_files:
        print(file)

# Fusionner uniquement les nouveaux fichiers
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

# Supprimer les fichiers fusionnés
for file in new_files:
    try:
        os.remove(file)
        print(f"Le fichier {file} a été supprimé.")
    except OSError as e:
        print(f"Erreur lors de la suppression du fichier {file}: {e.strerror}")
