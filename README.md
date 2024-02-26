# EUInvoiceFusion

# Fusion de Factures Excel

Ce projet fournit un script Python pour automatiser la fusion de multiples fichiers Excel contenant des factures par fournisseur au sein d'une organisation. Il identifie et fusionne les nouveaux fichiers, gère l'état des fichiers pour éviter les duplications et supprime les fichiers sources après la fusion.

## Fonctionnalités

- **Détection automatique des nouveaux fichiers Excel** : Identifie les nouveaux fichiers basés sur un motif spécifique dans un répertoire donné.
- **Fusion des données** : Combine les données de plusieurs fichiers Excel dans un seul fichier.
- **Gestion des états de fichiers** : Garde une trace des fichiers déjà traités pour éviter les traitements répétitifs.
- **Nettoyage** : Supprime les fichiers sources après leur fusion pour maintenir l'organisation du répertoire.

## Prérequis

- Python 3.x
- pandas
- openpyxl

## Installation

1. Clonez ce dépôt GitLab :
   ```
   git clone <url_du_dépôt>
   ```
2. Installez les dépendances requises :
   ```
   pip install pandas openpyxl
   ```

## Utilisation

Pour exécuter le script, naviguez dans le répertoire du projet et lancez :

```
python script_fusion_factures.py
```

Assurez-vous que les fichiers Excel à fusionner correspondent au motif défini dans le script et sont placés dans le même répertoire que le script.

## Structure du Projet

- `script_fusion_factures.py` : Script principal pour la fusion des fichiers.
- `files_state.json` : Stocke l'état des fichiers traités.
- `fused_files.json` : Stocke les noms des fichiers déjà fusionnés.

## Analyse et Tests

### Analyse de Code

Il est recommandé d'utiliser des outils tels que pylint ou flake8 pour analyser le code et s'assurer qu'il respecte les bonnes pratiques de codage en Python.


## Contribution

Les contributions sont les bienvenues ! Veuillez soumettre une demande de fusion (Merge Request) pour toute contribution.

## Licence

Ce projet est sous licence MIT. Voir le fichier LICENSE pour plus de détails.
