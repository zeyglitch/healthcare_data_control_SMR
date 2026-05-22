# 🩺 Conciliation DIM - Contrôle d'exhaustivité SMR

Bienvenue dans l'outil de conciliation dédié aux Soins Médicaux et de Réadaptation (SMR) pour le département d'Information Médicale (DIM). Dans la lignée des autres outils, ce petit logiciel a été conçu pour simplifier la vie des codeurs en automatisant la comparaison entre les exports **Orbis SMR** et **Hexagone SMR**.

L'objectif ? Repérer en un clin d'œil les venues manquantes dans l'un ou l'autre système en convertissant intelligemment les numéros de semaines et les grilles de présence en dates exactes !

## 🌟 Ce que fait l'outil pour vous

- **Transformation magique** : Convertit les présences sous forme de grille hebdomadaire (ex: "L.M.V..") en dates de venues individuelles exactes.
- **Nettoyage automatique** : Filtre les lignes inutiles (comme les doublons de fin de séjour "SD" dans l'export Hexagone).
- **Analyse croisée pointue** : Compare automatiquement chaque venue via la clé "NDA + Date" entre vos deux fichiers.
- **Rapports clairs** : Génère des fichiers Excel colorés identifiant clairement l'origine de l'anomalie (manquant dans Orbis ou dans Hexagone), ainsi qu'une synthèse chiffrée.

## 📁 Comment est organisé le projet ?

Pour que tout soit bien rangé, voici la structure de travail :
- `src/` : Contient la "mécanique" (les scripts Python `controle_smr.py` et `interface_smr.py` dans le dossier `code/`). Vous y trouverez aussi le mode d'emploi utilisateur dans `mode_logiciel/`.
- `fichiers_tests/` : Les jeux de données d'exemple pour tester l'outil sereinement (`Orbis_SMR_Test.xlsx` et `Hexagone_SMR_Test.xlsx`).
- `data_test/` : C'est le dossier de destination par défaut pour la sortie de vos analyses (les synthèses et rapports d'écarts).
- `docs/` : Vous y trouverez le **Guide Développeur** si vous souhaitez plonger sous le capot et modifier le comportement (notamment la logique métier sur les dates).

## 🛠️ Installation (pour la première fois)

1. Assurez-vous d'avoir **Python (3.9 ou plus)** installé sur votre ordinateur.
2. Dans votre terminal, à la racine du projet, installez les outils nécessaires (de préférence dans un environnement virtuel) :
   ```bash
   python -m venv venv
   # Sur Windows (PowerShell) :
   venv\Scripts\Activate.ps1 
   pip install -r requirements.txt
   ```

## 🚀 Comment l'utiliser ?

### Mode Classique (Interface Graphique)
C'est la méthode la plus simple pour l'utiliser au quotidien. Lancez cette commande depuis le dossier principal :
```bash
python src/code/interface_smr.py
```
Une fenêtre s'ouvrira : choisissez vos fichiers d'entrée Orbis et Hexagone, indiquez un dossier d'export et cliquez sur "Lancer le traitement". Et voilà !

### Créer un fichier .exe (Application autonome)
Si vous voulez partager l'outil à un collègue qui n'a pas Python d'installé, vous pouvez transformer le script en une vraie application Windows :
```bash
pip install pyinstaller
cd src/code
pyinstaller --noconsole --onefile --name "Controle_SMR_DIM" interface_smr.py
```
Le résultat (`Controle_SMR_DIM.exe`) sera généré dans le dossier `dist/`. Vous pouvez le récupérer et le distribuer sans soucis !
