# Contrôle d'exhaustivité PMSI — SMR (Orbis SMR ↔ Hexagone SMR)

Outil Python de rapprochement automatique des venues entre les exports Orbis SMR et Hexagone SMR, pour un Département d'Information Médicale (DIM) hospitalier — activité de Soins Médicaux et de Réadaptation (SMR).

Version SMR complémentaire du dépôt [healthcare-data-reconciliation](https://github.com/zeyglitch/healthcare-data-reconciliation) (MCO).

## Contexte

Développé pendant un stage de 3ème année d'école d'ingénieur (ISIS Castres, partenaire INSA) au Département d'Information Médicale d'un Centre Hospitalier public, entre avril et juin 2026.

L'activité SMR pose une difficulté spécifique : Orbis exprime les venues sous forme de **grilles hebdomadaires de présence** (par exemple `L.M.V..` pour lundi/mardi/vendredi), tandis qu'Hexagone les stocke à la venue individuelle. Le rapprochement manuel imposait de reconstruire mentalement les dates pour chaque semaine, chaque patient — travail long et à haut risque d'erreur.

## Résultats

- Traitement de **2 fichiers d'export SMR** (Orbis SMR, Hexagone SMR).
- **Conversion automatique** des grilles hebdomadaires en dates exactes de venues.
- Rapprochement par clé composite **NDA + Date**.
- Rapports Excel identifiant clairement l'origine de chaque anomalie (manquant dans Orbis, dans Hexagone, ou incohérence).
- **Validé et déployé en production**.

## Fonctionnalités

- **Conversion grille → dates** : à partir du numéro de semaine et de la grille de présence, génère la liste exacte des dates de venue.
- **Nettoyage automatique** : filtrage des lignes de type "SD" (fin de séjour) dans Hexagone qui polluent la comparaison brute.
- **Rapprochement pointu** : compare chaque venue individuelle via la clé composite NDA + date.
- **Rapports Excel formatés** avec synthèse chiffrée (nombre de venues manquantes par système, pourcentages).
- **Interface graphique** pour sélection des fichiers et paramètres.

## Stack technique

- Python 3.9+
- Pandas
- openpyxl (rapports Excel formatés)
- Tkinter (interface graphique)
- PyInstaller (distribution `.exe`)

## Confidentialité et publication

- **Publication autorisée** par le maître de stage.
- Code **entièrement générique**, sans référence à un établissement.
- Les fichiers de test inclus (`Orbis_SMR_Test.xlsx`, `Hexagone_SMR_Test.xlsx`) sont **entièrement synthétiques**.

## Structure

```
├── src/
│   ├── code/                # controle_smr.py, interface_smr.py
│   └── mode_logiciel/       # Mode d'emploi utilisateur
├── fichiers_tests/          # Jeux de données synthétiques
├── docs/                    # Guide développeur
├── requirements.txt
└── .gitignore
```

## Installation

```bash
python -m venv venv
venv\Scripts\Activate.ps1     # Windows
# ou : source venv/bin/activate  # macOS/Linux
pip install -r requirements.txt
```

## Utilisation

### Interface graphique

```bash
python src/code/interface_smr.py
```

Choisir les exports Orbis et Hexagone, indiquer le dossier d'export, cliquer sur **Lancer le traitement**.

### Génération d'un exécutable autonome

```bash
pip install pyinstaller
cd src/code
pyinstaller --noconsole --onefile --name "Controle_SMR_DIM" interface_smr.py
```

Le `.exe` est dans `dist/`.

## Licence

MIT — voir [LICENSE](LICENSE).
