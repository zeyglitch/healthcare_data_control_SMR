import pandas as pd
import random
from datetime import datetime, timedelta
import os

def generate_data():
    num_patients = 30
    
    orbis_data = []
    hexagone_data = []

    first_names = ["CHARLES", "LUCIEN", "MARIE", "JEAN", "PIERRE", "PAUL", "JACQUES", "ANNE", "SOPHIE", "JULIE", "TIMEO", "DUNE"]
    last_names = ["DUPUIT", "BERNARD", "MARTIN", "DURAND", "DUBOIS", "MOREAU", "LAURENT", "SIMON", "MICHEL", "GARCIA", "CHARLES", "DUNE"]
    
    for i in range(num_patients):
        nda = f"{random.randint(400000000, 499999999)}"
        first_name = random.choice(first_names)
        last_name = random.choice(last_names)
        sexe = random.choice(["M", "F"])
        dob = (datetime(1950, 1, 1) + timedelta(days=random.randint(0, 20000))).strftime("%d/%m/%Y")
        
        # Le patient peut venir de 2 à 10 fois pour sa rééducation
        num_visits = random.randint(2, 10)
        
        # Le séjour se termine-t-il par un SD ? (50% de chances)
        has_sd = random.random() < 0.5
        
        # Date de début aléatoire
        start_date = datetime(2026, 1, 1) + timedelta(days=random.randint(0, 300))
        visits = []
        current_date = start_date
        
        for v in range(num_visits):
            if v > 0:
                if v == num_visits - 1 and has_sd:
                    # La dernière visite (SD) doit être le même jour que la précédente, décalée de qq heures
                    current_date += timedelta(hours=random.randint(2, 4))
                else:
                    # Les autres visites s'enchaînent avec potentiellement des jours d'écart
                    days_advance = random.randint(0, 3)
                    if days_advance == 0:
                        current_date += timedelta(hours=random.randint(1, 5))
                    else:
                        current_date += timedelta(days=days_advance, hours=random.randint(-5, 5))
            visits.append(current_date)
            
        # Groupement des visites par semaine pour Orbis
        # (année, semaine_iso) -> liste des dates
        weeks_map = {}
        for d in visits:
            y, w, _ = d.isocalendar()
            if (y, w) not in weeks_map:
                weeks_map[(y, w)] = []
            weeks_map[(y, w)].append(d)
            
        # Probabilités d'erreurs pour tester le tri plus tard
        has_nda_error = random.random() < 0.1
        has_date_error = random.random() < 0.1

        orbis_nda = nda
        hex_nda = nda
        if has_nda_error:
            # Hexagone a une faute de frappe sur le NDA par rapport à Orbis
            hex_nda = f"{int(nda) + 1}"

        # --- Génération des lignes ORBIS ---
        for (y, w), dates_in_week in weeks_map.items():
            # Construction de la chaîne de présence (ex: L.MM.V..)
            presence_array = ['.', '.', '.', '.', '.', '.', '.']
            presence_letters = ['L', 'M', 'M', 'J', 'V', 'S', 'D']
            
            for d in dates_in_week:
                idx = d.isocalendar()[2] - 1  # Lundi=1 -> 0, Dimanche=7 -> 6
                presence_array[idx] = presence_letters[idx]
            
            presence_str = "".join(presence_array)
            
            # Format aléatoire du numéro de semaine dans Orbis
            week_str = f"{y}{w:02d}" if random.random() > 0.5 else f"{w}"
            
            # Introduction d'une erreur de semaine dans Orbis
            if has_date_error:
                w_error = w + 1 if w < 52 else 1
                week_str = f"{y}{w_error:02d}"
                
            orbis_data.append({
                "N° Hospit": orbis_nda,
                "N° semaine": week_str,
                "Présence": presence_str,
                "Nom": last_name,
                "Prénom": first_name,
                "Né(e) le": dob
            })
            
        # --- Génération des lignes HEXAGONE ---
        for v_idx, d in enumerate(visits):
            if v_idx == 0:
                v_type = "ED"
                mod_nat = "DOM"
                heber = "901"
                bat = "Bâtiment A"
                chamb = "101"
            elif v_idx == len(visits) - 1 and has_sd:
                v_type = "SD"
                mod_nat = "RDOM"
                heber = ""
                bat = ""
                chamb = ""
            else:
                v_type = "V"
                mod_nat = ""
                heber = "901"
                bat = "Bâtiment A"
                chamb = "101"
                
            hexagone_data.append({
                "Nom/Prénom": f"{last_name}/{first_name}",
                "Nom de Naissance": last_name,
                "Sexe": sexe,
                "Date de naissance": dob,
                "N° Dossier": hex_nda,
                "Date": d.strftime("%d/%m/%Y %H:%M"),
                "Mod/Nat": mod_nat,
                "Commentaire": "pas de réaction",
                "Type": v_type,
                "UF": "901",
                "Héber": heber,
                "Bâtiment": bat,
                "Chambre/Lit": chamb
            })

    df_orbis = pd.DataFrame(orbis_data)
    df_hexagone = pd.DataFrame(hexagone_data)

    output_dir = os.path.dirname(os.path.abspath(__file__))

    # Sauvegarde Orbis
    orbis_path = os.path.join(output_dir, "Orbis_SMR_Test.xlsx")
    df_orbis.to_excel(orbis_path, index=False)

    # Sauvegarde Hexagone avec format spécifique
    hexa_path = os.path.join(output_dir, "Hexagone_SMR_Test.xlsx")
    with pd.ExcelWriter(hexa_path, engine='openpyxl') as writer:
        df_hexagone.to_excel(writer, index=False, startrow=2)
        workbook = writer.book
        worksheet = writer.sheets['Sheet1']
        worksheet['A1'] = "Liste des mouvements par séjour"
        
    print(f"Fichiers de test SMR recréés avec succès dans : {output_dir}")
    print(f"- {os.path.basename(orbis_path)} (Lignes: {len(df_orbis)})")
    print(f"- {os.path.basename(hexa_path)} (Lignes: {len(df_hexagone)})")

if __name__ == "__main__":
    generate_data()
