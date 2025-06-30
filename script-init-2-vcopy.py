import os
import shutil
import pandas as pd

MODELE_DOC = "CERTIFICAT ADMINISTRATIF.doc"
FICHIER_EXCEL = "factures_filtrees_et_regularisations_sans_negatifs.xlsx"
DOSSIER_SORTIE = "régularisation"

# Lecture du fichier Excel
df = pd.read_excel(FICHIER_EXCEL)
df.columns = df.columns.str.strip()

os.makedirs(DOSSIER_SORTIE, exist_ok=True)

redevable_courant = None

for index, ligne in df.iterrows():
    redevable = ligne["Redevable"]
    montant = ligne["avoir final TTC"]

    # Si on a un redevable non vide (non NaN)
    if pd.notna(redevable):
        redevable_courant = str(redevable).strip()
        # On ne génère pas de certificat ici, on attend la ligne suivante
        continue

    # Ici, redevable est NaN, on est potentiellement sur la ligne avec montant négatif

    # On vérifie qu'on a un redevable courant valide et un montant négatif
    if redevable_courant is not None and pd.notna(montant) and montant < 0:
        # Ignorer si le redevable commence par un code postal (5 chiffres)
        if redevable_courant[:5].isdigit():
            print(f"⏭️  Ignoré : {redevable_courant} (semble être un code postal)")
            continue

        # Création dossier client
        dossier_client = os.path.join(DOSSIER_SORTIE, redevable_courant.replace("/", "_"))
        os.makedirs(dossier_client, exist_ok=True)

        # ❌ Partie modifiée : au lieu de manipuler un .odt, on copie un .doc global
        certificat_nom = f"CERTIFICAT ADMINISTRATIF - {redevable_courant.replace(' ', '_')}.doc"
        certificat_path = os.path.join(dossier_client, certificat_nom)
        shutil.copyfile(MODELE_DOC, certificat_path)

        print(f"✅ Certificat copié pour {redevable_courant} avec montant {montant:.2f} €")

        # On réinitialise le redevable courant pour éviter de générer plusieurs fois
        redevable_courant = None

print("✅ Certificats copiés dans le dossier 'régularisation'.")