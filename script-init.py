import os
import zipfile
import shutil
import pandas as pd

MODELE_ODT = "certificat.administratif.odt"
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

        # Décompression modèle ODT
        dossier_temp = os.path.join(dossier_client, "temp_odt")
        os.makedirs(dossier_temp, exist_ok=True)

        with zipfile.ZipFile(MODELE_ODT, 'r') as zip_ref:
            zip_ref.extractall(dossier_temp)

        # Modifier le content.xml
        content_path = os.path.join(dossier_temp, "content.xml")
        with open(content_path, "r", encoding="utf-8") as f:
            contenu = f.read()

        contenu = contenu.replace("……", redevable_courant, 1)
        contenu = contenu.replace("……", f"{montant:.2f} € TTC", 1)

        with open(content_path, "w", encoding="utf-8") as f:
            f.write(contenu)

        # Recréation du .odt personnalisé
        certificat_nom = f"certificat_{redevable_courant.replace(' ', '_')}.odt"
        certificat_path = os.path.join(dossier_client, certificat_nom)

        with zipfile.ZipFile(certificat_path, 'w', zipfile.ZIP_DEFLATED) as doc_odt:
            for dossier_racine, _, fichiers in os.walk(dossier_temp):
                for fichier in fichiers:
                    chemin_complet = os.path.join(dossier_racine, fichier)
                    chemin_relatif = os.path.relpath(chemin_complet, dossier_temp)
                    doc_odt.write(chemin_complet, arcname=chemin_relatif)

        # Nettoyage du dossier temporaire
        shutil.rmtree(dossier_temp)

        print(f"✅ Certificat généré pour {redevable_courant} avec montant {montant:.2f} €")

        # On réinitialise le redevable courant pour éviter de générer plusieurs fois
        redevable_courant = None

print("✅ Certificats générés dans le dossier 'régularisation'.")

