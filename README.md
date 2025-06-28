# Traitement de factures et génération de certificats administratifs personnalisés

Ce projet Python permet de traiter un fichier Excel contenant des lignes de facturation (avec lignes de régularisation), et de générer automatiquement pour chaque redevable un certificat administratif personnalisé au format .odt, en créant pour chacun un sous-dossier dans un dossier régularisation.

## 🔧 Fonctionnalités principales

* Lecture d'un fichier Excel contenant les données de facturation.
* Traitement des régularisations :
* Chaque redevable apparaît sur 2 lignes :
  * 1ère ligne avec le nom du redevable, montant = 0
  * 2ème ligne avec montant négatif (régularisation)
* Pour chaque régularisation valide :
  * Création d'un dossier régularisation/<nom_redevable>
  * Décompression d'un modèle ODT (certificat.administratif.odt)
  * Remplacement dans le content.xml des champs personnalisables par :
    * Nom du redevable
    * Montant à rembourser (en TTC)
* Reconstitution du fichier .odt personnalisé pour le client

## ✅ Exemple de répertoire généré
```
régularisation/
├── M. DUPONT JEAN/
│   ├── certificat_M._DUPONT_JEAN.odt
├── Mme MARTIN CLAIRE/
│   ├── certificat_Mme_MARTIN_CLAIRE.odt
...
```
## ▶️ Exécution du script

Place dans le dossier :

factures_filtrees_et_regularisations_sans_negatifs.xlsx

certificat.administratif.odt

Lance le script avec :

python script-init.py

Le dossier régularisation sera créé automatiquement avec tous les sous-dossiers et certificats.

## 🧹 Dépendances
```bash
pip install pandas openpyxl
```
## 🔍 Points techniques clés

Le script mémorise le dernier redevable lu avec pandas.

Il ne génère un certificat que lorsque la ligne suivante a un montant négatif (ligne de régularisation).

Le modèle .odt est un fichier compressé : il est dézippé, modifié, puis recompressé.

Le champ …… est utilisé comme repère de remplacement dans le fichier XML du modèle.

## 📂 Fichier d'entrée Excel (extrait)

| N° facture | Redevable | AVOIR HT | avoir final HT | avoir final TTC |
| ---------- | --------- | -------- | -------------- | --------------- |
| F12345     | Dupont    | 0    | 0         | 0           |
|  |  | 14.15 | -3.20  | -3.38 |



Chaque redevable apparaît sur deux lignes. Le traitement est appliqué uniquement à la deuxième ligne avec montant < 0.

## 📚 Auteurs

Script développé par [Hamza Meneceur](https://github.com/HamzaMeneceur)

