# Pose CO V3 – base Windows

Cette première base technique reprend le fichier HTML source et ajoute une enveloppe Electron destinée à produire une application Windows portable.

## Stockage

Les fichiers sont créés à côté de `Pose CO V3.exe`, dans `Donnees` :
- `PoseCO_Enregistrements.xlsx`
- `PoseCO_Magasin.xlsx`
- `PoseCO_Liste_Abonnes.xlsx`
- `PoseCO_Parametres.xlsx`

Les données métier restent dans des classeurs Excel. Une feuille conserve les enregistrements JSON ligne par ligne afin de préserver les champs actuels sans les renommer ni perdre les données imbriquées. Une copie `.bak` est créée avant remplacement du classeur.

## Attention – état de cette base

Ce projet est un premier portage technique, pas encore une version validée pour exploitation. Il faut notamment tester les feuilles Excel, l'import/export de la liste complète, les scans, l'impression A5 et les opérations magasin. Le stockage de la liste complète est actuellement placé dans une cellule Excel JSON, ce qui doit être optimisé en feuille tabulaire avant validation sur des listes de 30 000 à 50 000 abonnés. Les fonctions de scan reposent encore sur la disponibilité des fichiers locaux des bibliothèques copiés pendant `npm install`.

## Compilation

Sur GitHub Actions, lancer le workflow `Build Pose CO V3 Windows 7 x64`. L'EXE produit est un build portable x64. Electron 22 est choisi pour la compatibilité historique avec Windows 7 SP1 ; cette plateforme n'est plus prise en charge par Microsoft et ce runtime est ancien.
