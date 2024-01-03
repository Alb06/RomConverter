# RomConverter
 convert your ROMs in bulk easily

## Fonctionnalités

Le script PowerShell effectue les tâches suivantes :

1. **Gestion des Logs** : Crée et gère un fichier de log pour enregistrer les activités du script.
2. **Vérification de 7-Zip** : Vérifie si 7-Zip est installé sur la machine et guide pour l'installation si nécessaire.
3. **Vérification de chdman.exe** : Contrôle la présence de `chdman.exe` nécessaire pour la conversion des fichiers.
4. **Sélection des Dossiers de Travail** : Permet à l'utilisateur de sélectionner différents dossiers pour les fichiers source, décompression, stockage .chd et recompression.
5. **Chargement et Sauvegarde des Paramètres** : Charge et enregistre les paramètres de l'utilisateur pour une utilisation future.
6. **Décompression des Fichiers avec 7-Zip** : Décompresse les fichiers `.7z` dans un dossier spécifié.
7. **Conversion et Traitement des Fichiers** : Convertit les fichiers décompressés en format `.chd`.
8. **Interface Utilisateur Graphique** : Fournit une interface graphique avec une barre de progression et un bouton de pause/continuation.

## Utilisation

Comment utiliser le script, par exemple :

1. Clonez le dépôt dans votre système local.
2. Ouvrez le terminal ou PowerShell et naviguez vers le répertoire du script.
3. Exécutez le script avec la commande : `.\NomDuScript.ps1`
4. Suivez les instructions à l'écran pour sélectionner les dossiers et effectuer les opérations.

## Prérequis

- Windows PowerShell 5.1 ou plus récent.
- 7-Zip installé sur la machine.
- `chdman.exe` présent dans le même répertoire que le script.

## Licence

Ce projet est sous licence. Voir le fichier `LICENSE` pour plus de détails.
