Auteur  
Prenom Nom : Jean-Matthieu Charre  
Rôle : Développeur VBA senior  
Contexte : CACIB Direction Financière - DFI / GTVA  
Année : 2024  
________________________________________
Licence  
Projet interne CACIB Fast-IT / DFI - Reproduction interdite.  
Le code présenté sur GitHub a uniquement pour objectif de démontrer mes compétences techniques.  
________________________________________
Notes  
Ce développement illustre ma capacité à :  
•	concevoir des automatisations Excel robustes et compatibles RPA  
•	intégrer PowerQuery, des logs, des KPIs, et des gestions d’erreurs structurées  
•	produire un code fiable, maintenable, conforme aux standards industriels, et s’insérant dans un framework de développement entreprise  
•	travailler en collaboration directe avec des équipes de Business Analysts  
<br>
<br>
<br>
Développement VBA  
Génération de l’État des Écarts Intragroupes  
<br>
<br>
<br>
** Technologies et normes utilisées  
  
-	Excel VBA (compatible Office 32 bits et 64 bits)  
-	Requêtes PowerQuery
-	RPA integration via CMD + fichiers d’état 
-	Logging textuel en temps réel 
-	export de données en JSON (KPI)
-	Gestion des erreurs différenciée en fonction du mode de lancement (RPA ou manuel)
  
** Fichiers utilisés  
  
-	« Classeur « 361 - v1.2.2.xlsm » : Classeur contenant le programme VBA
-	« Masterfile - IG v10.8.xlsx » : fichier source
  
** Modes de lancement  
  
1.	Mode RPA (automatique)
-	Lancement via `cmd.bat`
-	Aucun message à l’écran
-	Fin silencieuse, fermeture automatique de l’application et des fichiers sources
-	Logs + KPI + GO.txt + Rapport.txt générés

2.	Mode manuel
-	Lancement par clic sur bouton Excel
-	Si erreur dans le traitement, MsgBox affichant le type d’erreur
-	MsgBox de fin de traitement affichée
<br>
<br>
A)	Contexte et objectif
<br>
Ce développement VBA/Excel vise à automatiser la génération de fichiers d’écarts intragroupes (environ 280 fichiers en output) pour le département DFI / GTVA à partir de données issues du process GTVA.  
Le traitement, historiquement manuel et chronophage, a été entièrement automatisé pour être exécuté en autonomie par un robot RPA (sans intervention humaine).
<br>
<br>
B)	Architecture du process
<br>
1.	Le robot RPA crée un fichier vide ‘GO.txt’ et lance un script ’cmd.bat’,
2.	Le script démarre Excel + VBA,
3.	Le programme VBA exécute le process principal :
3.1.	Vérifie les chemins et sources nécessaires,
3.2.	Crée le dossier de travail du jour,
3.3.	Sélectionne les entités listées dans l’onglet « Déclarants » du Classeur « Masterfile - IG v10.8 - ORIGINAL.xlsm » (classeur en input du programme),
3.4.	Lance la procédure « Sub_Main » localisée dans le Module « PROCESS_MAIN » du Classeur « 361 - v1.2.2.xlsm » (classeur contenant le programme principal),
3.5.	Crée le rapport d’exécution (‘Rapport.txt’),
3.6.	Renseigne le fichier ‘GO.txt’  avec le statut final (‘OK’ ou ‘KO’, signifiant le bon déroulement ou pas jusqu’à la fin du programme).
3.7.	Le robot lit le contenu du fichier ‘GO.txt’ envoie un rapport par e-mail et clôture le traitement.
<br>
<br>
C)	Fonctionnalités principales
<br>
-	Exécution autonome : Lancement par `cmd.bat` sans message ni interaction utilisateur,<br>
-	Compatibilité RPA : Gestion des erreurs, logs, KPI et fichiers d’état normalisés,<br>
-	Vérification préliminaire : Contrôle de la présence des onglets, plages nommées, dossiers et fichiers sources,<br>
-	Création du dossier du jour : Génération automatique d’un répertoire daté (`AA.MM.JJ`) selon le modèle paramétré dans Central/Prm_ModeleDestination,<br>
-	Paramétrage automatique des entités : Renseignement automatique d’un “X” dans la colonne A du tableau `Déclarants_IG`,<br>
-	Lancement du process métier : Exécution de la procédure `Export`, responsable de la création des fichiers par entité,<br>
-	Logs détaillés : Journalisation temps réel des actions dans `.\Log\YYYYMMDD_HHMM.txt`,<br>
-	Rapport d’exécution : Génération de `Rapport.txt` résumant le résultat global du traitement,<br>
-	Fichier d’état GO.txt : Statut `OK` ou `KO` en fin d’exécution, lu par le robot pour poursuivre ou interrompre le flux,<br>
-	KPI Fast-IT : Génération d’un fichier JSON contenant les métriques du traitement (durée, statut, entités, etc.).<br>
<br>
<br>
D)	Détails du fonctionnement
<br>
1)	Vérification des sources
Le programme contrôle la présence :
- des onglets *Déclarants* et *Central* ;
- des plages nommées : `Déclarants_IG`, `Prm_Tables`, `Prm_Temp2`, `Prm_Temp3`, `Prm_Destination`, `Prm_ModeleDestination` ;
- des fichiers attendus dans les répertoires (`Index.xlsx`, `Périmètre.xlsx`, `Plans.xlsx`, etc.).
Toute anomalie est reportée dans `Rapport.txt` et mène à un `KO`.
<br>
2)	Création du dossier du jour
Le chemin indiqué dans `Prm_ModeleDestination` peut contenir des variables :
- `AA` → année sur 2 chiffres  
- `MM` → mois  
- `JJ` → jour  
Exemple :  
`N:\Projets01\ROBOTISATION_DFI\361\1. Production\2. Etat des écarts\AA.MM.JJ`  
→ devient `N:\Projets01\ROBOTISATION_DFI\361\1. Production\2. Etat des écarts\25.11.13`
<br>
3)	Paramétrage des entités
Le tableau « Déclarants_IG » est parcouru et toutes les lignes de la colonne A sont renseignées avec un “X”.
<br>
4)	Exécution du process
La procédure « Export » est appelée :
-	actualisation PowerQuery ;
-	génération d’un fichier par entité ;
-	suivi des erreurs métier ;
-	ajout de logs spécifiques.
<br>
5)	Reporting & fin de process
- Si le programme s’est déroulé correctement :
-	« Rapport.txt » : « Traitement terminé sans anomalie » et envoi des KPIs par email.
-	« GO.txt » : ‘OK’
- En cas d’erreur :
-	« Rapport.txt » : message d’erreur explicite et envoi des KPIs par email.
-	« GO.txt » → ‘KO’
<br>
6)	KPI - Suivi de performance
Un fichier JSON est généré à la fin du traitement avec des indicateurs clefs sur le process réalisé :

| Clé | Exemple | Description |
|------|----------|-------------|
| `Code process` | `361` | Identifiant principal |
| `Sous code process` | `361-2` | Numéro de lot |
| `Nom du process` | `GTVA_Generation_Ecarts_IG` | Nom technique sans accent |
| `Direction` | `DFI` | Direction métier |
| `Département` | `GTVA` | Département |
| `Jour/homme passé` | `1.32` | Calculé : 0,002 par entité traitée |
| `Technologie` | `VBA` | En dur |
| `Statut` | `OK` / `KO` | Statut global |
| `Date/heure début` | `2025-11-13T21:00:00.000Z` | Timestamp ISO |
| `Date/heure fin` | `2025-11-13T23:30:00.000Z` | Timestamp ISO |
| `Nb occurrences lues` | `278` | Entités totales |
| `Nb occurrences traitées` | `278` | Entités réussies |
| `Nb occurrences rejetées` | `0` | Différence |
| `Nb actions` | `1500` | Nombre d’actions automatisées |
| `Environnement` | `Production` | Test / Production |
