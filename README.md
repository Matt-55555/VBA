<strong>Auteur</strong>  
Jean-Matthieu Charre  
Développeur VBA senior  
CACIB Direction Financière - DFI / GTVA  
Année 2024  
________________________________________
<strong>Licence</strong>  
Projet interne CACIB Fast-IT / DFI - Reproduction interdite.  
Le code présenté sur GitHub a uniquement pour objectif de présenter mes compétences techniques.  
________________________________________
<strong>Notes</strong>  
Ce développement illustre ma capacité à :  
•	concevoir des automatisations Excel robustes et compatibles RPA  
•	intégrer PowerQuery, des logs, des KPIs, et des gestions d’erreurs structurées  
•	produire un code fiable, maintenable, et conforme aux standards industriels  
•	utiliser un framework de développement entreprise existant  
•	travailler en collaboration directe avec des équipes de Business Analysts  
<br>
<br>
<br>
<h1>Développement VBA<br>
Génération de l’État des Écarts Intragroupes</h1>
<br>
<strong>*** Technologies et normes utilisées</strong>  
<br>
- 	Excel VBA (compatible Office 32 bits et 64 bits)<br>
- 	Requêtes PowerQuery<br>
- 	Intégration RPA via CMD + fichiers d’état<br>
- 	Logging textuel en temps réel<br>
- 	export de données en JSON (KPI)<br>
- 	Gestion différenciée des erreurs en fonction du mode de lancement (RPA ou manuel)</strong>
<br>
<br>
<strong>*** Fichiers utilisés</strong>
<br>
- 	Classeur Excel "361 - v1.2.2.xlsm" : classeur contenant le programme VBA<br>
- 	Classeur "Masterfile - IG v10.8.xlsx" : fichier source<br>
<br>
<br>
<strong>*** Modes de lancement</strong>
<br>
1.	Mode RPA (automatique)<br>
- 	Lancement via "cmd.bat"<br>
- 	Aucun message à l’écran<br>
- 	Fin silencieuse : fermeture propre et automatique de l’application et du fichier source (même en cas de bugs - car application tourne sur une VDI)<br>
- 	Logs + KPI + GO.txt + Rapport.txt générés<br>
2.	Mode manuel<br>
- 	Lancement par clic sur bouton Excel<br>
- 	Si erreur dans le traitement, MsgBox affichant le type d’erreur<br>
- 	En fin de programme, MsgBox de fin de traitement (l'application reste ouverte ainsi que le fichier source)<br>
<br>
<br>
<strong>A)	Contexte et objectif</strong>
<br>
<br>
Ce développement VBA/Excel vise à automatiser la génération de fichiers d’écarts intragroupes (environ 280 classeurs Excel à générer) pour le département DFI / GTVA à partir de données issues du process GTVA.  
Le traitement, historiquement manuel et chronophage, a été entièrement automatisé pour être exécuté en autonomie par un robot RPA (sans intervention humaine).
<br>
<br>
<br>
<strong>B)	Architecture du process</strong>
<br>
<br>
1.	Le robot RPA crée un fichier vide "GO.txt" et lance un script "cmd.bat"<br>
2.	Le script ouvre l'application Excel, et une procédure évènementielle dans l'application lance le programme VBA (procédure "Sub_Main" du Module "PROCESS_MAIN" du Classeur "361 - v1.2.2.xlsm")<br>
3.	La procédure "Sub_Main" appelle de manière séquentielle des procédures secondaires pour réaliser les traitements nécessaires sur les données :<br>
• Vérification de l'ensemble des paramètres necessaires au bon fonctionnement du programme (fichiers, dossiers, worksheets, tableaux structurés, colonnes de tableaux, variables, etc)<br>
•	Paramétrage des entités intragroupe à traiter<br>
•	Lancement des requêtes Power Query<br>
•	Création du rapport d’exécution ("Rapport.txt")<br>
•	Envoi des KPIs<br>
•	Renseignement du fichier "GO.txt" avec le statut final (‘OK’ ou ‘KO’ signifiant le bon déroulement ou pas jusqu’à la fin du programme)<br>
•	Le robot lit le contenu du fichier ‘GO.txt’ et envoie un rapport par e-mail et clôture le traitement.<br>
<br>
<br>
<br>
<strong>C)	Détails du fonctionnement</strong>
<br>
<br>
1)	Vérification des sources<br>
Le programme contrôle la présence :<br>
- des onglets `Déclarants` et `Central`,<br>
- des plages nommées : `Déclarants_IG`, `Prm_Tables`, `Prm_Temp2`, `Prm_Temp3`, `Prm_Destination`, `Prm_ModeleDestination`,<br>
- des fichiers attendus dans les répertoires (`Index.xlsx`, `Périmètre.xlsx`, `Plans.xlsx`, etc.).<br>
Toute anomalie est reportée dans `Rapport.txt` et mène à un `KO`.<br>
<br>
<br>
2)	Création du dossier du jour
<br>
Le chemin indiqué dans `Prm_ModeleDestination` peut contenir des variables :
- `AA` → année sur 2 chiffres  
- `MM` → mois  
- `JJ` → jour  
Exemple :  
`N:\Projets01\ROBOTISATION_DFI\361\1. Production\2. Etat des écarts\AA.MM.JJ`  
→ devient `N:\Projets01\ROBOTISATION_DFI\361\1. Production\2. Etat des écarts\25.11.13`
<br>
<br>
3)	Paramétrage des entités
<br>
Le tableau « Déclarants_IG » est parcouru et toutes les lignes de la colonne A sont renseignées avec un “X”.
<br>
<br>
4)	Exécution du process
<br>
La procédure « Export » est appelée :
-	actualisation PowerQuery ;
-	génération d’un fichier par entité ;
-	suivi des erreurs métier ;
-	ajout de logs spécifiques.
<br>
<br>
5)	Reporting & fin de process
<br>
- Si le programme s’est déroulé correctement :
-	« Rapport.txt » : « Traitement terminé sans anomalie » et envoi des KPIs par email.
-	« GO.txt » : ‘OK’
- En cas d’erreur :
-	« Rapport.txt » : message d’erreur explicite et envoi des KPIs par email.
-	« GO.txt » → ‘KO’
<br>
<br>
6)	KPI - Suivi de performance
<br>
Un fichier JSON est généré à la fin du traitement avec des indicateurs clefs sur le process réalisé :
<br>
| Clé | Exemple | Description |<br>
|------|----------|-------------|<br>
| `Code process` | `361` | Identifiant principal |<br>
| `Sous code process` | `361-2` | Numéro de lot |<br>
| `Nom du process` | `GTVA_Generation_Ecarts_IG` | Nom technique sans accent |<br>
| `Direction` | `DFI` | Direction métier |<br>
| `Département` | `GTVA` | Département |<br>
| `Jour/homme passé` | `1.32` | Calculé : 0,002 par entité traitée |<br>
| `Technologie` | `VBA` | En dur |<br>
| `Statut` | `OK` / `KO` | Statut global |<br>
| `Date/heure début` | `2025-11-13T21:00:00.000Z` | Timestamp ISO |<br>
| `Date/heure fin` | `2025-11-13T23:30:00.000Z` | Timestamp ISO |<br>
| `Nb occurrences lues` | `278` | Entités totales |<br>
| `Nb occurrences traitées` | `278` | Entités réussies |<br>
| `Nb occurrences rejetées` | `0` | Différence |<br>
| `Nb actions` | `1500` | Nombre d’actions automatisées |<br>
| `Environnement` | `Production` | Test / Production |<br>
