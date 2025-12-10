<strong>Auteur</strong>  
Jean-Matthieu Charre  
Développeur VBA senior  
CACIB Direction Financière - DFI / GTVA  
Année 2024  
________________________________________
<strong>Licence</strong>  
Projet interne CACIB Fast-IT / DFI - Reproduction interdite  
Le code présenté sur GitHub a uniquement pour objectif de mettre en valeur mes compétences techniques  
________________________________________
<strong>Notes</strong>  
Ce développement illustre ma capacité à :  
•	concevoir des automatisations Excel robustes et compatibles RPA  
•	intégrer des logs détaillés, des KPIs de fin de traitement, et des gestions d’erreurs structurées et différenciées  
•	produire un code fiable, maintenable, et conforme aux standards industriels  
•	intégrer mon code dans un framework de développement entreprise existant  
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
- 	Gestion différenciée des erreurs en fonction du mode de lancement
<br>
<br>
<strong>*** Fichiers utilisés</strong>
<br>
- 	Classeur Excel "361 - v1.2.2.xlsm" : application contenant le programme VBA<br>
- 	Classeur "Masterfile - IG v10.8.xlsx" : fichier en input<br>
<br>
<strong>*** Modes de lancement et spécificités</strong>
<br>
1.	Mode RPA (automatique)<br>
- 	Lancement via "cmd.bat"<br>
-   Gestion des erreurs silencieuse et fermeture propre de l'application en fin de programme (notamment en cas de bug car le programme tourne sur une VDI et aucune personne physique ne peut interagir avec l'application)<br>
2.	Mode manuel<br>
- 	Lancement par clic sur bouton Excel<br>
- 	Si erreur dans le traitement, une MsgBox apparaît et informe l'utilisateur sur le type d'erreur rencontré et sur la démarche à suivre<br>
- 	En fin de programme, une MsgBox informe l'utilisateur de la fin du traitement<br>
<br>
<br>
<strong>A)	Contexte et objectif</strong>
<br>
<br>
Ce développement VBA/Excel vise à automatiser la génération de fichiers d’écarts intragroupes (environ 280 classeurs en output) pour le département DFI / GTVA à partir de données issues du process GTVA.  
Le traitement, historiquement manuel et chronophage, a été entièrement automatisé pour être exécuté en autonomie par un robot RPA sur une VDI.
<br>
<br>
<br>
<strong>B)	Architecture du process</strong>
<br>
<br>
Le programme VBA s’appuie sur une architecture modulaire segmentée, organisée autour d’un point d’entrée unique (procédure "Main" dans le module "PROCESS_MAIN" du Classeur "361 - v1.2.2 - 2024-11-20.xlsm") qui pilote l’ensemble du workflow et coordonne les différentes étapes du traitement. Le code s’intègre dans un framework VBA interne fournissant des fonctionnalités transverses : cadre de gestion centralisée des erreurs (ErrorManagment), journalisation structurée (ALGOLOG), alimentation et diffusion des indicateurs de suivi (KPIs), ainsi qu’un ensemble d’outils permettant d’encapsuler les opérations de manipulation de fichiers, de dossiers, et d’objets Excel. Chaque responsabilité fonctionnelle est définie dans un module dédié : initialisation du contexte, des KPIs, et des variables globales (InitialisationGlobales), création et remise à zéro du fichier de rapport (CreateRapport), vérification de la complétude et de la cohérence des sources (VérificationsPréalables), génération du répertoire d’exécution journalier (CréationDossierJour), configuration des entités à traiter (ParamétrageEntités), puis exécution du traitement métier pilotée par une boucle intégrant les rafraîchissements PowerQuery avec une gestion asynchrone des requêtes permettant de synchroniser l’avancement du workflow VBA avec la finalisation effective des chargements de données, nettoyages intermédiaires et exports finalisés (Export). Enfin, les modules de clôture (End_Clean et End_Clean_OnError) garantissent une terminaison propre du processus, la mise à jour des KPIs ainsi que la production du statut final (OK ou KO). Cette organisation modulaire assure une séparation nette des responsabilités, une meilleure maintenabilité, et une fiabilité conforme aux exigences d’un environnement VBA professionnel.
<br>
<br>
<br>
<strong>C)	Worflow d'exécution du programme</strong>
<br>
<br>
<strong>&nbsp;&nbsp;1. Initialisation et préparation du contexte</strong>
<br>
<br>
À l’exécution, le processus est lancé par la procédure "Main" située dans le module "PROCESS_MAIN", qui initialise le contexte applicatif via l'appel de la procédure "Init", active le mode de gestion des erreurs centralisé et journalise l’amorçage du workflow dans le système de logging interne (ALGOLOG). Cette phase prépare les variables globales, configure le mode automatique éventuel (CMD/RPA) et établit la séquence d’appel des modules métier.

Le module "InitialisationGlobales" est ensuite appelé : il récupère l’ensemble des paramètres dynamiques nécessaires au traitement (chemins des fichiers sources, onglets requis, tableaux structurés obligatoires, plages nommées, répertoires d’entrée et de sortie, métadonnées KPI, etc.). Cette étape construit le contexte d'exécution du programme et initialise les compteurs opérationnels ainsi que la configuration KPI via "KPI_CONFIG".

<strong>&nbsp;&nbsp;2. Création du rapport et vérification de l’environnement</strong>

Le module "CreateRapport" supprime puis recrée le fichier "Rapport.txt", garantissant un espace de log propre pour la session d’exécution courante.
Le module "VérificationsPréalables" réalise ensuite un pipeline complet de validation de l’environnement. Il contrôle :
l’existence des fichiers essentiels ("Masterfile.xlsm", "GO.txt", "Rapport.txt")
la présence des onglets requis
la disponibilité des tableaux structurés attendus
la cohérence des plages nommées
la validité des répertoires spécifiés dans les paramètres
la présence des fichiers obligatoires dans chaque dossier source

L’ensemble repose sur une série de sous-modules spécialisés ("VérifFichiers", "VérifOnglets", "VérifTableauxStructurés", "VérifPlagesNommées", "VérifExistenceFichiers", "VérifExistenceRépertoires") articulés de façon à garantir un enchaînement fiable et logique des opérations de vérification. En cas d’erreur, les anomalies sont consolidées dans le rapport et entraînent une interruption contrôlée du processus.

<strong>&nbsp;&nbsp;3. Génération du répertoire journalier</strong>

Une fois l’environnement validé, le module "CréationDossierJour" génère le répertoire d’exécution du jour à partir d’un chemin modèle contenant des jetons dynamiques (AA/MM/JJ). Le système substitue ces jetons par la date courante, normalise le chemin final, puis crée le dossier s’il n’existe pas. Ce répertoire deviendra l’emplacement de sortie de l’ensemble des fichiers générés.

<strong>&nbsp;&nbsp;4. Préparation des entités à traiter</strong>

Le module "ParamétrageEntités" prépare le périmètre de traitement en marquant par un « X » l’ensemble des lignes du tableau structuré "Déclarants_IG".
L’opération n’altère pas la structure du tableau, mais prépare une liste de travail parfaitement déterministe pour le module d’export.

<strong>&nbsp;&nbsp;5. Phase d’export métier (boucle principale)</strong>

Le module "Export" constitue le cœur opérationnel du processus. Il commence par :
déterminer le nombre d’entités à traiter
alimenter les KPIs correspondants
masquer les feuilles non essentielles pour sécuriser l’environnement d’exécution
Pour chaque entité marquée :
la feuille "Entité" est renseignée avec les paramètres correspondants
les 6 connexions PowerQuery critiques ("Membre", "Imports", "Final", "Réponses", "Rejets", "Clé_de_lettrage") sont rafraîchies séquentiellement avec un processus asynchrone
les données intermédiaires du tableau "Imports" sont supprimées
les colonnes calculées problématiques ("Assistant_Lettrage" et "Statut_Final") sont reconstruites pour garantir la cohérence métier
l’ensemble des caches pivots du classeur est régénéré
le fichier final est produit dans le répertoire journalier en incluant le nom de l’entité dans son intitulé
Chaque rafraîchissement PowerQuery est chronométré et sécurisé : en cas d’erreur sur une connexion, un module dédié ("End_Clean_OnError_Connection") interrompt immédiatement le processus et journalise l’anomalie.

<strong>&nbsp;&nbsp;6. Clôture contrôlée (End_Clean)</strong>

À l’issue de la boucle :
"End_Clean" consolide et transmet les KPIs
sauvegarde le classeur maître
met à jour "Rapport.txt"
écrit le statut final "OK" dans le fichier d'état "GO.txt" faisant le lien entre le process VBA et l'automatisation RPA
journalise la durée totale du traitement
ferme proprement l’application Excel
Cette phase garantit une termination propre de l’ensemble du processus.

<strong>&nbsp;&nbsp;7. Gestion d’erreurs et arrêt sécurisé</strong>

En cas d’exception (anomalie métier, erreur PowerQuery, chemin manquant, structure non conforme…), les modules :
"End_Clean_OnError"
"End_Clean_OnError_Connection"
prennent automatiquement le relais. Ils assurent :
la mise à jour du statut final en "KO"
la journalisation complète de l’erreur
la fermeture sécurisée des fichiers
la préservation de l’intégrité du classeur et des sources
un fail-safe shutdown conforme aux standards de production VBA/BFI

<strong>&nbsp;&nbsp;8. Structure des KPIs</strong>

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
