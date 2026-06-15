# BOA : Better Office Accessibility (Meilleure Accessibilité Office)

BOA est une puissante suite d'améliorations d'accessibilité pour Microsoft Office, conçue pour améliorer considérablement l'expérience du lecteur d'écran pour les utilisateurs de NVDA. Elle corrige directement les composants d'interface utilisateur inaccessibles et introduit des outils de navigation rapide pour Excel et PowerPoint.

---

## ⌨️ Référence des raccourcis clavier

| Fonctionnalité | Combinaison de touches | Contexte / Notes |
| :--- | :--- | :--- |
| **Entrer en Mode de commande** | `NVDA+E` | Active le Mode préfixe de commande (déclenche un bip aigu) |
| **Analyser la disposition de la Feuille** | `NVDA+E`, puis `L` | À exécuter dans Excel avant de naviguer dans les Blocs de données |
| **Sauter au Bloc de données le plus proche** | `NVDA+E`, puis `J` /  | Nécessite d'abord une analyse de la disposition |
| **Ouvrir l'organisateur de Feuille en masse** | `NVDA+E`, puis `X` | Ouvre la boîte de dialogue accessible de réorganisation de Feuille |
| **Déplacer la Feuille active vers la gauche** | `NVDA+Shift+LeftArrow` | Décale la Feuille active d'une position vers le haut |
| **Déplacer la Feuille active vers la droite** | `NVDA+Shift+RightArrow` | Décale la Feuille de calcul active d'une position vers le bas |
| **Déplacer la Feuille au début/à la fin** | `NVDA+Shift+Home` / `End` | Envoie la Feuille de calcul aux limites absolues |
| **Mise en forme conditionnelle détaillée**| `NVDA+E`, puis `F` | Annonce les détails complets de la mise en forme de la Cellule avec le Focus |
| **Assigner la Cellule à un emplacement mémoire** | `NVDA+E`, puis `Shift+1` à `Shift+9` | Assigne la Cellule actuelle à un emplacement de surveillance en arrière-plan |
| **Lire l'emplacement de la Cellule surveillée** | `NVDA+E`, puis `1` à `9` | Rappelle et lit la valeur de l'emplacement assigné |
| **Basculer la surveillance en arrière-plan** | `NVDA+E`, puis `M` | Bascule manuellement le suivi des calculs en arrière-plan |
| **Effacer tous les emplacements mémoire** | `NVDA+E`, puis `Backspace` | Purge tous les moniteurs de Cellule enregistrés en arrière-plan |
| **Annuler le Mode de commande** | `Escape` | Quitte le Mode préfixe de commande |

---

## 🚀 Fonctionnalités

### Améliorations Excel

#### 1. Analyseur de disposition de Feuille et Mise en cache
Analysez instantanément n'importe quelle Feuille de calcul Excel pour comprendre sa structure, ses éléments masqués et ses Blocs de données.
* **Comment ça marche :** BOA analyse rapidement la Feuille et annonce les Blocs de données actifs. Il vous avertit également des **Onglets de Feuille masqués**, des **Filtres** actifs, des **Modes protégés** et des **Limites extérieures masquées** (par exemple, si les Colonnes près du bord droit de la Feuille sont masquées, vous évitant de manquer des données hors écran).
* **Navigation dans les données :** Après l'analyse, vous pouvez utiliser les raccourcis de saut de Bloc de données pour téléporter instantanément votre curseur entre les Blocs de données découverts, contournant ainsi sans effort des milliers de Cellules vides.

#### 2. Organisateur de Feuille en masse
Réorganisez et arrangez instantanément plusieurs Feuilles à la fois à l'aide d'une boîte de dialogue entièrement accessible.
* **Comment ça marche :** Ouvre une boîte de dialogue où vous pouvez sélectionner une Feuille et lui attribuer une nouvelle position. Les déplacements planifiés sont répertoriés dans un tableau de données (appuyez sur `Del` pour supprimer une erreur). Cliquez sur `OK` et votre Classeur est réorganisé instantanément.

#### 3. Déplacement rapide de Feuille
Déplacez la Feuille active vers la gauche, la droite, tout au début ou tout à la fin instantanément à l'aide de vos raccourcis clavier.

#### 4. Renommage accessible de Feuille
* Lors du renommage d'une Feuille, NVDA a nativement du mal à lire les caractères que vous tapez.
* BOA injecte une classe `ExcelSheetRenameEdit` personnalisée qui utilise le moteur `SafeRichEdit`, ce qui signifie que vous pouvez lire précisément par caractère, mot ou Ligne pendant le renommage. Cela sert d'amélioration au comportement de renommage par défaut existant.

#### 5. Suivi des Lignes/Colonnes masquées
* Suit de manière proactive vos mouvements sur la grille pour vous éviter de manquer des données masquées ou filtrées.
* **Cellules fragmentées traversées :** Si vous sautez à travers une section fortement fragmentée ou masquée de la grille (par exemple, en passant de la Ligne 3 à la Ligne 10 car les Lignes 4 à 9 sont masquées), BOA annonce explicitement "Lignes 4 à 9 masquées". Cela garantit que vous savez toujours quand des données ont été ignorées dans la structure.

#### 6. Annonceur de mise en forme conditionnelle
* Lit automatiquement la couleur, le style de police et la nuance d'arrière-plan des Cellules qui ont été modifiées dynamiquement par les règles de Mise en forme conditionnelle d'Excel.
* Vous donne le véritable état visuel de la Cellule plutôt que juste la valeur sous-jacente brute. Initialement, lors de la prise de Focus sur la Cellule, il annonce "a une mise en forme conditionnelle, et quelques autres détails mineurs". Pour des informations complètes, utilisez la configuration détaillée des raccourcis clavier qui est NVDA E puis F.

#### 7. Meilleure annonce de sélection
Lit si la Cellule ou la plage est sélectionnée ou désélectionnée.

#### 8 Moniteur de Cellule :
* **Moniteur de Cellule :** Utilisez les chemins de commande pour assigner des Cellules spécifiques à des emplacements de mémoire. Vous pouvez y revenir et les lire à tout moment en utilisant l'emplacement numérique assigné.
* **Surveillance continue :** Les Cellules assignées sont surveillées automatiquement en arrière-plan. Si Excel déclenche un recalcul ou une modification de Cellule, BOA annonce instantanément la nouvelle valeur. Basculez manuellement ou effacez tout via les emplacements de commande.

### Améliorations PowerPoint

#### 1. Sélecteurs de couleurs accessibles
* Déverrouille la boîte de dialogue Couleur personnalisée dans PowerPoint.
* Identifie et lit explicitement et correctement les zones d'édition "Rouge", "Vert" et "Bleu" (en remplaçant `PowerPointRGBEdit`).
* Mappe le champ de saisie Hexadécimal précédemment invisible afin que NVDA puisse lire proprement la valeur de couleur Hexadécimale complète.

#### 2. Prise en charge de la grille de couleurs standard
* Naviguer dans la grille hexagonale de couleurs "Standard" de PowerPoint est normalement lu comme "Graphique" ou du silence.
* BOA suit vos touches fléchées à travers l'hexagone et récupère silencieusement la valeur de couleur masquée, en vous l'annonçant en temps réel (par exemple, "Couleur #FF0000").

### Infrastructure et Mécanismes techniques

#### Le Mode préfixe de commande
Pour éviter les conflits de frappe avec d'autres plugins NVDA, BOA utilise un **Mode préfixe de commande** :
1. Appuyez sur le raccourci d'activation pour entrer en Mode de commande. Vous entendrez un bip aigu.
2. Appuyez sur une touche secondaire pour déclencher une fonctionnalité spécifique.
3. Si vous appuyez sur une touche non valide, vous entendrez un bip d'erreur.

#### Personnalisation et Panneau de paramètres
* Les fonctionnalités de BOA sont entièrement modulaires et peuvent être activées ou désactivées à tout moment. Allez dans `Menu NVDA -> Préférences -> Paramètres -> Améliorations Office BOA` pour activer ou désactiver les fonctionnalités individuelles.
* **Touches d'accélération intelligentes :** Chaque paramètre possède un raccourci d'accélérateur `Alt+Touche` mathématiquement unique dans le panneau. Par exemple, appuyez sur `Alt+E` pour sauter instantanément au groupe Excel, `Alt+P` pour PowerPoint et `Alt+W` pour Word.
* Les paramètres sont enregistrés de manière sécurisée dans un fichier JSON autonome (`boa_settings.json`), garantissant que votre configuration NVDA principale n'est jamais corrompue.
* Si Microsoft Office corrige officiellement un bogue d'accessibilité à l'avenir, vous pouvez désactiver en toute sécurité le crochet de remplacement spécifique de BOA sans perdre le reste des fonctionnalités du Module complémentaire.

#### Sécurité et Limites d'intégration
* Les injections dans le presse-papiers vérifient strictement les identifiants de processus de premier plan de la fenêtre pour éviter la fuite de données vers d'autres applications.
* Certains raccourcis clavier personnalisés sont entièrement exposés dans la boîte de dialogue Gestes de commande de NVDA sous la catégorie "Better Office Accessibility".

---

## 📋 Prérequis

* **NVDA :** Version 2026.1.0 ou ultérieure.
* **Applications :** Microsoft Excel et Microsoft PowerPoint.

---

## 💾 Installation

1. Téléchargez le dernier fichier de version `.nvda-addon`, ou trouvez-le dans la boutique native des modules complémentaires de NVDA.
2. Si vous installez à partir d'un fichier, ouvrez le fichier ou utilisez `Boutique des modules complémentaires de NVDA -> Installer à partir d'un fichier externe`.
3. Redémarrez NVDA.

---

## 🛠️ Journal des modifications

### Nouveautés de la v1.6.1
* **Localisation approfondie** : Correction des traductions manquantes au cœur des modules d'amélioration d'Excel (tels que l'Analyseur de disposition de feuille et le Déplaceur rapide de feuille) pour garantir une couverture de localisation à 100 %.
* **Support de traduction élargi** : Ajout de 7 nouvelles langues au système (Turc, Polonais, Coréen, Ukrainien, Tchèque, Ourdou et Pendjabi).
  *(Note : Ces traductions ont été générées par une IA, de légères erreurs de traduction ou imprécisions peuvent donc être présentes.)*

### v1.6.0
* **Prise en charge complète de la traduction** : Le Module complémentaire est désormais entièrement localisé avec la prise en charge de 17 langues mondiales. 
  *(Remarque : Ces traductions ont été générées par l'IA, de sorte que des erreurs de traduction mineures ou des inexactitudes peuvent être présentes.)*
* **Gouvernance stricte du code** : Application des en-têtes de droits d'auteur GPL-2.0 sur l'ensemble du code source.

### Version 1.5.0 
#### Nouvelles fonctionnalités
##### Radar de fin de données
Lors de la navigation dans de grandes Feuilles de calcul, il peut être difficile de savoir si une Cellule vide signifie que vous avez atteint la fin d'une liste, ou s'il y a simplement une lacune dans les données. Le **Radar de fin de données** agit comme une vérification intelligente du périmètre pour vous éviter de parcourir l'espace vide à l'aveugle.
Chaque fois que vous naviguez vers une Cellule vide, BOA analyse instantanément les Cellules restantes dans votre direction de déplacement. S'il n'y a absolument aucune donnée restante, il annoncera de manière proactive :
* *"Plus de données en dessous"*
* *"Plus de données au-dessus"*
* *"Plus de données à droite"*
* *"Plus de données à gauche"*
**Options de configuration :**
Vous pouvez configurer cette fonctionnalité via `Préférences NVDA -> Paramètres -> Améliorations Office BOA`. Parce que les Feuilles de calcul peuvent contenir des complexités masquées (comme des formules invisibles ou des Lignes réduites), le radar propose trois modes de fonctionnement :
1. **Désactivé** : Désactive entièrement le radar.
2. **Vérification stricte de la mémoire (Nbval) [Par défaut]** : L'approche la plus sûre et la plus rapide. Elle vérifie la mémoire brute de la Feuille de calcul. Si elle détecte *quoi que ce soit* en dessous de vous (y compris des Lignes masquées, du texte, des nombres ou des formules invisibles), elle reste complètement silencieuse pour éviter les fausses alarmes. Elle n'annonce "Plus de données" que lorsque le reste de la Feuille est 100 % mathématiquement vide.
3. **Données visibles uniquement (Moteur mathématique)** : Un moteur très avancé conçu pour les Feuilles complexes. Il filtre intelligemment les Lignes masquées et les formules invisibles (par exemple, `=""`). Il ne restera silencieux que s'il reste de vrais nombres ou du texte visible sur votre chemin.

### Version 1.4 - 2026-06-12
#### Nouvelles fonctionnalités
* **Moniteur de Cellule :** Utilisez les chemins de commande pour assigner des Cellules spécifiques à des emplacements de mémoire. Vous pouvez y revenir et les lire à tout moment en utilisant l'emplacement numérique assigné.
* **Surveillance continue :** Les Cellules assignées sont surveillées automatiquement en arrière-plan. Si Excel déclenche un recalcul ou une modification de Cellule, BOA annonce instantanément la nouvelle valeur. Basculez manuellement ou effacez tout via les emplacements de commande.

#### Corrections de bogues

### Version 1.3.0 — 2026-06-05
*Version finale.*

#### Nouvelles fonctionnalités
* **Analyseur de disposition de Feuille :** Ajout d'une puissante infrastructure d'analyse de disposition. Détecte instantanément la Protection de la Feuille de calcul, les Filtres de Colonne actifs, les Onglets de Feuille masqués et les bordures absolues masquées tout en mettant en cache les Blocs de données découverts.
* **Navigation guidée dans les Blocs de données :** La navigation post-analyse permet des sauts de curseur immédiats entre les principaux groupes de données, en contournant les Cellules vides de manière transparente.
* **Annonceur de mise en forme conditionnelle :** Détecte et lit automatiquement la couleur dynamique, le style de police et la nuance d'arrière-plan des Cellules modifiées par les règles de Mise en forme conditionnelle d'Excel.
* **Accélérateurs de paramètres explicites :** Refonte complète de l'interface graphique des paramètres BOA pour se conformer strictement à l'architecture NVDA. Chaque case à cocher de fonctionnalité possède désormais un raccourci `Alt+Lettre` globalement unique, empêchant le cycle du clavier et éliminant les échecs de navigation de la première lettre.

#### Corrections de bogues
* **Détection des limites des bords absolus :** Remplacement des vérifications de bord `UsedRange` natives de COM par des vérifications de limites mathématiques 1D absolues (`Ligne 1048576` et `Colonne 16384`) pour garantir la détection des Lignes/Colonnes masquées même si elles se trouvent loin en dehors du Bloc de données actif.
* **Sauvegardes sécurisées des propriétés COM paresseuses :** Renforcement des boucles de propriétés COM pour éviter les blocages de threads NVDA lors de l'évaluation de millions de structures masquées contiguës.

### Version 1.2.0 — 2026-06-03
*Version finale.*

#### Nouvelles fonctionnalités
* **Mise en cache au lancement de l'application :** Refonte architecturale majeure. Les modules de base sont désormais chargés de manière asynchrone exactement lorsque vous vous concentrez sur les applications Office, éliminant ainsi le décalage de démarrage, résolvant complètement le problème de Focus de l'objet "inconnu" sur les boîtes de dialogue de renommage et préservant la structure du code multi-fichiers.
* **Suivi amélioré des Cellules (Mathématiques COM 1D) :** Réécriture de la logique de détection des espaces de Cellules masquées pour évaluer uniquement les sections transversales unidimensionnelles (`current_col` ou `current_row`). Cela réduit la charge de calcul COM de plus de 16 millions de Cellules, éliminant instantanément les blocages de navigation lors du saut de plages masquées.
* **Nettoyage de la mémoire de processus :** Implémentation du suivi des descripteurs de fenêtre Excel (`Hwnd`) pour détecter lorsque l'utilisateur ferme et rouvre Excel. Cela efface activement la mémoire d'état global obsolète et résout complètement la fausse annonce "Feuille masquée" lors de l'ouverture d'un nouveau "Classeur1".

#### Corrections de bogues
* **Double annonce de sélection :** Migration de l'imprévisible et asynchrone `winUser.getKeyState` vers l'implémentation de `api.getLastInputGesture()` pour supprimer parfaitement les doubles annonces lors de l'utilisation des touches Maj+Flèche.
* **Désactivation du détecteur de limites :** Le détecteur proactif de limites a été désactivé pour protéger la stabilité de la navigation native de NVDA, s'appuyant entièrement sur le système de suivi de saut d'espace.

### Version 1.1.0 — 2026-05-30
*Version finale.*

#### Nouvelles fonctionnalités
* **Interface graphique de paramètres :** Ajout d'un panneau natif Améliorations Office BOA dans `NVDA -> Préférences -> Paramètres` pour activer ou désactiver facilement les fonctionnalités.
* **Crochet SafeRichEdit :** Empêche les plantages silencieux de NVDA lors de l'interaction avec les contrôles RichEdit dans Office 2024.
* **Raccourcis clavier personnalisables :** Tous les raccourcis clavier de BOA sont désormais entièrement exposés dans la boîte de dialogue Gestes de commande de NVDA sous la catégorie "Better Office Accessibility".
* **Excel : Détection de saut de Ligne/Colonne masquée :** Annonce de manière proactive la navigation au-delà de Lignes ou de Colonnes masquées, garantissant que vous ne manquez jamais les données filtrées. Peut être désactivé dans les paramètres.

#### Corrections de bogues
* **Sécurité des threads :** Suppression de tous les délais de blocage (`time.sleep`) et remplacement par des rappels asynchrones non bloquants NVDA pour garantir que le lecteur d'écran ne bégaie jamais pendant les opérations en arrière-plan.

### Version 1.0.0 — 2026-05-24
*Sortie publique initiale.*

#### Nouvelles fonctionnalités
* **Excel : Organisateur de Feuille en masse :** Réorganisez instantanément plusieurs Feuilles à la fois à l'aide d'une boîte de dialogue entièrement accessible.
* **Excel : Déplacement rapide de Feuille :** Déplacez la Feuille active vers la gauche, la droite, au début ou à la fin via des commandes au clavier.
* **Excel : Renommage accessible de Feuille :** Intercepte le champ de renommage natif inaccessible et le remplace par une boîte de dialogue accessible et fiable.
* **Excel : Suivi intelligent de la sélection :** Annonce avec précision les sélections et désélections de plages multi-Cellules.
* **PowerPoint : Sélecteurs de couleurs accessibles :** Permet à NVDA de lire avec précision les valeurs RVB et Hexadécimales à l'intérieur de la boîte de dialogue Couleur personnalisée.
* **PowerPoint : Prise en charge de la grille de couleurs standard :** Intercepte la navigation avec les touches fléchées pour lire les codes Hexadécimaux masqués depuis la grille hexagonale de couleurs inaccessible.
