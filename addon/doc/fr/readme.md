# BOA: Better Office Accessibility

BOA est une suite puissante d'améliorations d'accessibilité pour Microsoft Office, conçue pour améliorer considérablement l'expérience des lecteurs d'écran pour les utilisateurs de NVDA. Elle corrige directement les composants d'interface utilisateur inaccessibles et introduit des outils de navigation rapide pour Excel et PowerPoint.

---

## ⌨️ Hotkey Reference

| Feature | Key Combination | Context / Notes |
| :--- | :--- | :--- |
| **Entrer en mode commande** | `[Prefix]` (Par défaut : `NVDA+E`) | Active le mode de préfixe de commande (émet un bip aigu) |
| **Annuler le mode commande** | `Escape` | Quitte le mode de préfixe de commande |
| **AMÉLIORATIONS EXCEL** | | |
| **Analyser la disposition de la feuille** | `[Prefix]`, puis `L` | À exécuter dans Excel avant de naviguer dans les blocs de données |
| **Aller au bloc de données le plus proche** | `[Prefix]`, puis `J` | Nécessite d'abord une analyse de la disposition |
| **Ouvrir l'organisateur de feuilles en masse** | `[Prefix]`, puis `X` | Ouvre la boîte de dialogue accessible de réorganisation des feuilles |
| **Annonceur de formule brute** | `[Prefix]`, puis `F2` | Appuyer une fois pour entendre la formule brute |
| **Éditeur de formule puissant** | `[Prefix]`, puis `F2` deux fois | Double appui pour ouvrir l'éditeur de formule multiligne accessible |
| **Tracer les antécédents** | `[Prefix]`, puis `Shift+P` | Fonctionnalité de traçage des antécédents de manière accessible.|
| **Tracer les dépendants** | `[Prefix]`, puis `Shift+D` | Fonctionnalité de traçage des dépendants de manière accessible, appuyer sur Entrée sur une cellule vous y téléporte.|
| **Mise en forme conditionnelle détaillée**| `[Prefix]`, puis `F` | Annonce les détails complets de mise en forme de la cellule active |
| **Déplacer la feuille active vers la gauche** | `NVDA+Shift+LeftArrow` | Déplace la feuille active d'une position vers le haut |
| **Déplacer la feuille active vers la droite** | `NVDA+Shift+RightArrow` | Déplace la feuille active d'une position vers le bas |
| **Déplacer la feuille au début/à la fin** | `NVDA+Shift+Home` / `End` | Envoie la feuille de calcul aux limites absolues |
| **Masquer / Afficher la ligne** | `Ctrl+9` / `Ctrl+Shift+9` | Raccourci natif ; BOA annonce explicitement le changement de visibilité |
| **Masquer / Afficher la colonne** | `Ctrl+0` / `Ctrl+Shift+0` | Raccourci natif ; BOA annonce explicitement le changement de visibilité |
| **Afficher la colonne (secours)** | `NVDA+Ctrl+Shift+0` | Contourne les conflits de raccourcis clavier de langue de saisie Windows |
| **Associer une cellule à un emplacement mémoire** | `[Prefix]`, puis `Shift+1` à `Shift+9` | Assigne la cellule active à un emplacement de surveillance en arrière-plan |
| **Lire l'emplacement de cellule surveillée** | `[Prefix]`, puis `1` à `9` | Rappelle et lit la valeur de l'emplacement assigné |
| **Saut direct vers un emplacement** | `Alt` + `1` à `9` | Déplace instantanément votre curseur vers un emplacement surveillé |
| **Retourner à la cellule précédente** | `[Prefix]`, puis `\` | Vous téléporte instantanément en arrière après avoir vérifié un emplacement |
| **Boîte de dialogue du gestionnaire d'emplacements** | `[Prefix]`, puis `Alt+M` | Ouvre une boîte de dialogue pour afficher et gérer tous les surveillants actifs |
| **Activer/désactiver la surveillance en arrière-plan** | `[Prefix]`, puis `M` | Active ou désactive manuellement le suivi des calculs en arrière-plan |
| **Effacer tous les emplacements mémoire** | `[Prefix]`, puis `Backspace` | Purge toutes les cellules surveillées en arrière-plan enregistrées |
| **AMÉLIORATIONS POWERPOINT** | | |
| **Analyseur de disposition de diapositive** | `[Prefix]`, puis `L` | Analyse et annonce la disposition spatiale de la diapositive actuelle |
| **Analyseur de document** | `[Prefix]`, puis `D` | Génère une table des matières complète et un rapport d'état |
| **Organisateur de diapositives en masse** | `[Prefix]`, puis `X` | Ouvre la boîte de dialogue accessible pour réorganiser plusieurs diapositives |
| **AMÉLIORATIONS WORD** | | |
| **Auditeur de mise en forme** | `[Prefix]`, puis `F` | Audite le document actuel pour détecter les incohérences de mise en forme |
| **Analyseur de document** | `[Prefix]`, puis `D` | Analyse la disposition et la structure du document Word actuel |

---

## 🚀 Features

### Excel Enhancements

#### 1. Sheet Layout Analyzer & Caching
Scannez instantanément n'importe quelle feuille de calcul Excel pour comprendre sa structure, ses éléments masqués et ses blocs de données.
* **Comment ça marche :** BOA scanne rapidement la feuille et annonce les blocs de données actifs. Il vous avertit également des **onglets de feuille masqués**, des **filtres** actifs, des **modes protégés** et des **limites externes masquées** (par exemple, si des colonnes proches du bord droit de la feuille sont masquées, vous évitant de manquer des données hors écran).
* **Navigation dans les données :** Après le scan, vous pouvez utiliser les raccourcis clavier de saut de bloc de données pour déplacer instantanément votre curseur d'un bloc à l'autre, évitant ainsi sans effort des milliers de cellules vides.

#### 2. Bulk Sheet Organizer
Réorganisez et arrangez instantanément plusieurs feuilles à la fois à l'aide d'une boîte de dialogue entièrement accessible.
* **Comment ça marche :** Ouvre une boîte de dialogue dans laquelle vous pouvez sélectionner une feuille et l'associer à une nouvelle position. Les déplacements planifiés sont répertoriés dans un tableau de données (appuyez sur `Del` pour corriger une erreur). Cliquez sur `OK` et votre classeur est réorganisé instantanément.

#### 3. Quick Sheet Mover
Déplacez instantanément la feuille active vers la gauche, la droite, tout au début ou tout à la fin à l'aide de vos raccourcis clavier.

#### 4. Accessible Sheet Renaming
* Lors du renommage d'une feuille, NVDA a nativement des difficultés à lire les caractères que vous tapez.
* BOA injecte une classe personnalisée `ExcelSheetRenameEdit` qui utilise le moteur `SafeRichEdit`, ce qui signifie que vous pouvez lire précisément par caractère, mot ou ligne tout en renommant. Cela sert d'amélioration au comportement de renommage par défaut existant.

#### 5. Hidden Row/Column Tracker
* Suit de manière proactive vos mouvements sur la grille pour vous éviter de manquer des données masquées ou filtrées.
* **Cellules fragmentées traversées :** Si vous sautez à travers une section fortement fragmentée ou masquée de la grille (par exemple, passer de la ligne 3 à la ligne 10 parce que les lignes 4 à 9 sont masquées), BOA annonce explicitement "Rows 4 through 9 hidden". Cela garantit que vous savez toujours quand des données ont été sautées dans la structure.

#### 6. Conditional Formatting Announcer
* Lit automatiquement la couleur, le style de police et la couleur d'arrière-plan des cellules qui ont été modifiées dynamiquement par les règles de mise en forme conditionnelle d'Excel.
* Vous donne le véritable état visuel de la cellule plutôt que la simple valeur brute sous-jacente. Initialement, lors de la mise au point sur la cellule, il annonce "has conditional formatting, and some other minor details" (contient une mise en forme conditionnelle, et d'autres détails mineurs). Pour des informations complètes, utilisez la configuration détaillée des raccourcis clavier qui est NVDA E et F.

#### 7. Better selection announcement
lit si la cellule ou la plage est sélectionnée ou désélectionnée.

#### 8 Cell monitor:
* **Cell Monitor :** Utilisez les chemins de commande pour associer des cellules spécifiques à des emplacements mémoire. Vous pouvez y revenir et les lire à tout moment en utilisant l'emplacement numérique assigné.
* **Surveillance continue :** Les cellules associées sont automatiquement surveillées en arrière-plan. Si Excel déclenche un recalcul ou une modification de cellule, BOA annonce instantanément la nouvelle valeur. Activez/désactivez manuellement ou effacez tout via les emplacements de commande.
* **Excel : Mises à niveau Cell Monitor Pro :** 
  - **Boîte de dialogue du gestionnaire d'emplacements (`NVDA+E`, puis `Alt+M`) :** Ouvre une boîte de dialogue répertoriant toutes vos cellules activement surveillées. Appuyez sur `Enter` pour accéder instantanément à l'une d'elles.
  - **Retour rapide (`NVDA+E`, puis `\`) :** Vous téléporte instantanément à votre cellule de travail précédente après avoir vérifié un emplacement.
  - **Saut direct vers un emplacement (`Prefix + Alt` + `Slot Number`) :** Contourne le préfixe pour sauter instantanément vers un emplacement de cellule assigné.

#### 9 Power editor
* **Excel : L'éditeur puissant (Éditeur de formules accessible) :** Un changement radical pour modifier des formules volumineuses.
  - **Simple appui sur `NVDA+E`, puis `F2` :** Annonce instantanément la chaîne de formule brute de la cellule active (ou annonce "No formula" - Pas de formule).
  - **Double appui sur `NVDA+E`, puis `F2` :** Ouvre un éditeur multiligne entièrement accessible pour modifier en toute sécurité des formules imbriquées et massives. La touche `Enter` native ajoute des sauts de ligne pour une lecture facile, et `Ctrl+Enter` enregistre les modifications dans Excel.
  - *Contrôles de sécurité :* Intercepte en toute sécurité les erreurs de syntaxe avant qu'elles ne corrompent votre feuille, et détecte les erreurs post-calcul (comme `#NAME?` ou `#DIV/0!`) pour vous avertir instantanément si une formule est erronée.

#### 10 Formula auditing and evaluation enhancements:
* **Excel : Audit et évaluation des formules :** Ajout de raccourcis personnalisés (`NVDA+E`, puis `Shift+P` et `NVDA+E`, puis `Shift+D`) pour tracer de manière fiable les antécédents et les dépendants. De plus, la boîte de dialogue native d'Excel « Évaluer la formule » est désormais entièrement accessible ; NVDA lit automatiquement les résultats évalués à chaque étape du calcul !

### PowerPoint Enhancements

#### 1. Accessible Color Pickers
* Déverrouille la boîte de dialogue de couleur personnalisée dans PowerPoint.
* Identifie et lit explicitement et correctement les zones d'édition "Red", "Green" et "Blue" (en remplaçant `PowerPointRGBEdit`).
* Associe le champ de saisie Hex (hexadécimal) auparavant invisible afin que NVDA puisse lire proprement la valeur de couleur hexadécimale complète.

#### 2. Standard Color Grid Support
* La navigation dans la grille hexagonale de couleurs « Standard » de PowerPoint lit normalement « Graphique » ou reste silencieuse.
* BOA suit vos touches fléchées à travers l'hexagone et récupère silencieusement la valeur de couleur masquée, vous l'annonçant en temps réel (par ex., "Color #FF0000").

#### 3 Bulk Slide Organizer:
* **PowerPoint : Organisateur de diapositives en masse (Expérimental) (`NVDA+E`, puis `X`) :** Semblable à la fonctionnalité Excel, vous pouvez désormais réordonner, déplacer et organiser instantanément plusieurs diapositives PowerPoint à la fois à l'aide d'une boîte de dialogue entièrement accessible.

#### 4 Slide lay out analyzer
* **PowerPoint : Analyseur de disposition de diapositive (Expérimental) (`NVDA+E`, puis `L`) :** Scanne instantanément votre diapositive actuellement active pour comprendre sa disposition spatiale et ses contraintes d'accessibilité, garantissant une expérience de lecture d'écran fluide et réactive. C'est-à-dire que vous obtiendrez ici des détails sur la diapositive actuelle similaires à l'analyseur de disposition de feuille d'Excel.


#### 5 Complete Document [PPT] analyzer
* **PowerPoint : Analyseur de document complet (Expérimental) (`NVDA+E`, puis `D`) :** Un outil d'accessibilité très avancé, traité en arrière-plan, qui cartographie l'intégralité d'une présentation sans figer le moteur de parole de NVDA. Il fournit une table des matières virtuelle hautement navigable, détecte les discordances dans l'ordre de lecture (ordre visuel vs ordre Z), signale les diapositives de type « pavé de texte » (Wall of Text) et cartographie les objets complexes comme SmartArt et les tableaux de données.

#### 6 shape movement [adjustment] enhancements:
* **PowerPoint : Mode audio de déplacement de forme (Expérimental) :** Introduit des repères audio spatiaux 3D sur le canevas de PowerPoint. Fournit un retour sonore indiquant la direction et les limites d'un objet au fur et à mesure que vous le déplacez, améliorant considérablement la conscience spatiale.

### Word Enhancements:
#### 1. Document Analyzer inspired and derived from Paul's word access addon:
* **Word : Analyseur de document (`NVDA+E`, puis `D`) :** Affichez instantanément un aperçu structurel de votre document Word. *(Une note spéciale de crédit et de remerciements à Paul : cette fonctionnalité a été directement inspirée par son brillant module complémentaire « Word Access ». Nous lui sommes profondément reconnaissants pour son travail fondateur dans ce domaine !)*

#### 2 Formatting Auditor
* **Word : Auditeur de mise en forme (`NVDA+E`, puis `F`) :** Analyse votre document Word pour détecter les incohérences de mise en forme afin de garantir le respect des normes visuelles.

#### 3 Foot note reader:
* **Word : Annonceur automatique de notes de bas de page :** Les notes de bas de page seront désormais annoncées automatiquement au fil de la lecture, en fonction de vos paramètres BOA personnalisés. *(Note : La prise en charge des notes de fin et des commentaires est prévue pour une version ultérieure).*

### Infrastructure & Technical Mechanisms

#### The Command Prefix Mode
Pour éviter les conflits de raccourcis clavier avec d'autres extensions NVDA, BOA utilise un **mode de préfixe de commande** :
1. Appuyez sur le raccourci d'activation pour entrer en mode commande. Vous entendrez un bip aigu. Par défaut, il s'agit de NVDA plus E.
2. Appuyez sur une touche secondaire pour déclencher une fonctionnalité spécifique.
3. Si vous appuyez sur une touche invalide, vous entendrez un bip d'erreur.

#### Customization & Settings Panel
* Les fonctionnalités de BOA sont entièrement modulaires et peuvent être activées ou désactivées à tout moment. Allez dans `NVDA Menu -> Preferences -> Settings -> BOA Office Enhancements` pour activer ou désactiver des fonctionnalités individuelles.
* **Touches de raccourci intelligentes :** Chaque paramètre dispose d'un raccourci d'accélérateur `Alt+Touche` mathématiquement unique dans le panneau. Par exemple, appuyez sur `Alt+E` pour accéder instantanément au groupe Excel, `Alt+P` pour PowerPoint et `Alt+W` pour Word.
* Les paramètres sont enregistrés en toute sécurité dans un fichier JSON autonome (`boa_settings.json`), garantissant que votre configuration principale de NVDA n'est jamais corrompue.
* Si Microsoft Office corrige officiellement un bogue d'accessibilité à l'avenir, vous pouvez désactiver en toute sécurité le crochet de contournement spécifique de BOA sans perdre le reste des fonctionnalités du module complémentaire.
* **Personnalisation des raccourcis de saisie :** Toutes les fonctionnalités de toutes les applications Office ont été explicitement exposées dans la boîte de dialogue native des raccourcis de saisie de NVDA, vous offrant une liberté totale pour personnaliser chaque raccourci clavier.

#### Security & Integration Boundaries
* Les injections de presse-papiers vérifient strictement les ID de processus au premier plan des fenêtres pour empêcher la fuite de données vers d'autres applications.
* Certains raccourcis personnalisés sont entièrement exposés dans la boîte de dialogue Raccourcis de saisie de NVDA sous la catégorie « Better Office Accessibility ».

---

## 📋 Requirements

* **NVDA :** Version 2026.1.0 ou ultérieure.
* **Applications :** Microsoft Excel & Microsoft PowerPoint.

---

## 💾 Installation

1. Téléchargez le dernier fichier de version `.nvda-addon`, ou recherchez-le dans la boutique d'extensions NVDA native.
2. Si vous installez depuis un fichier, ouvrez le fichier ou utilisez `Boutique d'extensions NVDA -> Installer depuis un fichier externe`.
3. Redémarrez NVDA.

---

## 🛠️ Changelog

### Version 2.0.1
#### Améliorations UX/UI
* **Boîte de dialogue des paramètres avec onglets :** Réorganisation du panneau de paramètres BOA en onglets accessibles (&Excel, &Word et &PowerPoint) à l'aide de `wx.Notebook`, améliorant considérablement la navigation du lecteur d'écran et éliminant les longues listes de défilement. Vous pouvez basculer rapidement entre les onglets en utilisant `Alt+E`, `Alt+W`, `Alt+P`, ou les raccourcis standard `Ctrl+PageDown`/`Ctrl+PageUp`.
* **Compatibilité NVDA 2026.2 :** Testé et certifié pour NVDA 2026.2.

### Version 2.0.0
#### Nouvelles fonctionnalités
* **PowerPoint : Analyseur de document complet (Expérimental) (`NVDA+E`, puis `D`) :** Un outil d'accessibilité très avancé, traité en arrière-plan, qui cartographie l'intégralité d'une présentation sans figer le moteur de parole de NVDA. Il fournit une table des matières virtuelle hautement navigable, détecte les discordances dans l'ordre de lecture (ordre visuel vs ordre Z), signale les diapositives de type « pavé de texte » (Wall of Text) et cartographie les objets complexes comme SmartArt et les tableaux de données.
* **PowerPoint : Analyseur de disposition de diapositive (Expérimental) (`NVDA+E`, puis `L`) :** Scanne instantanément votre diapositive actuellement active pour comprendre sa disposition spatiale et ses contraintes d'accessibilité, garantissant une expérience de lecture d'écran fluide et réactive. C'est-à-dire que vous obtiendrez ici des détails sur la diapositive actuelle similaires à l'analyseur de disposition de feuille d'Excel.
* **PowerPoint : Organisateur de diapositives en masse (Expérimental) (`NVDA+E`, puis `X`) :** Semblable à la fonctionnalité Excel, vous pouvez désormais réordonner, déplacer et organiser instantanément plusieurs diapositives PowerPoint à la fois à l'aide d'une boîte de dialogue entièrement accessible.
* **PowerPoint : Mode audio de déplacement de forme (Expérimental) :** Introduit des repères audio spatiaux 3D sur le canevas de PowerPoint. Fournit un retour sonore indiquant la direction et les limites d'un objet au fur et à mesure que vous le déplacez, améliorant considérablement la conscience spatiale. Comme mentionné, ceci est expérimental, nous attendons vos retours pour l'améliorer.
* **Word : Auditeur de mise en forme (`NVDA+E`, puis `F`) :** Analyse votre document Word pour détecter les incohérences de mise en forme afin de garantir le respect des normes visuelles.
* **Word : Analyseur de document (`NVDA+E`, puis `D`) :** Affichez instantanément un aperçu structurel de votre document Word. *(Une note spéciale de crédit et de remerciements à Paul : cette fonctionnalité a été directement inspirée par son brillant module complémentaire « Word Access ». Nous lui sommes profondément reconnaissants pour son travail fondateur dans ce domaine !)*
* **Word : Annonceur automatique de notes de bas de page :** Les notes de bas de page seront désormais annoncées automatiquement au fil de la lecture, en fonction de vos paramètres BOA personnalisés. *(Note : La prise en charge des notes de fin et des commentaires est prévue pour une version ultérieure).*
* **Excel : L'éditeur puissant (Éditeur de formules accessible) :** Un changement radical pour modifier des formules volumineuses.
  - **Simple appui sur `NVDA+E`, puis `F2` :** Annonce instantanément la chaîne de formule brute de la cellule active (ou annonce "No formula" - Pas de formule).
  - **Double appui sur `NVDA+E`, puis `F2` :** Ouvre un éditeur multiligne entièrement accessible pour modifier en toute sécurité des formules imbriquées et massives. La touche `Enter` native ajoute des sauts de ligne pour une lecture facile, et `Ctrl+Enter` enregistre les modifications dans Excel.
  - *Contrôles de sécurité :* Intercepte en toute sécurité les erreurs de syntaxe avant qu'elles ne corrompent votre feuille, et détecte les erreurs post-calcul (comme `#NAME?` ou `#DIV/0!`) pour vous avertir instantanément si une formule est erronée.
* **Excel : Audit et évaluation des formules :** Ajout de raccourcis personnalisés (`NVDA+E`, puis `Shift+P` et `NVDA+E`, puis `Shift+D`) pour tracer de manière fiable les antécédents et les dépendants. De plus, la boîte de dialogue native d'Excel « Évaluer la formule » est désormais entièrement accessible ; NVDA lit automatiquement les résultats évalués à chaque étape du calcul !
* **Excel : Mises à niveau Cell Monitor Pro :** 
  - **Boîte de dialogue du gestionnaire d'emplacements (`NVDA+E`, puis `Alt+M`) :** Ouvre une boîte de dialogue répertoriant toutes vos cellules activement surveillées. Appuyez sur `Enter` pour accéder instantanément à l'une d'elles.
  - **Retour rapide (`NVDA+E`, puis `\`) :** Vous téléporte instantanément à votre cellule de travail précédente après avoir vérifié un emplacement.
  - **Saut direct vers un emplacement (`Alt` + `Slot Number`) :** Contourne complètement le préfixe et saute instantanément vers un emplacement de cellule assigné.
* **Personnalisation des raccourcis de saisie :** Toutes les fonctionnalités de toutes les applications Office ont été explicitement exposées dans la boîte de dialogue native des raccourcis de saisie de NVDA, vous offrant une liberté totale pour personnaliser chaque raccourci clavier.

#### Améliorations de l'expérience utilisateur et de l'interface graphique
* **Rapports navigables unifiés :** Nous avons adopté un système de rapports HTML unifié dans l'ensemble de l'extension. Les fonctionnalités telles que l'annonceur de mise en forme conditionnelle d'Excel, les analyseurs de disposition et les analyseurs de documents ne se contentent plus de lire de gros blocs de texte ; leurs résultats s'ouvrent désormais dans une fenêtre HTML native et navigable, vous permettant d'examiner les données à votre propre rythme.
* **Excel : Suivi amélioré des dépendants/antécédents :** Amélioration considérable de la sortie vocale pour les raccourcis de traçage de formules natifs d'Excel (`Ctrl+[` pour les antécédents directs et `Ctrl+]` pour les dépendants directs). NVDA annonce désormais explicitement exactement quelles cellules ont été sélectionnées.
* **Excel : Prise en charge des cellules fusionnées :** Les cellules fusionnées sont désormais correctement détectées et explicitement annoncées par le suivi de cellule avec saut d'espaces.

#### Corrections de bogues
* **Word : Double lecture des éléments de liste :** Implémentation d'un correctif temporaire pour résoudre le bogue où NVDA lit deux fois les éléments de liste de paragraphes dans certaines vues Word.
* **Excel : Bogue de localisation du surveillant de cellule :** Résolution des bogues de suivi sous-jacents causés par les récentes mises à jour de localisation des traductions.

### Quoi de neuf dans la v1.6.1
* **Localisation approfondie des fichiers :** Correction des traductions de chaînes manquantes au sein des modules d'amélioration Excel (tels que l'analyseur de disposition de feuille et le déplacement rapide de feuille) pour garantir une couverture de localisation de 100 %.
* **Prise en charge étendue des traductions :** Ajout de 7 nouvelles langues au système (turc, polonais, coréen, ukrainien, tchèque, ourdou et pendjabi). 
  *(Note : Ces traductions ont été générées par IA, de sorte que de légères erreurs ou inexactitudes de traduction peuvent être présentes.)*

### v1.6.0
* **Prise en charge complète des traductions :** Le module complémentaire est désormais entièrement localisé avec la prise en charge de 17 langues mondiales. 
  *(Note : Ces traductions ont été générées par IA, de sorte que de légères erreurs ou inexactitudes de traduction peuvent être présentes.)*
* **Strict Code Governance** : Application des en-têtes de droit d'auteur GPL-2.0 dans l'ensemble de la base de code.""",

### Version 1.5.0 
#### Nouvelles fonctionnalités
##### Radar de fin de données (End of Data Radar)
Lors de la navigation dans de grandes feuilles de calcul, il peut être difficile de savoir si une cellule vide signifie que vous avez atteint la fin d'une liste, ou s'il s'agit simplement d'un espace vide dans les données. Le **Radar de fin de données** agit comme un contrôle de périmètre intelligent pour vous éviter de naviguer à l'aveugle à l'aide des flèches dans un espace vide.
Chaque fois que vous naviguez dans une cellule vide, BOA scanne instantanément les cellules restantes dans votre direction de déplacement. S'il ne reste absolument aucune donnée, il l'annoncera de manière proactive :
* *"Plus de données en dessous"*
* *"Plus de données au-dessus"*
* *"Plus de données à droite"*
* *"Plus de données à gauche"*
**Options de configuration :**
Vous pouvez configurer cette fonctionnalité via `NVDA Preferences -> Settings -> BOA Office Enhancements`. Comme les feuilles de calcul peuvent contenir des complexités masquées (comme des formules invisibles ou des lignes réduites), le radar propose trois modes de fonctionnement :
1. **Désactivé** : Désactive entièrement le radar.
2. **Contrôle strict de la mémoire (CountA) [Par défaut]** : L'approche la plus sûre et la plus rapide. Elle vérifie la mémoire brute de la feuille de calcul. Si elle détecte *quoi que ce soit* en dessous de vous (y compris des lignes masquées, du texte, des chiffres ou des formules invisibles), elle reste complètement silencieuse pour éviter les fausses alertes. Elle annonce seulement "No more data" (Plus de données) lorsque le reste de la feuille est à 100 % mathématiquement vide.
3. **Données visibles uniquement (moteur mathématique)** : Un moteur très avancé conçu pour les feuilles complexes. Il filtre intelligemment les lignes masquées et les formules invisibles (par exemple, `=""`). Il ne restera silencieux que s'il reste des chiffres ou du texte visibles sur votre chemin.

### Version 1.4 - 2026-06-12
#### Nouvelles fonctionnalités
* **Cell Monitor :** Utilisez les chemins de commande pour associer des cellules spécifiques à des emplacements mémoire. Vous pouvez y revenir et les lire à tout moment en utilisant l'emplacement numérique assigné.
* **Surveillance continue :** Les cellules associées sont automatiquement surveillées en arrière-plan. Si Excel déclenche un recalcul ou une modification de cellule, BOA annonce instantanément la nouvelle valeur. Activez/désactivez manuellement ou effacez tout via les emplacements de commande.

#### Bug Fixes

### Version 1.3.0 — 2026-06-05
*Version finale.*

#### Nouvelles fonctionnalités
* **Analyseur de disposition de feuille :** Ajout d'une puissante infrastructure d'analyse de disposition. Détecte instantanément la protection des feuilles de calcul, les filtres de colonnes actifs, les onglets de feuilles masqués et les bordures absolues masquées tout en mettant en cache les blocs de données découverts.
* **Navigation guidée dans les blocs de données :** La navigation post-analyse permet de déplacer immédiatement le curseur entre les principaux groupes de données, en contournant de manière transparente les cellules vides.
* **Annonceur de mise en forme conditionnelle :** Détecte et lit automatiquement la couleur dynamique, le style de police et la couleur d'arrière-plan des cellules modifiées par les règles de mise en forme conditionnelle d'Excel.
* **Accélérateurs de paramètres explicites :** Refonte complète de l'interface graphique des paramètres BOA pour se conformer strictement à l'architecture de NVDA. Chaque case à cocher de fonctionnalité possède désormais un raccourci `Alt+Lettre` unique au niveau mondial, empêchant le défilement clavier et éliminant les échecs de navigation par première lettre.

#### Corrections de bogues
* **Détection absolue des limites extérieures :** Remplacement des vérifications de limites natives COM `UsedRange` par des vérifications de limites mathématiques 1D absolues (`Row 1048576` et `Column 16384`) pour garantir la détection des lignes/colonnes masquées même si elles se situent bien en dehors du bloc de données actif.
* **Sorties sécurisées pour les propriétés COM à chargement différé :** Renforcement des boucles de propriétés COM pour éviter les blocages de thread NVDA lors de l'évaluation de millions de structures masquées contiguës.

### Version 1.2.0 — 2026-06-03
*Version finale.*

#### Nouvelles fonctionnalités
* **Mise en cache au lancement de l'application :** Refonte architecturale majeure. Les modules de base sont désormais chargés à la demande (lazy-loaded) exactement lorsque vous mettez au point sur les applications Office, éliminant le délai de démarrage, résolvant complètement le problème de mise au point d'objet « inconnu » sur les boîtes de dialogue de renommage, et préservant la structure du code multi-fichiers.
* **Suivi de cellule amélioré (calculs mathématiques COM 1D) :** Réécriture de la logique de détection des espaces de cellules masquées pour évaluer uniquement les coupes unidimensionnelles (`current_col` or `current_row`). Cela réduit la charge de calcul COM de plus de 16 millions de cellules, éliminant instantanément les blocages de navigation lors des sauts de plages masquées.
* **Nettoyage de la mémoire du processus :** Implémentation du suivi de la poignée de fenêtre Excel (`Hwnd`) pour détecter quand l'utilisateur ferme et rouvre Excel. Cela efface activement la mémoire globale d'état obsolète et résout complètement la fausse annonce « Feuille masquée » lors de l'ouverture d'un nouveau classeur « Classeur1 ».

#### Corrections de bogues
* **Annonce de double sélection :** Abandon de l'utilisation peu fiable d'asynchrones `winUser.getKeyState` et implémentation de `api.getLastInputGesture()` pour supprimer parfaitement les doubles annonces lors de l'utilisation des touches Shift+Flèches.
* **Désactivation du détecteur de limites :** Le détecteur de limites proactif a été désactivé pour protéger la stabilité de la navigation native de NVDA, se rabattant entièrement sur le suivi avec saut d'espaces.

### Version 1.1.0 — 2026-05-30
*Version finale.*

#### Nouvelles fonctionnalités
* **Interface graphique des paramètres :** Ajout d'un panneau BOA Office Enhancements natif dans `NVDA -> Préférences -> Paramètres` pour activer ou désactiver facilement les fonctionnalités.
* **Crochet SafeRichEdit :** Empêche les plantages silencieux de NVDA lors de l'interaction avec les contrôles RichEdit dans Office 2024.
* **Raccourcis personnalisables :** Tous les raccourcis BOA sont désormais entièrement exposés dans la boîte de dialogue Raccourcis de saisie de NVDA sous la catégorie « Better Office Accessibility ».
* **Excel : Détection de saut de lignes/colonnes masquées :** Annonce de manière proactive lorsque vous passez des lignes ou colonnes masquées, garantissant que vous ne manquez jamais de données filtrées. Peut être activé/désactivé dans les paramètres.

#### Corrections de bogues
* **Sécurité des threads :** Suppression de tous les délais bloquants (`time.sleep`) et remplacement par des rappels asynchrones NVDA non bloquants pour garantir que le lecteur d'écran ne bégaie jamais pendant les opérations en arrière-plan.

### Version 1.0.0 — 2026-05-24
*Première version publique.*

#### Nouvelles fonctionnalités
* **Excel : Organisateur de feuilles en masse :** Réorganisez instantanément plusieurs feuilles à la fois à l'aide d'une boîte de dialogue entièrement accessible.
* **Excel : Déplacement rapide de feuille :** Déplacez la feuille active vers la gauche, la droite, le début ou la fin via des raccourcis clavier.
* **Excel : Renommage de feuille accessible :** Intercepte le champ de renommage natif inaccessible et le remplace par une boîte de dialogue accessible et fiable.
* **Excel : Suivi intelligent de la sélection :** Annonce avec précision les sélections et désélection de plages multicellulaires.
* **PowerPoint : Sélecteurs de couleurs accessibles :** Permet à NVDA de lire avec précision les valeurs RVB et Hexadécimales dans la boîte de dialogue de couleur personnalisée.
* **PowerPoint : Prise en charge de la grille de couleurs standard :** Intercepte la navigation avec les touches fléchées pour lire les codes hexadécimaux masqués de la grille hexagonale de couleurs inaccessible.
