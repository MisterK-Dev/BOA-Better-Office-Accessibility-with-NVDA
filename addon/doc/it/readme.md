# BOA: Better Office Accessibility

BOA è una potente suite di miglioramenti dell'accessibilità per Microsoft Office, progettata per migliorare notevolmente l'esperienza con i lettori dello schermo per gli utenti NVDA. Corregge direttamente i componenti dell'interfaccia utente non accessibili e introduce strumenti di navigazione rapida per Excel e PowerPoint.

---

## ⌨️ Riferimento scorciatoie da tastiera

| Funzionalità | Combinazione di tasti | Contesto / Note |
| :--- | :--- | :--- |
| **Accedi alla modalità comando** | `[Prefix]` (Predefinito: `NVDA+E`) | Attiva la modalità prefisso comando (emette un segnale acustico acuto) |
| **Annulla la modalità comando** | `Escape` | Esce dalla modalità prefisso comando |
| **MIGLIORAMENTI PER EXCEL** | | |
| **Analizza layout foglio** | `[Prefix]`, poi `L` | Da eseguire in Excel prima di navigare tra i blocchi di dati |
| **Salta al blocco di dati più vicino** | `[Prefix]`, poi `J` | Richiede prima l'analisi del layout |
| **Apri organizzatore di fogli multipli** | `[Prefix]`, poi `X` | Apre la finestra di dialogo accessibile per riordinare i fogli |
| **Annunciatore formula grezza** | `[Prefix]`, poi `F2` | Tocco singolo per ascoltare la stringa della formula così com'è |
| **Editor di formule avanzato** | `[Prefix]`, poi `F2` due volte | Doppio tocco per aprire l'editor di formule multilinea accessibile |
| **Individua precedenti** | `[Prefix]`, poi `Shift+P` | La stessa funzionalità "Individua precedenti" in modo accessibile.|
| **Individua dipendenti** | `[Prefix]`, poi `Shift+D` | La stessa funzionalità "Individua dipendenti" in modo accessibile, premendo Invio su una cella verrai teletrasportato lì.|
| **Formattazione condizionale dettagliata**| `[Prefix]`, poi `F` | Annuncia i dettagli completi sulla formattazione della cella focalizzata |
| **Sposta foglio attivo a sinistra** | `NVDA+Shift+LeftArrow` | Sposta il foglio attivo di una posizione in alto |
| **Sposta foglio attivo a destra** | `NVDA+Shift+RightArrow` | Sposta il foglio di lavoro attivo di una posizione in basso |
| **Sposta foglio all'inizio/alla fine** | `NVDA+Shift+Home` / `End` | Invia il foglio di lavoro ai limiti assoluti |
| **Nascondi / Scopri riga** | `Ctrl+9` / `Ctrl+Shift+9` | Scorciatoia nativa; BOA annuncia esplicitamente il cambio di visibilità |
| **Nascondi / Scopri colonna** | `Ctrl+0` / `Ctrl+Shift+0` | Scorciatoia nativa; BOA annuncia esplicitamente il cambio di visibilità |
| **Scopri colonna (alternativo)** | `NVDA+Ctrl+Shift+0` | Aggira i conflitti di scorciatoie da tastiera della lingua di input di Windows |
| **Mappa cella su slot di memoria** | `[Prefix]`, poi da `Shift+1` a `Shift+9` | Assegna la cella corrente a uno slot di monitoraggio in background |
| **Leggi slot cella monitorata** | `[Prefix]`, poi da `1` a `9` | Richiama e legge il valore dello slot assegnato |
| **Salto diretto allo slot** | `Alt` + da `1` a `9` | Salta istantaneamente con il cursore a uno slot monitorato |
| **Ritorna alla cella precedente** | `[Prefix]`, poi `\` | Ti teletrasporta istantaneamente indietro dopo aver controllato uno slot |
| **Finestra di dialogo gestore slot** | `[Prefix]`, poi `Alt+M` | Apre una finestra di dialogo per visualizzare e gestire tutti i monitoraggi attivi |
| **Attiva/Disattiva monitoraggio in background** | `[Prefix]`, poi `M` | Attiva o disattiva manualmente il tracciamento del calcolo in background |
| **Cancella tutti gli slot di memoria** | `[Prefix]`, poi `Backspace` | Rimuove tutti i monitoraggi di celle salvati in background |
| **MIGLIORAMENTI PER POWERPOINT** | | |
| **Analizzatore layout diapositiva** | `[Prefix]`, poi `L` | Analizza e annuncia il layout spaziale della diapositiva corrente |
| **Analizzatore documento** | `[Prefix]`, poi `D` | Genera un sommario completo e un rapporto sullo stato di salute |
| **Organizzatore di diapositive multiplo** | `[Prefix]`, poi `X` | Apre la finestra di dialogo accessibile per riordinare più diapositive |
| **MIGLIORAMENTI PER WORD** | | |
| **Verificatore formattazione** | `[Prefix]`, poi `F` | Controlla il documento corrente per incongruenze di formattazione |
| **Analizzatore documento** | `[Prefix]`, poi `D` | Analizza il layout e la struttura del documento Word corrente |

---

## 🚀 Funzionalità

### Miglioramenti per Excel

#### 1. Analizzatore layout foglio e caching
Scansiona istantaneamente qualsiasi foglio di lavoro Excel per comprenderne la struttura, gli elementi nascosti e i blocchi di dati.
* **Come funziona:** BOA scansiona rapidamente il foglio e annuncia i blocchi di dati attivi. Ti avvisa anche di **Schede dei fogli di lavoro nascoste**, **Filtri** attivi, **Modalità protette** e **Limiti esterni nascosti** (ad esempio, se le colonne vicino al bordo destro del foglio sono nascoste, impedendoti di perdere dati fuori dallo schermo).
* **Navigazione dei dati:** Dopo la scansione, puoi utilizzare i tasti di scelta rapida per il salto del blocco di dati per spostare istantaneamente il cursore tra i blocchi di dati rilevati, bypassando senza sforzo migliaia di celle vuote.

#### 2. Organizzatore di fogli multiplo
Riordina e organizza istantaneamente più fogli contemporaneamente utilizzando una finestra di dialogo completamente accessibile.
* **Come funziona:** Apre una finestra di dialogo in cui puoi selezionare un foglio e mapparlo su una nuova posizione. Gli spostamenti pianificati sono elencati in una tabella di dati (premi `Del` per rimuovere un errore). Fai clic su `OK` e la tua cartella di lavoro verrà riorganizzata istantaneamente.

#### 3. Spostamento rapido del foglio
Sposta il foglio attivo a sinistra, a destra, all'inizio o alla fine istantaneamente utilizzando le tue scorciatoie da tastiera.

#### 4. Ridenominazione accessibile del foglio
* Durante la ridenominazione di un foglio, NVDA ha difficoltà native a leggere i caratteri che stai digitando.
* BOA inserisce una classe personalizzata `ExcelSheetRenameEdit` che utilizza il motore `SafeRichEdit`, il che significa che puoi leggere con precisione per carattere, parola o riga durante la ridenominazione. Questo funge da miglioramento al comportamento di ridenominazione predefinito esistente.

#### 5. Tracciamento di righe/colonne nascoste
* Traccia in modo proattivo i tuoi spostamenti sulla griglia per evitare che ti sfuggano dati nascosti o filtrati.
* **Celle frammentate attraversate:** Se salti attraverso una sezione della griglia fortemente frammentata o nascosta (ad esempio, passando dalla Riga 3 alla Riga 10 perché le Righe 4–9 sono nascoste), BOA annuncia esplicitamente "Righe da 4 a 9 nascoste". Questo ti assicura di sapere sempre quando i dati sono stati saltati nella struttura.

#### 6. Annunciatore formattazione condizionale
* Legge automaticamente il colore, lo stile del carattere e la sfumatura dello sfondo delle celle che sono state modificate dinamicamente dalle regole di Formattazione condizionale di Excel.
* Ti fornisce il vero stato visivo della cella anziché solo il valore grezzo sottostante. Inizialmente, quando ti sposti sulla cella, viene annunciato "ha formattazione condizionale e altri piccoli dettagli". Per informazioni dettagliate, utilizza la configurazione dei tasti di scelta rapida dettagliata che è NVDA+E e poi F.

#### 7. Miglioramento dell'annuncio della selezione
legge se la cella o l'intervallo sono selezionati o deselezionati.

#### 8 cell monitor:
* **Cell Monitor:** Utilizza i percorsi di comando per mappare celle specifiche su slot di memoria. Puoi tornare indietro e leggerle in qualsiasi momento utilizzando lo slot numerico assegnato.
* **Continuous Monitoring:** Le celle assegnate agli slot vengono monitorate automaticamente in background. Se Excel avvia un ricalcolo o una modifica della cella, BOA annuncia istantaneamente il nuovo valore. Attiva/disattiva manualmente o cancella tutto tramite gli slot di comando.
* **Excel: Cell Monitor Pro Upgrades:** 
  - **Slot Manager Dialog (`NVDA+E`, poi `Alt+M`):** Apre una finestra di dialogo che elenca tutte le celle monitorate attivamente. Premi `Enter` per saltare istantaneamente a una di esse.
  - **Warp Back (`NVDA+E`, poi `\`,):** Ti teletrasporta istantaneamente indietro alla cella di lavoro precedente dopo aver controllato uno slot.
  - **Direct Slot Jump (`Prefix + Alt` + `Slot Number`):** Salta istantaneamente a uno slot di cella assegnato bypassando il prefisso.

#### 9 Power editor
* **Excel: The Power Editor (Accessible Formula Editor):** Una svolta assoluta per la modifica di formule massicce.
  - **Single-Tap `NVDA+E`, poi `F2`:** Annuncia istantaneamente la stringa della formula grezza della cella attiva (o annuncia "Nessuna formula").
  - **Double-Tap `NVDA+E`, poi `F2`:** Apre un editor multilinea completamente accessibile per modificare in sicurezza formule massicce e nidificate. Il tasto `Enter` nativo aggiunge interruzioni di riga per una facile lettura e `Ctrl+Enter` salva le modifiche in Excel.
  - *Controlli di sicurezza:* Rileva in sicurezza gli errori di sintassi prima che corrompano il foglio e individua gli errori post-calcolo (come `#NAME?` o `#DIV/0!`) per avvisarti immediatamente se una formula si è interrotta.

#### 10 Formula auditing and evaluation enhancements:
* **Excel: Controllo e valutazione delle formule:** Aggiunte scorciatoie personalizzate (`NVDA+E`, poi `Shift+P` e `NVDA+E`, poi `Shift+D`) per tracciare in modo affidabile Precedenti e Dipendenti. Inoltre, la finestra di dialogo nativa "Valuta formula" di Excel è ora completamente accessibile; NVDA legge automaticamente i risultati valutati man mano che procedi nel calcolo!

### Miglioramenti per PowerPoint

#### 1. Selettori colore accessibili
* Sblocca la finestra di dialogo Colori personalizzati in PowerPoint.
* Identifica e legge esplicitamente in modo corretto le caselle di modifica "Red", "Green" e "Blue" (eseguendo l'override di `PowerPointRGBEdit`).
* Mappa il campo di inserimento Hex (esadecimale) precedentemente invisibile in modo che NVDA possa leggerne chiaramente l'intero valore di colore Hex.

#### 2. Supporto per la griglia dei colori standard
* La navigazione nella griglia esagonale dei colori "Standard" di PowerPoint normalmente viene letta come "Grafico" o silenzio.
* BOA traccia i tasti freccia attraverso l'esagono e recupera silenziosamente il valore del colore nascosto, annunciandotelo in tempo reale (ad es., "Color #FF0000").

#### 3 Bulk Slide Organizer:
* **PowerPoint: Bulk Slide Organizer (Experimental) (`NVDA+E`, poi `X`):** Simile alla funzionalità per Excel, ora puoi riordinare, spostare e organizzare istantaneamente più diapositive PowerPoint contemporaneamente utilizzando una finestra di dialogo completamente accessibile.

#### 4 Slide lay out analyzer
* **PowerPoint: Slide Layout Analyzer (Experimental) (`NVDA+E`, poi `L`):** Scansiona istantaneamente la diapositiva attualmente attiva per comprenderne il layout spaziale e i vincoli di accessibilità, garantendo un'esperienza del lettore di schermo fluida e reattiva. In questo modo otterrai dettagli sulla diapositiva corrente simili a quelli dell'analizzatore del layout del foglio di Excel.


#### 5 Complete Document [PPT] analyzer
* **PowerPoint: Complete Document Analyzer (Experimental) (`NVDA+E`, poi `D`):** Uno strumento di accessibilità altamente avanzato, elaborato in background, che mappa un'intera presentazione senza bloccare il motore di sintesi vocale di NVDA. Fornisce un Sommario Virtuale facilmente navigabile, rileva incongruenze nell'ordine di lettura (Ordine Visivo vs. Ordine Z), segnala le diapositive piene di testo ("Wall of Text") e mappa oggetti complessi come SmartArt e Tabelle dati.

#### 6 shape movement [adjustment] enhancements:
* **PowerPoint: Shape Movement Audio Mode (Experimental):** Introduce segnali audio spaziali 3D nell'area di lavoro di PowerPoint. Fornisce un feedback uditivo che indica la direzione e i limites di bordo di un oggetto mentre lo sposti, migliorando notevolmente la percezione dello spazio.

### Miglioramenti per Word:
#### 1. Document Analyzer inspired and derived from Paul's word access addon:
* **Word: Document Analyzer (`NVDA+E`, poi `D`):** Mostra istantaneamente una panoramica strutturale del tuo documento Word. *(Una nota speciale di ringraziamento a Paul: questa funzionalità è stata direttamente ispirata dal suo brillante add-on "Word Access". Siamo profondamente grati per il suo lavoro fondamentale in questo ambito!)*

#### 2 Formatting Auditor
* **Word: Formatting Auditor (`NVDA+E`, poi `F`):** Scansiona il documento Word per rilevare incongruenze di formattazione al fine di garantire gli standard visivi.

#### 3 Foot note reader:
* **Word: Automated Footnote Announcer:** Le note a piè di pagina verranno ora annunciate automaticamente in linea durante la lettura, in base alle impostazioni BOA personalizzate. *(Nota: il supporto per le note di chiusura e i commenti è previsto per una versione futura).*

### Infrastruttura e meccanismi tecnici

#### La modalità prefisso comando
Per prevenire conflitti di tasti con altri plug-in di NVDA, BOA utilizza una **modalità prefisso comando**:
1. Premi il tasto di attivazione per accedere alla modalità comando. Sentirai un segnale acustico acuto. Il valore predefinito è NVDA più E.
2. Premi un tasto secondario per attivare una funzionalità specifica.
3. Se premi un tasto non valido, sentirai un segnale acustico di errore.

#### Pannello di personalizzazione e impostazioni
* Le funzionalità di BOA sono completamente modulari e possono essere abilitate o disabilitate in qualsiasi momento. Vai su `Menu NVDA -> Preferenze -> Impostazioni -> BOA Office Enhancements` per attivare o disattivare le singole funzionalità.
* **Tasti di scelta rapida intelligenti:** Ogni singola impostazione presenta una scorciatoia con acceleratore `Alt+Tasto` matematicamente unica all'interno del pannello. Ad esempio, premi `Alt+E` per saltare istantaneamente al gruppo Excel, `Alt+P` per PowerPoint e `Alt+W` per Word.
* Le impostazioni sono salvate in modo sicuro in un file JSON autonomo (`boa_settings.json`), garantendo che la configurazione principale di NVDA non venga mai corrotta.
* Se Microsoft Office risolverà ufficialmente un bug di accessibilità in futuro, potrai disabilitare in sicurezza l'override specifico di BOA senza perdere le altre funzionalità dell'add-on.
* **Personalizzazione dei gesti di immissione:** Tutte le funzionalità per tutte le app di Office sono state esplicitamente esposte alla finestra di dialogo nativa dei gesti di immissione di NVDA, offrendoti la completa libertà di personalizzare ogni scorciatoia da tastiera.

#### Limiti di sicurezza e integrazione
* Le iniezioni negli appunti verificano rigorosamente gli ID dei processi delle finestre in primo piano per prevenire la fuga di dati verso altre applicazioni.
* Alcune scorciatoie da tastiera personalizzate sono completamente esposte nella finestra di dialogo dei gesti di immissione di NVDA nella categoria "Better Office Accessibility".

---

## 📋 Requisiti

* **NVDA:** versione 2026.1.0 o successiva.
* **Applicazioni:** Microsoft Excel e Microsoft PowerPoint.

---

## 💾 Installazione

1. Scarica l'ultimo file di rilascio `.nvda-addon` o trovalo nel Catalogo dei componenti aggiuntivi nativo di NVDA (NVDA Add-on Store).
2. Se installi da file, apri il file o utilizza `Store dei componenti aggiuntivi di NVDA -> Installa da file esterno`.
3. Riavvia NVDA.

---

## 🛠️ Changelog

### Versione 2.0.0
#### Nuove funzionalità
* **PowerPoint: Complete Document Analyzer (Experimental) (`NVDA+E`, poi `D`):** Uno strumento di accessibilità altamente avanzato, elaborato in background, che mappa un'intera presentazione senza bloccare il motore di sintesi vocale di NVDA. Fornisce un Sommario Virtuale facilmente navigabile, rileva incongruenze nell'ordine di lettura (Ordine Visivo vs. Ordine Z), segnala le diapositive piene di testo ("Wall of Text") e mappa oggetti complessi come SmartArt e Tabelle dati.
* **PowerPoint: Slide Layout Analyzer (Experimental) (`NVDA+E`, poi `L`):** Scansiona istantaneamente la diapositiva attualmente attiva per comprenderne il layout spaziale e i vincoli di accessibilità, garantendo un'esperienza del lettore di schermo fluida e reattiva. In questo modo otterrai dettagli sulla diapositiva corrente simili a quelli dell'analizzatore del layout del foglio di Excel.
* **PowerPoint: Bulk Slide Organizer (Experimental) (`NVDA+E`, poi `X`):** Simile alla funzionalità per Excel, ora puoi riordinare, spostare e organizzare istantaneamente più diapositive PowerPoint contemporaneamente utilizzando una finestra di dialogo completamente accessibile.
* **PowerPoint: Shape Movement Audio Mode (Experimental):** Introduce segnali audio spaziali 3D nell'area di lavoro di PowerPoint. Fornisce un feedback uditivo che indica la direzione e i limiti di bordo di un oggetto mentre lo sposti, migliorando notevolmente la percezione dello spazio. Come accennato, questa funzione è sperimentale, si attendono feedback per migliorarla.
* **Word: Formatting Auditor (`NVDA+E`, poi `F`):** Scansiona il documento Word per rilevare incongruenze di formattazione al fine di garantire gli standard visivi.
* **Word: Document Analyzer (`NVDA+E`, poi `D`):** Mostra istantaneamente una panoramica strutturale del tuo documento Word. *(Una nota speciale di ringraziamento a Paul: questa funzionalità è stata direttamente ispirata dal suo brillante add-on "Word Access". Siamo profondamente grati per il suo lavoro fondamentale in questo ambito!)*
* **Word: Automated Footnote Announcer:** Le note a piè di pagina verranno ora annunciate automaticamente in linea durante la lettura, in base alle impostazioni BOA personalizzate. *(Nota: il supporto per le note di chiusura e i commenti è previsto per una versione futura).*
* **Excel: The Power Editor (Accessible Formula Editor):** Una svolta assoluta per la modifica di formule massicce.
  - **Single-Tap `NVDA+E`, poi `F2`:** Annuncia istantaneamente la stringa della formula grezza della cella attiva (o annuncia "Nessuna formula").
  - **Double-Tap `NVDA+E`, poi `F2`:** Apre un editor multilinea completamente accessibile per modificare in sicurezza formule massicce e nidificate. Il tasto `Enter` nativo aggiunge interruzioni di riga per una facile lettura e `Ctrl+Enter` salva le modifiche in Excel.
  - *Controlli di sicurezza:* Rileva in sicurezza gli errori di sintassi prima che corrompano il foglio e individua gli errori post-calcolo (come `#NAME?` o `#DIV/0!`) per avvisarti immediatamente se una formula si è interrotta.
* **Excel: Formula Auditing & Evaluation:** Aggiunte scorciatoie personalizzate (`NVDA+E`, poi `Shift+P` e `NVDA+E`, poi `Shift+D`) per tracciare in modo affidabile Precedenti e Dipendenti. Inoltre, la finestra di dialogo nativa "Valuta formula" di Excel è ora completamente accessibile; NVDA legge automaticamente i risultati valutati man mano che procedi nel calcolo!
* **Excel: Cell Monitor Pro Upgrades:** 
  - **Slot Manager Dialog (`NVDA+E`, poi `Alt+M`):** Apre una finestra di dialogo che elenca tutte le celle monitorate attivamente. Premi `Enter` per saltare istantaneamente a una di esse.
  - **Warp Back (`NVDA+E`, poi `\`):** Ti teletrasporta istantaneamente indietro alla cella di lavoro precedente dopo aver controllato uno slot.
  - **Direct Slot Jump (`Alt` + `Slot Number`):** Salta istantaneamente a uno slot di cella assegnato bypassando completamente il prefisso.
* **Input Gestures Customization:** Tutte le funzionalità per tutte le app di Office sono state esplicitamente esposte alla finestra di dialogo nativa dei gesti di immissione di NVDA, offrendoti la completa libertà di personalizzare ogni scorciatoia da tastiera.

#### Miglioramenti all'interfaccia utente (UX/UI)
* **Unified Browseable Reports:** Abbiamo adottato un sistema unificato di rapporti HTML in tutto l'add-on. Funzionalità come l'annunciatore della formattazione condizionale di Excel, gli analizzatori di layout e gli analizzatori di documenti non si limitano più a leggere enormi blocchi di testo; i loro risultati ora si aprono in una finestra HTML nativa e navigabile, consentendoti di rivedere i dati con il tuo ritmo.
* **Excel: Enhanced Dependents/Precedents Tracking:** Migliorato notevolmente l'output vocale per le scorciatoie di tracciamento formule native di Excel (`Ctrl+[` per i Precedenti diretti e `Ctrl+]` per i Dipendenti diretti). Ora NVDA annuncerà esplicitamente quali celle sono state selezionate.
* **Excel: Merge Cell Support:** Le celle unite sono ora rilevate correttamente e annunciate esplicitamente dal tracciatore di celle che salta i vuoti.

#### Risoluzione di bug
* **Word: List Item Double Reading:** Implementata una patch temporanea per correggere il bug in cui NVDA legge due volte gli elementi dell'elenco di paragrafi in alcune visualizzazioni di Word.
* **Excel: Cell Monitor Localization Bug:** Risolti i bug di tracciamento sottostanti causati dai recenti aggiornamenti di localizzazione delle traduzioni.

### Novità nella versione 1.6.1
* **Localizzazione approfondita dei file**: Corrette le traduzioni delle stringhe mancanti all'interno dei moduli di miglioramento di Excel (come l'analizzatore del layout del foglio e lo spostamento rapido del foglio) per garantire una copertura di localizzazione al 100%.
* **Supporto di traduzione ampliato**: Aggiunte 7 nuove lingue al sistema (turco, polacco, coreano, ucraino, ceco, urdu e punjabi). 
  *(Nota: queste traduzioni sono state generate dall'IA, pertanto potrebbero essere presenti piccoli errori o imprecisioni di traduzione).*

### v1.6.0
* **Supporto completo per le traduzioni**: L'add-on è ora completamente localizzato con supporto per 17 lingue globali. 
  *(Nota: queste traduzioni sono state generate dall'IA, pertanto potrebbero essere presenti piccoli errori o imprecisioni di traduzione).*
* **Rigorosa gestione del codice**: Applicate le intestazioni di copyright GPL-2.0 a tutta la base di codice."`,

### Versione 1.5.0 
#### Nuove funzionalità
##### Radar di fine dati
Quando si naviga in fogli di calcolo di grandi dimensioni, può essere difficile capire se una cella vuota indica il raggiungimento della fine di un elenco o se c'è semplicemente un'interruzione nei dati. Il **Radar di fine dati** funge da controllo intelligente del perimetro per evitare di spostarsi alla cieca con le frecce nello spazio vuoto.
Ogni volta che ti sposti su una cella vuota, BOA scansiona istantaneamente le celle rimanenti nella tua direzione di spostamento. Se non ci sono assolutamente più dati, annuncerà in modo proattivo:
* *\"Nessun altro dato sotto\"*
* *\"Nessun altro dato sopra\"*
* *\"Nessun altro dato a destra\"*
* *\"Nessun altro dato a sinistra\"*
**Opzioni di configurazione:**
Puoi configurare questa funzionalità tramite `Preferenze di NVDA -> Impostazioni -> BOA Office Enhancements`. Poiché i fogli di calcolo possono contenere complessità nascoste (come formule invisibili o righe compresse), il radar fornisce tre modalità operative:
1. **Off**: Disabilita completamente il radar.
2. **Strict Memory Check (CountA) [Predefinito]**: L'approccio più sicuro e veloce. Controlla la memoria grezza del foglio di calcolo. Se rileva *qualsiasi cosa* sotto di te (comprese righe nascoste, testo, numeri o formule invisibili), rimane completamente silenzioso per evitare falsi allarmi. Annuncia \"Nessun altro dato\" solo quando il resto del foglio è matematicamente vuoto al 100%.
3. **Visible Data Only (Math Engine)**: Un motore altamente avanzato progettato per fogli complessi. Filtra in modo intelligente le righe nascoste e le formule invisibili (ad es. `=""`). Rimarrà in silenzio solo se rimangono numeri o testo effettivi e visibili sul tuo percorso.

### Versione 1.4 - 2026-06-12
#### Nuove funzionalità
* **Monitoraggio celle:** Utilizza i percorsi di comando per mappare celle specifiche su slot di memoria. Puoi tornare indietro e leggerle in qualsiasi momento utilizzando lo slot numerico assegnato.
* **Monitoraggio continuo:** Le celle assegnate agli slot vengono monitorate automaticamente in background. Se Excel avvia un ricalcolo o una modifica della cella, BOA annuncia istantaneamente il nuovo valore. Attiva/disattiva manualmente o cancella tutto tramite gli slot di comando.

#### Risoluzione di bug

### Versione 1.3.0 — 2026-06-05
*Rilascio finale.*

#### Nuove funzionalità
* **Sheet Layout Analyzer:** Aggiunta una potente infrastruttura di scansione del layout. Rileva istantaneamente la protezione del foglio di lavoro, i filtri di colonna attivi, le schede dei fogli di lavoro nascoste e i bordi assoluti nascosti, eseguendo al contempo il caching dei blocchi di dati rilevati.
* **Guided Data Block Navigation:** La navigazione post-analisi consente spostamenti immediati del cursore tra i principali cluster di dati, bypassando le celle vuote in modo continuo.
* **Conditional Formatting Announcer:** Rileva e legge automaticamente il colore dinamico, lo stile del carattere e la sfumatura dello sfondo delle celle modificate dalle regole di formattazione condizionale di Excel.
* **Explicit Settings Accelerators:** Riprogettata completamente la GUI delle impostazioni di BOA per conformarsi rigorosamente all'architettura di NVDA. Ogni casella di controllo delle funzionalità possiede ora una scorciatoia `Alt+Lettera` unica a livello globale, evitando il ciclo della tastiera ed eliminando gli errori di navigazione con la prima lettera.

#### Risoluzione di bug
* **Rilevamento del limite del bordo assoluto:** Sostituiti i controles nativi dei bordi COM `UsedRange` con controlli matematici unidimensionali assoluti dei limiti (`Row 1048576` e `Column 16384`) per garantire il rilevamento di righe/colonne nascoste anche se si trovano molto al di fuori del blocco di dati attivo.
* **Lazy COM Property Safe Bailouts:** Rafforzati i cicli delle proprietà COM per prevenire blocchi del thread di NVDA durante la valutazione di milioni di strutture nascoste contigue.

### Versione 1.2.0 — 2026-06-03
*Rilascio finale.*

#### Nuove funzionalità
* **Caching all'avvio dell'applicazione:** Importante revisione dell'architettura. I moduli principali vengono ora caricati in modo differito (lazy-load) esattamente quando ti sposti sulle applicazioni Office, eliminando il ritardo di avvio, risolvendo completamente il problema di messa a fuoco dell'oggetto 'sconosciuto' nelle finestre di ridenominazione e preservando la struttura della base di codice multi-file.
* **Enhanced Cell Tracker (1D COM Math):** Riscritta la logica di rilevamento delle interruzioni delle celle nascoste per valutare solo sezioni unidimensionali (`current_col` o `current_row`). Ciò ordina il carico di calcolo COM di oltre 16 milioni di celle, eliminando istantaneamente i blocchi di navigazione quando si saltano intervalli nascosti.
* **Pulizia della memoria dei processi:** Implementato il tracciamento dell'Handle della finestra di Excel (`Hwnd`) per rilevare quando l'utente chiude e riapre Excel. Questo cancella attivamente la memoria dello stato globale obsoleto e risolve completamente il falso annuncio "Foglio nascosto" all'apertura di un nuovo "Book1".

#### Risoluzione di bug
* **Annuncio di doppia selezione:** Migrazione dal non affidabile `winUser.getKeyState` asincrono e implementazione di `api.getLastInputGesture()` per sopprimere perfettamente i doppi annunci quando si utilizzano i tasti Shift+Frecce.
* **Disattivazione del rilevatore di limiti:** Il rilevatore proattivo dei limiti (Proactive Boundary Detector) è stato disattivato per proteggere la stabilità della navigazione nativa di NVDA, affidandosi interamente al tracciatore che salta i vuoti.

### Versione 1.1.0 — 2026-05-30
*Rilascio finale.*

#### Nuove funzionalità
* **GUI delle impostazioni:** Aggiunto un pannello nativo BOA Office Enhancements all'interno di `NVDA -> Preferenze -> Impostazioni` per attivare o disattivare facilmente le funzionalità.
* **SafeRichEdit Hook:** Previene crash silenziosi di NVDA durante l'interazione con i controlli RichEdit in Office 2024.
* **Scorciatoie personalizzabili:** Tutte le scorciatoie da tastiera di BOA sono ora completamente esposte nella finestra di dialogo dei gesti di immissione di NVDA sotto la categoria "Better Office Accessibility".
* **Excel: Rilevamento salto righe/colonne nascoste:** Annuncia in modo proattivo quando si naviga oltre righe o colonne nascoste, garantendo di non perdere mai i dati filtrati. Può essere attivato/disattivato nelle impostazioni.

#### Risoluzione di bug
* **Thread Safety:** Rimosso ogni ritardo bloccante (`time.sleep`) e sostituito con callback asincrone di NVDA non bloccanti per garantire che il lettore di schermo non si inceppi mai durante le operazioni in background.

### Versione 1.0.0 — 2026-05-24
*Rilascio pubblico iniziale.*

#### Nuove funzionalità
* **Excel: Organizzatore di fogli multiplo:** Riordina istantaneamente più fogli contemporaneamente utilizzando una finestra di dialogo completamente accessibile.
* **Excel: Spostamento rapido del foglio:** Sposta il foglio attivo a sinistra, a destra, all'inizio o alla fine tramite comandi da tastiera.
* **Excel: Ridenominazione accessibile del foglio:** Intercetta il campo di ridenominazione nativo non accessibile e lo sostituisce con una finestra di dialogo accessibile affidabile.
* **Excel: Tracciamento intelligente della selezione:** Annuncia accuratamente le selezioni e le deselezioni di intervalli multi-cella.
* **PowerPoint: Selettori colore accessibili:** Consente a NVDA di leggere accuratamente i valori RGB ed esadecimali (Hex) all'interno della finestra di dialogo Colori personalizzati.
* **PowerPoint: Supporto per la griglia dei colori standard:** Intercetta la navigazione con i tasti freccia per leggere i codici Hex nascosti dalla griglia esagonale dei colori non accessibile.
