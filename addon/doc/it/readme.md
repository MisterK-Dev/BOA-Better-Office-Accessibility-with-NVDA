# BOA: Migliore Accessibilità per Office (Better Office Accessibility)

BOA è una potente suite di miglioramenti dell'accessibilità per Microsoft Office, progettata per migliorare notevolmente l'esperienza dello screen reader per gli utenti di NVDA. Applica correzioni dirette ai componenti dell'interfaccia utente inaccessibili e introduce strumenti di navigazione rapida per Excel e PowerPoint.

---

## ⌨️ Riferimento dei tasti di scelta rapida

| Funzionalità | Combinazione di tasti | Contesto / Note |
| :--- | :--- | :--- |
| **Entra in Modalità Comando** | `NVDA+E` | Attiva la Modalità Prefisso Comando (emette un segnale acustico acuto) |
| **Analizza il layout del foglio** | `NVDA+E`, poi `L` | Esegui all'interno di Excel prima di navigare nei blocchi dati |
| **Vai al Blocco Dati più vicino** | `NVDA+E`, poi `J` /  | Richiede prima l'Analisi del Layout |
| **Apri l'Organizzatore Fogli in Blocco** | `NVDA+E`, poi `X` | Apre la finestra di dialogo accessibile per il riordino dei fogli |
| **Sposta il foglio attivo a sinistra** | `NVDA+Maiusc+FrecciaSinistra` | Sposta il foglio attivo di una posizione verso l'alto|
| **Sposta il foglio attivo a destra** | `NVDA+Maiusc+FrecciaDestra` | Sposta il foglio di lavoro attivo di una posizione verso il basso|
| **Sposta il foglio all'inizio/fine** | `NVDA+Maiusc+Home` / `Fine` | Invia il foglio di lavoro ai limiti assoluti |
| **Formattazione condizionale dettagliata**| `NVDA+E`, poi `F` | Annuncia i dettagli completi di formattazione della cella focalizzata |
| **Mappa la cella a uno slot di memoria** | `NVDA+E`, poi `Maiusc+1` a `Maiusc+9` | Assegna la cella corrente a uno slot del monitor in background |
| **Leggi lo slot della cella monitorata** | `NVDA+E`, poi `1` a `9` | Richiama e legge il valore dello slot assegnato |
| **Attiva/Disattiva monitoraggio in background** | `NVDA+E`, poi `M` | Attiva/disattiva manualmente il tracciamento dei ricalcoli in background |
| **Cancella tutti gli slot di memoria** | `NVDA+E`, poi `Backspace` | Elimina tutti i monitoraggi delle celle in background salvati |
| **Esci dalla Modalità Comando** | `Esc` | Esce dalla Modalità Prefisso Comando |

---

## 🚀 Funzionalità

### Miglioramenti per Excel

#### 1. Analizzatore del Layout del Foglio e Cache
Analizza istantaneamente qualsiasi foglio di lavoro di Excel per comprenderne la struttura, gli elementi nascosti e i blocchi dati.
* **Come funziona:** BOA analizza rapidamente il foglio e annuncia i blocchi dati attivi. Ti avverte anche riguardo le **Schede del Foglio di Lavoro Nascoste**, **Filtri** attivi, **Modalità Protette** e **Bordi Esterni Nascosti** (ad esempio, se le colonne vicino al bordo destro del foglio sono nascoste, impedendoti di perdere dati fuori dallo schermo).
* **Navigazione Dati:** Dopo la scansione, puoi utilizzare i tasti di scelta rapida di salto del blocco dati per teletrasportare istantaneamente il cursore tra i blocchi dati scoperti, aggirando senza sforzo migliaia di celle vuote.

#### 2. Organizzatore Fogli in Blocco
Riordina e organizza istantaneamente più fogli contemporaneamente utilizzando una finestra di dialogo completamente accessibile.
* **Come funziona:** Apre una finestra di dialogo in cui puoi selezionare un foglio e mapparlo in una nuova posizione. Gli spostamenti pianificati sono elencati in una tabella dati (premi `Canc` per rimuovere un errore). Fai clic su `OK` e la tua cartella di lavoro verrà riorganizzata istantaneamente.

#### 3. Spostamento Rapido del Foglio
Sposta il foglio attivo a sinistra, a destra, all'inizio o alla fine in modo istantaneo utilizzando le scorciatoie da tastiera.

#### 4. Rinomina Accessibile del Foglio
* Quando si rinomina un foglio, NVDA ha difficoltà in modo nativo a leggere i caratteri che stai digitando.
* BOA inietta una classe `ExcelSheetRenameEdit` personalizzata che utilizza il motore `SafeRichEdit`, il che significa che puoi leggere in modo preciso carattere per carattere, parola per parola o per riga durante la rinomina. Questo serve come miglioramento del comportamento di rinomina predefinito esistente.

#### 5. Rilevatore di Righe/Colonne Nascoste
* Traccia proattivamente il tuo movimento sulla griglia per impedirti di perdere dati nascosti o filtrati.
* **Celle Frammentate Attraversate:** Se salti attraverso una sezione fortemente frammentata o nascosta della griglia (ad esempio, passando dalla Riga 3 alla Riga 10 perché le Righe 4–9 sono nascoste), BOA annuncia esplicitamente "Righe dalla 4 alla 9 nascoste". Questo ti assicura di sapere sempre quando dei dati sono stati saltati nella struttura.

#### 6. Annunciatore di Formattazione Condizionale
* Legge automaticamente il colore, lo stile del carattere e la sfumatura di sfondo delle celle che sono state modificate dinamicamente dalle regole di formattazione condizionale di Excel.
* Ti fornisce il vero stato visivo della cella piuttosto che solo il suo valore grezzo sottostante. Inizialmente, quando ti focalizzi sulla cella, annuncia "ha formattazione condizionale e alcuni altri dettagli minori". Per informazioni complete, usa la configurazione dettagliata dei tasti di scelta rapida che è NVDA E e F.

#### 7. Migliore annuncio della selezione
legge se una cella o un intervallo è selezionato o deselezionato.

#### 8 Monitor delle celle:
* **Monitor delle Celle:** Usa i percorsi di comando per mappare celle specifiche agli slot di memoria. Puoi tornare indietro e leggerle in qualsiasi momento utilizzando lo slot numerico assegnato.
* **Monitoraggio Continuo:** Le celle assegnate agli slot vengono automaticamente monitorate in background. Se Excel attiva un ricalcolo o la modifica di una cella, BOA annuncia istantaneamente il nuovo valore. Attivalo manualmente o cancella tutto tramite gli slot di comando.

### Miglioramenti per PowerPoint

#### 1. Selettori Colore Accessibili
* Sblocca la finestra di dialogo Colori personalizzati in PowerPoint.
* Identifica e legge esplicitamente le caselle di modifica "Rosso", "Verde" e "Blu" correttamente (sovrascrivendo `PowerPointRGBEdit`).
* Mappa il campo di input Esadecimale in precedenza invisibile, in modo che NVDA possa leggere il valore esadecimale completo in modo chiaro.

#### 2. Supporto per la Griglia dei Colori Standard
* Navigando la griglia esagonale dei colori "Standard" di PowerPoint, questa viene normalmente letta come "Grafico" o col silenzio.
* BOA traccia i tasti freccia attraverso l'esagono e recupera silenziosamente il valore esadecimale nascosto del colore, annunciandolo in tempo reale (ad esempio, "Colore #FF0000").

### Infrastruttura e Meccanismi Tecnici

#### La Modalità Prefisso Comando
Per prevenire conflitti di sequenze di tasti con altri plugin NVDA, BOA utilizza una **Modalità Prefisso Comando**:
1. Premi il tasto di attivazione per entrare in Modalità Comando. Sentirai un segnale acustico acuto.
2. Premi un tasto secondario per attivare una funzione specifica.
3. Se premi un tasto non valido, sentirai un segnale acustico di errore.

#### Personalizzazione e Pannello Impostazioni
* Le funzionalità di BOA sono completamente modulari e possono essere abilitate o disabilitate in qualsiasi momento. Vai su `Menu NVDA -> Preferenze -> Impostazioni -> Miglioramenti per Office BOA` per attivare o disattivare le singole funzionalità.
* **Tasti Acceleratori Intelligenti:** Ogni singola impostazione ha una scorciatoia univoca matematicamente testata `Alt+Tasto` all'interno del pannello. Ad esempio, premi `Alt+E` per passare istantaneamente al gruppo Excel, `Alt+P` per PowerPoint e `Alt+W` per Word.
* Le impostazioni sono salvate in modo sicuro in un file JSON autonomo (`boa_settings.json`), garantendo che la tua configurazione principale di NVDA non venga mai corrotta.
* Se Microsoft Office corregge ufficialmente un bug di accessibilità in futuro, puoi disabilitare in sicurezza il blocco di override specifico di BOA senza perdere il resto delle funzionalità del componente aggiuntivo.

#### Sicurezza e Limiti di Integrazione
* Le iniezioni negli appunti verificano rigorosamente gli ID del processo in primo piano della finestra per evitare la fuga di dati verso altre applicazioni.
* Alcuni tasti di scelta rapida personalizzati sono completamente esposti nella finestra di dialogo Gesti di immissione di NVDA sotto la categoria "Miglioramenti per l'Accessibilità in Office" (Better Office Accessibility).

---

## 📋 Requisiti

* **NVDA:** Versione 2026.1.0 o successive.
* **Applicazioni:** Microsoft Excel & Microsoft PowerPoint.

---

## 💾 Installazione

1. Scarica l'ultimo file di rilascio `.nvda-addon` o cercalo nel Negozio Componenti aggiuntivi nativo di NVDA.
2. Se lo installi da file, apri il file o usa `NVDA -> Negozio Componenti aggiuntivi -> Installa da file esterno`.
3. Riavvia NVDA.

---

## 🛠️ Log delle modifiche

### Novità della v1.6.1
* **Localizzazione approfondita dei file**: Sono state corrette le traduzioni mancanti delle stringhe all'interno dei moduli di potenziamento di Excel (come l'Analizzatore layout del foglio e lo Spostamento rapido del foglio) per garantire una copertura della localizzazione del 100%.
* **Supporto alle traduzioni espanso**: Sono state aggiunte 7 nuove lingue al sistema (Turco, Polacco, Coreano, Ucraino, Ceco, Urdu e Punjabi).
  *(Nota: Queste traduzioni sono state generate da un'IA, quindi potrebbero essere presenti alcuni lievi errori o imprecisioni nella traduzione.)*

### v1.6.0
* **Supporto completo per le traduzioni**: Il componente aggiuntivo è ora completamente localizzato con supporto per 17 lingue globali. 
  *(Nota: Queste traduzioni sono state generate dall'IA, pertanto potrebbero essere presenti alcuni errori o imprecisioni minori nelle traduzioni.)*
* **Gestione rigorosa del codice**: Applicati gli header di copyright GPL-2.0 in tutto il codice base.

### Versione 1.5.0 
#### Nuove Funzionalità
##### Radar di Fine Dati
Quando si naviga in grandi fogli di calcolo, può essere difficile capire se una cella vuota significa che si è giunti alla fine di una lista, o se c'è semplicemente un vuoto nei dati. Il **Radar di Fine Dati** agisce come un controllo perimetrale intelligente per salvarti dal navigare alla cieca attraverso lo spazio vuoto.
Ogni volta che navighi in una cella vuota, BOA analizza istantaneamente le celle rimanenti nella tua direzione di viaggio. Se non c'è assolutamente alcun dato rimasto, annuncerà proattivamente:
* *"Nessun dato in basso"*
* *"Nessun dato in alto"*
* *"Nessun dato a destra"*
* *"Nessun dato a sinistra"*
**Opzioni di configurazione:**
Puoi configurare questa funzionalità tramite `Preferenze NVDA -> Impostazioni -> Miglioramenti per Office BOA`. Poiché i fogli di calcolo possono contenere complessità nascoste (come formule invisibili o righe compresse), il radar offre tre modalità operative:
1. **Disattivato**: Disattiva completamente il radar.
2. **Controllo rigoroso della memoria (Conta.Valori) [Predefinito]**: L'approccio più sicuro e veloce. Controlla la memoria grezza del foglio di calcolo. Se rileva *qualsiasi* cosa sotto di te (incluse righe nascoste, testo, numeri o formule invisibili), rimane completamente in silenzio per evitare falsi allarmi. Annuncia "Nessun dato" solo quando il resto del foglio è matematicamente vuoto al 100%.
3. **Solo dati visibili (Motore matematico)**: Un motore altamente avanzato progettato per fogli complessi. Filtra in modo intelligente righe nascoste e formule invisibili (es., `=""`). Rimarrà in silenzio solo se ci sono numeri o testo visibili rimasti lungo il tuo percorso.

### Versione 1.4 - 2026-06-12
#### Nuove Funzionalità
* **Monitor delle celle:** Usa i percorsi di comando per mappare celle specifiche agli slot di memoria. Puoi tornare indietro e leggerle in qualsiasi momento utilizzando lo slot numerico assegnato.
* **Monitoraggio Continuo:** Le celle assegnate agli slot vengono automaticamente monitorate in background. Se Excel attiva un ricalcolo o la modifica di una cella, BOA annuncia istantaneamente il nuovo valore. Attivalo manualmente o cancella tutto tramite gli slot di comando.

#### Correzioni di bug

### Versione 1.3.0 — 2026-06-05
*Rilascio finale.*

#### Nuove Funzionalità
* **Analizzatore del layout del foglio:** Aggiunta di una potente infrastruttura di scansione del layout. Rileva istantaneamente la Protezione del Foglio di Lavoro, Filtri di Colonna attivi, Schede del Foglio Nascoste e bordi assoluti nascosti mentre memorizza in cache i blocchi dati scoperti.
* **Navigazione Guidata per Blocchi Dati:** La navigazione successiva all'analisi consente il teletrasporto immediato del cursore tra i principali cluster di dati, bypassando le celle vuote senza soluzione di continuità.
* **Annunciatore di formattazione condizionale:** Rileva e legge automaticamente il colore dinamico, lo stile del carattere e la sfumatura di sfondo delle celle alterate dalle regole di Formattazione Condizionale di Excel.
* **Acceleratori espliciti per le impostazioni:** Revisione completa della GUI delle Impostazioni di BOA per rispettare rigorosamente l'architettura di NVDA. Ogni casella di controllo delle funzionalità ora possiede una scorciatoia `Alt+Lettera` unica a livello globale, impedendo il ciclo della tastiera ed eliminando gli errori di navigazione tramite la prima lettera.

#### Correzioni di bug
* **Rilevamento del Limite del Bordo Assoluto:** Sostituiti i controlli nativi dei bordi COM `UsedRange` con controlli matematici dei limiti 1D assoluti (`Riga 1048576` e `Colonna 16384`) per garantire il rilevamento di righe/colonne nascoste anche se si trovano molto al di fuori del blocco dati attivo.
* **Uscite Sicure con Caricamento Ritardato delle Proprietà COM:** Rinforzati i cicli delle proprietà COM per impedire i blocchi del thread di NVDA quando si valutano milioni di strutture nascoste contigue.

### Versione 1.2.0 — 2026-06-03
*Rilascio finale.*

#### Nuove Funzionalità
* **Cache per l'Avvio delle App:** Importante revisione architettonica. I moduli principali ora vengono caricati in modo differito esattamente quando ti focalizzi sulle applicazioni di Office, eliminando il ritardo di avvio, risolvendo completamente il problema tecnico di focus sugli oggetti 'sconosciuti' nelle finestre di rinomina e preservando la struttura del codice base a più file.
* **Rilevatore Avanzato delle Celle (Matematica COM 1D):** Riscritta la logica di rilevamento degli spazi vuoti nelle celle nascoste per valutare solo sezioni trasversali unidimensionali (`current_col` o `current_row`). Ciò riduce il carico di calcolo COM di oltre 16 milioni di celle, eliminando istantaneamente i blocchi di navigazione quando si saltano intervalli nascosti.
* **Cancellazione della Memoria dei Processi:** Implementato il tracciamento dell'Handle della Finestra di Excel (`Hwnd`) per rilevare quando l'utente chiude e riapre Excel. Ciò elimina attivamente la memoria di stato globale obsoleta e risolve completamente l'annuncio falso di "Foglio nascosto" all'apertura di un nuovo "Cartel1".

#### Correzioni di bug
* **Annuncio Doppia Selezione:** Abbandonato il poco affidabile e asincrono `winUser.getKeyState` e implementato `api.getLastInputGesture()` per sopprimere perfettamente i doppi annunci quando si utilizzano i tasti Maiusc+Freccia.
* **Disattivazione del Rilevatore di Limiti:** Il Rilevatore Proattivo di Limiti è stato disattivato per proteggere la stabilità della navigazione nativa di NVDA, ripiegando interamente sul tracciatore di salto dei vuoti.

### Versione 1.1.0 — 2026-05-30
*Rilascio finale.*

#### Nuove Funzionalità
* **GUI delle Impostazioni:** Aggiunto un pannello nativo per i Miglioramenti per Office BOA all'interno di `NVDA -> Preferenze -> Impostazioni` per attivare o disattivare facilmente le funzionalità.
* **Hook SafeRichEdit:** Previene gli arresti anomali silenziosi di NVDA durante l'interazione con i controlli RichEdit in Office 2024.
* **Tasti di scelta rapida personalizzabili:** Tutte le scorciatoie di BOA sono ora completamente esposte nella finestra di dialogo Gesti di immissione di NVDA sotto la categoria "Miglioramenti per l'Accessibilità in Office" (Better Office Accessibility).
* **Excel: Rilevamento di Salto Riga/Colonna Nascosta:** Annuncia in modo proattivo durante la navigazione oltre righe o colonne nascoste, garantendo di non perdere mai i dati filtrati. Può essere attivato o disattivato nelle impostazioni.

#### Correzioni di bug
* **Sicurezza del Thread:** Rimossi tutti i ritardi bloccanti (`time.sleep`) e sostituiti con callback asincroni non bloccanti di NVDA per garantire che lo screen reader non subisca interruzioni durante le operazioni in background.

### Versione 1.0.0 — 2026-05-24
*Rilascio pubblico iniziale.*

#### Nuove Funzionalità
* **Excel: Organizzatore Fogli in Blocco:** Riordina istantaneamente più fogli contemporaneamente utilizzando una finestra di dialogo completamente accessibile.
* **Excel: Spostamento Rapido del Foglio:** Sposta il foglio attivo a sinistra, a destra, all'inizio o alla fine tramite comandi da tastiera.
* **Excel: Rinomina Accessibile del Foglio:** Intercetta l'inaccessibile campo di rinomina nativo e lo sostituisce con una finestra di dialogo accessibile e affidabile.
* **Excel: Tracciamento Intelligente della Selezione:** Annuncia con precisione le selezioni e deselezioni di intervalli di celle multiple.
* **PowerPoint: Selettori Colore Accessibili:** Consente a NVDA di leggere con precisione i valori RGB ed esadecimali (Hex) all'interno della finestra di dialogo Colori personalizzati.
* **PowerPoint: Supporto per la Griglia dei Colori Standard:** Intercetta la navigazione con i tasti freccia per leggere i codici esadecimali nascosti dalla griglia esagonale dei colori inaccessibile.
