# BOA: Better Office Accessibility

BOA ist eine leistungsstarke Sammlung von Barrierefreiheits-Erweiterungen für Microsoft Office, die entwickelt wurde, um die Nutzung von Bildschirmleseprogrammen (Screenreadern) für NVDA-Nutzer erheblich zu verbessern. Es korrigiert unzugängliche UI-Komponenten direkt und führt schnelle Navigationswerkzeuge für Excel und PowerPoint ein.

---

## ⌨️ Hotkey-Referenz

| Funktion | Tastenkombination | Kontext / Hinweise |
| :--- | :--- | :--- |
| **Befehlsmodus aufrufen** | `[Prefix]` (Standard: `NVDA+E`) | Aktiviert den Befehls-Präfixmodus (löst einen hohen Signalton aus) |
| **Befehlsmodus abbrechen** | `Escape` | Verlässt den Befehls-Präfixmodus |
| **EXCEL-ERWEITERUNGEN** | | |
| **Tabellenlayout analysieren** | `[Prefix]`, dann `L` | In Excel ausführen, bevor in Datenblöcken navigiert wird |
| **Zum nächsten Datenblock springen** | `[Prefix]`, dann `J` | Erfordert vorherige Layout-Analyse |
| **Sammel-Tabellenorganisator öffnen** | `[Prefix]`, dann `X` | Öffnet den barrierefreien Dialog zum Neuanordnen von Tabellenblättern |
| **Rohformel-Ansage** | `[Prefix]`, dann `F2` | Einmal drücken, um die rohe Formelzeichenkette zu hören |
| **Power-Formeleditor** | `[Prefix]`, dann zweimal `F2` | Zweimal drücken, um den barrierefreien mehrzeiligen Formeleditor zu öffnen |
| **Spur zum Vorgänger** | `[Prefix]`, dann `Shift+P` | Die gleiche Funktion „Spur zum Vorgänger“ in barrierefreier Form. |
| **Spur zum Nachfolger** | `[Prefix]`, dann `Shift+D` | Die gleiche Funktion „Spur zum Nachfolger“ in barrierefreier Form. Das Drücken von der Eingabetaste (Enter) auf einer Zelle teleportiert Sie dorthin. |
| **Detaillierte bedingte Formatierung**| `[Prefix]`, dann `F` | Kündigt die vollständigen Formatierungsdetails der fokussierten Zelle an |
| **Aktives Blatt nach links verschieben** | `NVDA+Shift+LeftArrow` | Verschiebt das aktive Blatt um eine Position nach oben |
| **Aktives Blatt nach rechts verschieben** | `NVDA+Shift+RightArrow` | Verschiebt das aktive Arbeitsblatt um eine Position nach unten |
| **Blatt zum Anfang/Ende verschieben** | `NVDA+Shift+Home` / `End` | Sendet das Arbeitsblatt an die absoluten Grenzen |
| **Zeile ausblenden / einblenden** | `Ctrl+9` / `Ctrl+Shift+9` | Nativer Shortcut; BOA kündigt die Sichtbarkeitsänderung explizit an |
| **Spalte ausblenden / einblenden** | `Ctrl+0` / `Ctrl+Shift+0` | Nativer Shortcut; BOA kündigt die Sichtbarkeitsänderung explizit an |
| **Spalte einblenden (Fallback)** | `NVDA+Ctrl+Shift+0` | Umgeht Konflikte bei Windows-Eingabesprachen-Tastenkombinationen |
| **Zelle einem Speicherplatz zuweisen** | `[Prefix]`, dann `Shift+1` bis `Shift+9` | Weist die aktuelle Zelle einem Hintergrund-Überwachungsplatz zu |
| **Überwachten Zellenplatz auslesen** | `[Prefix]`, dann `1` bis `9` | Ruft den Wert des zugewiesenen Platzes ab und liest ihn vor |
| **Direkter Platzsprung** | `Alt` + `1` bis `9` | Springt mit dem Cursor sofort zu einem überwachten Platz |
| **Zurück zur vorherigen Zelle springen** | `[Prefix]`, dann `\` | Teleportiert Sie sofort zurück, nachdem Sie einen Platz überprüft haben |
| **Platz-Manager-Dialog** | `[Prefix]`, dann `Alt+M` | Öffnet einen Dialog zum Anzeigen und Verwalten aller aktiven Überwachungen |
| **Hintergrundüberwachung umschalten** | `[Prefix]`, dann `M` | Schaltet die Verfolgung von Hintergrundberechnungen manuell um |
| **Alle Speicherplätze löschen** | `[Prefix]`, dann `Backspace` | Löscht alle gespeicherten Hintergrund-Zellüberwachungen |
| **POWERPOINT-ERWEITERUNGEN** | | |
| **Folienlayout-Analyse** | `[Prefix]`, dann `L` | Analysiert und kündigt das räumliche Layout der aktuellen Folie an |
| **Dokumenten-Analyse** | `[Prefix]`, dann `D` | Erstellt ein umfassendes Inhaltsverzeichnis und einen Integritätsbericht |
| **Sammel-Folienorganisator** | `[Prefix]`, dann `X` | Öffnet den barrierefreien Dialog zum Neuanordnen mehrerer Folien |
| **WORD-ERWEITERUNGEN** | | |
| **Formatierungs-Auditor** | `[Prefix]`, dann `F` | Überprüft das aktuelle Dokument auf Formatierungsinkonsistenzen |
| **Dokumenten-Analyse** | `[Prefix]`, dann `D` | Analysiert das Layout und die Struktur des aktuellen Word-Dokuments |

---

## 🚀 Funktionen

### Excel-Erweiterungen

#### 1. Tabellenlayout-Analyse & Caching
Scannen Sie jedes Excel-Arbeitsblatt sofort, um dessen Struktur, ausgeblendete Elemente und Datenblöcke zu verstehen.
* **Funktionsweise:** BOA scannt das Blatt schnell und kündigt aktive Datenblöcke an. Es warnt Sie auch vor **ausgeblendeten Arbeitsblatt-Reitern**, aktiven **Filtern**, **geschützten Modi** und **ausgeblendeten äußeren Grenzen** (z. B. wenn Spalten nahe dem rechten Rand des Blattes ausgeblendet sind, damit Sie keine Daten außerhalb des sichtbaren Bereichs verpassen).
* **Daten-Navigation:** Nach dem Scannen können Sie die Tastenkombinationen zum Springen zwischen Datenblöcken verwenden, um Ihren Cursor sofort zwischen den erkannten Datenblöcken zu bewegen, wodurch Sie mühelos Tausende von leeren Zellen umgehen.

#### 2. Sammel-Tabellenorganisator
Ordnen und sortieren Sie mehrere Tabellenblätter gleichzeitig mithilfe eines vollständig barrierefreien Dialogs neu.
* **Funktionsweise:** Öffnet einen Dialog, in dem Sie ein Blatt auswählen und ihm eine neue Position zuweisen können. Geplante Verschiebungen werden in einer Datentabelle aufgelistet (drücken Sie `Del`, um einen Fehler zu entfernen). Klicken Sie auf `OK` und Ihre Arbeitsmappe wird sofort neu angeordnet.

#### 3. Schneller Blatt-Verschieber
Verschieben Sie das aktive Blatt mit Ihren Tastenkombinationen sofort nach links, rechts, ganz an den Anfang oder ganz ans Ende.

#### 4. Barrierefreie Tabellenblatt-Umbenennung
* Beim Umbenennen eines Blattes hat NVDA von Natur aus Schwierigkeiten, die eingegebenen Zeichen vorzulesen.
* BOA fügt eine benutzerdefinierte `ExcelSheetRenameEdit`-Klasse ein, die die `SafeRichEdit`-Engine verwendet. Das bedeutet, dass Sie beim Umbenennen präzise nach Zeichen, Wort oder Zeile lesen können. Dies dient als Verbesserung des bestehenden Standard-Umbenennungsverhaltens.

#### 5. Verfolgung ausgeblendeter Zeilen/Spalten
* Verfolgt proaktiv Ihre Bewegung im Gitternetz, um zu verhindern, dass Sie ausgeblendete oder gefilterte Daten übersehen.
* **Übersprungene fragmentierte Zellen:** Wenn Sie über einen stark fragmentierten oder ausgeblendeten Bereich des Gitters springen (z. B. von Zeile 3 zu Zeile 10, weil die Zeilen 4–9 ausgeblendet sind), kündigt BOA explizit „Zeilen 4 bis 9 ausgeblendet“ an. So wissen Sie immer, wenn Daten in der Struktur übersprungen wurden.

#### 6. Ansage bedingter Formatierung
* Liest automatisch die Farbe, den Schriftstil und den Hintergrundschatten von Zellen vor, die durch die Regeln der bedingten Formatierung von Excel dynamisch geändert wurden.
* Gibt Ihnen den tatsächlichen visuellen Zustand der Zelle wieder und nicht nur den rohen zugrunde liegenden Wert. Beim Fokussieren der Zelle wird anfangs „hat bedingte Formatierung und einige andere kleinere Details“ angekündigt. Für umfassende Informationen verwenden Sie die detaillierte Tastenkombination `NVDA+E` und dann `F`.

#### 7. Verbesserte Auswahlankündigung
Liest vor, ob eine Zelle oder ein Bereich ausgewählt oder abgewählt wurde.

#### 8. Zellüberwachung:
* **Zellüberwachung:** Verwenden Sie Befehlspfade, um bestimmte Zellen Speicherplätzen zuzuweisen. Sie können jederzeit dorthin zurückspringen und sie über den zugewiesenen numerischen Platz auslesen.
* **Kontinuierliche Überwachung:** Zugewiesene Zellen werden im Hintergrund automatisch überwacht. Wenn Excel eine Neuberechnung oder Zellbearbeitung auslöst, kündigt BOA sofort den neuen Wert an. Manuell umschalten oder alle über Befehlsplätze löschen.
* **Excel: Upgrades für die Zellüberwachung Pro:** 
  - **Platz-Manager-Dialog (`NVDA+E`, dann `Alt+M`):** Öffnet einen Dialog, in dem alle aktiv überwachten Zellen aufgelistet sind. Drücken Sie `Enter`, um sofort zu einer Zelle zu springen.
  - **Zurückspringen (`NVDA+E`, dann `\`):** Teleportiert Sie sofort zurück zu Ihrer vorherigen Arbeitszelle, nachdem Sie einen Platz überprüft haben.
  - **Direkter Platzsprung (`Prefix + Alt` + `Slot Number`):** Umgeht das Präfix vollständig und springt sofort zu einem zugewiesenen Zellplatz.

#### 9. Power-Editor
* **Excel: Der Power-Editor (Barrierefreier Formeleditor):** Ein absoluter Meilenstein für das Ändern gewaltiger Formeln.
  - **Einmaliges Drücken von `NVDA+E`, dann `F2`:** Kündigt sofort die rohe Formelzeichenkette der aktiven Zelle an (oder kündigt „Keine Formel“ an).
  - **Zweimaliges Drücken von `NVDA+E`, dann `F2`:** Öffnet einen vollständig barrierefreien, mehrzeiligen Editor, um riesige, verschachtelte Formeln sicher zu bearbeiten. Ein natives `Enter` fügt Zeilenumbrüche für ein einfaches Lesen ein, und `Ctrl+Enter` speichert die Formel zurück in Excel.
  - *Sicherheitsprüfungen:* Fängt Syntaxfehler sicher ab, bevor sie Ihr Tabellenblatt beschädigen, und erkennt Fehler nach der Berechnung (wie `#NAME?` oder `#DIV/0!`), um Sie sofort zu warnen, wenn eine Formel fehlerhaft ist.

#### 10. Formelüberwachung & Verbesserungen bei der Auswertung:
* **Excel: Formelüberwachung & Auswertung:** Eigene Tastenkombinationen hinzugefügt (`NVDA+E`, dann `Shift+P` und `NVDA+E`, dann `Shift+D`), um Spur zum Vorgänger und Spur zum Nachfolger zuverlässig zu verfolgen. Darüber hinaus ist der native Excel-Dialog „Formelauswertung“ nun vollständig barrierefrei; NVDA liest die ausgewerteten Ergebnisse automatisch vor, während Sie die Berechnung schrittweise durchgehen!

### PowerPoint-Erweiterungen

#### 1. Barrierefreie Farbwähler
* Schaltet den Dialog für benutzerdefinierte Farben in PowerPoint frei.
* Identifiziert und liest die Eingabefelder „Rot“, „Grün“ und „Blau“ korrekt vor (durch Überschreiben von `PowerPointRGBEdit`).
* Ordnet das zuvor unsichtbare Hex-Eingabefeld so zu, dass NVDA den vollständigen Hex-Farbwert sauber auslesen kann.

#### 2. Unterstützung für das Standard-Farbraster
* Das Navigieren im PowerPoint-Standardfarbraster (Sechseck-Raster) wird normalerweise als „Grafik“ oder gar nicht vorgelesen.
* BOA verfolgt Ihre Pfeiltasten über das Sechseck, ruft den ausgeblendeten Farbwert im Hintergrund ab und kündigt ihn in Echtzeit an (z. B. „Farbe #FF0000“).

#### 3. Sammel-Folienorganisator:
* **PowerPoint: Sammel-Folienorganisator (Experimentell) (`NVDA+E`, dann `X`):** Ähnlich wie bei der Excel-Funktion können Sie jetzt mehrere PowerPoint-Folien gleichzeitig über einen vollständig barrierefreien Dialog neu anordnen, verschieben und organisieren.

#### 4. Folienlayout-Analyse
* **PowerPoint: Folienlayout-Analyse (Experimentell) (`NVDA+E`, dann `L`):** Scannt Ihre aktuell aktive Folie sofort, um deren räumliches Layout und Barrierefreiheitsbeschränkungen zu verstehen, was eine völlig reibungslose und reaktionsschnelle Bildschirmausgabe ermöglicht. Das bedeutet, dass Sie hier ähnliche Details über die aktuelle Folie erhalten wie bei der Tabellenlayout-Analyse von Excel.

#### 5. Komplette Dokumenten-Analyse [PPT]
* **PowerPoint: Komplette Dokumenten-Analyse (Experimentell) (`NVDA+E`, dann `D`):** Ein hochmodernes, im Hintergrund verarbeitetes Barrierefreiheitswerkzeug, das eine gesamte Präsentation abbildet, ohne die Sprachausgabe von NVDA einzufrieren. Es bietet ein tief navigierbares virtuelles Inhaltsverzeichnis, erkennt Abweichungen in der Lesereihenfolge (visuelle Reihenfolge vs. Z-Reihenfolge), kennzeichnet Folien mit zu viel Text („Wall of Text“) und erfasst komplexe Objekte wie SmartArt und Datentabellen.

#### 6. Verbesserungen der Formbewegung [Anpassung]:
* **PowerPoint: Audio-Modus für Formbewegung (Experimentell):** Führt räumliche 3D-Audio-Hinweise auf der PowerPoint-Arbeitsfläche ein. Bietet akustische Rückmeldungen, die die Richtung und die Begrenzungslinien eines Objekts anzeigen, während Sie es bewegen, was das räumliche Bewusstsein erheblich verbessert.

### Word-Erweiterungen:

#### 1. Dokumenten-Analyse (inspiriert und abgeleitet von Pauls „Word Access“-Add-on):
* **Word: Dokumenten-Analyse (`NVDA+E`, dann `D`):** Ruft sofort eine strukturelle Übersicht Ihres Word-Dokuments auf. *(Ein besonderer Dank geht an Paul: Diese Funktion wurde direkt von seinem brillanten „Word Access“-Add-on inspiriert. Wir sind sehr dankbar für seine bahnbrechende Arbeit in diesem Bereich!)*

#### 2. Formatierungs-Auditor
* **Word: Formatierungs-Auditor (`NVDA+E`, dann `F`):** Scannt Ihr Word-Dokument auf Formatierungsinkonsistenzen, um visuelle Standards zu gewährleisten.

#### 3. Fußnoten-Leser:
* **Word: Automatische Fußnotenansage:** Fußnoten werden nun beim Lesen automatisch im Text angekündigt, abhängig von Ihren benutzerdefinierten BOA-Einstellungen. *(Hinweis: Die Unterstützung für Endnoten und Kommentare ist für eine zukünftige Version geplant).*

### Infrastruktur & technische Mechanismen

#### Der Befehls-Präfixmodus
Um Konflikte bei Tastenkombinationen mit anderen NVDA-Add-ons zu vermeiden, verwendet BOA einen **Befehls-Präfixmodus**:
1. Drücken Sie die Aktivierungs-Tastenkombination, um den Befehlsmodus aufzurufen. Sie hören einen hohen Signalton. Standard ist `NVDA+E` (NVDA plus E).
2. Drücken Sie eine zweite Taste, um eine bestimmte Funktion auszulösen.
3. Wenn Sie eine ungültige Taste drücken, hören Sie einen Fehlerton.

#### Anpassungs- & Einstellungsfenster
* Die Funktionen von BOA sind vollständig modular aufgebaut und können jederzeit aktiviert oder deaktiviert werden. Gehen Sie zu `NVDA-Menü -> Einstellungen -> Einstellungen -> BOA Office Enhancements`, um einzelne Funktionen ein- oder auszuschalten.
* **Intelligente Tastenkombinationen:** Jede einzelne Einstellung verfügt über eine mathematisch eindeutige `Alt+Taste`-Tastenkombination innerhalb des Fensters. Drücken Sie beispielsweise `Alt+E`, um sofort zur Excel-Gruppe zu springen, `Alt+P` für PowerPoint und `Alt+W` für Word.
* Die Einstellungen werden sicher in einer eigenständigen JSON-Datei (`boa_settings.json`) gespeichert, sodass Ihre NVDA-Kernkonfiguration niemals beschädigt wird.
* Wenn Microsoft Office in Zukunft einen Barrierefreiheitsfehler offiziell behebt, können Sie die entsprechende Überschreibungsfunktion von BOA sicher deaktivieren, ohne die restliche Funktionalität des Add-ons zu verlieren.
* **Anpassung der Tastenbelegung:** Alle Funktionen in allen Office-Anwendungen wurden explizit im nativen NVDA-Dialog für die Tastenbelegung freigegeben, sodass Sie jede Tastenkombination völlig frei anpassen können.

#### Sicherheits- & Integrationsgrenzen
* Zwischenablage-Injektionen überprüfen streng die Prozess-IDs des Vordergrundfensters, um den Abfluss von Daten in andere Anwendungen zu verhindern.
* Einige benutzerdefinierte Tastenkombinationen sind im NVDA-Dialog für die Tastenbelegung unter der Kategorie „Better Office Accessibility“ vollständig freigegeben.

---

## 📋 Anforderungen

* **NVDA:** Version 2026.1.0 oder höher.
* **Anwendungen:** Microsoft Excel & Microsoft PowerPoint.

---

## 💾 Installation

1. Laden Sie die neueste `.nvda-addon`-Installationsdatei herunter oder suchen Sie diese im nativen NVDA-Add-on-Store.
2. Wenn Sie aus einer Datei installieren, öffnen Sie die Datei oder verwenden Sie den `NVDA Add-on-Store -> Aus externer Datei installieren`.
3. Starten Sie NVDA neu.

---

## 🛠️ Änderungsprotokoll

### Version 2.0.0
#### Neue Funktionen
* **PowerPoint: Komplette Dokumenten-Analyse (Experimentell) (`NVDA+E`, dann `D`):** Ein hochmodernes, im Hintergrund verarbeitetes Barrierefreiheitswerkzeug, das eine gesamte Präsentation abbildet, ohne die Sprachausgabe von NVDA einzufrieren. Es bietet ein tief navigierbares virtuelles Inhaltsverzeichnis, erkennt Abweichungen in der Lesereihenfolge (visuelle Reihenfolge vs. Z-Reihenfolge), kennzeichnet Folien mit zu viel Text („Wall of Text“) und erfasst komplexe Objekte wie SmartArt und Datentabellen.
* **PowerPoint: Folienlayout-Analyse (Experimentell) (`NVDA+E`, dann `L`):** Scannt Ihre aktuell aktive Folie sofort, um deren räumliches Layout und Barrierefreiheitsbeschränkungen zu verstehen, was eine völlig reibungslose und reaktionsschnelle Bildschirmausgabe ermöglicht. Das bedeutet, dass Sie hier ähnliche Details über die aktuelle Folie erhalten wie bei der Tabellenlayout-Analyse von Excel.
* **PowerPoint: Sammel-Folienorganisator (Experimentell) (`NVDA+E`, dann `X`):** Ähnlich wie bei der Excel-Funktion können Sie jetzt mehrere PowerPoint-Folien gleichzeitig über einen vollständig barrierefreien Dialog neu anordnen, verschieben und organisieren.
* **PowerPoint: Audio-Modus für Formbewegung (Experimentell):** Führt räumliche 3D-Audio-Hinweise auf der PowerPoint-Arbeitsfläche ein. Bietet akustische Rückmeldungen, die die Richtung und die Begrenzungslinien eines Objekts anzeigen, während Sie es bewegen, was das räumliche Bewusstsein erheblich verbessert. Wie bereits erwähnt, handelt es sich hierbei um eine experimentelle Funktion. Wir freuen uns auf Rückmeldungen, um sie zu verbessern.
* **Word: Formatierungs-Auditor (`NVDA+E`, dann `F`):** Scannt Ihr Word-Dokument auf Formatierungsinkonsistenzen, um visuelle Standards zu gewährleisten.
* **Word: Dokumenten-Analyse (`NVDA+E`, dann `D`):** Ruft sofort eine strukturelle Übersicht Ihres Word-Dokuments auf. *(Ein besonderer Dank geht an Paul: Diese Funktion wurde direkt von seinem brillanten „Word Access“-Add-on inspiriert. Wir sind sehr dankbar für seine bahnbrechende Arbeit in diesem Bereich!)*
* **Word: Automatische Fußnotenansage:** Fußnoten werden nun beim Lesen automatisch im Text angekündigt, abhängig von Ihren benutzerdefinierten BOA-Einstellungen. *(Hinweis: Die Unterstützung für Endnoten und Kommentare ist für eine zukünftige Version geplant).*
* **Excel: Der Power-Editor (Barrierefreier Formeleditor):** Ein absoluter Meilenstein für das Ändern gewaltiger Formeln.
  - **Einmaliges Drücken von `NVDA+E`, dann `F2`:** Kündigt sofort die rohe Formelzeichenkette der aktiven Zelle an (oder kündigt „Keine Formel“ an).
  - **Zweimaliges Drücken von `NVDA+E`, dann `F2`:** Öffnet einen vollständig barrierefreien, mehrzeiligen Editor, um riesige, verschachtelte Formeln sicher zu bearbeiten. Ein natives `Enter` fügt Zeilenumbrüche für ein einfaches Lesen ein, und `Ctrl+Enter` speichert die Formel zurück in Excel.
  - *Sicherheitsprüfungen:* Fängt Syntaxfehler sicher ab, bevor sie Ihr Tabellenblatt beschädigen, und erkennt Fehler nach der Berechnung (wie `#NAME?` oder `#DIV/0!`), um Sie sofort zu warnen, wenn eine Formel fehlerhaft ist.
* **Excel: Formelüberwachung & Auswertung:** Eigene Tastenkombinationen hinzugefügt (`NVDA+E`, dann `Shift+P` und `NVDA+E`, dann `Shift+D`), um Spur zum Vorgänger und Spur zum Nachfolger zuverlässig zu verfolgen. Darüber hinaus ist der native Excel-Dialog „Formelauswertung“ nun vollständig barrierefrei; NVDA liest die ausgewerteten Ergebnisse automatisch vor, während Sie die Berechnung schrittweise durchgehen!
* **Excel: Upgrades für die Zellüberwachung Pro:** 
  - **Platz-Manager-Dialog (`NVDA+E`, dann `Alt+M`):** Öffnet einen Dialog, in dem alle aktiv überwachten Zellen aufgelistet sind. Drücken Sie `Enter`, um sofort zu einer Zelle zu springen.
  - **Zurückspringen (`NVDA+E`, dann `\`):** Teleportiert Sie sofort zurück zu Ihrer vorherigen Arbeitszelle, nachdem Sie einen Platz überprüft haben.
  - **Direkter Platzsprung (`Alt` + `Slot-Nummer`):** Umgeht das Präfix vollständig und springt sofort zu einem zugewiesenen Zellplatz.
* **Anpassung der Tastenbelegung:** Alle Funktionen in allen Office-Anwendungen wurden explizit im nativen NVDA-Dialog für die Tastenbelegung freigegeben, sodass Sie jede Tastenkombination völlig frei anpassen können.

#### UX/UI-Verbesserungen
* **Einheitliche navigierbare Berichte:** Wir haben ein einheitliches HTML-Berichtssystem für das gesamte Add-on eingeführt. Funktionen wie die Excel-Ansage der bedingten Formatierung, die Layout-Analysatoren und Dokumenten-Analysatoren geben nicht mehr nur riesige Textblöcke über die Sprachausgabe aus; ihre Ergebnisse öffnen sich nun in einem nativen, navigierbaren HTML-Fenster, sodass Sie die Daten in Ihrem eigenen Tempo überprüfen können.
* **Excel: Verbesserte Nachfolger-/Vorgänger-Verfolgung:** Die Sprachausgabe für die nativen Excel-Tastenkombinationen zur Formelüberwachung (`Ctrl+[` für direkte Vorgänger und `Ctrl+]` für direkte Nachfolger) wurde erheblich verbessert. NVDA kündigt nun explizit an, welche Zellen genau ausgewählt wurden.
* **Excel: Unterstützung für verbundene Zellen:** Verbundene Zellen werden nun korrekt erkannt und vom lückenüberspringenden Zell-Tracker explizit angesagt.

#### Fehlerbehebungen
* **Word: Doppeltes Vorlesen von Listeneinträgen:** Es wurde ein temporärer Patch implementiert, um den Fehler zu beheben, bei dem NVDA in bestimmten Word-Ansichten Listenelemente von Absätzen doppelt vorliest.
* **Excel: Lokalisierungsfehler der Zellüberwachung:** Grundlegende Verfolgungsfehler behoben, die durch die jüngsten Übersetzungs-Lokalisierungsupdates verursacht wurden.

### Was ist neu in v1.6.1
* **Tiefe Dateilokalisierung**: Fehlende Zeichenkettenübersetzungen tief in den Excel-Erweiterungsmodulen (wie der Tabellenlayout-Analyse und dem schnellen Blatt-Verschieber) wurden behoben, um eine 100%ige Lokalisierungsabdeckung zu gewährleisten.
* **Erweiterte Übersetzungsunterstützung**: 7 neue Sprachen zum System hinzugefügt (Türkisch, Polnisch, Koreanisch, Ukrainisch, Tschechisch, Urdu und Punjabi). 
  *(Hinweis: Diese Übersetzungen wurden von einer KI erstellt, daher können einige kleinere Übersetzungsfehler oder Ungenauigkeiten vorliegen).*

### v1.6.0
* **Umfassende Übersetzungsunterstützung**: Das Add-on ist jetzt vollständig lokalisiert und unterstützt 17 globale Sprachen. 
  *(Hinweis: Diese Übersetzungen wurden von einer KI erstellt, daher können einige kleinere Übersetzungsfehler oder Ungenauigkeiten vorliegen).*
* **Strikte Code-Governance**: GPL-2.0-Urheberrechtshinweise im gesamten Quellcode angewendet."""),

### Version 1.5.0 
#### Neue Funktionen
##### Datenende-Radar
Beim Navigieren in großen Tabellenkalkulationen kann es schwierig sein zu sagen, ob eine leere Zelle bedeutet, dass Sie das Ende einer Liste erreicht haben, oder ob es sich lediglich um eine Lücke in den Daten handelt. Das **Datenende-Radar** fungiert als intelligente Umfeldprüfung, um zu verhindern, dass Sie blind mit den Pfeiltasten durch leere Bereiche navigieren.
Wann immer Sie in eine leere Zelle navigieren, scannt BOA sofort die verbleibenden Zellen in Ihrer Bewegungsrichtung. Wenn absolut keine Daten mehr vorhanden sind, wird proaktiv Folgendes angesagt:
* *"Keine weiteren Daten unten"*
* *"Keine weiteren Daten oben"*
* *"Keine weiteren Daten rechts"*
* *"Keine weiteren Daten links"*
**Konfigurationsoptionen:**
Sie können diese Funktion unter `NVDA-Einstellungen -> Einstellungen -> BOA Office Enhancements` konfigurieren. Da Tabellenblätter versteckte Komplexitäten enthalten können (wie unsichtbare Formeln oder ausgeblendete Zeilen), bietet das Radar drei Betriebsmodi:
1. **Aus**: Deaktiviert das Radar vollständig.
2. **Strikte Speicherprüfung (CountA) [Standard]**: Der sicherste und schnellste Ansatz. Er prüft den Rohspeicher des Arbeitsblatts. Wenn er *irgendetwas* unter Ihnen erkennt (einschließlich ausgeblendeter Zeilen, Text, Zahlen oder unsichtbarer Formeln), bleibt er völlig stumm, um Fehlalarme zu vermeiden. Er kündigt „Keine Daten mehr“ nur dann an, wenn der Rest des Blattes mathematisch zu 100 % leer ist.
3. **Nur sichtbare Daten (Math Engine)**: Eine hochentwickelte Engine, die für komplexe Blätter konzipiert wurde. Sie filtert ausgeblendete Zeilen und unsichtbare Formeln (z. B. `=""`) intelligent heraus. Sie bleibt nur dann stumm, wenn sich tatsächlich sichtbare Zahlen oder Texte auf Ihrem Weg befinden.

### Version 1.4 - 2026-06-12
#### Neue Funktionen
* **Zellüberwachung:** Verwenden Sie Befehlspfade, um bestimmte Zellen Speicherplätzen zuzuweisen. Sie können jederzeit dorthin zurückspringen und sie über den zugewiesenen numerischen Platz auslesen.
* **Kontinuierliche Überwachung:** Zugewiesene Zellen werden im Hintergrund automatisch überwacht. Wenn Excel eine Neuberechnung oder Zellbearbeitung auslöst, kündigt BOA sofort den neuen Wert an. Manuell umschalten oder alle über Befehlsplätze löschen.

#### Fehlerbehebungen

### Version 1.3.0 — 2026-06-05
*Finale Version.*

#### Neue Funktionen
* **Tabellenlayout-Analyse:** Leistungsstarke Infrastruktur zum Scannen des Layouts hinzugefügt. Erkennt sofort Blattschutz, aktive Spaltenfilter, ausgeblendete Arbeitsblatt-Reiter und ausgeblendete absolute Ränder, während erkannte Datenblöcke im Cache gespeichert werden.
* **Geführte Datenblock-Navigation:** Die Navigation nach der Analyse ermöglicht sofortige Cursorsprünge zwischen größeren Datenclustern, wodurch leere Zellen nahtlos umgangen werden.
* **Ansage bedingter Formatierung:** Erkennt und liest automatisch die dynamische Farbe, den Schriftstil und den Hintergrundschatten von Zellen, die durch die Regeln für die bedingten Formatierungen in Excel geändert wurden.
* **Explizite Tastenkombinationen für Einstellungen:** Die Benutzeroberfläche der BOA-Einstellungen wurde komplett überarbeitet, um den NVDA-Richtlinien strikt zu entsprechen. Jedes Kontrollkästchen für Funktionen besitzt nun eine global eindeutige Tastenkombination aus `Alt+Buchstabe`, was das Durchwechseln der Tastatur verhindert und Fehler bei der Navigation mit dem Anfangsbuchstaben ausschließt.

#### Fehlerbehebungen
* **Absolute Randerkennung:** Ersetzte native COM-`UsedRange`-Randprüfungen durch absolute mathematische 1D-Grenzprüfungen (`Row 1048576` und `Column 16384`), um die Erkennung von ausgeblendeten Zeilen/Spalten zu garantieren, selbst wenn diese weit außerhalb des aktiven Datenblocks liegen.
* **Sichere Ausstiege bei verzögerten COM-Eigenschaften:** COM-Eigenschaftsschleifen wurden gehärtet, um das Einfrieren des NVDA-Threads bei der Auswertung von Millionen zusammenhängender ausgeblendeter Strukturen zu verhindern.

### Version 1.2.0 — 2026-06-03
*Finale Version.*

#### Neue Funktionen
* **Caching beim Anwendungsstart:** Große architektonische Überarbeitung. Kernmodule werden nun verzögert genau dann geladen, wenn Sie Office-Anwendungen fokussieren, was Startverzögerungen eliminiert, den Fokusfehler bei „unbekannten“ Objekten in Umbenennungsdialogen vollständig behebt und die Struktur des mehrteiligen Quellcodes bewahrt.
* **Verbesserter Zell-Tracker (1D COM Math):** Die Logik zur Erkennung ausgeblendeter Zelllücken wurde umgeschrieben, um nur noch eindimensionale Querschnitte zu bewerten (`current_col` oder `current_row`). Dies reduziert die Rechenlast von COM um über 16 Millionen Zellen und verhindert sofort das Einfrieren der Navigation beim Überspringen ausgeblendeter Bereiche.
* **Bereinigung des Prozessspeichers:** Excel Window Handle (`Hwnd`)-Verfolgung implementiert, um zu erkennen, wenn der Benutzer Excel schließt und wieder öffnet. Dies löscht aktiv veralteten globalen Zustandsspeicher und behebt vollständig die fälschliche Ansage „Blatt ausgeblendet“ beim Öffnen einer neuen „Mappe1“.

#### Fehlerbehebungen
* **Doppelte Auswahlankündigung:** Wechsel von unzuverlässigem asynchronen `winUser.getKeyState` zur Implementierung von `api.getLastInputGesture()`, um doppelte Ansagen bei der Verwendung von Shift+Pfeiltasten perfekt zu unterdrücken.
* **Deaktivierung des Grenzendetektors:** Der proaktive Grenzendetektor wurde deaktiviert, um die Stabilität der nativen NVDA-Navigation zu schützen, und fällt vollständig auf den lückenüberspringenden Tracker zurück.

### Version 1.1.0 — 2026-05-30
*Finale Version.*

#### Neue Funktionen
* **Einstellungs-GUI:** Ein natives Einstellungsfenster für „BOA Office Enhancements“ unter `NVDA -> Einstellungen -> Einstellungen` hinzugefügt, um Funktionen einfach ein- oder auszuschalten.
* **SafeRichEdit-Hook:** Verhindert stille NVDA-Abstürze bei der Interaktion mit RichEdit-Steuerelementen in Office 2024.
* **Anpassbare Tastenkombinationen:** Alle BOA-Tastenkombinationen sind nun vollständig im NVDA-Dialog für die Tastenbelegung unter der Kategorie „Better Office Accessibility“ verfügbar.
* **Excel: Erkennung ausgeblendeter Zeilen/Spalten beim Überspringen:** Kündigt proaktiv an, wenn an ausgeblendeten Zeilen oder Spalten vorbeigesteuert wird, um sicherzustellen, dass Sie gefilterte Daten nie verpassen. Kann in den Einstellungen umgeschaltet werden.

#### Fehlerbehebungen
* **Thread-Sicherheit:** Alle blockierenden Verzögerungen (`time.sleep`) entfernt und durch nicht-blockierende asynchrone NVDA-Rückrufe (Callbacks) ersetzt, um sicherzustellen, dass der Screenreader bei Hintergrundoperationen niemals stottert.

### Version 1.0.0 — 2026-05-24
*Erste öffentliche Version.*

#### Neue Funktionen
* **Excel: Sammel-Tabellenorganisator:** Mehrere Blätter gleichzeitig über einen vollständig barrierefreien Dialog neu anordnen.
* **Excel: Schneller Blatt-Verschieber:** Das aktive Blatt per Tastatur nach links, rechts, an den Anfang oder ans Ende verschieben.
* **Excel: Barrierefreie Blatt-Umbenennung:** Fängt das unzugängliche native Umbenennungsfeld ab und ersetzt es durch einen zuverlässigen, barrierefreien Dialog.
* **Excel: Intelligente Auswahlverfolgung:** Kündigt Auswahlen und Abwahlen von Bereichen mit mehreren Zellen präzise an.
* **PowerPoint: Barrierefreie Farbwähler:** Ermöglicht es NVDA, RGB- und Hex-Werte im Dialog für benutzerdefinierte Farben genau vorzulesen.
* **PowerPoint: Unterstützung für das Standard-Farbraster:** Fängt die Pfeiltastennavigation ab, um ausgeblendete Hex-Codes aus dem unzugänglichen Farbraster (Sechseck-Raster) vorzulesen.
