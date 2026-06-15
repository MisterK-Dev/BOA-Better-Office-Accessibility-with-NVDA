# BOA: Better Office Accessibility

BOA ist eine leistungsstarke Suite von Zugangserweiterungen für Microsoft Office, die entwickelt wurde, um die Bildschirmleser-Erfahrung für NVDA-Benutzer erheblich zu verbessern. Sie korrigiert direkt unzugängliche UI-Komponenten und führt schnelle Navigationswerkzeuge für Excel und PowerPoint ein.

---

## ⌨️ Hotkey-Referenz

| Funktion | Tastenkombination | Kontext / Notizen |
| :--- | :--- | :--- |
| **Befehlsmodus aktivieren** | `NVDA+E` | Aktiviert den Befehlspräfix-Modus (löst einen hohen Piepton aus) |
| **Tabellenlayout analysieren** | `NVDA+E`, dann `L` | Innerhalb von Excel ausführen, bevor in Datenblöcken navigiert wird |
| **Zum nächsten Datenblock springen** | `NVDA+E`, dann `J` /  | Erfordert vorherige Layout-Analyse |
| **Bulk-Tabellen-Organisator öffnen** | `NVDA+E`, dann `X` | Öffnet den zugänglichen Dialog zur Tabellen-Neuanordnung |
| **Aktive Tabelle nach links verschieben** | `NVDA+Umschalt+PfeilLinks` | Verschiebt die aktive Tabelle um eine Position nach oben |
| **Aktive Tabelle nach rechts verschieben** | `NVDA+Umschalt+PfeilRechts` | Verschiebt das aktive Arbeitsblatt um eine Position nach unten |
| **Tabelle an den Anfang/das Ende verschieben** | `NVDA+Umschalt+Pos1` / `Ende` | Sendet das Arbeitsblatt an die absoluten Grenzen |
| **Detaillierte bedingte Formatierung** | `NVDA+E`, dann `F` | Sagt vollständige Formatierungsdetails der fokussierten Zelle an |
| **Zelle einem Speichersteckplatz zuweisen** | `NVDA+E`, dann `Umschalt+1` bis `Umschalt+9` | Weist die aktuelle Zelle einem Hintergrund-Überwachungssteckplatz zu |
| **Überwachten Zellen-Steckplatz lesen** | `NVDA+E`, dann `1` bis `9` | Ruft den Wert des zugewiesenen Steckplatzes ab und liest ihn |
| **Hintergrundüberwachung umschalten** | `NVDA+E`, dann `M` | Schaltet die Hintergrundberechnungs-Verfolgung manuell um |
| **Alle Speichersteckplätze löschen** | `NVDA+E`, dann `Rücktaste` | Löscht alle gespeicherten Hintergrund-Zellen-Monitore |
| **Befehlsmodus abbrechen** | `Escape` | Verlässt den Befehlspräfix-Modus |

---

## 🚀 Funktionen

### Excel-Erweiterungen

#### 1. Tabellenlayout-Analysator & Caching
Scannen Sie sofort jede Excel-Tabelle, um ihre Struktur, verborgene Elemente und Datenblöcke zu verstehen.
* **Wie es funktioniert:** BOA scannt die Tabelle schnell und sagt aktive Datenblöcke an. Es warnt Sie auch vor **ausgeblendeten Tabellen-Tabs**, aktiven **Filtern**, **geschützten Modi** und **verborgenen äußeren Rändern** (z.B. wenn Spalten nahe dem rechten Rand der Tabelle ausgeblendet sind, um zu verhindern, dass Sie Daten außerhalb des Bildschirms übersehen).
* **Daten-Navigation:** Nach dem Scannen können Sie die Datenblock-Sprung-Hotkeys verwenden, um Ihren Cursor sofort zwischen den entdeckten Datenblöcken zu bewegen und mühelos Tausende leerer Zellen zu überspringen.

#### 2. Bulk-Tabellen-Organisator
Ordnen und arrangieren Sie mehrere Tabellen sofort und gleichzeitig über einen vollständig zugänglichen Dialog.
* **Wie es funktioniert:** Öffnet einen Dialog, in dem Sie eine Tabelle auswählen und sie auf eine neue Position abbilden können. Geplante Verschiebungen werden in einer Datentabelle aufgelistet (drücken Sie `Entf`, um einen Fehler zu entfernen). Klicken Sie auf `OK` und Ihre Arbeitsmappe wird sofort neu angeordnet.

#### 3. Schneller Tabellenverschieber
Verschieben Sie die aktive Tabelle sofort mit Ihren Tastaturkürzeln nach links, nach rechts, ganz an den Anfang oder ganz an das Ende.

#### 4. Zugängliches Umbenennen von Tabellen
* Beim Umbenennen einer Tabelle hat NVDA nativ Probleme, die von Ihnen eingegebenen Zeichen zu lesen.
* BOA fügt eine benutzerdefinierte `ExcelSheetRenameEdit`-Klasse ein, die die `SafeRichEdit`-Engine verwendet, was bedeutet, dass Sie beim Umbenennen präzise nach Zeichen, Wort oder Zeile lesen können. Dies dient als Verbesserung des bestehenden Standard-Umbenennungsverhaltens.

#### 5. Tracker für versteckte Zeilen/Spalten
* Verfolgt proaktiv Ihre Bewegung über das Raster, um zu verhindern, dass Sie verborgene oder gefilterte Daten übersehen.
* **Überquerte fragmentierte Zellen:** Wenn Sie einen stark fragmentierten oder verborgenen Abschnitt des Rasters überspringen (z.B. von Zeile 3 zu Zeile 10 wechseln, da die Zeilen 4–9 ausgeblendet sind), sagt BOA explizit „Zeilen 4 bis 9 ausgeblendet" an. Dies stellt sicher, dass Sie immer wissen, wenn Daten in der Struktur übersprungen wurden.

#### 6. Ansage für bedingte Formatierung
* Liest automatisch die Farbe, den Schriftstil und die Hintergrundschattierung von Zellen, die durch die bedingten Formatierungsregeln von Excel dynamisch geändert wurden.
* Gibt Ihnen den wahren visuellen Zustand der Zelle anstelle nur des rohen zugrunde liegenden Werts. Wenn Sie die Zelle fokussieren, wird zunächst „hat bedingte Formatierung und einige andere kleinere Details" angesagt. Für umfassende Informationen verwenden Sie die detaillierte Hotkey-Konfiguration, die NVDA E und F ist.

#### 7. Bessere Auswahl-Ansage
Liest, ob eine Zelle oder ein Bereich ausgewählt oder abgewählt wird.

#### 8 Zellen-Monitor:
* **Zellen-Monitor:** Verwenden Sie Befehlspfade, um bestimmte Zellen auf Speichersteckplätze abzubilden. Sie können jederzeit zurückspringen und sie über den zugewiesenen numerischen Steckplatz lesen.
* **Kontinuierliche Überwachung:** Zugewiesene Zellen werden automatisch im Hintergrund überwacht. Wenn Excel eine Neuberechnung oder Zellenbearbeitung auslöst, sagt BOA sofort den neuen Wert an. Manuelles Umschalten oder Löschen aller über Befehlssteckplätze.

### PowerPoint-Erweiterungen

#### 1. Zugängliche Farbwähler
* Entsperrt das Dialogfeld „Benutzerdefinierte Farbe" in PowerPoint.
* Identifiziert und liest die Bearbeitungsfelder „Rot", „Grün" und „Blau" explizit und korrekt vor (durch Überschreiben von `PowerPointRGBEdit`).
* Bildet das zuvor unsichtbare Hex-Eingabefeld ab, damit NVDA den vollständigen Hex-Farbwert sauber lesen kann.

#### 2. Standardfarbraster-Unterstützung
* Das Navigieren im PowerPoint „Standard"-Farbsechseckraster wird normalerweise als „Grafik" oder durch Stille gelesen.
* BOA verfolgt Ihre Pfeiltasten über das Sechseck und ruft den versteckten Farbwert im Hintergrund ab und sagt ihn Ihnen in Echtzeit an (z.B. „Farbe #FF0000").

### Infrastruktur & Technische Mechanismen

#### Der Befehlspräfix-Modus
Um Tastaturkürzel-Konflikte mit anderen NVDA-Plugins zu vermeiden, verwendet BOA einen **Befehlspräfix-Modus**:
1. Drücken Sie den Aktivierungs-Hotkey, um in den Befehlsmodus zu gelangen. Sie hören einen hohen Piepton.
2. Drücken Sie eine zweite Taste, um eine bestimmte Funktion auszulösen.
3. Wenn Sie eine ungültige Taste drücken, hören Sie einen Fehler-Piepton.

#### Anpassungs- & Einstellungsfeld
* Die BOA-Funktionen sind vollständig modular und können jederzeit aktiviert oder deaktiviert werden. Gehen Sie zu `NVDA Menü -> Optionen -> Einstellungen -> BOA Office Erweiterungen`, um einzelne Funktionen ein- oder auszuschalten.
* **Intelligente Beschleunigertasten:** Jede einzelne Einstellung verfügt über ein mathematisch eindeutiges `Alt+Taste`-Beschleuniger-Kürzel innerhalb des Bedienfelds. Drücken Sie beispielsweise `Alt+E`, um sofort zur Excel-Gruppe zu springen, `Alt+P` für PowerPoint und `Alt+W` für Word.
* Die Einstellungen werden sicher in einer eigenständigen JSON-Datei (`boa_settings.json`) gespeichert, sodass Ihre NVDA-Kernkonfiguration niemals beschädigt wird.
* Wenn Microsoft Office in Zukunft einen Zugänglichkeitsfehler offiziell behebt, können Sie den spezifischen BOA-Überschreibungs-Hook sicher deaktivieren, ohne den Rest der Add-on-Funktionalität zu verlieren.

#### Sicherheit & Integrationsgrenzen
* Zwischenablage-Einfügungen überprüfen strikt die Vordergrund-Prozess-IDs der Fenster, um ein Durchsickern von Daten in andere Anwendungen zu verhindern.
* Einige benutzerdefinierte Hotkeys sind vollständig im Dialogfeld für NVDA-Eingabegesten in der Kategorie „Better Office Accessibility" verfügbar.

---

## 📋 Anforderungen

* **NVDA:** Version 2026.1.0 oder neuer.
* **Anwendungen:** Microsoft Excel & Microsoft PowerPoint.

---

## 💾 Installation

1. Laden Sie die neueste `.nvda-addon`-Release-Datei herunter oder suchen Sie sie im nativen NVDA-Add-on-Store.
2. Wenn Sie aus einer Datei installieren, öffnen Sie die Datei oder verwenden Sie `NVDA Add-on-Store -> Aus externer Datei installieren`.
3. Starten Sie NVDA neu.

---

## 🛠️ Änderungsprotokoll

### Neuerungen in v1.6.1
* **Tiefe Dateilokalisierung**: Fehlende Zeichenfolgen-Übersetzungen tief in den Excel-Erweiterungsmodulen (wie dem Blattlayout-Analysator und der Schnellen Blattverschiebung) wurden behoben, um eine 100%ige Abdeckung der Lokalisierung sicherzustellen.
* **Erweiterte Übersetzungsunterstützung**: Es wurden 7 neue Sprachen zum System hinzugefügt (Türkisch, Polnisch, Koreanisch, Ukrainisch, Tschechisch, Urdu und Punjabi).
  *(Hinweis: Diese Übersetzungen wurden von einer KI erstellt, daher können kleinere Übersetzungsfehler oder Ungenauigkeiten vorhanden sein.)*

### v1.6.0
* **Umfassende Übersetzungsunterstützung**: Das Add-on ist jetzt vollständig lokalisiert und unterstützt 17 globale Sprachen. 
  *(Hinweis: Diese Übersetzungen wurden durch KI generiert, daher können geringfügige Übersetzungsfehler oder Ungenauigkeiten vorhanden sein.)*
* **Strikte Code-Governance**: GPL-2.0-Copyright-Header im gesamten Codebestand angewendet.

### Version 1.5.0 
#### Neue Funktionen
##### Datenende-Radar
Beim Navigieren durch große Tabellen kann es schwierig sein zu erkennen, ob eine leere Zelle das Ende einer Liste bedeutet oder ob es nur eine Lücke in den Daten gibt. Das **Datenende-Radar** fungiert als intelligente Umkreiskontrolle, um Sie davor zu bewahren, blind durch leeren Raum zu navigieren.
Wenn Sie in eine leere Zelle navigieren, scannt BOA sofort die restlichen Zellen in Ihrer Bewegungsrichtung. Wenn absolut keine Daten mehr übrig sind, wird proaktiv angesagt:
* *„Keine weiteren Daten darunter"*
* *„Keine weiteren Daten darüber"*
* *„Keine weiteren Daten rechts"*
* *„Keine weiteren Daten links"*
**Konfigurationsoptionen:**
Sie können diese Funktion über `NVDA Optionen -> Einstellungen -> BOA Office Erweiterungen` konfigurieren. Da Tabellen verborgene Komplexitäten (wie unsichtbare Formeln oder reduzierte Zeilen) enthalten können, bietet das Radar drei Betriebsmodi:
1. **Aus**: Deaktiviert das Radar vollständig.
2. **Strenge Speicherprüfung (Anzahl2) [Standard]**: Der sicherste und schnellste Ansatz. Es überprüft den Rohspeicher der Tabelle. Wenn es *irgendetwas* unter Ihnen erkennt (einschließlich versteckter Zeilen, Text, Zahlen oder unsichtbarer Formeln), bleibt es völlig stumm, um Fehlalarme zu vermeiden. Es sagt nur „Keine weiteren Daten" an, wenn der Rest der Tabelle zu 100% mathematisch leer ist.
3. **Nur sichtbare Daten (Mathematische Engine)**: Eine hochentwickelte Engine für komplexe Tabellen. Sie filtert intelligent ausgeblendete Zeilen und unsichtbare Formeln (z.B. `=""`) heraus. Es bleibt nur dann stumm, wenn sich in Ihrem Pfad tatsächliche, sichtbare Zahlen oder Text befinden.

### Version 1.4 - 2026-06-12
#### Neue Funktionen
* **Zellen-Monitor:** Verwenden Sie Befehlspfade, um bestimmte Zellen auf Speichersteckplätze abzubilden. Sie können jederzeit zurückspringen und sie über den zugewiesenen numerischen Steckplatz lesen.
* **Kontinuierliche Überwachung:** Zugewiesene Zellen werden automatisch im Hintergrund überwacht. Wenn Excel eine Neuberechnung oder Zellenbearbeitung auslöst, sagt BOA sofort den neuen Wert an. Manuelles Umschalten oder Löschen aller über Befehlssteckplätze.

#### Fehlerbehebungen

### Version 1.3.0 — 2026-06-05
*Finale Veröffentlichung.*

#### Neue Funktionen
* **Tabellenlayout-Analysator:** Leistungsstarke Layout-Scan-Infrastruktur hinzugefügt. Erkennt sofort den Tabellenschutz, aktive Spaltenfilter, ausgeblendete Tabellen-Tabs und versteckte absolute Ränder, während entdeckte Datenblöcke zwischengespeichert werden.
* **Geführte Datenblock-Navigation:** Die Navigation nach der Analyse ermöglicht sofortige Cursor-Sprünge zwischen großen Daten-Clustern, wodurch leere Zellen nahtlos umgangen werden.
* **Ansage für bedingte Formatierung:** Erkennt und liest automatisch die dynamische Farbe, den Schriftstil und die Hintergrundschattierung von Zellen, die durch die bedingten Formatierungsregeln von Excel geändert wurden.
* **Explizite Einstellungs-Beschleuniger:** Die BOA-Einstellungs-GUI wurde komplett überarbeitet, um strikt der NVDA-Architektur zu entsprechen. Jedes Funktionskontrollkästchen besitzt nun ein global eindeutiges `Alt+Buchstabe`-Kürzel, das Tastaturwechsel verhindert und Fehler bei der Navigation mit dem ersten Buchstaben beseitigt.

#### Fehlerbehebungen
* **Absolute Randgrenzenerkennung:** Native COM-`UsedRange`-Randprüfungen wurden durch absolute mathematische 1D-Grenzprüfungen (`Zeile 1048576` und `Spalte 16384`) ersetzt, um die Erkennung verborgener Zeilen/Spalten zu gewährleisten, selbst wenn sie weit außerhalb des aktiven Datenblocks liegen.
* **Sicheres Abbrechen von faulen COM-Eigenschaften:** COM-Eigenschaftsschleifen wurden gehärtet, um das Einfrieren des NVDA-Threads bei der Auswertung von Millionen zusammenhängender versteckter Strukturen zu verhindern.

### Version 1.2.0 — 2026-06-03
*Finale Veröffentlichung.*

#### Neue Funktionen
* **App-Start-Caching:** Große Architekturüberholung. Kernmodule werden jetzt exakt dann geladen, wenn Sie Office-Anwendungen fokussieren, was Boot-Verzögerungen eliminiert, den 'unbekannten' Objektfokus-Fehler in Umbenennungsdialogen vollständig löst und die Multi-File-Codebasis-Struktur erhält.
* **Verbesserter Tracker für versteckte Zellen (1D COM Math):** Die Logik zur Erkennung von Lücken durch versteckte Zellen wurde neu geschrieben, um nur eindimensionale Querschnitte (`current_col` oder `current_row`) auszuwerten. Dies reduziert die COM-Berechnungslast um über 16 Millionen Zellen und eliminiert sofort das Einfrieren der Navigation beim Überspringen verborgener Bereiche.
* **Löschen des Prozessspeichers:** Excel Window Handle (`Hwnd`)-Verfolgung implementiert, um zu erkennen, wenn der Benutzer Excel schließt und wieder öffnet. Dies löscht aktiv den veralteten globalen Zustandsspeicher und löst die falsche Ansage „Tabelle ausgeblendet" beim Öffnen einer frischen „Mappe1" vollständig.

#### Fehlerbehebungen
* **Doppelte Auswahl-Ansage:** Abkehr vom unzuverlässigen asynchronen `winUser.getKeyState` und Implementierung von `api.getLastInputGesture()`, um doppelte Ansagen bei Verwendung der Umschalt+Pfeil-Tasten perfekt zu unterdrücken.
* **Deaktivierung des Grenzwertdetektors:** Der proaktive Grenzwertdetektor wurde deaktiviert, um die native Navigationsstabilität von NVDA zu schützen, und greift nun vollständig auf den Lücken überspringenden Tracker zurück.

### Version 1.1.0 — 2026-05-30
*Finale Veröffentlichung.*

#### Neue Funktionen
* **Einstellungs-GUI:** Ein natives BOA-Office-Erweiterungsfeld wurde unter `NVDA -> Optionen -> Einstellungen` hinzugefügt, um Funktionen einfach ein- oder auszuschalten.
* **SafeRichEdit-Hook:** Verhindert lautlose NVDA-Abstürze bei der Interaktion mit RichEdit-Steuerelementen in Office 2024.
* **Anpassbare Hotkeys:** Alle BOA-Hotkeys sind jetzt vollständig im Dialogfeld für NVDA-Eingabegesten unter der Kategorie „Better Office Accessibility" zugänglich.
* **Excel: Erkennung von versteckten Zeilen/Spalten-Sprüngen:** Sagt proaktiv an, wenn an ausgeblendeten Zeilen oder Spalten vorbeinavigiert wird, und stellt sicher, dass Sie gefilterte Daten nie übersehen. Kann in den Einstellungen umgeschaltet werden.

#### Fehlerbehebungen
* **Thread-Sicherheit:** Alle blockierenden Verzögerungen (`time.sleep`) wurden entfernt und durch nicht blockierende asynchrone NVDA-Rückrufe ersetzt, um sicherzustellen, dass der Bildschirmleser bei Hintergrundvorgängen niemals stottert.

### Version 1.0.0 — 2026-05-24
*Erste öffentliche Veröffentlichung.*

#### Neue Funktionen
* **Excel: Bulk-Tabellen-Organisator:** Ordnen Sie mehrere Tabellen gleichzeitig über einen vollständig zugänglichen Dialog neu an.
* **Excel: Schneller Tabellenverschieber:** Verschieben Sie die aktive Tabelle mit Tastaturbefehlen nach links, nach rechts, an den Anfang oder an das Ende.
* **Excel: Zugängliches Umbenennen von Tabellen:** Fängt das unzugängliche native Umbenennungsfeld ab und ersetzt es durch einen zuverlässigen, zugänglichen Dialog.
* **Excel: Intelligente Auswahlverfolgung:** Sagt Mehrfachzellenauswahlen und -abwahlen präzise an.
* **PowerPoint: Zugängliche Farbwähler:** Ermöglicht NVDA das genaue Lesen von RGB- und Hex-Werten im Dialogfeld 'Benutzerdefinierte Farbe'.
* **PowerPoint: Standardfarbraster-Unterstützung:** Fängt die Pfeiltastennavigation ab, um versteckte Hex-Codes aus dem unzugänglichen Farbsechseckraster zu lesen.
