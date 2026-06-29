# BOA: Better Office Accessibility

BOA to potężny pakiet ulepszeń dostępności dla pakietu Microsoft Office, zaprojektowany w celu znacznego usprawnienia korzystania z czytnika ekranu przez użytkowników NVDA. Bezpośrednio naprawia niedostępne elementy interfejsu użytkownika oraz wprowadza narzędzia szybkiej nawigacji w programach Excel i PowerPoint.

---

## ⌨️ Spis skrótów klawiszowych

| Funkcja | Kombinacja klawiszy | Kontekst / Uwagi |
| :--- | :--- | :--- |
| **Wejście w tryb poleceń** | `[Prefix]` (Domyślnie: `NVDA+E`) | Aktywuje tryb prefiksu poleceń (wyzwala wysoki sygnał dźwiękowy) |
| **Anulowanie trybu poleceń** | `Escape` | Wychodzi z trybu prefiksu poleceń |
| **ULEPSZENIA PROGRAMU EXCEL** | | |
| **Analiza układu arkusza** | `[Prefix]`, then `L` | Uruchom w programie Excel przed nawigacją po blokach danych |
| **Skok do najbliższego bloku danych** | `[Prefix]`, then `J` | Wymaga wcześniejszego wykonania analizy układu |
| **Otwarcie organizatora arkuszy** | `[Prefix]`, then `X` | Otwiera dostępny dialog zmiany kolejności arkuszy |
| **Odczytywanie czystej formuły** | `[Prefix]`, then `F2` | Jednokrotne naciśnięcie pozwala usłyszeć czysty ciąg formuły |
| **Edytor formuł Power Editor** | `[Prefix]`, then `F2` twice | Dwukrotne naciśnięcie otwiera dostępny, wielolinijkowy edytor formuł |
| **Śledzenie zależności (poprzedniki)** | `[Prefix]`, then `Shift+P` | Śledzenie poprzedników – ta sama funkcja w dostępny sposób. |
| **Śledzenie zależności (zależne)** | `[Prefix]`, then `Shift+D` | Śledzenie zależności – ta sama funkcja w dostępny sposób; naciśnięcie klawisza Enter na komórce przeniesie Cię do niej. |
| **Szczegóły formatowania warunkowego** | `[Prefix]`, then `F` | Odczytuje pełne szczegóły formatowania aktywnej komórki |
| **Przeniesienie aktywnego arkusza w lewo** | `NVDA+Shift+LeftArrow` | Przesuwa aktywny arkusz o jedną pozycję w górę |
| **Przeniesienie aktywnego arkusza w prawo** | `NVDA+Shift+RightArrow` | Przesuwa aktywny arkusz o jedną pozycję w dół |
| **Przeniesienie arkusza na początek/koniec** | `NVDA+Shift+Home` / `End` | Przesuwa arkusz na same krańce |
| **Ukrywanie / Pokazywanie wiersza** | `Ctrl+9` / `Ctrl+Shift+9` | Skrót natywny; BOA jawnie ogłasza zmianę widoczności |
| **Ukrywanie / Pokazywanie kolumny** | `Ctrl+0` / `Ctrl+Shift+0` | Skrót natywny; BOA jawnie ogłasza zmianę widoczności |
| **Pokazywanie kolumny (skrót alternatywny)** | `NVDA+Ctrl+Shift+0` | Omija konflikty skrótów klawiszowych języka wprowadzania systemu Windows |
| **Przypisanie komórki do slotu pamięci** | `[Prefix]`, then `Shift+1` do `Shift+9` | Przypisuje bieżącą komórkę do slotu monitorowania w tle |
| **Odczytanie monitorowanej komórki** | `[Prefix]`, then `1` do `9` | Przywołuje i odczytuje wartość z przypisanego slotu |
| **Bezpośredni skok do slotu** | `Alt` + `1` do `9` | Natychmiast przenosi kursor do monitorowanego slotu |
| **Powrót do poprzedniej komórki** | `[Prefix]`, then `\` | Natychmiast przenosi z powrotem po sprawdzeniu slotu |
| **Okno menedżera slotów** | `[Prefix]`, then `Alt+M` | Otwiera okno dialogowe do przeglądania i zarządzania wszystkimi aktywnymi monitorami |
| **Przełączanie monitorowania w tle** | `[Prefix]`, then `M` | Ręcznie przełącza śledzenie obliczeń w tle |
| **Czyszczenie wszystkich slotów pamięci** | `[Prefix]`, then `Backspace` | Usuwa wszystkie zapisane monitory komórek w tle |
| **ULEPSZENIA PROGRAMU POWERPOINT** | | |
| **Analiza układu slajdu** | `[Prefix]`, then `L` | Analizuje i odczytuje przestrzenny układ bieżącego slajdu |
| **Analiza dokumentu** | `[Prefix]`, then `D` | Generuje szczegółowy spis treści i raport o stanie dostępności |
| **Organizator slajdów** | `[Prefix]`, then `X` | Otwiera dostępny dialog zmiany kolejności wielu slajdów |
| **ULEPSZENIA PROGRAMU WORD** | | |
| **Audytor formatowania** | `[Prefix]`, then `F` | Kontroluje bieżący dokument pod kątem niespójności formatowania |
| **Analiza dokumentu** | `[Prefix]`, then `D` | Analizuje układ i strukturę bieżącego dokumentu Word |

---

## 🚀 Funkcje

### Ulepszenia programu Excel

#### 1. Analizator układu arkusza i buforowanie
Natychmiastowe skanowanie dowolnego arkusza programu Excel w celu poznania jego struktury, ukrytych elementów i bloków danych.
* **Jak to działa:** BOA szybko skanuje arkusz i ogłasza aktywne bloki danych. Ostrzega również o **ukrytych kartach arkuszy**, aktywnych **filtrach**, **trybach chronionych** oraz **ukrytych granicach zewnętrznych** (np. jeśli kolumny w pobliżu prawej krawędzi arkusza są ukryte, co zapobiega przeoczeniu danych poza ekranem).
* **Nawigacja po danych:** Po zakończeniu skanowania możesz użyć skrótów skoku do bloku danych, aby natychmiast przenosić kursor między wykrytymi blokami danych, bez wysiłku omijając tysiące pustych komórek.

#### 2. Zbiorczy organizator arkuszy
Natychmiastowa zmiana kolejności i układu wielu arkuszy jednocześnie za pomocą w pełni dostępnego okna dialogowego.
* **Jak to działa:** Otwiera okno dialogowe, w którym można wybrać arkusz i przypisać go do nowej pozycji. Zaplanowane przeniesienia są wymienione w tabeli danych (naciśnij `Del`, aby usunąć pomyłkę). Kliknij `OK`, a kolejność arkuszy w skoroszycie zostanie natychmiast zaktualizowana.

#### 3. Szybkie przenoszenie arkusza
Przenosń aktywny arkusz w lewo, w prawo, na sam początek lub na sam koniec natychmiast, korzystając ze skrótów klawiszowych.

#### 4. Dostępna zmiana nazwy arkusza
* Podczas zmiany nazwy arkusza czytnik NVDA ma natywne problemy z odczytywaniem wpisywanych znaków.
* BOA wstrzykuje niestandardową klasę `ExcelSheetRenameEdit` która korzysta z silnika `SafeRichEdit`, co oznacza, że podczas zmiany nazwy możesz precyzyjnie czytać tekst znak po znaku, słowo po słowie lub linia po linii. Stanowi to ulepszenie istniejącego, domyślnego zachowania zmiany nazwy.

#### 5. Śledzenie ukrytych wierszy/kolumn
* Aktywnie śledzi ruch po siatce, aby zapobiec przeoczeniu ukrytych lub przefiltrowanych danych.
* **Przekraczanie pofragmentowanych komórek:** Jeśli przeskoczysz przez mocno pofragmentowaną lub ukrytą sekcję siatki (np. przechodząc z wiersza 3 do wiersza 10, ponieważ wiersze 4–9 są ukryte), BOA jawnie ogłosi „Wiersze od 4 do 9 są ukryte”. Dzięki temu zawsze wiesz, kiedy w strukturze pominięto dane.

#### 6. Informowanie o formatowaniu warunkowym
* Automatycznie odczytuje kolor, styl czcionki i odcień tła komórek, które zostały dynamicznie zmienione przez reguły formatowania warunkowego programu Excel.
* Podaje rzeczywisty stan wizualny komórki, a nie tylko jej surową, bazową wartość. Początkowo, po ustawieniu ostrości na komórce, ogłasza komunikat typu „ma formatowanie warunkowe i inne drobne szczegóły”. Aby uzyskać pełne informacje, użyj szczegółowej konfiguracji skrótów klawiszowych, którą jest NVDA E i F.

#### 7. Lepsze ogłaszanie zaznaczania
odczytuje, czy komórka lub zakres zostały zaznaczone lub odznaczone.

#### 8 Monitor komórek:
* **Monitor komórek:** Użyj ścieżek poleceń, aby przypisać określone komórki do slotów pamięci. Możesz do nich wrócić i odczytać je w dowolnym momencie, korzystając z przypisanego slotu numerycznego.
* **Ciągłe monitorowanie:** Komórki przypisane do slotów są automatycznie monitorowane w tle. Jeśli program Excel wyzwoli ponowne przeliczanie lub edycję komórki, BOA natychmiast ogłosi nową wartość. Przełączaj ręcznie lub wyczyść wszystko za pomocą slotów poleceń.
* **Excel: Ulepszenia Cell Monitor Pro:** 
  - **Menedżer slotów (`NVDA+E`, then `Alt+M`):** Otwiera okno dialogowe z listą wszystkich aktywnie monitorowanych komórek. Naciśnij `Enter`, aby natychmiast przejść do jednej z nich.
  - **Powrót (`NVDA+E`, then `\`):** Natychmiast przenosi z powrotem do poprzednio edytowanej komórki po sprawdzeniu slotu.
  - **Bezpośredni skok do slotu (`Prefix + Alt` + `Slot Number`):** Omiń prefiks i natychmiast przejdź do przypisanego slotu komórki.

#### 9 Power editor
* **Excel: Power Editor (dostępny edytor formuł):** Prawdziwy przełom w modyfikowaniu ogromnych formuł.
  - **Jednokrotne naciśnięcie `NVDA+E`, then `F2`:** Natychmiast odczytuje czysty ciąg formuły aktywnej komórki (lub ogłasza „Brak formuły”).
  - **Dwukrotne naciśnięcie `NVDA+E`, then `F2`:** Otwiera w pełni dostępny, wielolinijkowy edytor pozwalający na bezpieczną modyfikację potężnych, zagnieżdżonych formuł. Natywne naciśnięcie klawisza `Enter` wstawia podziały linii ułatwiające czytanie, a skrót `Ctrl+Enter` zapisuje zmiany z powrotem w programie Excel.
  - *Kontrola bezpieczeństwa:* Bezpiecznie przechwytuje błędy składniowe, zanim uszkodzą one arkusz, oraz wykrywa błędy po obliczeniach (takie jak `#NAME?` lub `#DIV/0!`), aby natychmiast ostrzec Cię, jeśli formuła uległa uszkodzeniu.

#### 10 Formula auditing and evaluation enhancements:
* **Excel: Audytowanie i ewaluacja formuł:** Dodano niestandardowe skróty klawiszowe (`NVDA+E`, then `Shift+P` oraz `NVDA+E`, then `Shift+D`), aby niezawodnie śledzić poprzedniki i zależności. Co więcej, natywne okno dialogowe programu Excel „Szacuj formułę” jest teraz w pełni dostępne; NVDA automatycznie odczytuje oszacowane wyniki w miarę przechodzenia przez kolejne kroki obliczeń!

### PowerPoint Enhancements

#### 1. Dostępne narzędzia wyboru kolorów
* Odblokowuje okno dialogowe Kolory niestandardowe w programie PowerPoint.
* Prawidłowo identyfikuje i bezpośrednio odczytuje pola edycji „Czerwony”, „Zielony” i „Niebieski” (poprzez nadpisanie klasy `PowerPointRGBEdit`).
* Mapuje wcześniej niewidoczne pole wprowadzania wartości Hex, dzięki czemu NVDA może czysto odczytać pełną wartość koloru w tym formacie.

#### 2. Obsługa standardowej siatki kolorów
* Nawigacja po „standardowej” sześciokątnej siatce kolorów programu PowerPoint zwykle skutkuje odczytaniem komunikatu „Grafika” lub ciszą.
* BOA śledzi ruchy klawiszy strzałek po sześciokącie, po cichu pobiera ukrytą wartość koloru i ogłasza ją w czasie rzeczywistym (np. „Kolor #FF0000”).

#### 3 Bulk Slide Organizer:
* **PowerPoint: Zbiorczy organizator slajdów (eksperymentalny) (`NVDA+E`, then `X`):** Podobnie jak w programie Excel, możesz teraz natychmiast zmieniać kolejność, przenosić i układać wiele slajdów programu PowerPoint jednocześnie za pomocą w pełni dostępnego okna dialogowego.

#### 4 Slide lay out analyzer
* **PowerPoint: Analizator układu slajdu (eksperymentalny) (`NVDA+E`, then `L`):** Natychmiast skanuje aktualnie aktywny slajd, aby przeanalizować jego układ przestrzenny i ograniczenia dostępności, zapewniając w pełni płynne i responsywne działanie czytnika ekranu. Oznacza to, że uzyskasz tutaj szczegółowe informacje o bieżącym slajdzie, podobnie jak w przypadku analizatora układu arkusza w programie Excel.


#### 5 Complete Document [PPT] analyzer
* **PowerPoint: Kompletny analizator dokumentu (eksperymentalny) (`NVDA+E`, then `D`):** Wysoce zaawansowane, działające w tle narzędzie ułatwiające dostępność, które mapuje całą prezentację bez zamrażania silnika mowy NVDA. Zapewnia ono szczegółowy, w pełni nawigowalny wirtualny spis treści, wykrywa niezgodności kolejności czytania (kolejność wizualna kontra kolejność na osi Z), oznacza slajdy typu „ściana tekstu” oraz mapuje złożone obiekty, takie jak SmartArt i tabele danych.

#### 6 shape movement [adjustment] enhancements:
* **PowerPoint: Dźwiękowy tryb przesuwania kształtów (eksperymentalny):** Wprowadza wskazówki dźwięku przestrzennego 3D na obszarze roboczym programu PowerPoint. Zapewnia sprzężenie zwrotne w postaci dźwięków wskazujących kierunek i granice obiektu podczas jego przesuwania, znacznie poprawiając świadomość przestrzenną.

### Word Enhancements:
#### 1. Analizator dokumentu (zainspirowany i oparty na dodatku Word Access autorstwa Paula):
* **Word: Analizator dokumentu (`NVDA+E`, then `D`):** Natychmiast wyświetla strukturalny przegląd dokumentu Word. *(Specjalne podziękowania i wyrazy uznania dla Paula: Ta funkcja została bezpośrednio zainspirowana jego genialnym dodatkiem „Word Access”. Jesteśmy głęboko wdzięczni za jego fundamentalną pracę w tym obszarze!)*

#### 2 Formatting Auditor
* **Word: Audytor formatowania (`NVDA+E`, then `F`):** Skanuje dokument Word w poszukiwaniu niespójności formatowania, aby zapewnić zgodność ze standardami wizualnymi.

#### 3 Foot note reader:
* **Word: Automatyczne ogłaszanie przypisów dolnych:** Przypisy dolne będą teraz automatycznie ogłaszane w tekście podczas czytania, w zależności od niestandardowych ustawień BOA. *(Uwaga: Obsługa przypisów końcowych i komentarzy jest planowana w przyszłym wydaniu).*

### Infrastructure & Technical Mechanisms

#### Tryb prefiksu poleceń
Aby zapobiec konfliktom skrótów klawiszowych z innymi wtyczkami NVDA, BOA używa **trybu prefiksu poleceń**:
1. Naciśnij skrót aktywacyjny, aby wejść w tryb poleceń. Usłyszysz wysoki sygnał dźwiękowy. Domyślnie jest to NVDA plus E.
2. Naciśnij drugi klawisz, aby wywołać określoną funkcję.
3. W przypadku naciśnięcia nieprawidłowego klawisza usłyszysz dźwięk błędu.

#### Panel personalizacji i ustawień
* Funkcje BOA są w pełni modularne i można je włączać lub wyłączać w dowolnym momencie. Przejdź do `Menu NVDA -> Preferencje -> Ustawienia -> BOA Office Enhancements`, aby przełączać poszczególne funkcje.
* **Inteligentne klawisze dostępu (skróty Alt):** Każde pojedyncze ustawienie w panelu posiada unikalny skrót klawisza dostępu `Alt+Klawisz`. Na przykład naciśnij `Alt+E`, aby natychmiast przejść do grupy ustawień programu Excel, `Alt+P` dla programu PowerPoint i `Alt+W` dla programu Word.
* Ustawienia są bezpiecznie zapisywane w osobnym pliku JSON (`boa_settings.json`), co gwarantuje, że główna konfiguracja NVDA nigdy nie zostanie uszczuplona ani uszkodzona.
* Jeśli firma Microsoft w przyszłości oficjalnie naprawi dany błąd dostępności w pakiecie Office, możesz bezpiecznie wyłączyć konkretny punkt przechwytujący (override hook) w BOA, nie tracąc przy tym pozostałych funkcji dodatku.
* **Dostosowywanie gestów wprowadzania:** Wszystkie funkcje we wszystkich aplikacjach pakietu Office zostały jawnie udostępnione w natywnym oknie dialogowym Zdarzenia wejściowe czytnika NVDA, co daje pełną swobodę dostosowywania każdego skrótu klawiszowego.

#### Bezpieczeństwo i granice integracji
* Operacje na schowku ściśle weryfikują identyfikatory procesów okna pierwszoplanowego, aby zapobiec wyciekowi danych do innych aplikacji.
* Niektóre niestandardowe skróty klawiszowe są w pełni udostępnione w oknie dialogowym Zdarzenia wejściowe czytnika NVDA w kategorii „Better Office Accessibility”.

---

## 📋 Wymagania

* **NVDA:** Wersja 2026.1.0 lub nowsza.
* **Aplikacje:** Microsoft Excel i Microsoft PowerPoint.

---

## 💾 Instalacja

1. Pobierz najnowszy plik wydania `.nvda-addon` lub znajdź go bezpośrednio w sklepie z dodatkami NVDA.
2. W przypadku instalacji z pliku otwórz plik lub użyj opcji `Sklep z dodatkami NVDA -> Zainstaluj z zewnętrznego pliku`.
3. Uruchom ponownie NVDA.

---

## 🛠️ Dziennik zmian

### Wersja 2.0.0
#### Nowe funkcje
* **PowerPoint: Kompletny analizator dokumentu (eksperymentalny) (`NVDA+E`, then `D`):** Wysoce zaawansowane, działające w tle narzędzie ułatwiające dostępność, które mapuje całą prezentację bez zamrażania silnika mowy NVDA. Zapewnia ono szczegółowy, w pełni nawigowalny wirtualny spis treści, wykrywa niezgodności kolejności czytania (kolejność wizualna kontra kolejność na osi Z), oznacza slajdy typu „ściana tekstu” oraz mapuje złożone obiekty, takie jak SmartArt i tabele danych.
* **PowerPoint: Analizator układu slajdu (eksperymentalny) (`NVDA+E`, then `L`):** Natychmiast skanuje aktualnie aktywny slajd, aby przeanalizować jego układ przestrzenny i ograniczenia dostępności, zapewniając w pełni płynne i responsywne działanie czytnika ekranu. Oznacza to, że uzyskasz tutaj szczegółowe informacje o bieżącym slajdzie, podobnie jak w przypadku analizatora układu arkusza w programie Excel.
* **PowerPoint: Zbiorczy organizator slajdów (eksperymentalny) (`NVDA+E`, then `X`):** Podobnie jak w programie Excel, możesz teraz natychmiast zmieniać kolejność, przenosić i układać wiele slajdów programu PowerPoint jednocześnie za pomocą w pełni dostępnego okna dialogowego.
* **PowerPoint: Dźwiękowy tryb przesuwania kształtów (eksperymentalny):** Wprowadza wskazówki dźwięku przestrzennego 3D na obszarze roboczym programu PowerPoint. Zapewnia sprzężenie zwrotne w postaci dźwięków wskazujących kierunek i granice obiektu podczas jego przesuwania, znacznie poprawiając świadomość przestrzenną. Ponieważ funkcja ta ma charakter eksperymentalny, czekamy na opinie w celu jej ulepszenia.
* **Word: Audytor formatowania (`NVDA+E`, then `F`):** Skanuje dokument Word w poszukiwaniu niespójności formatowania, aby zapewnić zgodność ze standardami wizualnymi.
* **Word: Analizator dokumentu (`NVDA+E`, then `D`):** Natychmiast wyświetla strukturalny przegląd dokumentu Word. *(Specjalne podziękowania i wyrazy uznania dla Paula: Ta funkcja została bezpośrednio zainspirowana jego genialnym dodatkiem „Word Access”. Jesteśmy głęboko wdzięczni za jego fundamentalną pracę w tym obszarze!)*
* **Word: Automatyczne ogłaszanie przypisów dolnych:** Przypisy dolne będą teraz automatycznie ogłaszane w tekście podczas czytania, w zależności od niestandardowych ustawień BOA. *(Uwaga: Obsługa przypisów końcowych i komentarzy jest planowana w przyszłym wydaniu).*
* **Excel: Power Editor (dostępny edytor formuł):** Prawdziwy przełom w modyfikowaniu ogromnych formuł.
  - **Jednokrotne naciśnięcie `NVDA+E`, then `F2`:** Natychmiast odczytuje czysty ciąg formuły aktywnej komórki (lub ogłasza „Brak formuły”).
  - **Dwukrotne naciśnięcie `NVDA+E`, then `F2`:** Otwiera w pełni dostępny, wielolinijkowy edytor pozwalający na bezpieczną modyfikację potężnych, zagnieżdżonych formuł. Natywne naciśnięcie klawisza `Enter` wstawia podziały linii ułatwiające czytanie, a skrót `Ctrl+Enter` zapisuje zmiany z powrotem w programie Excel.
  - *Kontrola bezpieczeństwa:* Bezpiecznie przechwytuje błędy składniowe, zanim uszkodzą one arkusz, oraz wykrywa błędy po obliczeniach (takie jak `#NAME?` lub `#DIV/0!`), aby natychmiast ostrzec Cię, jeśli formuła uległa uszkodzeniu.
* **Excel: Audytowanie i ewaluacja formuł:** Dodano niestandardowe skróty klawiszowe (`NVDA+E`, then `Shift+P` oraz `NVDA+E`, then `Shift+D`), aby niezawodnie śledzić poprzedniki i zależności. Co więcej, natywne okno dialogowe programu Excel „Szacuj formułę” jest teraz w pełni dostępne; NVDA automatycznie odczytuje oszacowane wyniki w miarę przechodzenia przez kolejne kroki obliczeń!
* **Excel: Ulepszenia Cell Monitor Pro:** 
  - **Menedżer slotów (`NVDA+E`, then `Alt+M`):** Otwiera okno dialogowe z listą wszystkich aktywnie monitorowanych komórek. Naciśnij `Enter`, aby natychmiast przejść do jednej z nich.
  - **Powrót (`NVDA+E`, then `\`):** Natychmiast przenosi z powrotem do poprzednio edytowanej komórki po sprawdzeniu slotu.
  - **Bezpośredni skok do slotu (`Alt` + `Slot Number`):** Omiń całkowicie prefiks i natychmiast przejdź do przypisanego slotu komórki.
* **Dostosowywanie gestów wprowadzania:** Wszystkie funkcje we wszystkich aplikacjach pakietu Office zostały jawnie udostępnione w natywnym oknie dialogowym Zdarzenia wejściowe czytnika NVDA, co daje pełną swobodę dostosowywania każdego skrótu klawiszowego.

#### Ulepszenia interfejsu i doświadczenia użytkownika (UX/UI)
* **Jednolite raporty w oknie przeglądarki:** W całym dodatku wprowadziliśmy jednolity system raportów HTML. Funkcje takie jak informowanie o formatowaniu warunkowym programu Excel, analizatory układu oraz analizatory dokumentów nie odczytują już ogromnych bloków tekstu; ich wyniki otwierają się teraz w natywnym, nawigowalnym oknie HTML, co pozwala na przeglądanie danych we własnym tempie.
* **Excel: Ulepszone śledzenie zależności/poprzedników:** Znacznie ulepszono ogłaszanie mowy dla natywnych skrótów śledzenia formuł programu Excel (`Ctrl+[` dla bezpośrednich poprzedników oraz `Ctrl+]` dla bezpośrednich zależności). NVDA będzie teraz jawnie ogłaszać, jakie dokładnie komórki zostały zaznaczone.
* **Excel: Obsługa scalonych komórek:** Scalone komórki są teraz prawidłowo wykrywane i jawnie ogłaszane przez monitor komórek omijający puste przestrzenie.

#### Poprawki błędów
* **Word: Podwójne odczytywanie elementów listy:** Wprowadzono tymczasową poprawkę usuwającą błąd polegający na tym, że NVDA dwukrotnie odczytuje elementy listy akapitowej w niektórych widokach programu Word.
* **Excel: Błąd lokalizacji monitora komórek:** Rozwiązano podstawowe błędy śledzenia spowodowane niedawnymi aktualizacjami lokalizacji tłumaczeń.

### Co nowego w v1.6.1
* **Głęboka lokalizacja plików**: Naprawiono brakujące tłumaczenia ciągów znaków głęboko w modułach ulepszeń programu Excel (takich jak analizator układu arkusza i szybkie przenoszenie arkusza), aby zapewnić 100% pokrycia lokalizacji.
* **Rozszerzona obsługa tłumaczeń**: Dodano 7 nowych języków do systemu (turecki, polski, koreański, ukraiński, czeski, urdu i pendżabski). 
  *(Uwaga: Tłumaczenia te zostały wygenerowane przez sztuczną inteligencję, więc mogą występować drobne błędy lub nieścisłości).*

### v1.6.0
* **Wszechstronna obsługa tłumaczeń**: Dodatek jest teraz w pełni zlokalizowany i obsługuje 17 języków globalnych. 
  *(Uwaga: Tłumaczenia te zostały wygenerowane przez sztuczną inteligencję, więc mogą występować drobne błędy lub nieścisłości).*
* **Rygorystyczne zarządzanie kodem**: Zastosowano nagłówki praw autorskich GPL-2.0 w całym kodzie źródłowym.""",

### Wersja 1.5.0 
#### Nowe funkcje
##### Radar końca danych (End of Data Radar)
Podczas nawigacji w dużych arkuszach kalkulacyjnych trudność może sprawić określenie, czy pusta komórka oznacza koniec listy, czy też jest to po prostu przerwa w danych. **Radar końca danych** działa jako inteligentne sprawdzanie obwodu, chroniąc przed bezcelowym poruszaniem się strzałkami po pustej przestrzeni.
Ilekroć przejdziesz do pustej komórki, BOA natychmiast skanuje pozostałe komórki w kierunku Twojego ruchu. Jeśli nie pozostały już absolutnie żadne dane, aktywnie ogłosi:
* *"Brak danych poniżej"*
* *"Brak danych powyżej"*
* *"Brak danych po prawej"*
* *"Brak danych po lewej"*
**Opcje konfiguracji:**
Możesz skonfigurować tę funkcję w menu `Preferencje NVDA -> Ustawienia -> BOA Office Enhancements`. Ponieważ arkusze kalkulacyjne mogą kryć różne złożoności (takie jak niewidoczne formuły lub zwinięte wiersze), radar oferuje trzy tryby pracy:
1. **Wyłączony**: Całkowicie wyłącza radar.
2. **Rygorystyczne sprawdzanie pamięci (CountA) [Domyślne]**: Najbezpieczniejsza i najszybsza metoda. Sprawdza czystą pamięć arkusza kalkulacyjnego. Jeśli wykryje *cokolwiek* w dalszej części (w tym ukryte wiersze, tekst, liczby lub niewidoczne formuły), pozostaje całkowicie cicha, aby zapobiec fałszywym alarmom. Ogłasza komunikat o braku danych tylko wtedy, gdy reszta arkusza jest w 100% matematycznie pusta.
3. **Tylko widoczne dane (silnik matematyczny)**: Wysoce zaawansowany silnik przeznaczony dla skomplikowanych arkuszy. Inteligentnie odfiltrowuje ukryte wiersze i niewidoczne formuły (np. `=""`). Zachowa milczenie tylko wtedy, gdy na Twojej drodze pozostaną rzeczywiste, widoczne liczby lub teksty.

### Wersja 1.4 – 2026-06-12
#### Nowe funkcje
* **Monitor komórek:** Użyj ścieżek poleceń, aby przypisać określone komórki do slotów pamięci. Możesz do nich wrócić i odczytać je w dowolnym momencie, korzystając z przypisanego slotu numerycznego.
* **Ciągłe monitorowanie:** Komórki przypisane do slotów są automatycznie monitorowane w tle. Jeśli program Excel wyzwoli ponowne przeliczanie lub edycję komórki, BOA natychmiast ogłosi nową wartość. Przełączaj ręcznie lub wyczyść wszystko za pomocą slotów poleceń.

#### Poprawki błędów

### Wersja 1.3.0 — 2026-06-05
*Wydanie ostateczne.*

#### Nowe funkcje
* **Analizator układu arkusza:** Dodano wydajną infrastrukturę skanowania układu. Natychmiast wykrywa ochronę arkusza, aktywne filtry kolumn, ukryte karty arkuszy i ukryte granice bezwzględne, buforując jednocześnie wykryte bloki danych.
* **Nawigacja z przewodnikiem po blokach danych:** Nawigacja po analizie umożliwia natychmiastowe przenoszenie kursora między głównymi skupiskami danych, bezproblemowo omijając puste komórki.
* **Informowanie o formatowaniu warunkowym:** Automatycznie wykrywa i odczytuje dynamiczne zmiany koloru, stylu czcionki i odcienia tła komórek zmienionych przez reguły formatowania warunkowego programu Excel.
* **Jasne klawisze dostępu w ustawieniach:** Całkowicie przebudowano graficzny interfejs ustawień BOA, aby ściśle dostosować go do architektury NVDA. Każde pole wyboru funkcji posiada teraz unikalny w skali globalnej skrót `Alt+Litera`, co zapobiega pętli przełączania klawiatury i eliminuje błędy nawigacji po pierwszej literze.

#### Poprawki błędów
* **Wykrywanie bezwzględnych granic krawędzi:** Zastąpiono natywne sprawdzanie krawędzi za pomocą właściwości COM `UsedRange` bezwzględnymi jednowymiarowymi matematycznymi sprawdzeniami granic (`Wiersz 1048576` i `Kolumna 16384`), aby zagwarantować wykrywanie ukrytych wierszy/kolumn, nawet jeśli znajdują się daleko poza aktywnym blokiem danych.
* **Bezpieczne wyjścia z właściwości Lazy COM:** Wzmocniono pętle właściwości COM, aby zapobiec zamrażaniu wątku NVDA podczas sprawdzania milionów sąsiadujących ukrytych struktur.

### Wersja 1.2.0 — 2026-06-03
*Wydanie ostateczne.*

#### Nowe funkcje
* **Buforowanie przy uruchamianiu aplikacji:** Główna przebudowa architektury. Główne moduły są teraz ładowane w sposób opóźniony (lazy-load) dokładnie wtedy, gdy ustawiasz ostrość na aplikacjach pakietu Office, co eliminuje opóźnienia rozruchowe, całkowicie rozwiązuje problem z ostrością na „nieznanym” obiekcie w oknach zmiany nazwy i zachowuje strukturę kodu opartego na wielu plikach.
* **Ulepszone śledzenie komórek (1D COM Math):** Przepisano logikę wykrywania przerw między ukrytymi komórkami, aby oceniała tylko jednowymiarowe przekroje (`current_col` lub `current_row`). Zmniejsza to obciążenie obliczeniowe COM o ponad 16 milionów komórek, natychmiast eliminując zawieszanie się nawigacji podczas przeskakiwania ukrytych zakresów.
* **Czyszczenie pamięci procesu:** Zaimplementowano śledzenie uchwytu okna Excela (`Hwnd`), aby wykrywać, kiedy użytkownik zamyka i ponownie otwiera program Excel. Aktywnie usuwa to nieaktualną pamięć stanu globalnego i całkowicie eliminuje błędne ogłaszanie komunikatu „Arkusz ukryty” podczas otwierania nowego skoroszytu „Zeszyt1”.

#### Poprawki błędów
* **Podwójne ogłaszanie zaznaczenia:** Zrezygnowano z niezaufanego asynchronicznego `winUser.getKeyState` i zaimplementowano `api.getLastInputGesture()`, aby doskonale tłumić podwójne komunikaty przy użyciu klawiszy Shift+Strzałki.
* **Dezaktywacja detektora granic:** Aktywny detektor granic (Proactive Boundary Detector) został wyłączony w celu ochrony stabilności natywnej nawigacji NVDA, a program opiera się teraz całkowicie na module śledzącym omijającym puste przestrzenie.

### Wersja 1.1.0 — 2026-05-30
*Wydanie ostateczne.*

#### Nowe funkcje
* **GUI ustawień:** Dodano natywny panel BOA Office Enhancements w `NVDA -> Preferencje -> Ustawienia`, aby łatwo włączać i wyłączać funkcje.
* **SafeRichEdit Hook:** Zapobiega cichym awariom czytnika NVDA podczas interakcji z kontrolkami RichEdit w pakiecie Office 2024.
* **Dostosowywalne skróty:** Wszystkie skróty klawiszowe BOA są teraz w pełni widoczne w oknie dialogowym Zdarzenia wejściowe czytnika NVDA w kategorii „Better Office Accessibility”.
* **Excel: Wykrywanie pominięcia ukrytego wiersza/kolumny:** Aktywnie ogłasza przechodzenie obok ukrytych wierszy lub kolumn, upewniając się, że nigdy nie przeoczysz przefiltrowanych danych. Można to przełączyć w ustawieniach.

#### Poprawki błędów
* **Bezpieczeństwo wątków:** Usunięto wszystkie blokujące opóźnienia (`time.sleep`) i zastąpiono je nieblokującymi asynchronicznymi wywołaniami zwrotnymi NVDA, aby czytnik ekranu nigdy nie zacinał się podczas operacji w tle.

### Wersja 1.0.0 — 2026-05-24
*Pierwsze publiczne wydanie.*

#### Nowe funkcje
* **Excel: Zbiorczy organizator arkuszy:** Natychmiastowa zmiana kolejności wielu arkuszy jednocześnie za pomocą w pełni dostępnego okna dialogowego.
* **Excel: Szybkie przenoszenie arkusza:** Przenoszenie aktywnego arkusza w lewo, w prawo, na początek lub na koniec za pomocą poleceń klawiaturowych.
* **Excel: Dostępna zmiana nazwy arkusza:** Przechwytuje niedostępne, natywne pole zmiany nazwy i zastępuje je niezawodnym, dostępnym oknem dialogowym.
* **Excel: Inteligentne śledzenie zaznaczenia:** Dokładnie ogłasza zaznaczanie i odznaczanie zakresów wielokomórkowych.
* **PowerPoint: Dostępne narzędzia wyboru kolorów:** Umożliwia czytnikowi NVDA dokładne odczytywanie wartości RGB i Hex w oknie dialogowym Kolory niestandardowe.
* **PowerPoint: Obsługa standardowej siatki kolorów:** Przechwytuje nawigację klawiszami strzałek, aby odczytać ukryte kody Hex z niedostępnej, sześciokątnej siatki kolorów.
