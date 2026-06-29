# BOA: Better Office Accessibility

BOA je výkonná sada vylepšení přístupnosti pro Microsoft Office, která je navržena tak, aby výrazně zlepšila práci s odečítačem obrazovky pro uživatele NVDA. Přímo opravuje nepřístupné součásti uživatelského rozhraní a přináší nástroje pro rychlou navigaci v aplikacích Excel a PowerPoint.

---

## ⌨️ Přehled klávesových zkratek

| Funkce | Klávesová zkratka | Kontext / Poznámky |
| :--- | :--- | :--- |
| **Vstup do příkazového režimu** | `[Prefix]` (Výchozí: `NVDA+E`) | Aktivuje režim předpony příkazů (spustí vysoké pípnutí) |
| **Zrušit příkazový režim** | `Escape` | Ukončí režim předpony příkazů |
| **VYLEPŠENÍ EXCELU** | | |
| **Analyzovat rozložení listu** | `[Prefix]`, pak `L` | Spustit v Excelu před procházením datových bloků |
| **Přejít na nejbližší datový blok** | `[Prefix]`, pak `J` | Vyžaduje nejprve provedení analýzy rozložení |
| **Otevřít hromadný organizér listů** | `[Prefix]`, pak `X` | Otevře přístupný dialog pro změnu pořadí listů |
| **Oznámení čistého vzorce** | `[Prefix]`, pak `F2` | Jedním stisknutím si poslechnete čistý řetězec vzorce |
| **Pokročilý editor vzorců** | `[Prefix]`, pak dvakrát `F2` | Dvojitým stisknutím otevřete přístupný víceřádkový editor vzorců |
| **Sledování předchůdců** | `[Prefix]`, pak `Shift+P` | Funkce sledování předchůdců v přístupné podobě. |
| **Sledování závislostí** | `[Prefix]`, pak `Shift+D` | Funkce sledování závislostí v přístupné podobě; stisknutím Enter na buňce se tam přemístíte. |
| **Podrobné podmíněné formátování** | `[Prefix]`, pak `F` | Oznámí kompletní podrobnosti o formátování zaměřené buňky |
| **Přesunout aktivní list vlevo** | `NVDA+Shift+LeftArrow` | Posune aktivní list o jednu pozici nahoru |
| **Přesunout aktivní list vpravo** | `NVDA+Shift+RightArrow` | Posune aktivní list o jednu pozici dolů |
| **Přesunout list na začátek/konec** | `NVDA+Shift+Home` / `End` | Odešle list na absolutní hranice |
| **Skrýt / Zobrazit řádek** | `Ctrl+9` / `Ctrl+Shift+9` | Výchozí zkratka; BOA explicitně oznamuje změnu viditelnosti |
| **Skrýt / Zobrazit sloupec** | `Ctrl+0` / `Ctrl+Shift+0` | Výchozí zkratka; BOA explicitně oznamuje změnu viditelnosti |
| **Zobrazit sloupec (náhradní)** | `NVDA+Ctrl+Shift+0` | Obchází konflikty klávesových zkratek s jazykem vstupu ve Windows |
| **Přiřadit buňku do paměťového slotu** | `[Prefix]`, pak `Shift+1` až `Shift+9` | Přiřadí aktuální buňku do slotu pro sledování na pozadí |
| **Přečíst sledovaný paměťový slot** | `[Prefix]`, pak `1` až `9` | Vyvolá a přečte hodnotu přiřazeného slotu |
| **Přímý skok na slot** | `Alt` + `1` až `9` | Okamžitě přemístí kurzor na sledovaný slot |
| **Návrat na předchozí buňku** | `[Prefix]`, pak `\` | Okamžitě vás přenese zpět po kontrole slotu |
| **Dialog Správce slotů** | `[Prefix]`, pak `Alt+M` | Otevře dialog pro zobrazení a správu všech aktivních sledování |
| **Přepnout sledování na pozadí** | `[Prefix]`, pak `M` | Ručně přepíná sledování výpočtů na pozadí |
| **Vymazat všechny paměťové sloty** | `[Prefix]`, pak `Backspace` | Vymaže všechny uložené sledované buňky na pozadí |
| **VYLEPŠENÍ POWERPOINTU** | | |
| **Analyzátor rozložení snímku** | `[Prefix]`, pak `L` | Analyzuje a oznámí prostorové rozložení aktuálního snímku |
| **Analyzátor dokumentu** | `[Prefix]`, pak `D` | Generuje komplexní obsah a zprávu o stavu |
| **Hromadný organizér snímků** | `[Prefix]`, pak `X` | Otevře přístupný dialog pro změnu pořadí více snímků |
| **VYLEPŠENÍ WORDU** | | |
| **Auditor formátování** | `[Prefix]`, pak `F` | Prověří aktuální dokument na nesrovnalosti ve formátování |
| **Analyzátor dokumentu** | `[Prefix]`, pak `D` | Analyzuje strukturu a rozložení aktuálního dokumentu Word |

---

## 🚀 Funkce

### Vylepšení Excelu

#### 1. Analyzátor rozložení listu a mezipaměť
Okamžitě naskenuje jakýkoli list v Excelu, abyste porozuměli jeho struktuře, skrytým prvkům a datovým blokům.
* **Jak to funguje:** BOA rychle naskenuje list a oznámí aktivní datové bloky. Upozorní vás také na **Skryté záložky listů**, aktivní **Filtry**, **Chráněné režimy** a **Skryté vnější hranice** (např. pokud jsou skryté sloupce poblíž pravého okraje listu, což vám zabrání v přehlédnutí dat mimo obrazovku).
* **Navigace v datech:** Po naskenování můžete použít klávesové zkratky pro skok na datový blok a okamžitě přemístit kurzor mezi nalezenými datovými bloky, čímž bez námahy přeskočíte tisíce prázdných buněk.

#### 2. Hromadný organizér listů
Okamžitě změní pořadí a uspořádá více listů najednou pomocí plně přístupného dialogu.
* **Jak to funguje:** Otevře dialog, kde můžete vybrat list a přiřadit mu novou pozici. Naplánované přesuny jsou vypsány v datové tabulce (stisknutím `Del` odstraníte chybu). Klikněte na `OK` a váš sešit se okamžitě přeorganizuje.

#### 3. Rychlé přesunutí listu
Přesuňte aktivní list vlevo, vpravo, na úplný začátek nebo na úplný konec okamžitě pomocí klávesových zkratek.

#### 4. Přístupné přejmenování listu
* Při přejmenování listu má NVDA nativně potíže se čtením znaků, které píšete.
* BOA vkládá vlastní třídu `ExcelSheetRenameEdit`, která používá modul `SafeRichEdit`, což znamená, že při přejmenování můžete přesně číst po znacích, slovech nebo řádcích. To slouží jako vylepšení stávajícího výchozího chování při přejmenování.

#### 5. Sledování skrytých řádků/sloupců
* Proaktivně sleduje váš pohyb po mřížce, aby vám zabránil přehlédnout skrytá nebo filtrovaná data.
* **Přechod přes fragmentované buňky:** Pokud přeskočíte silně fragmentovanou nebo skrytou část mřížky (např. přesun z řádku 3 na řádek 10, protože řádky 4–9 jsou skryté), BOA explicitně oznámí „Řádky 4 až 9 jsou skryté“. To zajišťuje, že vždy víte, kdy byla data ve struktuře přeskočena.

#### 6. Oznamování podmíněného formátování
* Automaticky čte barvu, styl písma a stínování pozadí buněk, které byly dynamicky změněny pravidly podmíněného formátování aplikace Excel.
* Poskytuje vám skutečný vizuální stav buňky namísto pouhé hrubé podkladové hodnoty. Na začátku, když se zaměříte na buňku, oznámí „má podmíněné formátování a některé další drobné podrobnosti“. Pro získání úplných informací použijte podrobnou konfiguraci klávesových zkratek, což je NVDA E a F.

#### 7. Lepší oznamování výběru
čte, zda je vybrána nebo zrušena volba buňky nebo rozsahu.

#### 8 Sledování buněk:
* **Sledování buněk (Cell Monitor):** Použijte příkazové cesty k přiřazení konkrétních buněk do paměťových slotů. Můžete se k nim kdykoli vrátit a přečíst je pomocí přiřazeného číselného slotu.
* **Nepřetržité sledování:** Buňky ve slotech jsou automaticky sledovány na pozadí. Pokud Excel spustí přepočet nebo úpravu buňky, BOA okamžitě oznámí novou hodnotu. Sledování můžete přepnout ručně nebo vymazat vše pomocí příkazových slotů.
* **Excel: Vylepšení Sledování buněk Pro:** 
  - **Dialog Správce slotů (`NVDA+E`, pak `Alt+M`):** Otevře dialog se seznamem všech aktivně sledovaných buněk. Stisknutím `Enter` na jednu z nich okamžitě přejdete.
  - **Rychlý návrat (`NVDA+E`, pak `\`):** Okamžitě vás přenese zpět na předchozí pracovní buňku po kontrole slotu.
  - **Přímý skok na slot (`Prefix + Alt` + `Slot Number`):** Obchází prefix a okamžitě skočí na přiřazenou buňku.

#### 9 Pokročilý editor
* **Excel: Pokročilý editor (přístupný editor vzorců):** Naprostá revoluce pro úpravu rozsáhlých vzorců.
  - **Jedno stisknutí `NVDA+E`, pak `F2`:** Okamžitě oznámí čistý řetězec vzorce aktivní buňky (nebo oznámí „Žádný vzorec“).
  - **Dvojité stisknutí `NVDA+E`, pak `F2`:** Otevře plně přístupný, víceřádkový editor pro bezpečnou úpravu masivních, vnořených vzorců. Klávesa `Enter` přidává zalomení řádků pro snadné čtení a `Ctrl+Enter` jej uloží zpět do Excelu.
  - *Bezpečnostní kontroly:* Bezpečně zachycuje syntaktické chyby dříve, než poškodí váš list, a detekuje chyby po výpočtu (jako `#NAME?` nebo `#DIV/0!`), aby vás okamžitě varoval, pokud se vzorec porušil.

#### 10 Vylepšení auditu a vyhodnocování vzorců:
* **Excel: Audit a vyhodnocování vzorců:** Byly přidány vlastní klávesové zkratky (`NVDA+E`, pak `Shift+P` a `NVDA+E`, pak `Shift+D`) pro spolehlivé sledování předchůdců a závislostí. Nativní dialogové okno aplikace Excel „Vyhodnocení vzorce“ je navíc nyní plně přístupné; NVDA automaticky čte vyhodnocené výsledky při procházení výpočtu!

### Vylepšení PowerPointu

#### 1. Přístupné výběry barev
* Zpřístupňuje dialog Vlastní barva v PowerPointu.
* Správně identifikuje a explicitně čte editační pole „Červená“, „Zelená“ a „Modrá“ (přepsáním `PowerPointRGBEdit`).
* Mapuje dříve neviditelné pole pro zadání Hex hodnoty, takže NVDA může čistě přečíst celou Hex hodnotu barvy.

#### 2. Podpora standardní mřížky barev
* Navigace v mřížce „Standardních“ barev v PowerPointu běžně čte pouze „grafika“ nebo je ticho.
* BOA sleduje vaše šipky po hexagonu a tiše načítá skrytou hodnotu barvy, kterou vám oznamuje v reálném čase (např. „Barva #FF0000“).

#### 3 Hromadný organizér snímků:
* **PowerPoint: Hromadný organizér snímků (experimentální) (`NVDA+E`, pak `X`):** Podobně jako u funkce v Excelu nyní můžete okamžitě měnit pořadí, přesouvat a uspořádávat více snímků PowerPoint najednou pomocí plně přístupného dialogového okna.

#### 4 Analyzátor rozložení snímku
* **PowerPoint: Analyzátor rozložení snímku (experimentální) (`NVDA+E`, pak `L`):** Okamžitě naskenuje váš aktuálně aktivní snímek, aby porozuměl jeho prostorovému rozložení a omezením přístupnosti, což zajišťuje zcela plynulou a responzivní práci s odečítačem. To znamená, že zde získáte podrobnosti o aktuálním snímku podobně jako u analyzátoru rozložení listu v Excelu.


#### 5 Komplexní analyzátor dokumentu [PPT]
* **PowerPoint: Komplexní analyzátor dokumentu (experimentální) (`NVDA+E`, pak `D`):** Velmi pokročilý nástroj přístupnosti běžící na pozadí, který zmapuje celou prezentaci bez zamrznutí řečového syntetizéru NVDA. Poskytuje snadno procházelný virtuální obsah, detekuje nesrovnalosti v pořadí čtení (vizuální pořadí versus pořadí vykreslování Z-Order), označuje snímky s velkým množstvím textu a mapuje složité objekty jako SmartArt a datové tabulky.

#### 6 Vylepšení pohybu [úprav] tvarů:
* **PowerPoint: Zvukový režim pohybu tvarů (experimentální):** Zavádí 3D prostorové zvukové signály na plátno PowerPointu. Poskytuje sluchovou zpětnou vazbu indikující směr a limity ohraničení objektu při jeho přesouvání, čímž výrazně zlepšuje prostorovou orientaci.

### Vylepšení Wordu:
#### 1. Analyzátor dokumentu inspirovaný a odvozený od doplňku Word Access od Paula:
* **Word: Analyzátor dokumentu (`NVDA+E`, pak `D`):** Okamžitě zobrazí strukturální přehled vašeho dokumentu ve Wordu. *(Zvláštní poděkování a uznání Paulovi: Tato funkce byla přímo inspirována jeho skvělým doplňkem „Word Access“. Jsme hluboce vděční za jeho základní práci v této oblasti!)*

#### 2 Auditor formátování
* **Word: Auditor formátování (`NVDA+E`, pak `F`):** Prohledá váš dokument Word na nesrovnalosti ve formátování, aby zajistil dodržení vizuálních standardů.

#### 3 Čtečka poznámek pod čarou:
* **Word: Automatické oznamování poznámek pod čarou:** Poznámky pod čarou budou nyní automaticky oznamovány přímo během čtení v závislosti na vašem vlastním nastavení BOA. *(Poznámka: Podpora pro vysvětlivky a komentáře je plánována pro budoucí verzi).*

### Infrastruktura a technické mechanismy

#### Režim předpony příkazů
Aby se zabránilo konfliktům klávesových zkratek s jinými doplňky NVDA, používá BOA **režim předpony příkazů**:
1. Stisknutím aktivační klávesové zkratky vstoupíte do příkazového režimu. Uslyšíte vysoké pípnutí. Výchozí je NVDA plus E.
2. Stisknutím druhé klávesy spustíte konkrétní funkci.
3. Pokud stisknete neplatnou klávesu, uslyšíte chybové pípnutí.

#### Panel přizpůsobení a nastavení
* Funkce BOA jsou plně modulární a lze je kdykoli povolit nebo zakázat. Přejděte do `Nabídka NVDA -> Možnosti -> Nastavení -> BOA Office Enhancements` pro zapnutí nebo vypnutí jednotlivých funkcí.
* **Inteligentní klávesové zkratky (akcelerátory):** Každé jednotlivé nastavení má v panelu matematicky unikátní klávesovou zkratku `Alt+Klávesa`. Stisknutím `Alt+E` například okamžitě přejdete na skupinu Excel, `Alt+P` na PowerPoint a `Alt+W` na Word.
* Nastavení se bezpečně ukládají do samostatného souboru JSON (`boa_settings.json`), což zajišťuje, že vaše hlavní konfigurace NVDA nebude nikdy poškozena.
* Pokud Microsoft Office v budoucnu oficiálně opraví nějakou chybu přístupnosti, můžete bezpečně zakázat konkrétní přepsání (override hook) v BOA, aniž byste přišli o zbytek funkcí doplňku.
* **Přizpůsobení vstupních gest:** Všechny funkce ve všech aplikacích Office byly explicitně zpřístupněny v nativním dialogu Vstupní gesta NVDA, což vám dává plnou svobodu přizpůsobit si každou klávesovou zkratku.

#### Zabezpečení a hranice integrace
* Vkládání do schránky přísně ověřuje ID procesů v popředí okna, aby se zabránilo úniku dat do jiných aplikací.
* Některé uživatelské klávesové zkratky jsou plně vystaveny v dialogu Vstupní gesta NVDA pod kategorií „Better Office Accessibility“.

---

## 📋 Požadavky

* **NVDA:** Verze 2026.1.0 nebo novější.
* **Aplikace:** Microsoft Excel & Microsoft PowerPoint.

---

## 💾 Instalace

1. Stáhněte si nejnovější vydaný soubor `.nvda-addon` nebo jej vyhledejte v nativním Obchodě s doplňky NVDA.
2. Pokud instalujete ze souboru, otevřete soubor nebo použijte `Obchod s doplňky NVDA -> Instalovat z externího souboru`.
3. Restartujte NVDA.

---

## 🛠️ Seznam změn

### Verze 2.0.0
#### Nové funkce
* **PowerPoint: Komplexní analyzátor dokumentu (experimentální) (`NVDA+E`, pak `D`):** Velmi pokročilý nástroj přístupnosti běžící na pozadí, který zmapuje celou prezentaci bez zamrznutí řečového syntetizéru NVDA. Poskytuje snadno procházelný virtuální obsah, detekuje nesrovnalosti v pořadí čtení (vizuální pořadí versus pořadí vykreslování Z-Order), označuje snímky s velkým množstvím textu a mapuje složité objekty jako SmartArt a datové tabulky.
* **PowerPoint: Analyzátor rozložení snímku (experimentální) (`NVDA+E`, pak `L`):** Okamžitě naskenuje váš aktuálně aktivní snímek, aby porozuměl jeho prostorovému rozložení a omezením přístupnosti, což zajišťuje zcela plynulou a responzivní práci s odečítačem. To znamená, že zde získáte podrobnosti o aktuálním snímku podobně jako u analyzátoru rozložení listu v Excelu.
* **PowerPoint: Hromadný organizér snímků (experimentální) (`NVDA+E`, pak `X`):** Podobně jako u funkce v Excelu nyní můžete okamžitě měnit pořadí, přesouvat a uspořádávat více snímků PowerPoint najednou pomocí plně přístupného dialogového okna.
* **PowerPoint: Zvukový režim pohybu tvarů (experimentální):** Zavádí 3D prostorové zvukové signály na plátno PowerPointu. Poskytuje sluchovou zpětnou vazbu indikující směr a limity ohraničení objektu při jeho přesouvání, čímž výrazně zlepšuje prostorovou orientaci. Jak již bylo zmíněno, jedná se o experimentální funkci a čekáme na zpětnou vazbu pro její vylepšení.
* **Word: Auditor formátování (`NVDA+E`, pak `F`):** Prohledá váš dokument Word na nesrovnalosti ve formátování, aby zajistil dodržení vizuálních standardů.
* **Word: Analyzátor dokumentu (`NVDA+E`, pak `D`):** Okamžitě zobrazí strukturální přehled vašeho dokumentu ve Wordu. *(Zvláštní poděkování a uznání Paulovi: Tato funkce byla přímo inspirována jeho skvělým doplňkem „Word Access“. Jsme hluboce vděční za jeho základní práci v této oblasti!)*
* **Word: Automatické oznamování poznámek pod čarou:** Poznámky pod čarou budou nyní automaticky oznamovány přímo během čtení v závislosti na vašem vlastním nastavení BOA. *(Poznámka: Podpora pro vysvětlivky a komentáře je plánována pro budoucí verzi).*
* **Excel: Pokročilý editor (přístupný editor vzorců):** Naprostá revoluce pro úpravu rozsáhlých vzorců.
  - **Jedno stisknutí `NVDA+E`, pak `F2`:** Okamžitě oznámí čistý řetězec vzorce aktivní buňky (nebo oznámí „Žádný vzorec“).
  - **Dvojité stisknutí `NVDA+E`, pak `F2`:** Otevře plně přístupný, víceřádkový editor pro bezpečnou úpravu masivních, vnořených vzorců. Klávesa `Enter` přidává zalomení řádků pro snadné čtení a `Ctrl+Enter` jej uloží zpět do Excelu.
  - *Bezpečnostní kontroly:* Bezpečně zachycuje syntaktické chyby dříve, než poškodí váš list, a detekuje chyby po výpočtu (jako `#NAME?` nebo `#DIV/0!`), aby vás okamžitě varoval, pokud se vzorec porušil.
* **Excel: Audit a vyhodnocování vzorců:** Byly přidány vlastní klávesové zkratky (`NVDA+E`, pak `Shift+P` a `NVDA+E`, pak `Shift+D`) pro spolehlivé sledování předchůdců a závislostí. Nativní dialogové okno aplikace Excel „Vyhodnocení vzorce“ je navíc nyní plně přístupné; NVDA automaticky čte vyhodnocené výsledky při procházení výpočtu!
* **Excel: Vylepšení Sledování buněk Pro:** 
  - **Dialog Správce slotů (`NVDA+E`, pak `Alt+M`):** Otevře dialog se seznamem všech aktivně sledovaných buněk. Stisknutím `Enter` na jednu z nich okamžitě přejdete.
  - **Rychlý návrat (`NVDA+E`, pak `\`):** Okamžitě vás přenese zpět na předchozí pracovní buňku po kontrole slotu.
  - **Přímý skok na slot (`Alt` + `Slot Number`):** Zcela obchází prefix a okamžitě skočí na přiřazenou buňku.
* **Přizpůsobení vstupních gest:** Všechny funkce ve všech aplikacích Office byly explicitně zpřístupněny v nativním dialogu Vstupní gesta NVDA, což vám dává plnou svobodu přizpůsobit si každou klávesovou zkratku.

#### Vylepšení UX/UI
* **Sjednocené prohlížitelné zprávy:** V celém doplňku jsme zavedli sjednocený systém HTML reportů. Funkce jako oznamování podmíněného formátování v Excelu, analyzátory rozložení a analyzátory dokumentů již nečtou pouze masivní bloky textu; jejich výsledky se nyní otevírají v nativním, procházelném HTML okně, což vám umožňuje procházet data vlastním tempem.
* **Excel: Vylepšené sledování závislostí/předchůdců:** Výrazně se zlepšil hlasový výstup pro nativní zkratky sledování vzorců v Excelu (`Ctrl+[` pro přímé předchůdce a `Ctrl+]` pro přímé závislosti). NVDA nyní bude explicitně oznamovat, které přesné buňky byly vybrány.
* **Excel: Podpora sloučených buněk:** Sloučené buňky jsou nyní správně detekovány a explicitně oznamovány sledováním buněk s přeskakováním mezer.

#### Opravy chyb
* **Word: Dvojité čtení položek seznamu:** Byl implementován dočasný patch pro opravu chyby, kdy NVDA v některých zobrazeních Wordu četl položky seznamu odstavců dvakrát.
* **Excel: Chyba lokalizace sledování buněk:** Byly vyřešeny interní chyby sledování způsobené nedávnými aktualizacemi překladů a lokalizace.

### Co je nového ve v1.6.1
* **Hluboká lokalizace souborů**: Byly opraveny chybějící překlady řetězců hluboko v modulech pro vylepšení Excelu (jako je Analyzátor rozložení listu a Rychlé přesunutí listu), aby se zajistilo 100% pokrytí lokalizace.
* **Rozšířená podpora překladů**: Do systému bylo přidáno 7 nových jazyků (turečtina, polština, korejština, ukrajinština, čeština, urdština a pandžábština). 
  *(Poznámka: Tyto překlady byly vytvořeny umělou inteligencí, proto se mohou vyskytnout drobné chyby nebo nepřesnosti v překladu.)*

### v1.6.0
* **Komplexní podpora překladů**: Doplněk je nyní plně lokalizován s podporou 17 globálních jazyků. 
  *(Poznámka: Tyto překlady byly vytvořeny umělou inteligencí, proto se mohou vyskytnout drobné chyby nebo nepřesnosti v překladu.)*
* **Přísná správa kódu**: V celé codebase byly aplikovány hlavičky autorských práv GPL-2.0."""),

### Verze 1.5.0 
#### Nové funkce
##### Radar konce dat (End of Data Radar)
Při procházení velkých tabulek může být obtížné zjistit, zda prázdná buňka znamená, že jste dosáhli konce seznamu, nebo zda se jedná o pouhou mezeru v datech. **Radar konce dat** funguje jako chytrá kontrola perimetru, která vás zachrání před slepým procházením prázdného prostoru.
Kdykoli přejdete do prázdné buňky, BOA okamžitě naskenuje zbývající buňky ve vašem směru pohybu. Pokud již nezbývají vůbec žádná data, proaktivně oznámí:
* *"Žáda další data pod"*
* *"Žáda další data nad"*
* *"Žáda další data vpravo"*
* *"Žáda další data vlevo"*
**Možnosti konfigurace:**
Tuto funkci můžete konfigurovat v `Možnosti NVDA -> Nastavení -> BOA Office Enhancements`. Vzhledem k tomu, že tabulky mohou obsahovat skryté složitosti (jako jsou neviditelné vzorce nebo sbalené řádky), radar nabízí tři provozní režimy:
1. **Vypnuto (Off)**: Zcela zakáže radar.
2. **Přísná kontrola paměti (CountA) [Výchozí]**: Nejbezpečnější a nejrychlejší přístup. Kontroluje čistou paměť tabulky. Pokud pod vámi detekuje *cokoli* (včetně skrytých řádků, textu, čísel nebo neviditelných vzorců), zůstane zcela potichu, aby se zabránilo falešným poplachům. Oznámí „Žádná další data“ pouze tehdy, když je zbytek listu 100% matematicky prázdný.
3. **Pouze viditelná data (Math Engine)**: Vysoce pokročilý systém navržený pro složité listy. Inteligentně odfiltruje skryté řádky a neviditelné vzorce (např. `=""`). Zůstane potichu pouze tehdy, pokud ve vaší cestě zůstanou skutečná, viditelná čísla nebo text.

### Verze 1.4 – 2026-06-12
#### Nové funkce
* **Sledování buněk (Cell Monitor):** Použijte příkazové cesty k přiřazení konkrétních buněk do paměťových slotů. Můžete se k nim kdykoli vrátit a přečíst je pomocí přiřazeného číselného slotu.
* **Nepřetržité sledování:** Buňky ve slotech jsou automaticky sledovány na pozadí. Pokud Excel spustí přepočet nebo úpravu buňky, BOA okamžitě oznámí novou hodnotu. Sledování můžete přepnout ručně nebo vymazat vše pomocí příkazových slotů.

#### Opravy chyb

### Verze 1.3.0 — 2026-06-05
*Finální verze.*

#### Nové funkce
* **Analyzátor rozložení listu:** Byla přidána výkonná infrastruktura pro skenování rozložení. Okamžitě detekuje zámek listu, aktivní filtry sloupců, skryté záložky listů a skrytá absolutní ohraničení při ukládání nalezených datových bloků do mezipaměti.
* **Řízená navigace po datových blocích:** Navigace po provedení analýzy umožňuje okamžité přesuny kurzoru mezi hlavními shluky dat, čímž hladce přeskakuje prázdné buňky.
* **Oznamování podmíněného formátování:** Automaticky detekuje a čte dynamickou barvu, styl písma a stín pozadí buněk změněných pravidly podmíněného formátování aplikace Excel.
* **Explicitní klávesové zkratky v nastavení:** Kompletně přepracováno grafické uživatelské rozhraní nastavení BOA, aby striktně splňovalo architekturu NVDA. Každé zaškrtávací políčko funkce nyní disponuje globálně unikátní zkratkou `Alt+Písmeno`, což zabraňuje cyklení klávesnice a odstraňuje selhání navigace podle prvního písmene.

#### Opravy chyb
* **Detekce absolutních okrajů hranic:** Nahrazeny nativní kontroly okrajů COM `UsedRange` za absolutní jednorozměrné matematické kontroly hranic (`Řádek 1048576` a `Sloupec 16384`), aby byla zaručena detekce skrytých řádků/sloupců, i když leží daleko mimo aktivní datový blok.
* **Bezpečné opuštění líných COM vlastností:** Zodolněny smyčky vlastností COM, aby se zabránilo zamrznutí vlákna NVDA při vyhodnocování milionů souvislých skrytých struktur.

### Verze 1.2.0 — 2026-06-03
*Finální verze.*

#### Nové funkce
* **Mezipaměť při spuštění aplikace:** Velká architektonická úprava. Jádrové moduly se nyní načítají líně (lazy-load) přesně ve chvíli, kdy se zaměříte na aplikace Office, což eliminuje zpoždění při spouštění, zcela řeší chybu zaměření na „neznámý“ objekt v dialozích přejmenování a zachovává strukturu codebase s více soubory.
* **Vylepšené sledování buněk (1D COM Math):** Přepsána logika detekce mezer skrytých buněk, aby vyhodnocovala pouze jednorozměrné průřezy (`current_col` nebo `current_row`). To snižuje zátěž výpočtů COM o více než 16 milionů buněk, čímž se okamžitě eliminuje zamrzání navigace při přeskakování skrytých rozsahů.
* **Mazání paměti procesu:** Implementováno sledování popisovače okna Excelu (`Hwnd`), které detekuje, kdy uživatel zavře a znovu otevře Excel. Tím se aktivně vymaže zastaralá globální paměť stavu a zcela se vyřeší falešné oznámení „List je skrytý“ při otevření nového sešitu „Sešit1“.

#### Opravy chyb
* **Dvojité oznamování výběru:** Převedeno z nespolehlivého asynchronního `winUser.getKeyState` na implementaci `api.getLastInputGesture()`, aby se dokonale potlačilo dvojité oznamování při použití šipek se Shiftem.
* **Deaktivace proaktivního detektoru hranic:** Proaktivní detektor hranic byl deaktivován z důvodu ochrany stability nativní navigace NVDA a plně se spoléhá na sledování s přeskakováním mezer.

### Verze 1.1.0 — 2026-05-30
*Finální verze.*

#### Nové funkce
* **Uživatelské rozhraní nastavení:** Přidán nativní panel BOA Office Enhancements uvnitř `Nabídka NVDA -> Možnosti -> Nastavení` pro snadné zapnutí nebo vypnutí funkcí.
* **SafeRichEdit Hook:** Zabraňuje tichým pádům NVDA při interakci s ovládacími prvky RichEdit v Office 2024.
* **Přizpůsobitelné klávesové zkratky:** Všechny klávesové zkratky BOA jsou nyní plně vystaveny v dialogu Vstupní gesta NVDA pod kategorií „Better Office Accessibility“.
* **Excel: Detekce přeskočení skrytých řádků/sloupců:** Proaktivně oznamuje, když přejdete přes skryté řádky nebo sloupce, čímž zajišťuje, že nikdy nepřehlédnete filtrovaná data. Lze přepnout v nastavení.

#### Opravy chyb
* **Bezpečnost vláken:** Odstraněna všechna blokující zpoždění (`time.sleep`) a nahrazena neblokujícími asynchronními zpětnými voláními (callbacks) NVDA, aby se zajistilo, že se odečítač obrazovky nikdy nezadrhne při operacích na pozadí.

### Verze 1.0.0 — 2026-05-24
*První veřejná verze.*

#### Nové funkce
* **Excel: Hromadný organizér listů:** Okamžitě změní pořadí více listů najednou pomocí plně přístupného dialogu.
* **Excel: Rychlé přesunutí listu:** Přesuňte aktivní list vlevo, vpravo, na začátek nebo na konec pomocí klávesových příkazů.
* **Excel: Přístupné přejmenování listu:** Zachytí nepřístupné nativní pole pro přejmenování a nahradí jej spolehlivým přístupným dialogem.
* **Excel: Chytré sledování výběru:** Přesně oznamuje výběr a zrušení výběru vícebuněčného rozsahu.
* **PowerPoint: Přístupné výběry barev:** Umožňuje NVDA přesně číst hodnoty RGB a Hex uvnitř dialogu Vlastní barva.
* **PowerPoint: Podpora standardní mřížky barev:** Zachytává navigaci směrovými šipkami za účelem čtení skrytých Hex kódů z nepřístupné šestiúhelníkové mřížky barev.
