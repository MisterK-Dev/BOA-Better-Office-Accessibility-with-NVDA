# -*- coding: UTF-8 -*-
import os
import polib

# Batch 1: es, fr, de, pt, it, ru, pl, uk

translations = {
    "es": {
        "ui": {
            "Enable &Bulk Sheet Organizer and Quick Sheet Mover": "Habilitar el Organizador de hojas en &masa y el Movimiento rápido de hoja",
            "Explicitly announce mer&ged cells": "Anunciar explícitamente las c&eldas combinadas",
            "&Excel": "&Excel",
            "Prevent NVDA &crashes in Word text fields (SafeRichEdit)": "Prevenir bloq&ueos de NVDA en los campos de texto de Word (SafeRichEdit)",
            "&Word": "&Word",
            "Read hidden &hex codes when navigating the Standard Color hexagon grid": "Leer los códigos &hex ocultos al navegar por la cuadrícula hexagonal de Color estándar",
            "Ensure the Hex Color edit &field is properly labeled": "Asegurar que el &campo de edición de Color Hex esté etiquetado correctamente",
            "Enable &Slide Layout Analyzer (Prefix + L)": "Habilitar el Analizador de &diseño de diapositiva (Prefix + L)",
            "Shape Movement &Audio Mode:": "Modo de &audio de movimiento de formas:",
            "&PowerPoint": "&PowerPoint"
        },
        "desc_intro": "BOA es una potente suite de mejoras de accesibilidad para Microsoft Office, fruto del desarrollo asistido por IA; diseñada para mejorar enormemente la experiencia del lector de pantalla para los usuarios de NVDA. Toda la base de código ha sido generada mediante Anti Gravity 2.0. Parchea directamente componentes de interfaz inaccesibles e introduce herramientas de navegación rápida para Excel y PowerPoint.",
        "cl_201": """### Version 2.0.1
#### UX/UI Enhancements
* **Tabbed Settings Dialog:** Reorganized the BOA Settings Panel into accessible tabs (&Excel, &Word, and &PowerPoint) using `wx.Notebook`, vastly improving screen reader navigation and eliminating long scrolling lists.
* **NVDA 2026.2 Compatibility:** Tested and certified for NVDA 2026.2.""",
        "readme_201": """### Version 2.0.1
#### Mejoras de UX/UI
* **Diálogo de configuración con pestañas:** Se reorganizó el panel de configuración de BOA en pestañas accesibles (&Excel, &Word y &PowerPoint) usando `wx.Notebook`, mejorando enormemente la navegación del lector de pantalla y eliminando largas listas de desplazamiento. Puede cambiar rápidamente entre pestañas usando `Alt+E`, `Alt+W`, `Alt+P` o los atajos estándar `Ctrl+AvanzarPágina`/`Ctrl+RetrocederPágina`.
* **Compatibilidad con NVDA 2026.2:** Probado y certificado para NVDA 2026.2."""
    },
    "fr": {
        "ui": {
            "Enable &Bulk Sheet Organizer and Quick Sheet Mover": "Activer l'&Organisateur de feuilles en masse et le Déplacement rapide de feuille",
            "Explicitly announce mer&ged cells": "Annoncer explicitement les cellules f&usionnées",
            "&Excel": "&Excel",
            "Prevent NVDA &crashes in Word text fields (SafeRichEdit)": "Empêcher les &plantages de NVDA dans les champs de texte Word (SafeRichEdit)",
            "&Word": "&Word",
            "Read hidden &hex codes when navigating the Standard Color hexagon grid": "Lire les codes &hex masqués lors de la navigation sur la grille hexagonale de couleurs standard",
            "Ensure the Hex Color edit &field is properly labeled": "S'assurer que le &champ d'édition de couleur Hex est correctement étiqueté",
            "Enable &Slide Layout Analyzer (Prefix + L)": "Activer l'Analyseur de &disposition de diapositive (Prefix + L)",
            "Shape Movement &Audio Mode:": "Mode &audio de déplacement de forme :",
            "&PowerPoint": "&PowerPoint"
        },
        "desc_intro": "BOA est une puissante suite d'améliorations d'accessibilité pour Microsoft Office, issue du développement assisté par IA ; conçue pour améliorer considérablement l'expérience des lecteurs d'écran pour les utilisateurs de NVDA. L'intégralité du code a été générée avec Anti Gravity 2.0. Elle corrige directement les composants d'interface inaccessibles et introduit des outils de navigation rapide pour Excel et PowerPoint.",
        "cl_201": """### Version 2.0.1
#### Améliorations UX/UI
* **Boîte de dialogue des paramètres avec onglets :** Réorganisation du panneau de paramètres BOA en onglets accessibles (&Excel, &Word et &PowerPoint) à l'aide de `wx.Notebook`, améliorant considérablement la navigation du lecteur d'écran et éliminant les longues listes de défilement.
* **Compatibilité NVDA 2026.2 :** Testé et certifié pour NVDA 2026.2.""",
        "readme_201": """### Version 2.0.1
#### Améliorations UX/UI
* **Boîte de dialogue des paramètres avec onglets :** Réorganisation du panneau de paramètres BOA en onglets accessibles (&Excel, &Word et &PowerPoint) à l'aide de `wx.Notebook`, améliorant considérablement la navigation du lecteur d'écran et éliminant les longues listes de défilement. Vous pouvez basculer rapidement entre les onglets en utilisant `Alt+E`, `Alt+W`, `Alt+P`, ou les raccourcis standard `Ctrl+PageDown`/`Ctrl+PageUp`.
* **Compatibilité NVDA 2026.2 :** Testé et certifié pour NVDA 2026.2."""
    },
    "de": {
        "ui": {
            "Enable &Bulk Sheet Organizer and Quick Sheet Mover": "&Massen-Tabellenorganisator und Schnellverschieber aktivieren",
            "Explicitly announce mer&ged cells": "Ver&bundene Zellen explizit ansagen",
            "&Excel": "&Excel",
            "Prevent NVDA &crashes in Word text fields (SafeRichEdit)": "NVDA-Abstürze in Word-Textfeldern verhindern (&SafeRichEdit)",
            "&Word": "&Word",
            "Read hidden &hex codes when navigating the Standard Color hexagon grid": "Versteckte &Hex-Codes beim Navigieren im Standardfarben-Sechseckraster vorlesen",
            "Ensure the Hex Color edit &field is properly labeled": "Sicherstellen, dass das Hex-Farb-Eingabe&feld korrekt beschriftet ist",
            "Enable &Slide Layout Analyzer (Prefix + L)": "Folien-&Layout-Analysator aktivieren (Prefix + L)",
            "Shape Movement &Audio Mode:": "Formbewegungs-&Audiomodus:",
            "&PowerPoint": "&PowerPoint"
        },
        "desc_intro": "BOA ist eine leistungsstarke Suite von Barrierefreiheitserweiterungen für Microsoft Office, entstanden aus KI-gestützter Entwicklung; entwickelt, um das Bildschirmleseerlebnis für NVDA-Benutzer drastisch zu verbessern. Die gesamte Codebasis wurde mit Anti Gravity 2.0 erstellt. Sie behebt unzugängliche UI-Komponenten direkt und führt schnelle Navigationstools für Excel und PowerPoint ein.",
        "cl_201": """### Version 2.0.1
#### UX/UI-Verbesserungen
* **Einstellungen mit Registerkarten:** Das BOA-Einstellungsfenster wurde mithilfe von `wx.Notebook` in barrierefreie Registerkarten (&Excel, &Word und &PowerPoint) unterteilt, was die Navigation mit Bildschirmlesern erheblich verbessert und lange Bildlauflisten überflüssig macht.
* **NVDA 2026.2-Kompatibilität:** Getestet und zertifiziert für NVDA 2026.2.""",
        "readme_201": """### Version 2.0.1
#### UX/UI-Verbesserungen
* **Einstellungen mit Registerkarten:** Das BOA-Einstellungsfenster wurde mithilfe von `wx.Notebook` in barrierefreie Registerkarten (&Excel, &Word und &PowerPoint) unterteilt, was die Navigation mit Bildschirmlesern erheblich verbessert und lange Bildlauflisten überflüssig macht. Sie können schnell mit `Alt+E`, `Alt+W`, `Alt+P` oder den Standard-Tastenkombinationen `Strg+BildAb`/`Strg+BildAuf` zwischen den Registerkarten wechseln.
* **NVDA 2026.2-Kompatibilität:** Getestet und zertifiziert für NVDA 2026.2."""
    },
    "pt": {
        "ui": {
            "Enable &Bulk Sheet Organizer and Quick Sheet Mover": "Ativar o Organizador de planilhas em &massa e o Movimentador rápido de planilha",
            "Explicitly announce mer&ged cells": "Anunciar explicitamente células mescl&adas",
            "&Excel": "&Excel",
            "Prevent NVDA &crashes in Word text fields (SafeRichEdit)": "Evitar f&alhas do NVDA em campos de texto do Word (SafeRichEdit)",
            "&Word": "&Word",
            "Read hidden &hex codes when navigating the Standard Color hexagon grid": "Ler códigos &hex ocultos ao navegar na grade hexagonal de Cores Padrão",
            "Ensure the Hex Color edit &field is properly labeled": "Garantir que o &campo de edição de Cor Hex esteja rotulado corretamente",
            "Enable &Slide Layout Analyzer (Prefix + L)": "Ativar o Analisador de &layout de slide (Prefix + L)",
            "Shape Movement &Audio Mode:": "Modo de &áudio de movimento de formas:",
            "&PowerPoint": "&PowerPoint"
        },
        "desc_intro": "O BOA é uma poderosa suíte de aprimoramentos de acessibilidade para o Microsoft Office, resultado de desenvolvimento assistido por IA; projetada para melhorar drasticamente a experiência de leitor de tela para usuários do NVDA. Toda a base de código foi gerada usando Anti Gravity 2.0. Ele corrige diretamente componentes de interface inacessíveis e introduz ferramentas de navegação rápida para Excel e PowerPoint.",
        "cl_201": """### Version 2.0.1
#### Melhorias de UX/UI
* **Diálogo de configurações em abas:** Reorganização do painel de configurações do BOA em abas acessíveis (&Excel, &Word e &PowerPoint) usando `wx.Notebook`, melhorando consideravelmente a navegação do leitor de tela e eliminando longas listas de rolagem.
* **Compatibilidade com o NVDA 2026.2:** Testado e certificado para o NVDA 2026.2.""",
        "readme_201": """### Version 2.0.1
#### Melhorias de UX/UI
* **Diálogo de configurações em abas:** Reorganização do painel de configurações do BOA em abas acessíveis (&Excel, &Word e &PowerPoint) usando `wx.Notebook`, melhorando consideravelmente a navegação do leitor de tela e eliminando longas listas de rolagem. Você pode alternar rapidamente entre as abas usando `Alt+E`, `Alt+W`, `Alt+P` ou os atalhos padrão `Ctrl+PageDown`/`Ctrl+PageUp`.
* **Compatibilidade com o NVDA 2026.2:** Testado e certificado para o NVDA 2026.2."""
    },
    "it": {
        "ui": {
            "Enable &Bulk Sheet Organizer and Quick Sheet Mover": "Abilita l'Organizzatore di fogli in &massa e lo Spostamento rapido del foglio",
            "Explicitly announce mer&ged cells": "Annuncia esplicitamente le celle &unite",
            "&Excel": "&Excel",
            "Prevent NVDA &crashes in Word text fields (SafeRichEdit)": "Previeni i &blocchi di NVDA nei campi di testo di Word (SafeRichEdit)",
            "&Word": "&Word",
            "Read hidden &hex codes when navigating the Standard Color hexagon grid": "Leggi i codici &hex nascosti durante la navigazione nella griglia esagonale Colori Standard",
            "Ensure the Hex Color edit &field is properly labeled": "Assicurati che il &campo di modifica del Colore Hex sia etichettato correttamente",
            "Enable &Slide Layout Analyzer (Prefix + L)": "Abilita l'Analizzatore di &layout delle diapositive (Prefix + L)",
            "Shape Movement &Audio Mode:": "Modalità &audio di movimento delle forme:",
            "&PowerPoint": "&PowerPoint"
        },
        "desc_intro": "BOA è una potente suite di miglioramenti dell'accessibilità per Microsoft Office, frutto dello sviluppo assistito da IA; progettata per migliorare notevolmente l'esperienza dello screen reader per gli utenti di NVDA. L'intero codice sorgente è stato generato utilizzando Anti Gravity 2.0. Corregge direttamente i componenti dell'interfaccia utente inaccessibili e introduce strumenti di navigazione rapida per Excel e PowerPoint.",
        "cl_201": """### Version 2.0.1
#### Miglioramenti UX/UI
* **Finestra delle impostazioni a schede:** Il pannello delle impostazioni di BOA è stato riorganizzato in schede accessibili (&Excel, &Word e &PowerPoint) tramite `wx.Notebook`, migliorando notevolmente la navigazione con lo screen reader ed eliminando lunghi elenchi a scorrimento.
* **Compatibilità con NVDA 2026.2:** Testato e certificato per NVDA 2026.2.""",
        "readme_201": """### Version 2.0.1
#### Miglioramenti UX/UI
* **Finestra delle impostazioni a schede:** Il pannello delle impostazioni di BOA è stato riorganizzato in schede accessibili (&Excel, &Word e &PowerPoint) tramite `wx.Notebook`, migliorando notevolmente la navigazione con lo screen reader ed eliminando lunghi elenchi a scorrimento. È possibile passare rapidamente da una scheda all'altra premendo `Alt+E`, `Alt+W`, `Alt+P` o tramite le scorciatoie standard `Ctrl+PagGiù`/`Ctrl+PagSu`.
* **Compatibilità con NVDA 2026.2:** Testato e certificato per NVDA 2026.2."""
    },
    "ru": {
        "ui": {
            "Enable &Bulk Sheet Organizer and Quick Sheet Mover": "Включить &массовый органайзер листов и быстрое перемещение листа",
            "Explicitly announce mer&ged cells": "Явно объявлять &объединенные ячейки",
            "&Excel": "&Excel",
            "Prevent NVDA &crashes in Word text fields (SafeRichEdit)": "Предотвращать &сбои NVDA в текстовых полях Word (SafeRichEdit)",
            "&Word": "&Word",
            "Read hidden &hex codes when navigating the Standard Color hexagon grid": "Читать скрытые &hex-коды при навигации по шестиугольной сетке стандартных цветов",
            "Ensure the Hex Color edit &field is properly labeled": "Убедиться, что &поле ввода Hex-цвета правильно подписано",
            "Enable &Slide Layout Analyzer (Prefix + L)": "Включить анализатор &макета слайда (Prefix + L)",
            "Shape Movement &Audio Mode:": "Звуковой &режим перемещения фигур:",
            "&PowerPoint": "&PowerPoint"
        },
        "desc_intro": "BOA — это мощный набор средств улучшения доступности для Microsoft Office, созданный с помощью разработки на базе ИИ; разработан для значительного улучшения работы пользователей NVDA с программами чтения с экрана. Вся кодовая база создана с помощью Anti Gravity 2.0. Он напрямую исправляет недоступные элементы интерфейса и добавляет инструменты быстрой навигации для Excel и PowerPoint.",
        "cl_201": """### Version 2.0.1
#### Улучшения UX/UI
* **Диалог настроек с вкладками:** Панель настроек BOA реорганизована в удобные вкладки (&Excel, &Word и &PowerPoint) с помощью `wx.Notebook`, что значительно улучшает навигацию для программ чтения с экрана и устраняет длинные списки прокрутки.
* **Совместимость с NVDA 2026.2:** Протестировано и сертифицировано для NVDA 2026.2.""",
        "readme_201": """### Version 2.0.1
#### Улучшения UX/UI
* **Диалог настроек с вкладками:** Панель настроек BOA реорганизована в удобные вкладки (&Excel, &Word и &PowerPoint) с помощью `wx.Notebook`, что значительно улучшает навигацию для программ чтения с экрана и устраняет длинные списки прокрутки. Вы можете быстро переключаться между вкладками с помощью `Alt+E`, `Alt+W`, `Alt+P` или стандартных сочетаний `Ctrl+PageDown`/`Ctrl+PageUp`.
* **Совместимость с NVDA 2026.2:** Протестировано и сертифицировано для NVDA 2026.2."""
    },
    "pl": {
        "ui": {
            "Enable &Bulk Sheet Organizer and Quick Sheet Mover": "Włącz &masowy organizator arkuszy i szybkie przenoszenie arkusza",
            "Explicitly announce mer&ged cells": "Wyraźnie zapowiadaj s&calone komórki",
            "&Excel": "&Excel",
            "Prevent NVDA &crashes in Word text fields (SafeRichEdit)": "Zapobiegaj a&wariom NVDA w polach tekstowych Worda (SafeRichEdit)",
            "&Word": "&Word",
            "Read hidden &hex codes when navigating the Standard Color hexagon grid": "Odczytuj ukryte kody &hex podczas poruszania się po siatce kolorów standardowych",
            "Ensure the Hex Color edit &field is properly labeled": "Upewnij się, że &pole edycji koloru Hex jest prawidłowo etykietowane",
            "Enable &Slide Layout Analyzer (Prefix + L)": "Włącz analizator &układu slajdu (Prefix + L)",
            "Shape Movement &Audio Mode:": "Tryb &dźwiękowy przesuwania kształtów:",
            "&PowerPoint": "&PowerPoint"
        },
        "desc_intro": "BOA to zaawansowany pakiet ułatwień dostępu dla pakietu Microsoft Office, stworzony z pomocą sztucznej inteligencji; zaprojektowany z myślą o znacznym ułatwieniu pracy użytkownikom czytnika ekranu NVDA. Cały kod został wygenerowany przy użyciu Anti Gravity 2.0. Bezpośrednio naprawia niedostępne elementy interfejsu i wprowadza narzędzia szybkiej nawigacji w programach Excel i PowerPoint.",
        "cl_201": """### Version 2.0.1
#### Ulepszenia UX/UI
* **Okno ustawień z kartami:** Przebudowano panel ustawień BOA na dostępne karty (&Excel, &Word i &PowerPoint) za pomocą `wx.Notebook`, co znacznie poprawia nawigację czytnikiem ekranu i eliminuje długie listy przewijania.
* **Zgodność z NVDA 2026.2:** Przetestowano i certyfikowano pod kątem zgodności z NVDA 2026.2.""",
        "readme_201": """### Version 2.0.1
#### Ulepszenia UX/UI
* **Okno ustawień z kartami:** Przebudowano panel ustawień BOA na dostępne karty (&Excel, &Word i &PowerPoint) za pomocą `wx.Notebook`, co znacznie poprawia nawigację czytnikiem ekranu i eliminuje długie listy przewijania. Możesz szybko przełączać się między kartami za pomocą `Alt+E`, `Alt+W`, `Alt+P` lub standardowych skrótów `Ctrl+PageDown`/`Ctrl+PageUp`.
* **Zgodność z NVDA 2026.2:** Przetestowano i certyfikowano pod kątem zgodności z NVDA 2026.2."""
    },
    "uk": {
        "ui": {
            "Enable &Bulk Sheet Organizer and Quick Sheet Mover": "Увімкнути &масовий організатор аркушів та швидке переміщення аркуша",
            "Explicitly announce mer&ged cells": "Чітко оголошувати о&б'єднані клітинки",
            "&Excel": "&Excel",
            "Prevent NVDA &crashes in Word text fields (SafeRichEdit)": "Запобігати з&боям NVDA у текстових полях Word (SafeRichEdit)",
            "&Word": "&Word",
            "Read hidden &hex codes when navigating the Standard Color hexagon grid": "Читати приховані &hex-коди під час навігації шестикутною сіткою стандартних кольорів",
            "Ensure the Hex Color edit &field is properly labeled": "Переконатися, що &поле редагування кольору Hex має правильний підпис",
            "Enable &Slide Layout Analyzer (Prefix + L)": "Увімкнути аналізатор &макета слайда (Prefix + L)",
            "Shape Movement &Audio Mode:": "Звуковий &режим переміщення фігур:",
            "&PowerPoint": "&PowerPoint"
        },
        "desc_intro": "BOA — це потужний набір засобів покращення доступності для Microsoft Office, створений за допомогою розробки на базі ШІ; покликаний кардинально покращити досвід роботи з екранним читачем для користувачів NVDA. Весь код згенеровано за допомогою Anti Gravity 2.0. Він безпосередньо виправляє недоступні елементи інтерфейсу та додає інструменти швидкої навігації для Excel і PowerPoint.",
        "cl_201": """### Version 2.0.1
#### Покращення UX/UI
* **Діалогове вікно налаштувань із вкладками:** Панель налаштувань BOA реорганізовано у зручні вкладки (&Excel, &Word та &PowerPoint) за допомогою `wx.Notebook`, що значно покращує навігацію для програм зчитування з екрана та усуває довгі списки прокручування.
* **Сумісність з NVDA 2026.2:** Протестовано та сертифіковано для NVDA 2026.2.""",
        "readme_201": """### Version 2.0.1
#### Покращення UX/UI
* **Діалогове вікно налаштувань із вкладками:** Панель налаштувань BOA реорганізовано у зручні вкладки (&Excel, &Word та &PowerPoint) за допомогою `wx.Notebook`, що значно покращує навігацію для програм зчитування з екрана та усуває довгі списки прокручування. Ви можете швидко перемикатися між вкладками за допомогою `Alt+E`, `Alt+W`, `Alt+P` або стандартних комбінацій `Ctrl+PageDown`/`Ctrl+PageUp`.
* **Сумісність з NVDA 2026.2:** Протестовано та сертифіковано для NVDA 2026.2."""
    }
}

def apply_batch(data):
    pot = polib.pofile("BOA.pot")
    pot_desc = [e for e in pot if "Key Features:" in e.msgid][0].msgid
    pot_cl = [e for e in pot if "### Version 2.0.1" in e.msgid][0].msgid
    
    for lang, info in data.items():
        po_path = os.path.join("addon", "locale", lang, "LC_MESSAGES", "nvda.po")
        if not os.path.exists(po_path):
            print(f"[{lang}] PO file missing: {po_path}")
            continue
            
        po = polib.pofile(po_path)
        
        # 1. Update UI strings
        for msgid, msgstr in info["ui"].items():
            entry = po.find(msgid)
            if entry:
                entry.msgstr = msgstr
                if "fuzzy" in entry.flags:
                    entry.flags.remove("fuzzy")
            else:
                po.append(polib.POEntry(msgid=msgid, msgstr=msgstr))
                
        # 2. Update Description
        old_desc_entry = [e for e in po if "Key Features:" in e.msgid]
        if old_desc_entry:
            old_desc_str = old_desc_entry[0].msgstr
            # Keep features part of old translation
            features_part = ""
            for marker in ["\nKey Features:", "\n    Key Features:", "\n    Caractéristiques principales :", "\n    Características principales:", "\n    Hauptmerkmale:", "\n    Caratteristiche principali:", "\n    Основные возможности:", "\n    Główne funkcje:", "\n    Основні можливості:"]:
                if marker in old_desc_str:
                    features_part = old_desc_str[old_desc_str.find(marker):]
                    break
            if not features_part:
                idx = old_desc_str.find("- ")
                if idx != -1:
                    features_part = "\n" + old_desc_str[idx-4:]
                else:
                    features_part = "\n" + old_desc_str
            new_desc_str = info["desc_intro"].strip() + "\n" + features_part.strip()
        else:
            new_desc_str = info["desc_intro"]
            
        desc_entry = po.find(pot_desc)
        if desc_entry:
            desc_entry.msgstr = new_desc_str
            if "fuzzy" in desc_entry.flags:
                desc_entry.flags.remove("fuzzy")
        else:
            po.append(polib.POEntry(msgid=pot_desc, msgstr=new_desc_str))
            
        # 3. Update Changelog
        old_cl_entry = [e for e in po if "### Version 2.0.0" in e.msgid and e.msgstr]
        if old_cl_entry:
            existing_cl = old_cl_entry[0].msgstr
            new_cl_str = info["cl_201"].strip() + "\n\n" + existing_cl.strip()
        else:
            new_cl_str = info["cl_201"]
            
        cl_entry = po.find(pot_cl)
        if cl_entry:
            cl_entry.msgstr = new_cl_str
            if "fuzzy" in cl_entry.flags:
                cl_entry.flags.remove("fuzzy")
        else:
            po.append(polib.POEntry(msgid=pot_cl, msgstr=new_cl_str))
            
        po.save(po_path)
        print(f"[{lang}] Successfully updated nvda.po")
        
        # 4. Update Readme
        readme_path = os.path.join("addon", "doc", lang, "readme.md")
        if os.path.exists(readme_path):
            with open(readme_path, "r", encoding="utf-8") as f:
                content = f.read()
            if "Version 2.0.1" not in content:
                for target in ["### Version 2.0.0", "## 🛠️ Changelog", "## Changelog"]:
                    if target in content:
                        content = content.replace(target, f"{info['readme_201'].strip()}\n\n{target}", 1)
                        break
                with open(readme_path, "w", encoding="utf-8") as f:
                    f.write(content)
                print(f"[{lang}] Successfully updated readme.md")
            else:
                print(f"[{lang}] readme.md already has Version 2.0.1")

if __name__ == "__main__":
    apply_batch(translations)
