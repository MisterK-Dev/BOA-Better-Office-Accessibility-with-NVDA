import os
import re

pot_path = "BOA.pot"
po_path = "addon/locale/hi/LC_MESSAGES/nvda.po"
en_doc_path = "addon/doc/en/readme.md"
hi_doc_path = "addon/doc/hi/readme.md"

os.makedirs(os.path.dirname(po_path), exist_ok=True)
os.makedirs(os.path.dirname(hi_doc_path), exist_ok=True)

translations = {
    "BOA Office Enhancements": "BOA ऑफिस एन्हांसमेंट्स",
    "Off": "बंद",
    "One-time Announcement": "एक बार की घोषणा",
    "Guided Announcement": "निर्देशित घोषणा",
    "Strict Memory Check (CountA)": "सख्त मेमोरी जांच (CountA)",
    "Visible Data Only (Math Engine)": "केवल दृश्य डेटा (गणित इंजन)",
    "Select the engine used to detect the end of data in a row or column.": "किसी पंक्ति या कॉलम में डेटा के अंत का पता लगाने के लिए उपयोग किए जाने वाले इंजन का चयन करें।",
    "&Excel: Enable Bulk Sheet Organizer and Quick Sheet Mover": "&Excel: बल्क शीट ऑर्गेनाइज़र और क्विक शीट मूवर सक्षम करें",
    "Use accessible Sheet &Rename dialog instead of native edit field": "मूल संपादन फ़ील्ड के बजाय सुलभ शीट &Rename (नाम बदलें) डायलॉग का उपयोग करें",
    "Prevent NVDA &crashes in Excel text fields": "Excel टेक्स्ट फ़ील्ड्स में NVDA क्रैश को रोकें (&c)",
    "Announce when a m&ulti-cell selection is unexpectedly lost": "जब एक बहु-सेल (m&ulti-cell) चयन अप्रत्याशित रूप से खो जाता है तो घोषणा करें",
    "Proactively announce when navigating past &hidden rows or columns": "छिपी हुई पंक्तियों या कॉलम (&hidden) के पार नेविगेट करते समय सक्रिय रूप से घोषणा करें",
    "Enable Sheet &Layout Analyzer via NVDA+E, L": "NVDA+E, L के माध्यम से शीट &Layout विश्लेषक सक्षम करें",
    "Sheet Layout Auto-Announce &Mode:": "शीट लेआउट ऑटो-अनाउंस &Mode:",
    "Conditional &Formatting and color": "सशर्त &Formatting और रंग",
    "Enable Cell Moni&tor (slots 1-9 and continuous background monitoring)": "सेल मॉनिटर सक्षम करें (स्लॉट 1-9 और निरंतर पृष्ठभूमि निगरानी) (&t)",
    "Announce when there is no more data in the direction you are &navigating": "घोषणा करें जब आपके नेविगेट करने की दिशा में कोई और डेटा न हो (&n)",
    "Excel Enhancements": "Excel एन्हांसमेंट्स",
    "&PowerPoint: Read hidden Hex codes when navigating the Standard Color hexagon grid": "&PowerPoint: मानक रंग षट्भुज ग्रिड पर नेविगेट करते समय छिपे हुए हेक्स कोड पढ़ें",
    "Ensure the He&x Color edit field is properly labeled": "सुनिश्चित करें कि हेक्स (He&x) रंग संपादन फ़ील्ड ठीक से लेबल किया गया है",
    "Ensure the R&GB Color edit fields are properly labeled": "सुनिश्चित करें कि R&GB रंग संपादन फ़ील्ड ठीक से लेबल किए गए हैं",
    "Prevent NVDA cr&ashes in PowerPoint text fields": "PowerPoint टेक्स्ट फ़ील्ड्स में NVDA क्रैश (&a) को रोकें",
    "PowerPoint Enhancements": "PowerPoint एन्हांसमेंट्स",
    "&Word: Prevent NVDA crashes in Word text fields": "&Word: Word टेक्स्ट फ़ील्ड्स में NVDA क्रैश को रोकें",
    "Word Enhancements": "Word एन्हांसमेंट्स",
    "BOA: Better Office Accessibility": "BOA: बेहतर ऑफिस एक्सेसिबिलिटी"
}

# Add long strings
translations["""A powerful suite of accessibility enhancements for Microsoft Office, designed to vastly improve the screen reader experience for NVDA users.\\nI recommend to check whats new for complete change log and help file for features. just given below some details for refference. now onwards i avoide updating description section.\\nNew in v1.4.0:\\n- Cell Monitor : Press NVDA+E, Shift+1 to 9 to map specific cells to memory slots. You can jump back and read them anytime using NVDA+E and the number.\\n- Continuous Background Monitoring: Slotted cells are automatically monitored. If Excel triggers a recalculation or cell edit, BOA instantly announces the new value. Toggle manually with NVDA+E then M. Clear all with NVDA+E then Backspace.\\n\\nNew in v1.3.0:\\n- Sheet Layout Analyzer: Press NVDA+E, L to instantly scan any Excel worksheet for Data Blocks, Hidden Tabs, Active Filters, Protected states, and Hidden Edge Borders.\\n- Conditional Formatting Announcer: Automatically reads the dynamic visual color, font, and background shade of cells changed by Excel's Conditional Formatting rules.\\n- Intelligent Accelerator Keys: The BOA Settings panel has been deeply integrated with NVDA's GUI architecture. Every single feature now has a unique, explicit Alt-key shortcut.\\n\\nNew in v1.2.0:\\n- Lightning Fast App-Launch Caching: Core modules lazy-load exactly when you focus on Office, eliminating boot lag.\\n- Enhanced Hidden Cell Tracker: Redesigned with 1D COM Math to instantly detect crossed hidden rows or columns without freezing NVDA.\\n- Intelligent Process Tracking: Resolves false \\\"Sheet hidden\\\" alarms when reopening fresh workbooks.\\n\\nExcel Features:\\n- Bulk Sheet Organizer (NVDA+Alt+C): Instantly reorder multiple sheets at once using a fully accessible dialog.\\n- Quick Sheet Mover: Move the active sheet instantly via NVDA+Shift+Arrows (Left/Right) or Home/End.\\n- Accessible Sheet Renaming: Bypasses the inaccessible Excel rename edit field with a reliable dialog.\\n- Smart Selection Tracking: Accurately announces multi-cell range selections and deselections.\\n\\nPowerPoint Features:\\n- Accessible Color Pickers: Enables NVDA to accurately read RGB and Hex values inside the Custom Color dialog.\\n- Standard Color Grid Support: Intercepts arrow keys to read hidden color Hex codes from the inaccessible color hexagon."""] = """Microsoft Office के लिए एक्सेसिबिलिटी एन्हांसमेंट्स का एक शक्तिशाली सुइट, जो NVDA उपयोगकर्ताओं के लिए स्क्रीन रीडर अनुभव को बेहतर बनाने के लिए डिज़ाइन किया गया है।\\nमैं संपूर्ण चेंजलॉग के लिए व्हाट्स न्यू और सुविधाओं के लिए सहायता फ़ाइल की जांच करने की अनुशंसा करता हूं। संदर्भ के लिए नीचे कुछ विवरण दिए गए हैं। अब से मैं विवरण अनुभाग को अपडेट करने से बचता हूँ।\\nv1.4.0 में नया:\\n- सेल मॉनिटर: विशिष्ट सेल को मेमोरी स्लॉट में मैप करने के लिए NVDA+E, Shift+1 से 9 दबाएँ। आप NVDA+E और संख्या का उपयोग करके कभी भी वापस जा सकते हैं और उन्हें पढ़ सकते हैं।\\n- निरंतर पृष्ठभूमि निगरानी: स्लॉटेड सेल की पृष्ठभूमि में स्वचालित रूप से निगरानी की जाती है। यदि Excel पुनर्गणना या सेल संपादन को ट्रिगर करता है, तो BOA तुरंत नए मान की घोषणा करता है। NVDA+E फिर M के साथ मैन्युअल रूप से टॉगल करें। NVDA+E फिर Backspace के साथ सभी साफ़ करें।\\n\\nv1.3.0 में नया:\\n- शीट लेआउट विश्लेषक: डेटा ब्लॉक्स, छिपे हुए टैब्स, सक्रिय फ़िल्टर्स, संरक्षित स्थितियों और छिपे हुए किनारे की सीमाओं के लिए किसी भी Excel वर्कशीट को तुरंत स्कैन करने के लिए NVDA+E, L दबाएँ।\\n- सशर्त स्वरूपण उद्घोषक: Excel के सशर्त स्वरूपण नियमों द्वारा परिवर्तित सेल के गतिशील दृश्य रंग, फ़ॉन्ट और पृष्ठभूमि छाया को स्वचालित रूप से पढ़ता है।\\n- इंटेलिजेंट एक्सेलेरेटर कुंजियाँ: BOA सेटिंग्स पैनल को NVDA की GUI वास्तुकला के साथ गहराई से एकीकृत किया गया है। अब हर एक सुविधा में एक अद्वितीय, स्पष्ट Alt-key शॉर्टकट है।\\n\\nv1.2.0 में नया:\\n- लाइटनिंग फास्ट ऐप-लॉन्च कैशिंग: कोर मॉड्यूल बिल्कुल तभी लेज़ी-लोड होते हैं जब आप ऑफिस पर फ़ोकस करते हैं, जिससे बूट लैग समाप्त हो जाता है।\\n- उन्नत हिडन सेल ट्रैकर: NVDA को फ़्रीज़ किए बिना पार की गई छिपी हुई पंक्तियों या कॉलम का तुरंत पता लगाने के लिए 1D COM गणित के साथ फिर से डिज़ाइन किया गया।\\n- इंटेलिजेंट प्रोसेस ट्रैकिंग: ताज़ा वर्कबुक को फिर से खोलने पर झूठे \\\"शीट छिपी हुई\\\" अलार्म को हल करता है।\\n\\nExcel सुविधाएँ:\\n- बल्क शीट ऑर्गेनाइज़र (NVDA+Alt+C): पूरी तरह से सुलभ डायलॉग का उपयोग करके एक साथ कई शीट को तुरंत पुनर्व्यवस्थित करें।\\n- क्विक शीट मूवर: सक्रिय शीट को तुरंत NVDA+Shift+Arrows (बाएँ/दाएँ) या Home/End के माध्यम से ले जाएँ।\\n- सुलभ शीट का नाम बदलना: एक विश्वसनीय डायलॉग के साथ अगम्य Excel नाम बदलें संपादन फ़ील्ड को बायपास करता है।\\n- स्मार्ट चयन ट्रैकिंग: बहु-सेल श्रेणी चयनों और अचयन की सटीक घोषणा करता है।\\n\\nPowerPoint सुविधाएँ:\\n- सुलभ कलर पिकर्स: NVDA को कस्टम कलर डायलॉग के अंदर RGB और Hex मानों को सटीक रूप से पढ़ने में सक्षम बनाता है।\\n- स्टैंडर्ड कलर ग्रिड सपोर्ट: अगम्य कलर हेक्सागोन से छिपे हुए कलर हेक्स कोड को पढ़ने के लिए एरो कीज़ को रोकता है।"""

translations["""## What's New in v1.6.0-dev1\\n* **Comprehensive Code Governance**: Strict implementation of copyright headers and GPL-2.0 across the entire codebase.\\n* **Unified Build Process**: Fully autonomous `scons` build system generating local translation manifests and documentation.\\n* **I18n Translation Foundation**: The entire add-on has been strictly prepared for multi-language gettext translation."""] = """## v1.6.0-dev1 में नया क्या है\\n* **व्यापक कोड गवर्नेंस**: संपूर्ण कोडबेस में कॉपीराइट हेडर और GPL-2.0 का सख्त कार्यान्वयन।\\n* **यूनिफाइड बिल्ड प्रोसेस**: पूरी तरह से स्वायत्त `scons` बिल्ड सिस्टम जो स्थानीय अनुवाद मेनिफेस्ट और दस्तावेज़ उत्पन्न करता है।\\n* **I18n अनुवाद फाउंडेशन**: संपूर्ण ऐड-ऑन को बहु-भाषा गेटटेक्स्ट अनुवाद के लिए सख्ती से तैयार किया गया है।"""

with open(pot_path, "r", encoding="utf-8") as f:
    pot_content = f.read()

# Replace metadata
pot_content = pot_content.replace("#, fuzzy\\n", "")
pot_content = pot_content.replace("Language-Team: LANGUAGE <LL@li.org>\\n", "Language-Team: Hindi <hi@li.org>\\n")
pot_content = pot_content.replace("Language: \\n", "Language: hi\\n")
pot_content = pot_content.replace("charset=CHARSET", "charset=utf-8")

# Parse msgids and replace msgstrs
lines = pot_content.split("\\n")
out_lines = []
i = 0
while i < len(lines):
    line = lines[i]
    if line.startswith("msgid "):
        msgid_str = line[6:].strip('"')
        
        # Accumulate multi-line msgid
        j = i + 1
        while j < len(lines) and lines[j].startswith('"') and not lines[j].startswith('msgstr'):
            msgid_str += lines[j].strip('"')
            j += 1
            
        # Write msgid lines
        while i < j:
            out_lines.append(lines[i])
            i += 1
            
        # We are at msgstr. Find the msgid in translations
        if msgid_str in translations:
            trans_str = translations[msgid_str]
            # Write msgstr as single or multiple lines depending on length/newlines
            if "\\n" in trans_str:
                out_lines.append('msgstr ""')
                t_lines = trans_str.split("\\n")
                for k, t_line in enumerate(t_lines):
                    suffix = "\\n" if k < len(t_lines) - 1 else ""
                    out_lines.append(f'"{t_line}{suffix}"')
            else:
                out_lines.append(f'msgstr "{trans_str}"')
            
            # Skip original msgstr lines
            while i < len(lines) and lines[i].startswith('msgstr'):
                i += 1
            while i < len(lines) and lines[i].startswith('"'):
                i += 1
            continue
    
    out_lines.append(line)
    i += 1

with open(po_path, "w", encoding="utf-8") as f:
    f.write("\\n".join(out_lines))


# Now translate readme.md
with open(en_doc_path, "r", encoding="utf-8") as f:
    readme_content = f.read()

readme_translations = {
    "BOA is a powerful suite of accessibility enhancements for Microsoft Office, designed to vastly improve the screen reader experience for NVDA users. It directly patches inaccessible UI components and introduces rapid navigation tools for Excel and PowerPoint.": "BOA Microsoft Office के लिए एक्सेसिबिलिटी एन्हांसमेंट्स का एक शक्तिशाली सुइट है, जो NVDA उपयोगकर्ताओं के लिए स्क्रीन रीडर अनुभव को बेहतर बनाने के लिए डिज़ाइन किया गया है। यह सीधे अगम्य UI घटकों को पैच करता है और Excel और PowerPoint के लिए तीव्र नेविगेशन टूल पेश करता है।",
    "## ⌨️ Hotkey Reference": "## ⌨️ हॉटकी संदर्भ",
    "| Feature | Key Combination | Context / Notes |": "| सुविधा | कुंजी संयोजन | संदर्भ / नोट्स |",
    "| **Enter Command Mode** | `NVDA+E` | Activates Command Prefix Mode (triggers a high-pitched beep) |": "| **कमांड मोड दर्ज करें** | `NVDA+E` | कमांड प्रीफिक्स मोड को सक्रिय करता है (एक उच्च-पिच बीप ट्रिगर करता है) |",
    "| **Analyze Sheet Layout** | `NVDA+E`, then `L` | Run within Excel before navigating data blocks |": "| **शीट लेआउट का विश्लेषण करें** | `NVDA+E`, फिर `L` | डेटा ब्लॉक्स नेविगेट करने से पहले Excel के भीतर चलाएँ |",
    "| **Jump to Nearest Data Block** | `NVDA+E`, then `J` /  | Requires Layout Analysis first |": "| **निकटतम डेटा ब्लॉक पर जाएँ** | `NVDA+E`, फिर `J` /  | पहले लेआउट विश्लेषण की आवश्यकता है |",
    "| **Open Bulk Sheet Organizer** | `NVDA+E`, then `X` | Opens the accessible sheet reordering dialog |": "| **बल्क शीट ऑर्गेनाइज़र खोलें** | `NVDA+E`, फिर `X` | सुलभ शीट रीऑर्डरिंग डायलॉग खोलता है |",
    "| **Move Active Sheet Left** | `NVDA+Shift+LeftArrow` | Shifts the active sheet one position up|": "| **सक्रिय शीट को बाएँ ले जाएँ** | `NVDA+Shift+LeftArrow` | सक्रिय शीट को एक स्थान ऊपर खिसकाता है|",
    "| **Move Active Sheet Right** | `NVDA+Shift+RightArrow` | Shifts the active worksheet one position down|": "| **सक्रिय शीट को दाएँ ले जाएँ** | `NVDA+Shift+RightArrow` | सक्रिय वर्कशीट को एक स्थान नीचे खिसकाता है|",
    "| **Move Sheet to Start/End** | `NVDA+Shift+Home` / `End` | Sends worksheet to the absolute absolute boundaries |": "| **शीट को प्रारंभ/अंत में ले जाएँ** | `NVDA+Shift+Home` / `End` | वर्कशीट को पूर्ण सीमाओं पर भेजता है |",
    "| **Detailed Conditional Formatting**| `NVDA+E`, then `F` | Announces complete formatting details of focused cell |": "| **विस्तृत सशर्त स्वरूपण**| `NVDA+E`, फिर `F` | केंद्रित सेल के संपूर्ण स्वरूपण विवरण की घोषणा करता है |",
    "| **Map Cell to Memory Slot** | `NVDA+E`, then `Shift+1` to `Shift+9` | Assigns current cell to a background monitor slot |": "| **सेल को मेमोरी स्लॉट में मैप करें** | `NVDA+E`, फिर `Shift+1` से `Shift+9` | वर्तमान सेल को पृष्ठभूमि मॉनिटर स्लॉट में असाइन करता है |",
    "| **Read Monitored Cell Slot** | `NVDA+E`, then `1` to `9` | Recalls and reads the value of the assigned slot |": "| **मॉनिटर किया गया सेल स्लॉट पढ़ें** | `NVDA+E`, फिर `1` से `9` | असाइन किए गए स्लॉट के मान को याद करता है और पढ़ता है |",
    "| **Toggle Background Monitoring** | `NVDA+E`, then `M` | Manually toggles background calculation tracking |": "| **पृष्ठभूमि निगरानी टॉगल करें** | `NVDA+E`, फिर `M` | मैन्युअल रूप से पृष्ठभूमि गणना ट्रैकिंग टॉगल करता है |",
    "| **Clear All Memory Slots** | `NVDA+E`, then `Backspace` | Purges all saved background cell monitors |": "| **सभी मेमोरी स्लॉट साफ़ करें** | `NVDA+E`, फिर `Backspace` | सभी सहेजे गए पृष्ठभूमि सेल मॉनिटर को साफ़ करता है |",
    "| **Cancel Command Mode** | `Escape` | Exits Command Prefix Mode |": "| **कमांड मोड रद्द करें** | `Escape` | कमांड प्रीफिक्स मोड से बाहर निकलता है |",
    "## 🚀 Features": "## 🚀 सुविधाएँ",
    "### Excel Enhancements": "### Excel एन्हांसमेंट्स",
    "#### 1. Sheet Layout Analyzer & Caching": "#### 1. शीट लेआउट विश्लेषक और कैशिंग",
    "Instantly scan any Excel worksheet to understand its structure, hidden elements, and data blocks.": "किसी भी Excel वर्कशीट की संरचना, छिपे हुए तत्वों और डेटा ब्लॉक्स को समझने के लिए उसे तुरंत स्कैन करें।",
    "* **How it works:** BOA quickly scans the sheet and announces active data blocks. It also warns you about **Hidden Worksheet Tabs**, active **Filters**, **Protected Modes**, and **Hidden Outer Boundaries** (e.g., if columns near the right edge of the sheet are hidden, preventing you from missing off-screen data).": "* **यह कैसे काम करता है:** BOA जल्दी से शीट को स्कैन करता है और सक्रिय डेटा ब्लॉक्स की घोषणा करता है। यह आपको **छिपे हुए वर्कशीट टैब्स**, सक्रिय **फ़िल्टर्स**, **संरक्षित मोड**, और **छिपी हुई बाहरी सीमाओं** के बारे में भी चेतावनी देता है (उदाहरण के लिए, यदि शीट के दाहिने किनारे के पास के कॉलम छिपे हुए हैं, तो यह आपको ऑफ-स्क्रीन डेटा से चूकने से रोकता है)।",
    "* **Data Navigation:** After scanning, you can use the data block jump hotkeys to instantly warp your cursor between discovered data blocks, effortlessly bypassing thousands of empty cells.": "* **डेटा नेविगेशन:** स्कैन करने के बाद, आप डेटा ब्लॉक जंप हॉटकी का उपयोग करके अपने कर्सर को खोजे गए डेटा ब्लॉक्स के बीच तुरंत ले जा सकते हैं, जिससे हजारों खाली सेल आसानी से बायपास हो जाते हैं।",
    "#### 2. Bulk Sheet Organizer": "#### 2. बल्क शीट ऑर्गेनाइज़र",
    "Instantly reorder and arrange multiple sheets at once using a fully accessible dialog .": "पूरी तरह से सुलभ डायलॉग का उपयोग करके एक साथ कई शीट को तुरंत पुनर्व्यवस्थित और व्यवस्थित करें।",
    "* **How it works:** Opens a dialog where you can select a sheet and map it to a new position. Scheduled moves are listed in a data table (press `Del` to remove a mistake). Click `OK` and your workbook is rearranged instantly.": "* **यह कैसे काम करता है:** एक डायलॉग खोलता है जहाँ आप एक शीट का चयन कर सकते हैं और इसे एक नई स्थिति में मैप कर सकते हैं। निर्धारित चालें एक डेटा तालिका में सूचीबद्ध हैं (किसी गलती को हटाने के लिए `Del` दबाएं)। `OK` पर क्लिक करें और आपकी वर्कबुक तुरंत पुनर्व्यवस्थित हो जाएगी।",
    "#### 3. Quick Sheet Mover": "#### 3. क्विक शीट मूवर",
    "Move the active sheet left, right, to the very beginning, or to the very end instantly using your keyboard shortcuts.": "अपने कीबोर्ड शॉर्टकट का उपयोग करके सक्रिय शीट को बाएँ, दाएँ, बिल्कुल शुरुआत में, या बिल्कुल अंत में तुरंत ले जाएँ।",
    "#### 4. Accessible Sheet Renaming": "#### 4. सुलभ शीट का नाम बदलना",
    "* When renaming a sheet, NVDA natively struggles to read the characters you are typing.": "* किसी शीट का नाम बदलते समय, NVDA स्वाभाविक रूप से आपके द्वारा टाइप किए जा रहे वर्णों को पढ़ने के लिए संघर्ष करता है।",
    "* BOA injects a custom `ExcelSheetRenameEdit` class that uses the `SafeRichEdit` engine, meaning you can precisely read by character, word, or line while renaming. This serves as an enhancement to the existing default renaming behavior.": "* BOA एक कस्टम `ExcelSheetRenameEdit` क्लास इंजेक्ट करता है जो `SafeRichEdit` इंजन का उपयोग करता है, जिसका अर्थ है कि नाम बदलते समय आप वर्ण, शब्द या लाइन के अनुसार सटीक रूप से पढ़ सकते हैं। यह मौजूदा डिफ़ॉल्ट नाम बदलने के व्यवहार के लिए एक संवर्द्धन के रूप में कार्य करता है।",
    "#### 5. Hidden Row/Column Tracker": "#### 5. हिडन पंक्ति/कॉलम ट्रैकर",
    "* Proactively tracks your movement across the grid to prevent you from missing hidden or filtered data.": "* आपको छिपे हुए या फ़िल्टर किए गए डेटा को चूकने से रोकने के लिए ग्रिड में आपकी गति को सक्रिय रूप से ट्रैक करता है।",
    '* **Crossed Fragmented Cells:** If you jump across a heavily fragmented or hidden section of the grid (e.g., moving from Row 3 to Row 10 because Rows 4–9 are hidden), BOA explicitly announces "Rows 4 through 9 hidden". This ensures you always know when data has been skipped in the structure.': '* **क्रॉस्ड फ़्रेग्मेंटेड सेल:** यदि आप ग्रिड के अत्यधिक खंडित या छिपे हुए भाग में कूदते हैं (उदा., पंक्ति 3 से पंक्ति 10 तक जाना क्योंकि पंक्तियाँ 4-9 छिपी हुई हैं), तो BOA स्पष्ट रूप से घोषणा करता है "Rows 4 through 9 hidden"। यह सुनिश्चित करता है कि आपको हमेशा पता हो कि संरचना में डेटा कब छोड़ा गया है।',
    "#### 6. Conditional Formatting Announcer": "#### 6. सशर्त स्वरूपण उद्घोषक",
    "* Automatically reads the color, font style, and background shade of cells that have been dynamically changed by Excel's Conditional Formatting rules.": "* Excel के सशर्त स्वरूपण नियमों द्वारा गतिशील रूप से परिवर्तित किए गए सेल के रंग, फ़ॉन्ट शैली और पृष्ठभूमि छाया को स्वचालित रूप से पढ़ता है।",
    '* Gives you the true visual state of the cell rather than just the raw underlying value. Initially, when focusing on the cell, it announces "has conditional formatting, and some other minor details". For comprehensive info, use the detailed hotkey configuration which is NVDA E and F.': '* आपको केवल कच्चे अंतर्निहित मूल्य के बजाय सेल की वास्तविक दृश्य स्थिति देता है। प्रारंभ में, सेल पर ध्यान केंद्रित करते समय, यह "has conditional formatting, and some other minor details" की घोषणा करता है। व्यापक जानकारी के लिए, विस्तृत हॉटकी कॉन्फ़िगरेशन का उपयोग करें जो NVDA E और F है।',
    "#### 7. Better selection announcement": "#### 7. बेहतर चयन उद्घोषणा",
    "reads if cell or range selected or unselected.": "पढ़ता है कि क्या सेल या श्रेणी का चयन किया गया है या अचयनित किया गया है।",
    "#### 8 Cell monitor:": "#### 8 सेल मॉनिटर:",
    "* **Cell Monitor:** Use command paths to map specific cells to memory slots. You can jump back and read them anytime using the assigned numerical slot.": "* **सेल मॉनिटर:** विशिष्ट सेल को मेमोरी स्लॉट में मैप करने के लिए कमांड पथों का उपयोग करें। आप असाइन किए गए संख्यात्मक स्लॉट का उपयोग करके कभी भी वापस जा सकते हैं और उन्हें पढ़ सकते हैं।",
    "* **Continuous Monitoring:** Slotted cells are automatically monitored in the background. If Excel triggers a recalculation or cell edit, BOA instantly announces the new value. Toggle manually or clear all via command slots.": "* **निरंतर निगरानी:** पृष्ठभूमि में स्लॉटेड सेल की स्वचालित रूप से निगरानी की जाती है। यदि Excel पुनर्गणना या सेल संपादन को ट्रिगर करता है, तो BOA तुरंत नए मान की घोषणा करता है। कमांड स्लॉट के माध्यम से मैन्युअल रूप से टॉगल करें या सभी को साफ़ करें।",
    "### PowerPoint Enhancements": "### PowerPoint एन्हांसमेंट्स",
    "#### 1. Accessible Color Pickers": "#### 1. सुलभ कलर पिकर्स",
    "* Unlocks the Custom Color dialog in PowerPoint.": "* PowerPoint में कस्टम कलर डायलॉग को अनलॉक करता है।",
    '* Identifies and explicitly reads out the "Red", "Green", and "Blue" edit boxes correctly (by overriding `PowerPointRGBEdit`).': '* "Red", "Green", और "Blue" संपादन बॉक्स को सही ढंग से पहचानता है और स्पष्ट रूप से पढ़ता है (`PowerPointRGBEdit` को ओवरराइड करके)।',
    "* Maps the previously invisible Hex input field so NVDA can read the full Hex color value cleanly.": "* पहले से अदृश्य हेक्स इनपुट फ़ील्ड को मैप करता है ताकि NVDA पूर्ण हेक्स रंग मान को स्पष्ट रूप से पढ़ सके।",
    "#### 2. Standard Color Grid Support": "#### 2. मानक रंग ग्रिड समर्थन",
    '* Navigating the PowerPoint "Standard" color hexagon grid normally reads as "Graphic" or silence.': '* PowerPoint "Standard" रंग षट्भुज ग्रिड को नेविगेट करने से आमतौर पर "Graphic" या चुप्पी पढ़ी जाती है।',
    '* BOA tracks your arrow keys across the hexagon and silently fetches the hidden color value, announcing it to you in real-time (e.g., "Color #FF0000").': '* BOA षट्भुज के पार आपके तीर कुंजियों को ट्रैक करता है और चुपचाप छिपे हुए रंग मान को प्राप्त करता है, इसे वास्तविक समय में आपके लिए घोषित करता है (उदा., "Color #FF0000")।',
    "### Infrastructure & Technical Mechanisms": "### अवसंरचना और तकनीकी तंत्र",
    "#### The Command Prefix Mode": "#### कमांड प्रीफिक्स मोड",
    "To prevent keystroke conflicts with other NVDA plugins, BOA uses a **Command Prefix Mode**:": "अन्य NVDA प्लगइन्स के साथ कीस्ट्रोक संघर्ष को रोकने के लिए, BOA एक **कमांड प्रीफिक्स मोड** का उपयोग करता है:",
    "1. Press the activation hotkey to enter Command Mode. You will hear a high-pitched beep.": "1. कमांड मोड में प्रवेश करने के लिए सक्रियण हॉटकी दबाएं। आपको एक हाई-पिच बीप सुनाई देगी।",
    "2. Press a secondary key to trigger a specific feature.": "2. किसी विशिष्ट सुविधा को ट्रिगर करने के लिए एक द्वितीयक कुंजी दबाएँ।",
    "3. If you press an invalid key, you will hear an error beep.": "3. यदि आप एक अमान्य कुंजी दबाते हैं, तो आपको एक त्रुटि बीप सुनाई देगी।",
    "#### Customization & Settings Panel": "#### अनुकूलन और सेटिंग्स पैनल",
    "* BOA features are fully modular and can be enabled or disabled at any time. Go to `NVDA Menu -> Preferences -> Settings -> BOA Office Enhancements` to toggle individual features on or off.": "* BOA सुविधाएँ पूरी तरह से मॉड्यूलर हैं और इन्हें किसी भी समय सक्षम या अक्षम किया जा सकता है। व्यक्तिगत सुविधाओं को चालू या बंद करने के लिए `NVDA Menu -> Preferences -> Settings -> BOA Office Enhancements` पर जाएँ।",
    "* **Intelligent Accelerator Keys:** Every single setting features a mathematically unique `Alt+Key` accelerator shortcut within the panel. For example, press `Alt+E` to instantly jump to the Excel group, `Alt+P` for PowerPoint, and `Alt+W` for Word.": "* **इंटेलिजेंट एक्सेलेरेटर कुंजियाँ:** प्रत्येक सेटिंग में पैनल के भीतर एक गणितीय रूप से अद्वितीय `Alt+Key` एक्सेलेरेटर शॉर्टकट होता है। उदाहरण के लिए, Excel समूह पर तुरंत जाने के लिए `Alt+E`, PowerPoint के लिए `Alt+P`, और Word के लिए `Alt+W` दबाएँ।",
    "* Settings are saved securely to a standalone JSON file (`boa_settings.json`), ensuring your core NVDA configuration is never corrupted.": "* सेटिंग्स को एक स्टैंडअलोन JSON फ़ाइल (`boa_settings.json`) में सुरक्षित रूप से सहेजा जाता है, यह सुनिश्चित करते हुए कि आपका मुख्य NVDA कॉन्फ़िगरेशन कभी दूषित नहीं होता है।",
    "* If Microsoft Office officially fixes an accessibility bug in the future, you can safely disable BOA's specific override hook without losing the rest of the addon's functionality.": "* यदि Microsoft Office भविष्य में आधिकारिक तौर पर एक्सेसिबिलिटी बग को ठीक करता है, तो आप ऐडऑन की बाकी कार्यक्षमता को खोए बिना सुरक्षित रूप से BOA के विशिष्ट ओवरराइड हुक को अक्षम कर सकते हैं।",
    "#### Security & Integration Boundaries": "#### सुरक्षा और एकीकरण सीमाएँ",
    "* Clipboard injections strictly verify window foreground process IDs to prevent leakage of data into other applications.": "* अन्य अनुप्रयोगों में डेटा के रिसाव को रोकने के लिए क्लिपबोर्ड इंजेक्शन विंडो अग्रभूमि प्रक्रिया आईडी को सख्ती से सत्यापित करते हैं।",
    '* some Custom hotkeys are fully exposed in NVDA\'s Input Gestures dialog under the "Better Office Accessibility" category.': '* कुछ कस्टम हॉटकी पूरी तरह से NVDA के इनपुट जेस्चर डायलॉग में "Better Office Accessibility" श्रेणी के अंतर्गत उजागर किए गए हैं।',
    "## 📋 Requirements": "## 📋 आवश्यकताएँ",
    "* **NVDA:** Version 2026.1.0 or later.": "* **NVDA:** संस्करण 2026.1.0 या बाद का।",
    "* **Applications:** Microsoft Excel & Microsoft PowerPoint.": "* **अनुप्रयोग:** Microsoft Excel और Microsoft PowerPoint।",
    "## 💾 Installation": "## 💾 स्थापना",
    "1. Download the latest `.nvda-addon` release file, or locate it within the native NVDA Add-on Store.": "1. नवीनतम `.nvda-addon` रिलीज़ फ़ाइल डाउनलोड करें, या इसे मूल NVDA Add-on Store के भीतर खोजें।",
    "2. if installing from file, Open the file or use `NVDA's Add-on Store -> Install from external file`.": "2. यदि फ़ाइल से इंस्टॉल कर रहे हैं, तो फ़ाइल खोलें या `NVDA's Add-on Store -> Install from external file` का उपयोग करें।",
    "3. Restart NVDA.": "3. NVDA को पुनरारंभ करें।",
    "## 🛠️ Changelog": "## 🛠️ चेंजलॉग",
    "### Version 1.5.0 ": "### संस्करण 1.5.0 ",
    "#### New Features": "#### नई सुविधाएँ",
    "##### End of Data Radar": "##### एंड ऑफ डेटा रडार",
    "When navigating through large spreadsheets, it can be difficult to tell if an empty cell means you've reached the end of a list, or if there is simply a gap in the data. The **End of Data Radar** acts as a smart perimeter check to save you from blindly arrowing through empty space.": "बड़ी स्प्रेडशीट के माध्यम से नेविगेट करते समय, यह बताना मुश्किल हो सकता है कि क्या खाली सेल का मतलब है कि आप सूची के अंत तक पहुँच गए हैं, या डेटा में कोई अंतर है। **एंड ऑफ डेटा रडार** खाली जगह से आँख बंद करके नेविगेट करने से आपको बचाने के लिए एक स्मार्ट परिधि जांच के रूप में कार्य करता है।",
    "Whenever you navigate into an empty cell, BOA instantly scans the remaining cells in your direction of travel. If there is absolutely no data left, it will proactively announce:": "जब भी आप एक खाली सेल में नेविगेट करते हैं, तो BOA तुरंत यात्रा की दिशा में शेष सेल को स्कैन करता है। यदि कोई डेटा नहीं बचा है, तो यह सक्रिय रूप से घोषणा करेगा:",
    '* *"No more data below"*': '* *"No more data below"*',
    '* *"No more data above"*': '* *"No more data above"*',
    '* *"No more data to the right"*': '* *"No more data to the right"*',
    '* *"No more data to the left"*': '* *"No more data to the left"*',
    "**Configuration Options:**": "**कॉन्फ़िगरेशन विकल्प:**",
    "You can configure this feature via `NVDA Preferences -> Settings -> BOA Office Enhancements`. Because spreadsheets can contain hidden complexities (like invisible formulas or collapsed rows), the radar provides three operating modes:": "आप इस सुविधा को `NVDA Preferences -> Settings -> BOA Office Enhancements` के माध्यम से कॉन्फ़िगर कर सकते हैं। चूँकि स्प्रेडशीट्स में छिपी हुई जटिलताएँ हो सकती हैं (जैसे अदृश्य सूत्र या ढही हुई पंक्तियाँ), रडार तीन ऑपरेटिंग मोड प्रदान करता है:",
    "1. **Off**: Disables the radar entirely.": "1. **बंद**: रडार को पूरी तरह से अक्षम कर देता है।",
    "2. **Strict Memory Check (CountA) [Default]**: The safest and fastest approach. It checks the raw memory of the spreadsheet. If it detects *anything* below you (including hidden rows, text, numbers, or invisible formulas), it stays completely silent to prevent false alarms. It only announces \"No more data\" when the remainder of the sheet is 100% mathematically blank.": "2. **सख्त मेमोरी जांच (CountA) [डिफ़ॉल्ट]**: सबसे सुरक्षित और सबसे तेज़ तरीका। यह स्प्रेडशीट की रॉ मेमोरी की जांच करता है। यदि यह आपके नीचे *कुछ भी* पता लगाता है (छिपी हुई पंक्तियों, टेक्स्ट, संख्याओं या अदृश्य सूत्रों सहित), तो यह झूठे अलार्म को रोकने के लिए पूरी तरह से चुप रहता है। यह केवल \"No more data\" की घोषणा करता है जब शीट का शेष भाग 100% गणितीय रूप से रिक्त होता है।",
    "3. **Visible Data Only (Math Engine)**: A highly advanced engine designed for complex sheets. It intelligently filters out hidden rows and invisible formulas (e.g., `=\"\"`). It will only stay silent if there are actual, visible numbers or text left in your path.": "3. **केवल दृश्य डेटा (गणित इंजन)**: जटिल शीट के लिए डिज़ाइन किया गया एक अत्यधिक उन्नत इंजन। यह समझदारी से छिपी हुई पंक्तियों और अदृश्य सूत्रों (उदा., `=\"\"`) को फ़िल्टर करता है। यह केवल तभी चुप रहेगा जब आपके पथ में वास्तविक, दृश्य संख्याएँ या टेक्स्ट बचा हो।",
    "### Version 1.4 - 2026-06-12": "### संस्करण 1.4 - 2026-06-12",
    "#### Bug Fixes": "#### बग फिक्स",
    "### Version 1.3.0 — 2026-06-05": "### संस्करण 1.3.0 — 2026-06-05",
    "*Final release.*": "*अंतिम रिलीज़।*",
    "* **Sheet Layout Analyzer:** Added powerful layout scanning infrastructure. Instantly detects Worksheet Protection, active Column Filters, Hidden Worksheet Tabs, and hidden absolute borders while caching discovered data blocks.": "* **शीट लेआउट विश्लेषक:** शक्तिशाली लेआउट स्कैनिंग अवसंरचना जोड़ी गई। वर्कशीट संरक्षण, सक्रिय कॉलम फ़िल्टर, छिपे हुए वर्कशीट टैब और छिपी हुई पूर्ण सीमाओं का तुरंत पता लगाता है जबकि खोजे गए डेटा ब्लॉक्स को कैश करता है।",
    "* **Guided Data Block Navigation:** Post-analysis navigation allows immediate cursor warps between major clusters of data, bypassing empty cells seamlessly.": "* **निर्देशित डेटा ब्लॉक नेविगेशन:** विश्लेषण के बाद का नेविगेशन डेटा के प्रमुख समूहों के बीच कर्सर को तुरंत ले जाने की अनुमति देता है, खाली सेल को सहजता से दरकिनार करता है।",
    "* **Conditional Formatting Announcer:** Automatically detects and reads the dynamic color, font style, and background shade of cells altered by Excel's Conditional Formatting rules.": "* **सशर्त स्वरूपण उद्घोषक:** Excel के सशर्त स्वरूपण नियमों द्वारा परिवर्तित सेल के गतिशील रंग, फ़ॉन्ट शैली और पृष्ठभूमि छाया को स्वचालित रूप से पता लगाता है और पढ़ता है।",
    "* **Explicit Settings Accelerators:** Completely overhauled the BOA Settings GUI to strictly comply with NVDA architecture. Every feature checkbox now possesses a globally unique `Alt+Letter` shortcut, preventing keyboard cycling and eliminating first-letter navigation failures.": "* **स्पष्ट सेटिंग्स एक्सेलेरेटर:** NVDA वास्तुकला का सख्ती से पालन करने के लिए BOA सेटिंग्स GUI को पूरी तरह से सुधारा गया। अब प्रत्येक सुविधा चेकबॉक्स में विश्व स्तर पर अद्वितीय `Alt+Letter` शॉर्टकट है, जो कीबोर्ड साइकलिंग को रोकता है और प्रथम-अक्षर नेविगेशन विफलताओं को समाप्त करता है।",
    "* **Absolute Edge Boundary Detection:** Replaced native COM `UsedRange` edge checks with absolute 1D mathematical boundary checks (`Row 1048576` and `Column 16384`) to guarantee detection of hidden rows/columns even if they lie far outside the active data block.": "* **पूर्ण एज बाउंड्री डिटेक्शन:** सक्रिय डेटा ब्लॉक के बहुत बाहर स्थित होने पर भी छिपी हुई पंक्तियों/कॉलम का पता लगाने की गारंटी देने के लिए मूल COM `UsedRange` किनारे की जांच को निरपेक्ष 1D गणितीय सीमा जांच (`Row 1048576` और `Column 16384`) से बदल दिया गया।",
    "* **Lazy COM Property Safe Bailouts:** Hardened COM property loops to prevent NVDA thread freezes when evaluating millions of contiguous hidden structures.": "* **लेज़ी COM प्रॉपर्टी सेफ बेलआउट्स:** लाखों सन्निहित छिपी हुई संरचनाओं का मूल्यांकन करते समय NVDA थ्रेड फ़्रीज़ को रोकने के लिए COM प्रॉपर्टी लूप को कठोर किया गया।",
    "### Version 1.2.0 — 2026-06-03": "### संस्करण 1.2.0 — 2026-06-03",
    "* **App-Launch Caching:** Major architectural overhaul. Core modules are now lazy-loaded exactly when you focus on Office applications, eliminating boot lag, completely solving the 'unknown' object focus glitch on rename dialogs, and preserving multi-file codebase structure.": "* **ऐप-लॉन्च कैशिंग:** प्रमुख वास्तुशिल्प बदलाव। कोर मॉड्यूल अब आलसी-लोड होते हैं जब आप Office एप्लिकेशन पर ध्यान केंद्रित करते हैं, बूट लैग को समाप्त करते हैं, नाम बदलने वाले संवादों पर 'अज्ञात' ऑब्जेक्ट फ़ोकस गड़बड़ को पूरी तरह से हल करते हैं, और मल्टी-फ़ाइल कोडबेस संरचना को संरक्षित करते हैं।",
    "* **Enhanced Cell Tracker (1D COM Math):** Rewrote the hidden cell gap detection logic to only evaluate one-dimensional cross-sections (`current_col` or `current_row`). This reduces the COM calculation payload by over 16 million cells, instantly eliminating navigation freezes when jumping hidden ranges.": "* **उन्नत सेल ट्रैकर (1D COM गणित):** केवल एक-आयामी क्रॉस-सेक्शन (`current_col` या `current_row`) का मूल्यांकन करने के लिए छिपे हुए सेल गैप डिटेक्शन लॉजिक को फिर से लिखा। यह COM गणना पेलोड को 16 मिलियन से अधिक सेल तक कम कर देता है, छिपी हुई सीमाओं को कूदते समय नेविगेशन फ़्रीज़ को तुरंत समाप्त कर देता है।",
    "* **Process Memory Wiping:** Implemented Excel Window Handle (`Hwnd`) tracking to detect when the user closes and reopens Excel. This actively wipes out stale global state memory and completely solves the false \"Sheet hidden\" announcement when opening a fresh \"Book1\".": "* **प्रोसेस मेमोरी वाइपिंग:** जब उपयोगकर्ता Excel को बंद करता है और फिर से खोलता है, तो पता लगाने के लिए Excel Window Handle (`Hwnd`) ट्रैकिंग लागू की गई। यह सक्रिय रूप से पुरानी वैश्विक राज्य मेमोरी को मिटा देता है और एक ताज़ा \"Book1\" खोलते समय झूठी \"शीट हिडन\" घोषणा को पूरी तरह से हल कर देता है।",
    "* **Double Selection Announcement:** Migrated away from unreliable asynchronous `winUser.getKeyState` and implemented `api.getLastInputGesture()` to perfectly suppress double announcements when using Shift+Arrow keys.": "* **डबल सेलेक्शन अनाउंसमेंट:** अविश्वसनीय एसिंक्रोनस `winUser.getKeyState` से दूर स्थानांतरित हो गया और Shift+Arrow कुंजी का उपयोग करते समय दोहरे घोषणाओं को पूरी तरह से दबाने के लिए `api.getLastInputGesture()` लागू किया गया।",
    "* **Boundary Detector Deactivation:** The Proactive Boundary Detector has been deactivated to protect NVDA native navigation stability, falling back entirely to the gap-skipping tracker.": "* **बाउंड्री डिटेक्टर डिएक्टिवेशन:** प्रोएक्टिव बाउंड्री डिटेक्टर को NVDA नेटिव नेविगेशन स्थिरता की सुरक्षा के लिए निष्क्रिय कर दिया गया है, जो पूरी तरह से गैप-स्किपिंग ट्रैकर पर वापस आ गया है।",
    "### Version 1.1.0 — 2026-05-30": "### संस्करण 1.1.0 — 2026-05-30",
    "* **Settings GUI:** Added a native BOA Office Enhancements panel inside `NVDA -> Preferences -> Settings` to easily toggle features on or off.": "* **सेटिंग्स GUI:** सुविधाओं को आसानी से चालू या बंद करने के लिए `NVDA -> Preferences -> Settings` के अंदर एक देशी BOA Office Enhancements पैनल जोड़ा गया।",
    "* **SafeRichEdit Hook:** Prevents silent NVDA crashes when interacting with RichEdit controls in Office 2024.": "* **SafeRichEdit Hook:** Office 2024 में RichEdit नियंत्रणों के साथ इंटरैक्ट करते समय मूक NVDA क्रैश को रोकता है।",
    "* **Customizable Hotkeys:** All BOA hotkeys are now fully exposed in NVDA's Input Gestures dialog under the \"Better Office Accessibility\" category.": "* **अनुकूलन योग्य हॉटकीज़:** सभी BOA हॉटकीज़ अब पूरी तरह से \"Better Office Accessibility\" श्रेणी के अंतर्गत NVDA के इनपुट जेस्चर डायलॉग में उजागर किए गए हैं।",
    "* **Excel: Hidden Row/Column Skip Detection:** Proactively announces when navigating past hidden rows or columns, ensuring you never miss filtered data. Can be toggled in settings.": "* **Excel: हिडन पंक्ति/कॉलम स्किप डिटेक्शन:** छिपी हुई पंक्तियों या कॉलम को पार करते समय सक्रिय रूप से घोषणा करता है, यह सुनिश्चित करता है कि आप कभी भी फ़िल्टर किए गए डेटा को न चूकें। सेटिंग्स में टॉगल किया जा सकता है।",
    "* **Thread Safety:** Removed all blocking delays (`time.sleep`) and replaced them with non-blocking NVDA asynchronous callbacks to ensure the screen reader never stutters during background operations.": "* **थ्रेड सुरक्षा:** सभी ब्लॉकिंग देरी (`time.sleep`) को हटा दिया और उन्हें नॉन-ब्लॉकिंग NVDA एसिंक्रोनस कॉलबैक के साथ बदल दिया ताकि यह सुनिश्चित हो सके कि स्क्रीन रीडर पृष्ठभूमि संचालन के दौरान कभी नहीं हकलाता है।",
    "### Version 1.0.0 — 2026-05-24": "### संस्करण 1.0.0 — 2026-05-24",
    "*Initial public release.*": "*प्रारंभिक सार्वजनिक रिलीज़।*",
    "* **Excel: Bulk Sheet Organizer:** Instantly reorder multiple sheets at once using a fully accessible dialog.": "* **Excel: बल्क शीट ऑर्गेनाइज़र:** पूरी तरह से सुलभ डायलॉग का उपयोग करके एक साथ कई शीट को तुरंत पुनर्व्यवस्थित करें।",
    "* **Excel: Quick Sheet Mover:** Move the active sheet left, right, to start, or to end via keyboard commands.": "* **Excel: क्विक शीट मूवर:** सक्रिय शीट को कीबोर्ड कमांड के माध्यम से बाएँ, दाएँ, शुरू करने के लिए, या अंत में ले जाएँ।",
    "* **Excel: Accessible Sheet Renaming:** Intercepts the inaccessible native rename field and replaces it with a reliable accessible dialog.": "* **Excel: सुलभ शीट का नाम बदलना:** अगम्य मूल नाम बदलने वाले फ़ील्ड को रोकता है और इसे एक विश्वसनीय सुलभ डायलॉग से बदल देता है।",
    "* **Excel: Smart Selection Tracking:** Accurately announces multi-cell range selections and deselections.": "* **Excel: स्मार्ट चयन ट्रैकिंग:** बहु-सेल श्रेणी चयनों और अचयन की सटीक घोषणा करता है।",
    "* **PowerPoint: Accessible Color Pickers:** Enables NVDA to accurately read RGB and Hex values inside the Custom Color dialog.": "* **PowerPoint: सुलभ कलर पिकर्स:** NVDA को कस्टम कलर डायलॉग के अंदर RGB और Hex मानों को सटीक रूप से पढ़ने में सक्षम बनाता है।",
    "* **PowerPoint: Standard Color Grid Support:** Intercepts arrow key navigation to read hidden Hex codes from the inaccessible color hexagon grid.": "* **PowerPoint: स्टैंडर्ड कलर ग्रिड सपोर्ट:** अगम्य कलर हेक्सागोन ग्रिड से छिपे हुए हेक्स कोड को पढ़ने के लिए एरो की नेविगेशन को रोकता है।"
}

hi_readme = readme_content

for en_text, hi_text in readme_translations.items():
    hi_readme = hi_readme.replace(en_text, hi_text)

with open(hi_doc_path, "w", encoding="utf-8") as f:
    f.write(hi_readme)
