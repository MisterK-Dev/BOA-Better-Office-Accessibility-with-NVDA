# -*- coding: UTF-8 -*-
import os
import polib

# Batch 3: hi, ta, te, kn, ml, mr, bn, gu

translations = {
    "hi": {
        "ui": {
            "Enable &Bulk Sheet Organizer and Quick Sheet Mover": "बल्क शीट ऑर्गेनाइज़र और त्वरित शीट मूवर को सक्षम करें (&B)",
            "Explicitly announce mer&ged cells": "मर्ज किए गए (&G) सेल की स्पष्ट घोषणा करें",
            "&Excel": "&Excel",
            "Prevent NVDA &crashes in Word text fields (SafeRichEdit)": "Word टेक्स्ट फ़ील्ड में NVDA क्रैश को रोकें (&C) (SafeRichEdit)",
            "&Word": "&Word",
            "Read hidden &hex codes when navigating the Standard Color hexagon grid": "मानक रंग षट्भुज ग्रिड पर नेविगेट करते समय छिपे हुए &Hex कोड पढ़ें",
            "Ensure the Hex Color edit &field is properly labeled": "सुनिश्चित करें कि Hex रंग संपादन फ़ील्ड (&F) ठीक से लेबल किया गया है",
            "Enable &Slide Layout Analyzer (Prefix + L)": "स्लाइड लेआउट विश्लेषक को सक्षम करें (&S) (Prefix + L)",
            "Shape Movement &Audio Mode:": "आकार संचलन ऑडियो मोड (&A):",
            "&PowerPoint": "&PowerPoint"
        },
        "desc_intro": "BOA माइक्रोसॉफ्ट ऑफिस के लिए एक्सेसिबिलिटी एन्हांसमेंट का एक शक्तिशाली सुइट है, जो AI-सहायता प्राप्त विकास का परिणाम है; इसे NVDA उपयोगकर्ताओं के लिए स्क्रीन रीडर अनुभव को बेहतर बनाने के लिए डिज़ाइन किया गया है। संपूर्ण कोडबेस Anti Gravity 2.0 का उपयोग करके उत्पन्न किया गया है। यह दुर्गम UI घटकों को सीधे ठीक करता है और Excel व PowerPoint के लिए त्वरित नेविगेशन उपकरण प्रस्तुत करता है।",
        "cl_201": """### Version 2.0.1
#### UX/UI सुधार
* **टैबयुक्त सेटिंग्स संवाद:** `wx.Notebook` का उपयोग करके BOA सेटिंग्स पैनल को सुलभ टैब (&Excel, &Word, और &PowerPoint) में पुनर्गठित किया गया है, जिससे स्क्रीन रीडर नेविगेशन में अत्यधिक सुधार हुआ है और लंबी स्क्रॉलिंग सूचियां समाप्त हो गई हैं।
* **NVDA 2026.2 अनुकूलता:** NVDA 2026.2 के लिए परीक्षण किया गया और प्रमाणित किया गया।""",
        "readme_201": """### Version 2.0.1
#### UX/UI सुधार
* **टैबयुक्त सेटिंग्स संवाद:** `wx.Notebook` का उपयोग करके BOA सेटिंग्स पैनल को सुलभ टैब (&Excel, &Word, और &PowerPoint) में पुनर्गठित किया गया है, जिससे स्क्रीन रीडर नेविगेशन में अत्यधिक सुधार हुआ है और लंबी स्क्रॉलिंग सूचियां समाप्त हो गई हैं। आप `Alt+E`, `Alt+W`, `Alt+P` या मानक शॉर्टकट `Ctrl+PageDown`/`Ctrl+PageUp` का उपयोग करके टैब के बीच त्वरित स्विच कर सकते हैं।
* **NVDA 2026.2 अनुकूलता:** NVDA 2026.2 के लिए परीक्षण किया गया और प्रमाणित किया गया।"""
    },
    "ta": {
        "ui": {
            "Enable &Bulk Sheet Organizer and Quick Sheet Mover": "மொத்த தாள் அமைப்பாளர் மற்றும் விரைவு தாள் நகர்த்தியை இயக்கு (&B)",
            "Explicitly announce mer&ged cells": "ஒருங்கிணைக்கப்பட்ட கலங்களை வெளிப்படையாக அறிவிக்கவும் (&G)",
            "&Excel": "&Excel",
            "Prevent NVDA &crashes in Word text fields (SafeRichEdit)": "Word உரை புலங்களில் NVDA முறிவைத் தடுக்கவும் (&C) (SafeRichEdit)",
            "&Word": "&Word",
            "Read hidden &hex codes when navigating the Standard Color hexagon grid": "நிலையான வண்ண அறுகோண கட்டத்தில் செல்லும்போது மறைக்கப்பட்ட &Hex குறியீடுகளைப் படிக்கவும்",
            "Ensure the Hex Color edit &field is properly labeled": "Hex வண்ணத் திருத்தப் புலம் (&F) சரியாக லேபிளிடப்பட்டுள்ளதா என்பதை உறுதிப்படுத்தவும்",
            "Enable &Slide Layout Analyzer (Prefix + L)": "ஸ்லைடு தளவமைப்பு பகுப்பாய்வியை இயக்கு (&S) (Prefix + L)",
            "Shape Movement &Audio Mode:": "வடிவ இயக்க ஆடியோ பயன்முறை (&A):",
            "&PowerPoint": "&PowerPoint"
        },
        "desc_intro": "BOA என்பது மைக்ரோசாஃப்ட் ஆபிஸிற்கான அணுகல் மேம்பாடுகளின் சக்திவாய்ந்த தொகுப்பாகும், இது AI-உதவி வளர்ச்சியின் விளைவாகும்; NVDA பயனர்களுக்கான திரை வாசிப்பான் அனுபவத்தை மேம்படுத்த வடிவமைக்கப்பட்டுள்ளது. முழு குறியீட்டுத் தளமும் Anti Gravity 2.0 ஐப் பயன்படுத்தி உருவாக்கப்பட்டது. இது அணுக முடியாத UI கூறுகளை நேரடியாக சரிசெய்கிறது மற்றும் Excel மற்றும் PowerPoint க்கான விரைவான வழிசெலுத்தல் கருவிகளை அறிமுகப்படுத்துகிறது.",
        "cl_201": """### Version 2.0.1
#### UX/UI மேம்பாடுகள்
* **தாவல் அமைப்புகள் உரையாடல்:** `wx.Notebook` ஐப் பயன்படுத்தி BOA அமைப்புகள் பலகம் அணுகக்கூடிய தாவல்களாக (&Excel, &Word, மற்றும் &PowerPoint) மறுசீரமைக்கப்பட்டுள்ளது, இது திரை வாசிப்பான் வழிசெலுத்தலை பெரிதும் மேம்படுத்துகிறது மற்றும் நீண்ட பட்டியல்களை நீக்குகிறது.
* **NVDA 2026.2 பொருந்தக்கூடிய தன்மை:** NVDA 2026.2 க்கு சோதிக்கப்பட்டு சான்றளிக்கப்பட்டது.""",
        "readme_201": """### Version 2.0.1
#### UX/UI மேம்பாடுகள்
* **தாவல் அமைப்புகள் உரையாடல்:** `wx.Notebook` ஐப் பயன்படுத்தி BOA அமைப்புகள் பலகம் அணுகக்கூடிய தாவல்களாக (&Excel, &Word, மற்றும் &PowerPoint) மறுசீரமைக்கப்பட்டுள்ளது, இது திரை வாசிப்பான் வழிசெலுத்தலை பெரிதும் மேம்படுத்துகிறது மற்றும் நீண்ட பட்டியல்களை நீக்குகிறது. `Alt+E`, `Alt+W`, `Alt+P` அல்லது நிலையான குறுக்குவழிகள் `Ctrl+PageDown`/`Ctrl+PageUp` ஐப் பயன்படுத்தி தாவல்களுக்கு இடையில் விரைவாக மாறலாம்.
* **NVDA 2026.2 பொருந்தக்கூடிய தன்மை:** NVDA 2026.2 க்கு சோதிக்கப்பட்டு சான்றளிக்கப்பட்டது."""
    },
    "te": {
        "ui": {
            "Enable &Bulk Sheet Organizer and Quick Sheet Mover": "బల్క్ షీట్ ఆర్గనైజర్ మరియు త్వరిత షీట్ మూవర్‌ను ప్రారంభించండి (&B)",
            "Explicitly announce mer&ged cells": "విలీనం చేయబడిన (&G) కణాలను స్పష్టంగా ప్రకటించండి",
            "&Excel": "&Excel",
            "Prevent NVDA &crashes in Word text fields (SafeRichEdit)": "Word టెక్స్ట్ ఫీల్డ్‌లలో NVDA క్రాష్‌లను నిరోధించండి (&C) (SafeRichEdit)",
            "&Word": "&Word",
            "Read hidden &hex codes when navigating the Standard Color hexagon grid": "ప్రామాణిక రంగు షడ్భుజి గ్రిడ్‌లో నావిగేట్ చేస్తున్నప్పుడు దాచిన &Hex కోడ్‌లను చదవండి",
            "Ensure the Hex Color edit &field is properly labeled": "Hex రంగు సవరణ ఫీల్డ్ (&F) సరిగ్గా లేబుల్ చేయబడిందని నిర్ధారించుకోండి",
            "Enable &Slide Layout Analyzer (Prefix + L)": "స్లైడ్ లేఅవుట్ ఎనలైజర్‌ను ప్రారంభించండి (&S) (Prefix + L)",
            "Shape Movement &Audio Mode:": "ఆకార కదలిక ఆడియో మోడ్ (&A):",
            "&PowerPoint": "&PowerPoint"
        },
        "desc_intro": "BOA అనేది మైక్రోసాఫ్ట్ ఆఫీస్ కోసం ప్రాప్యత మెరుగుదలల యొక్క శక్తివంతమైన సూట్, ఇది AI-సహాయక అభివృద్ధి ఫలితం; NVDA వినియోగదారుల కోసం స్క్రీన్ రీడర్ అనుభవాన్ని మెరుగుపరచడానికి రూపొందించబడింది. మొత్తం కోడ్‌బేస్ Anti Gravity 2.0 ఉపయోగించి రూపొందించబడింది. ఇది ప్రాప్యత చేయలేని UI భాగాలను నేరుగా పరిష్కరిస్తుంది మరియు Excel, PowerPoint ల కోసం వేగవంతమైన నావిగేషన్ సాధనాలను పరిచయం చేస్తుంది.",
        "cl_201": """### Version 2.0.1
#### UX/UI మెరుగుదలలు
* **ట్యాబ్ చేసిన సెట్టింగ్‌ల డైలాగ్:** `wx.Notebook`ని ఉపయోగించి BOA సెట్టింగ్‌ల ప్యానెల్ యాక్సెస్ చేయగల ట్యాబ్‌లుగా (&Excel, &Word, మరియు &PowerPoint) పునర్వ్యవస్థీకరించబడింది, ఇది స్క్రీన్ రీడర్ నావిగేషన్‌ను గణనీయంగా మెరుగుపరుస్తుంది మరియు సుదీర్ఘ జాబితాలను తొలగిస్తుంది.
* **NVDA 2026.2 అనుకూలత:** NVDA 2026.2 కోసం పరీక్షించబడింది మరియు ధృవీకరించబడింది.""",
        "readme_201": """### Version 2.0.1
#### UX/UI మెరుగుదలలు
* **ట్యాబ్ చేసిన సెట్టింగ్‌ల డైలాగ్:** `wx.Notebook`ని ఉపయోగించి BOA సెట్టింగ్‌ల ప్యానెల్ యాక్సెస్ చేయగల ట్యాబ్‌లుగా (&Excel, &Word, మరియు &PowerPoint) పునర్వ్యవస్థీకరించబడింది, ఇది స్క్రీన్ రీడర్ నావిగేషన్‌ను గణనీయంగా మెరుగుపరుస్తుంది మరియు సుదీర్ఘ జాబితాలను తొలగిస్తుంది. మీరు `Alt+E`, `Alt+W`, `Alt+P` లేదా ప్రామాణిక షార్ట్‌కట్‌లు `Ctrl+PageDown`/`Ctrl+PageUp` ఉపయోగించి ట్యాబ్‌ల మధ్య త్వరగా మారవచ్చు.
* **NVDA 2026.2 అనుకూలత:** NVDA 2026.2 కోసం పరీక్షించబడింది మరియు ధృవీకరించబడింది."""
    },
    "kn": {
        "ui": {
            "Enable &Bulk Sheet Organizer and Quick Sheet Mover": "ಬಲ್ಕ್ ಶೀಟ್ ಆರ್ಗನೈಜರ್ ಮತ್ತು ತ್ವರಿತ ಶೀಟ್ ಮೂವರ್ ಅನ್ನು ಸಕ್ರಿಯಗೊಳಿಸಿ (&B)",
            "Explicitly announce mer&ged cells": "ವಿಲೀನಗೊಂಡ (&G) ಕೋಶಗಳನ್ನು ಸ್ಪಷ್ಟವಾಗಿ ಪ್ರಕಟಿಸಿ",
            "&Excel": "&Excel",
            "Prevent NVDA &crashes in Word text fields (SafeRichEdit)": "Word ಪಠ್ಯ ಕ್ಷೇತ್ರಗಳಲ್ಲಿ NVDA ಕ್ರ್ಯಾಶ್‌ಗಳನ್ನು ತಡೆಯಿರಿ (&C) (SafeRichEdit)",
            "&Word": "&Word",
            "Read hidden &hex codes when navigating the Standard Color hexagon grid": "ಪ್ರಮಾಣಿತ ಬಣ್ಣದ ಷಡ್ಭುಜಾಕೃತಿಯ ಗ್ರಿಡ್‌ನಲ್ಲಿ ನ್ಯಾವಿಗೇಟ್ ಮಾಡುವಾಗ ಗುಪ್ತ &Hex ಕೋಡ್‌ಗಳನ್ನು ಓದಿ",
            "Ensure the Hex Color edit &field is properly labeled": "Hex ಬಣ್ಣ ಸಂಪಾದನೆ ಕ್ಷೇತ್ರಕ್ಕೆ (&F) ಸರಿಯಾಗಿ ಲೇಬಲ್ ಮಾಡಲಾಗಿದೆಯೇ ಎಂದು ಖಚಿತಪಡಿಸಿಕೊಳ್ಳಿ",
            "Enable &Slide Layout Analyzer (Prefix + L)": "ಸ್ಲೈಡ್ ವಿನ್ಯಾಸ ವಿಶ್ಲೇಷಕವನ್ನು ಸಕ್ರಿಯಗೊಳಿಸಿ (&S) (Prefix + L)",
            "Shape Movement &Audio Mode:": "ಆಕಾರ ಚಲನೆಯ ಆಡಿಯೋ ಮೋಡ್ (&A):",
            "&PowerPoint": "&PowerPoint"
        },
        "desc_intro": "BOA ಮೈಕ್ರೋಸಾಫ್ಟ್ ಆಫೀಸ್‌ಗಾಗಿ ಪ್ರವೇಶಿಸುವಿಕೆ ವರ್ಧನೆಗಳ ಪ್ರಬಲ ಸೂಟ್ ಆಗಿದೆ, ಇದು AI-ಸಹಾಯದ ಅಭಿವೃದ್ಧಿಯ ಫಲಿತಾಂಶವಾಗಿದೆ; NVDA ಬಳಕೆದಾರರಿಗೆ ಸ್ಕ್ರೀನ್ ರೀಡರ್ ಅನುಭವವನ್ನು ನಾಟಕೀಯವಾಗಿ ಸುಧಾರಿಸಲು ವಿನ್ಯಾಸಗೊಳಿಸಲಾಗಿದೆ. ಸಂಪೂರ್ಣ ಕೋಡ್‌ಬೇಸ್ ಅನ್ನು Anti Gravity 2.0 ಬಳಸಿ ರಚಿಸಲಾಗಿದೆ. ಇದು ಪ್ರವೇಶಿಸಲಾಗದ UI ಘಟಕಗಳನ್ನು ನೇರವಾಗಿ ಸರಿಪಡಿಸುತ್ತದೆ ಮತ್ತು Excel ಮತ್ತು PowerPoint ಗಾಗಿ ವೇಗದ ನ್ಯಾವಿಗೇಷನ್ ಪರಿಕರಗಳನ್ನು ಪರಿಚಯಿಸುತ್ತದೆ.",
        "cl_201": """### Version 2.0.1
#### UX/UI ವರ್ಧನೆಗಳು
* **ಟ್ಯಾಬ್ ಮಾಡಲಾದ ಸೆಟ್ಟಿಂಗ್‌ಗಳ ಸಂವಾದ:** `wx.Notebook` ಬಳಸಿ BOA ಸೆಟ್ಟಿಂಗ್‌ಗಳ ಫಲಕವನ್ನು ಪ್ರವೇಶಿಸಬಹುದಾದ ಟ್ಯಾಬ್‌ಗಳಾಗಿ (&Excel, &Word, ಮತ್ತು &PowerPoint) ಮರುಸಂಘಟಿಸಲಾಗಿದೆ, ಇದು ಸ್ಕ್ರೀನ್ ರೀಡರ್ ನ್ಯಾವಿಗೇಷನ್ ಅನ್ನು ಹೆಚ್ಚು ಸುಧಾರಿಸುತ್ತದೆ.
* **NVDA 2026.2 ಹೊಂದಾಣಿಕೆ:** NVDA 2026.2 ಗಾಗಿ ಪರೀಕ್ಷಿಸಲಾಗಿದೆ ಮತ್ತು ಪ್ರಮಾಣೀಕರಿಸಲಾಗಿದೆ.""",
        "readme_201": """### Version 2.0.1
#### UX/UI ವರ್ಧನೆಗಳು
* **ಟ್ಯಾಬ್ ಮಾಡಲಾದ ಸೆಟ್ಟಿಂಗ್‌ಗಳ ಸಂವಾದ:** `wx.Notebook` ಬಳಸಿ BOA ಸೆಟ್ಟಿಂಗ್‌ಗಳ ಫಲಕವನ್ನು ಪ್ರವೇಶಿಸಬಹುದಾದ ಟ್ಯಾಬ್‌ಗಳಾಗಿ (&Excel, &Word, ಮತ್ತು &PowerPoint) ಮರುಸಂಘಟಿಸಲಾಗಿದೆ, ಇದು ಸ್ಕ್ರೀನ್ ರೀಡರ್ ನ್ಯಾವಿಗೇಷನ್ ಅನ್ನು ಹೆಚ್ಚು ಸುಧಾರಿಸುತ್ತದೆ. `Alt+E`, `Alt+W`, `Alt+P` ಅಥವಾ ಪ್ರಮಾಣಿತ ಶಾರ್ಟ್‌ಕಟ್‌ಗಳಾದ `Ctrl+PageDown`/`Ctrl+PageUp` ಬಳಸಿಕೊಂಡು ಟ್ಯಾಬ್‌ಗಳ ನಡುವೆ ತ್ವರಿತವಾಗಿ ಬದಲಾಯಿಸಬಹುದು.
* **NVDA 2026.2 ಹೊಂದಾಣಿಕೆ:** NVDA 2026.2 ಗಾಗಿ ಪರೀಕ್ಷಿಸಲಾಗಿದೆ ಮತ್ತು ಪ್ರಮಾಣೀಕರಿಸಲಾಗಿದೆ."""
    },
    "ml": {
        "ui": {
            "Enable &Bulk Sheet Organizer and Quick Sheet Mover": "ബൾക്ക് ഷീറ്റ് ഓർഗനൈസറും ദ്രുത ഷീറ്റ് മൂവറും പ്രവർത്തനക്ഷമമാക്കുക (&B)",
            "Explicitly announce mer&ged cells": "ലയിപ്പിച്ച (&G) സെല്ലുകൾ വ്യക്തമായി പ്രഖ്യാപിക്കുക",
            "&Excel": "&Excel",
            "Prevent NVDA &crashes in Word text fields (SafeRichEdit)": "Word ടെക്സ്റ്റ് ഫീൽഡുകളിൽ NVDA ക്രാഷുകൾ തടയുക (&C) (SafeRichEdit)",
            "&Word": "&Word",
            "Read hidden &hex codes when navigating the Standard Color hexagon grid": "സ്റ്റാൻഡേർഡ് കളർ ഷഡ്ഭുജ ഗ്രിഡിൽ നാവിഗേറ്റ് ചെയ്യുമ്പോൾ മറഞ്ഞിരിക്കുന്ന &Hex കോഡുകൾ വായിക്കുക",
            "Ensure the Hex Color edit &field is properly labeled": "Hex കളർ എഡിറ്റ് ഫീൽഡ് (&F) ശരിയായി ലേബൽ ചെയ്തിട്ടുണ്ടെന്ന് ഉറപ്പാക്കുക",
            "Enable &Slide Layout Analyzer (Prefix + L)": "സ്ലൈഡ് ലേഔട്ട് അനലൈസർ പ്രവർത്തനക്ഷമമാക്കുക (&S) (Prefix + L)",
            "Shape Movement &Audio Mode:": "ആകൃതി ചലന ഓഡിയോ മോഡ് (&A):",
            "&PowerPoint": "&PowerPoint"
        },
        "desc_intro": "മൈക്രോസോഫ്റ്റ് ഓഫീസിനായുള്ള പ്രവേശനക്ഷമത മെച്ചപ്പെടുത്തലുകളുടെ ശക്തമായ ഒരു സ്യൂട്ടാണ് BOA, AI-സഹായത്തോടെയുള്ള വികസനത്തിൻ്റെ ഫലമാണിത്; NVDA ഉപയോക്താക്കൾക്കായി സ്ക്രീൻ റീഡർ അനുഭവം ഗണ്യമായി മെച്ചപ്പെടുത്തുന്നതിനാണ് ഇത് രൂപകൽപ്പന ചെയ്തിരിക്കുന്നത്. പൂർണ്ണ കോഡ്ബേസ് Anti Gravity 2.0 ഉപയോഗിച്ചാണ് നിർമ്മിച്ചിരിക്കുന്നത്. ഇത് അപ്രാപ്യമായ UI ഘടകങ്ങൾ നേരിട്ട് പരിഹരിക്കുകയും Excel, PowerPoint എന്നിവയ്ക്കായി വേഗത്തിലുള്ള നാവിഗേഷൻ ടൂളുകൾ നൽകുകയും ചെയ്യുന്നു.",
        "cl_201": """### Version 2.0.1
#### UX/UI മെച്ചപ്പെടുത്തലുകൾ
* **ടാബ് ചെയ്ത ക്രമീകരണ ഡയലോഗ്:** `wx.Notebook` ഉപയോഗിച്ച് BOA ക്രമീകരണ പാനൽ ആക്‌സസ് ചെയ്യാവുന്ന ടാബുകളിലേക്ക് (&Excel, &Word, &PowerPoint) പുനഃസംഘടിപ്പിച്ചു, ഇത് സ്‌ക്രീൻ റീഡർ നാവിഗേഷൻ വളരെയധികം മെച്ചപ്പെടുത്തുന്നു.
* **NVDA 2026.2 അനുയോജ്യത:** NVDA 2026.2-നായി പരീക്ഷിക്കുകയും സാക്ഷ്യപ്പെടുത്തുകയും ചെയ്തു.""",
        "readme_201": """### Version 2.0.1
#### UX/UI മെച്ചപ്പെടുത്തലുകൾ
* **ടാബ് ചെയ്ത ക്രമീകരണ ഡയലോഗ്:** `wx.Notebook` ഉപയോഗിച്ച് BOA ക്രമീകരണ പാനൽ ആക്‌സസ് ചെയ്യാവുന്ന ടാബുകളിലേക്ക് (&Excel, &Word, &PowerPoint) പുനഃസംഘടിപ്പിച്ചു, ഇത് സ്‌ക്രീൻ റീഡർ നാവിഗേഷൻ വളരെയധികം മെച്ചപ്പെടുത്തുന്നു. `Alt+E`, `Alt+W`, `Alt+P` അല്ലെങ്കിൽ സാധാരണ കുറുക്കുവഴികളായ `Ctrl+PageDown`/`Ctrl+PageUp` ഉപയോഗിച്ച് നിങ്ങൾക്ക് ടാബുകൾക്കിടയിൽ വേഗത്തിൽ മാറാം.
* **NVDA 2026.2 അനുയോജ്യത:** NVDA 2026.2-നായി പരീക്ഷിക്കുകയും സാക്ഷ്യപ്പെടുത്തുകയും ചെയ്തു."""
    },
    "mr": {
        "ui": {
            "Enable &Bulk Sheet Organizer and Quick Sheet Mover": "बल्क शीट ऑर्गनायझर आणि द्रुत शीट मूव्हर सक्षम करा (&B)",
            "Explicitly announce mer&ged cells": "मर्ज केलेल्या (&G) सेलची स्पष्ट घोषणा करा",
            "&Excel": "&Excel",
            "Prevent NVDA &crashes in Word text fields (SafeRichEdit)": "Word मजकूर फील्डमध्ये NVDA क्रॅश रोखा (&C) (SafeRichEdit)",
            "&Word": "&Word",
            "Read hidden &hex codes when navigating the Standard Color hexagon grid": "मानक रंग षटकोनी ग्रिडवर नेव्हिगेट करताना लपवलेले &Hex कोड वाचा",
            "Ensure the Hex Color edit &field is properly labeled": "Hex रंग संपादन फील्ड (&F) योग्यरित्या लेबल केलेले असल्याची खात्री करा",
            "Enable &Slide Layout Analyzer (Prefix + L)": "स्लाइड लेआउट विश्लेषक सक्षम करा (&S) (Prefix + L)",
            "Shape Movement &Audio Mode:": "आकार हालचाल ऑडिओ मोड (&A):",
            "&PowerPoint": "&PowerPoint"
        },
        "desc_intro": "BOA हा मायक्रोसॉफ्ट ऑफिससाठी प्रवेशयोग्यता सुधारणांचा एक शक्तिशाली संच आहे, जो AI-सहाय्यित विकासाचा परिणाम आहे; NVDA वापरकर्त्यांसाठी स्क्रीन रीडर अनुभव मोठ्या प्रमाणात सुधारण्यासाठी डिझाइन केला गेला आहे. संपूर्ण कोडबेस Anti Gravity 2.0 वापरून व्युत्पन्न केला गेला आहे. हे थेट दुर्गम UI घटकांना दुरुस्त करते आणि Excel व PowerPoint साठी जलद नेव्हिगेशन साधने सादर करते.",
        "cl_201": """### Version 2.0.1
#### UX/UI सुधारणा
* **टॅब केलेला सेटिंग्ज डायलॉग:** `wx.Notebook` वापरून BOA सेटिंग्ज पॅनेल प्रवेशयोग्य टॅबमध्ये (&Excel, &Word, आणि &PowerPoint) पुनर्रचित केले गेले आहे, ज्यामुळे स्क्रीन रीडर नेव्हिगेशनमध्ये मोठ्या प्रमाणात सुधारणा झाली आहे.
* **NVDA 2026.2 सुसंगतता:** NVDA 2026.2 साठी चाचणी केली आणि प्रमाणित केली गेली.""",
        "readme_201": """### Version 2.0.1
#### UX/UI सुधारणा
* **टॅब केलेला सेटिंग्ज डायलॉग:** `wx.Notebook` वापरून BOA सेटिंग्ज पॅनेल प्रवेशयोग्य टॅबमध्ये (&Excel, &Word, आणि &PowerPoint) पुनर्रचित केले गेले आहे, ज्यामुळे स्क्रीन रीडर नेव्हिगेशनमध्ये मोठ्या प्रमाणात सुधारणा झाली आहे. तुम्ही `Alt+E`, `Alt+W`, `Alt+P` किंवा मानक शॉर्टकट `Ctrl+PageDown`/`Ctrl+PageUp` वापरून टॅब दरम्यान द्रुतपणे स्विच करू शकता.
* **NVDA 2026.2 सुसंगतता:** NVDA 2026.2 साठी चाचणी केली आणि प्रमाणित केली गेली."""
    },
    "bn": {
        "ui": {
            "Enable &Bulk Sheet Organizer and Quick Sheet Mover": "বাল্ক শীট অর্গানাইজার এবং দ্রুত শীট মুভার সক্ষম করুন (&B)",
            "Explicitly announce mer&ged cells": "মার্জ করা (&G) ঘরগুলি স্পষ্টভাবে ঘোষণা করুন",
            "&Excel": "&Excel",
            "Prevent NVDA &crashes in Word text fields (SafeRichEdit)": "Word টেক্সট ফিল্ডে NVDA ক্র্যাশ প্রতিরোধ করুন (&C) (SafeRichEdit)",
            "&Word": "&Word",
            "Read hidden &hex codes when navigating the Standard Color hexagon grid": "স্ট্যান্ডার্ড কালার হেক্সাগন গ্রিডে নেভিগেট করার সময় লুকানো &Hex কোড পড়ুন",
            "Ensure the Hex Color edit &field is properly labeled": "Hex রঙ সম্পাদনা ক্ষেত্রটি (&F) সঠিকভাবে লেবেলযুক্ত তা নিশ্চিত করুন",
            "Enable &Slide Layout Analyzer (Prefix + L)": "স্লাইড লেআউট বিশ্লেষক সক্ষম করুন (&S) (Prefix + L)",
            "Shape Movement &Audio Mode:": "আকৃতি চলন অডিও মোড (&A):",
            "&PowerPoint": "&PowerPoint"
        },
        "desc_intro": "BOA হল মাইক্রোসফ্ট অফিসের জন্য অ্যাক্সেসিবিলিটি বর্ধনের একটি শক্তিশালী স্যুট, যা AI-সহায়তাপ্রাপ্ত বিকাশের ফলাফল; NVDA ব্যবহারকারীদের জন্য স্ক্রিন রিডার অভিজ্ঞতাকে ব্যাপকভাবে উন্নত করার জন্য ডিজাইন করা হয়েছে। সম্পূর্ণ কোডবেস Anti Gravity 2.0 ব্যবহার করে তৈরি করা হয়েছে। এটি দুর্গম UI উপাদানগুলিকে সরাসরি ঠিক করে এবং Excel ও PowerPoint-এর জন্য দ্রুত নেভিগেশন সরঞ্জাম প্রবর্তন করে।",
        "cl_201": """### Version 2.0.1
#### UX/UI উন্নতি
* **ট্যাবযুক্ত সেটিংস ডায়ালগ:** `wx.Notebook` ব্যবহার করে BOA সেটিংস প্যানেলটি অ্যাক্সেসযোগ্য ট্যাবে (&Excel, &Word, এবং &PowerPoint) পুনর্গঠিত করা হয়েছে, যা স্ক্রিন রিডার নেভিগেশনকে ব্যাপকভাবে উন্নত করেছে।
* **NVDA 2026.2 সামঞ্জস্য:** NVDA 2026.2 এর জন্য পরীক্ষিত এবং প্রত্যয়িত।""",
        "readme_201": """### Version 2.0.1
#### UX/UI উন্নতি
* **ট্যাবযুক্ত সেটিংস ডায়ালগ:** `wx.Notebook` ব্যবহার করে BOA সেটিংস প্যানেলটি অ্যাক্সেসযোগ্য ট্যাবে (&Excel, &Word, এবং &PowerPoint) পুনর্গঠিত করা হয়েছে, যা স্ক্রিন রিডার নেভিগেশনকে ব্যাপকভাবে উন্নত করেছে। আপনি `Alt+E`, `Alt+W`, `Alt+P` বা স্ট্যান্ডার্ড শর্টকাট `Ctrl+PageDown`/`Ctrl+PageUp` ব্যবহার করে ট্যাবগুলির মধ্যে দ্রুত স্যুইচ করতে পারেন।
* **NVDA 2026.2 সামঞ্জস্য:** NVDA 2026.2 এর জন্য পরীক্ষিত এবং প্রত্যয়িত।"""
    },
    "gu": {
        "ui": {
            "Enable &Bulk Sheet Organizer and Quick Sheet Mover": "બલ્ક શીટ ઓર્ગેનાઇઝર અને ઝડપી શીટ મૂવર સક્ષમ કરો (&B)",
            "Explicitly announce mer&ged cells": "મર્જ કરેલા (&G) કોષોની સ્પષ્ટ જાહેરાત કરો",
            "&Excel": "&Excel",
            "Prevent NVDA &crashes in Word text fields (SafeRichEdit)": "Word ટેક્સ્ટ ફીલ્ડ્સમાં NVDA ક્રેશ અટકાવો (&C) (SafeRichEdit)",
            "&Word": "&Word",
            "Read hidden &hex codes when navigating the Standard Color hexagon grid": "પ્રમાણભૂત રંગ ષટ્કોણ ગ્રીડ પર નેવિગેટ કરતી વખતે છુપાયેલા &Hex કોડ વાંચો",
            "Ensure the Hex Color edit &field is properly labeled": "Hex રંગ સંપાદન ક્ષેત્ર (&F) યોગ્ય રીતે લેબલ થયેલ છે તેની ખાતરી કરો",
            "Enable &Slide Layout Analyzer (Prefix + L)": "સ્લાઇડ લેઆઉટ એનાલાઇઝર સક્ષમ કરો (&S) (Prefix + L)",
            "Shape Movement &Audio Mode:": "આકાર હલનચલન ઑડિઓ મોડ (&A):",
            "&PowerPoint": "&PowerPoint"
        },
        "desc_intro": "BOA એ માઇક્રોસોફ્ટ ઑફિસ માટે સુલભતા ઉન્નતીકરણોનો એક શક્તિશાળી સ્યુટ છે, જે AI-સહાયિત વિકાસનું પરિણામ છે; NVDA વપરાશકર્તાઓ માટે સ્ક્રીન રીડર અનુભવને નોંધપાત્ર રીતે સુધારવા માટે રચાયેલ છે. સમગ્ર કોડબેઝ Anti Gravity 2.0 નો ઉપયોગ કરીને બનાવવામાં આવ્યો છે. તે સીધા જ અગમ્ય UI ઘટકોને સુધારે છે અને Excel તથા PowerPoint માટે ઝડપી નેવિગેશન ટૂલ્સ રજૂ કરે છે.",
        "cl_201": """### Version 2.0.1
#### UX/UI સુધારાઓ
* **ટેબ કરેલ સેટિંગ્સ સંવાદ:** `wx.Notebook` નો ઉપયોગ કરીને BOA સેટિંગ્સ પેનલને સુલભ ટેબ્સમાં (&Excel, &Word, અને &PowerPoint) પુનર્ગઠિત કરવામાં આવી છે, જે સ્ક્રીન રીડર નેવિગેશનમાં મોટાપાયે સુધારો કરે છે.
* **NVDA 2026.2 સુસંગતતા:** NVDA 2026.2 માટે ચકાસાયેલ અને પ્રમાણિત.""",
        "readme_201": """### Version 2.0.1
#### UX/UI સુધારાઓ
* **ટેબ કરેલ સેટિંગ્સ સંવાદ:** `wx.Notebook` નો ઉપયોગ કરીને BOA સેટિંગ્સ પેનલને સુલભ ટેબ્સમાં (&Excel, &Word, અને &PowerPoint) પુનર્ગઠિત કરવામાં આવી છે, જે સ્ક્રીન રીડર નેવિગેશનમાં મોટાપાયે સુધારો કરે છે. તમે `Alt+E`, `Alt+W`, `Alt+P` અથવા સ્ટાન્ડર્ડ શૉર્ટકટ્સ `Ctrl+PageDown`/`Ctrl+PageUp` નો ઉપયોગ કરીને ટેબ્સ વચ્ચે ઝડપથી સ્વિચ કરી શકો છો.
* **NVDA 2026.2 સુસંગતતા:** NVDA 2026.2 માટે ચકાસાયેલ અને પ્રમાણિત."""
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
            features_part = ""
            for marker in ["\nKey Features:", "\n    Key Features:", "\n    मुख्य विशेषताएं:", "\n    મુખ્ય વિશેષતાઓ:", "\n    મુખ્ય લક્ષણો:", "\n    പ്രധാന സവിശേഷതകൾ:", "\n    ಮುಖ್ಯ ವೈಶಿಷ್ಟ್ಯಗಳು:", "\n    முக்கிய அம்சங்கள்:", "\n    ముఖ్య లక్షణాలు:", "\n    प्रमुख वैशिष्ट्ये:", "\n    মূল বৈশিষ্ট্যগুলি:"]:
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
