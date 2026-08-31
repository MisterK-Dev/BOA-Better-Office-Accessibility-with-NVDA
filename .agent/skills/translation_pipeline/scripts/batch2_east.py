# -*- coding: UTF-8 -*-
import os
import polib

# Batch 2: zh_CN, ja, ko, tr, cs, ar, ur, pa

translations = {
    "zh_CN": {
        "ui": {
            "Enable &Bulk Sheet Organizer and Quick Sheet Mover": "启用批量工作表整理器和快速工作表移动器(&B)",
            "Explicitly announce mer&ged cells": "明确播报合并单元格(&G)",
            "&Excel": "&Excel",
            "Prevent NVDA &crashes in Word text fields (SafeRichEdit)": "防止 NVDA 在 Word 文本字段中崩溃 (&C)(SafeRichEdit)",
            "&Word": "&Word",
            "Read hidden &hex codes when navigating the Standard Color hexagon grid": "在标准颜色六边形网格中导航时朗读隐藏的十六进制代码(&H)",
            "Ensure the Hex Color edit &field is properly labeled": "确保十六进制颜色编辑字段有正确的标签(&F)",
            "Enable &Slide Layout Analyzer (Prefix + L)": "启用幻灯片版式分析器 (&S)(Prefix + L)",
            "Shape Movement &Audio Mode:": "形状移动音频模式(&A)：",
            "&PowerPoint": "&PowerPoint"
        },
        "desc_intro": "BOA 是专为 Microsoft Office 设计的功能强大的无障碍增强套件，是 AI 辅助开发的成果；旨在极大地改善 NVDA 用户的屏幕朗读体验。整个代码库均使用 Anti Gravity 2.0 生成。它直接修补了不可访问的 UI 组件，并为 Excel 和 PowerPoint 引入了快速导航工具。",
        "cl_201": """### Version 2.0.1
#### UX/UI Enhancements
* **选项卡式设置对话框：** 使用 `wx.Notebook` 将 BOA 设置面板重组为可访问的选项卡（&Excel、&Word 和 &PowerPoint），大幅改善了屏幕朗读器导航并消除了冗长的滚动列表。
* **NVDA 2026.2 兼容性：** 经过全面测试并确保兼容 NVDA 2026.2。""",
        "readme_201": """### Version 2.0.1
#### 体验/界面增强
* **选项卡式设置对话框：** 使用 `wx.Notebook` 将 BOA 设置面板重组为可访问的选项卡（&Excel、&Word 和 &PowerPoint），大幅改善了屏幕朗读器导航并消除了冗长的滚动列表。您可以使用 `Alt+E`、`Alt+W`、`Alt+P` 或标准的 `Ctrl+PageDown`/`Ctrl+PageUp` 快捷键在选项卡之间快速切换。
* **NVDA 2026.2 兼容性：** 经过全面测试并确保兼容 NVDA 2026.2。"""
    },
    "ja": {
        "ui": {
            "Enable &Bulk Sheet Organizer and Quick Sheet Mover": "一括シートオーガナイザーとクイックシート移動を有効化(&B)",
            "Explicitly announce mer&ged cells": "結合されたセルを明示的に通知(&G)",
            "&Excel": "&Excel",
            "Prevent NVDA &crashes in Word text fields (SafeRichEdit)": "Word テキストフィールドでの NVDA クラッシュを防止 (&C)(SafeRichEdit)",
            "&Word": "&Word",
            "Read hidden &hex codes when navigating the Standard Color hexagon grid": "標準カラーの六角形グリッド移動時に非表示のカラーコードを読み上げ(&H)",
            "Ensure the Hex Color edit &field is properly labeled": "Hex カラー編集フィールドが適切にラベル付けされていることを確認(&F)",
            "Enable &Slide Layout Analyzer (Prefix + L)": "スライドレイアウトアナライザーを有効化 (&S)(Prefix + L)",
            "Shape Movement &Audio Mode:": "図形移動オーディオモード(&A):",
            "&PowerPoint": "&PowerPoint"
        },
        "desc_intro": "BOA は、NVDA ユーザーのスクリーンリーダー体験を劇的に向上させるために設計された、AI 支援開発による Microsoft Office 向けの強力なアクセシビリティ強化スイートです。コードベース全体が Anti Gravity 2.0 を使用して生成されています。アクセス不能な UI コンポーネントを直接修正し、Excel および PowerPoint 向けの高速ナビゲーションツールを導入します。",
        "cl_201": """### Version 2.0.1
#### UX/UI の改善
* **タブ付き設定ダイアログ:** `wx.Notebook` を使用して BOA 設定パネルをアクセス可能なタブ（&Excel、&Word、&PowerPoint）に再構成し、スクリーンリーダーのナビゲーションを大幅に改善し、長いスクロールリストを排除しました。
* **NVDA 2026.2 互換性:** NVDA 2026.2 での動作をテストし認定しました。""",
        "readme_201": """### Version 2.0.1
#### UX/UI の改善
* **タブ付き設定ダイアログ:** `wx.Notebook` を使用して BOA 設定パネルをアクセス可能なタブ（&Excel、&Word、&PowerPoint）に再構成し、スクリーンリーダーのナビゲーションを大幅に改善し、長いスクロールリストを排除しました。`Alt+E`、`Alt+W`、`Alt+P` または標準の `Ctrl+PageDown`/`Ctrl+PageUp` を使用してタブ間を素早く切り替えることができます。
* **NVDA 2026.2 互換性:** NVDA 2026.2 での動作をテストし認定しました。"""
    },
    "ko": {
        "ui": {
            "Enable &Bulk Sheet Organizer and Quick Sheet Mover": "대량 시트 정리 도구 및 빠른 시트 이동 활성화(&B)",
            "Explicitly announce mer&ged cells": "병합된 셀 명시적 알림(&G)",
            "&Excel": "&Excel",
            "Prevent NVDA &crashes in Word text fields (SafeRichEdit)": "Word 텍스트 필드에서 NVDA 충돌 방지 (&C)(SafeRichEdit)",
            "&Word": "&Word",
            "Read hidden &hex codes when navigating the Standard Color hexagon grid": "표준 색상 육각형 그리드 탐색 시 숨겨진 16진수 코드 읽기(&H)",
            "Ensure the Hex Color edit &field is properly labeled": "16진수 색상 편집 필드에 올바른 레이블 지정 확인(&F)",
            "Enable &Slide Layout Analyzer (Prefix + L)": "슬라이드 레이아웃 분석기 활성화 (&S)(Prefix + L)",
            "Shape Movement &Audio Mode:": "도형 이동 오디오 모드(&A):",
            "&PowerPoint": "&PowerPoint"
        },
        "desc_intro": "BOA는 NVDA 사용자의 스크린 리더 환경을 대폭 개선하도록 설계된 AI 지원 개발 기반의 강력한 Microsoft Office 접근성 향상 도구 모음입니다. 전체 코드베이스는 Anti Gravity 2.0을 사용하여 생성되었습니다. 접근 불가능한 UI 요소를 직접 수정하고 Excel 및 PowerPoint를 위한 빠른 탐색 도구를 도입합니다.",
        "cl_201": """### Version 2.0.1
#### UX/UI 개선 사항
* **탭 설정 대화 상자:** `wx.Notebook`을 사용하여 BOA 설정 패널을 접근 가능한 탭(&Excel, &Word, &PowerPoint)으로 재구성하여 스크린 리더 탐색을 대폭 개선하고 긴 스크롤 목록을 제거했습니다.
* **NVDA 2026.2 호환성:** NVDA 2026.2에 대해 테스트 및 호환성을 인증했습니다.""",
        "readme_201": """### Version 2.0.1
#### UX/UI 개선 사항
* **탭 설정 대화 상자:** `wx.Notebook`을 사용하여 BOA 설정 패널을 접근 가능한 탭(&Excel, &Word, &PowerPoint)으로 재구성하여 스크린 리더 탐색을 대폭 개선하고 긴 스크롤 목록을 제거했습니다. `Alt+E`, `Alt+W`, `Alt+P` 또는 표준 `Ctrl+PageDown`/`Ctrl+PageUp`을 사용하여 탭 간을 빠르게 전환할 수 있습니다.
* **NVDA 2026.2 호환성:** NVDA 2026.2에 대해 테스트 및 호환성을 인증했습니다."""
    },
    "tr": {
        "ui": {
            "Enable &Bulk Sheet Organizer and Quick Sheet Mover": "&Toplu Sayfa Düzenleyiciyi ve Hızlı Sayfa Taşıyıcıyı Etkinleştir",
            "Explicitly announce mer&ged cells": "Bir&leştirilmiş hücreleri açıkça seslendir",
            "&Excel": "&Excel",
            "Prevent NVDA &crashes in Word text fields (SafeRichEdit)": "Word metin alanlarında NVDA &çökmelerini önle (SafeRichEdit)",
            "&Word": "&Word",
            "Read hidden &hex codes when navigating the Standard Color hexagon grid": "Standart Renk altıgen ızgarasında gezinirken gizli &hex kodlarını oku",
            "Ensure the Hex Color edit &field is properly labeled": "Hex Renk düzenleme &alanının doğru etiketlendiğinden emin olun",
            "Enable &Slide Layout Analyzer (Prefix + L)": "&Slayt Düzeni Çözümleyicisini Etkinleştir (Prefix + L)",
            "Shape Movement &Audio Mode:": "Şekil Hareketi &Ses Modu:",
            "&PowerPoint": "&PowerPoint"
        },
        "desc_intro": "BOA, Microsoft Office için NVDA kullanıcılarının ekran okuyucu deneyimini büyük ölçüde geliştirmek üzere tasarlanmış, yapay zeka destekli geliştirmenin ürünü güçlü bir erişilebilirlik paketidir. Kod tabanının tamamı Anti Gravity 2.0 kullanılarak oluşturulmuştur. Erişilemeyen kullanıcı arabirimi bileşenlerini doğrudan onarır ve Excel ile PowerPoint için hızlı gezinti araçları sunar.",
        "cl_201": """### Version 2.0.1
#### Kullanıcı Deneyimi/Arayüz İyileştirmeleri
* **Sekmeli Ayarlar İletişim Kutusu:** BOA Ayarlar Paneli, `wx.Notebook` kullanılarak erişilebilir sekmeler (&Excel, &Word ve &PowerPoint) halinde yeniden düzenlendi; bu sayede ekran okuyucu gezintisi büyük ölçüde kolaylaştı ve uzun kaydırma listeleri ortadan kalktı.
* **NVDA 2026.2 Uyumluluğu:** NVDA 2026.2 için test edildi ve onaylandı.""",
        "readme_201": """### Version 2.0.1
#### Kullanıcı Deneyimi/Arayüz İyileştirmeleri
* **Sekmeli Ayarlar İletişim Kutusu:** BOA Ayarlar Paneli, `wx.Notebook` kullanılarak erişilebilir sekmeler (&Excel, &Word ve &PowerPoint) halinde yeniden düzenlendi; bu sayede ekran okuyucu gezintisi büyük ölçüde kolaylaştı ve uzun kaydırma listeleri ortadan kalktı. `Alt+E`, `Alt+W`, `Alt+P` veya standart `Ctrl+PageDown`/`Ctrl+PageUp` kısayollarını kullanarak sekmeler arasında hızlıca geçiş yapabilirsiniz.
* **NVDA 2026.2 Uyumluluğu:** NVDA 2026.2 için test edildi ve onaylandı."""
    },
    "cs": {
        "ui": {
            "Enable &Bulk Sheet Organizer and Quick Sheet Mover": "Povolit &hromadný organizér listů a rychlý přesun listu",
            "Explicitly announce mer&ged cells": "Explicitně oznamovat slou&čené buňky",
            "&Excel": "&Excel",
            "Prevent NVDA &crashes in Word text fields (SafeRichEdit)": "Zabránit &pádům NVDA v textových polích Wordu (SafeRichEdit)",
            "&Word": "&Word",
            "Read hidden &hex codes when navigating the Standard Color hexagon grid": "Při procházení šestiúhelníkové mřížky standardních barev číst skryté &hex kódy",
            "Ensure the Hex Color edit &field is properly labeled": "Zajistit, aby editační &pole barvy Hex mělo správný popisek",
            "Enable &Slide Layout Analyzer (Prefix + L)": "Povolit analyzátor rozvržení &snímku (Prefix + L)",
            "Shape Movement &Audio Mode:": "Zvukový &režim pohybu tvarů:",
            "&PowerPoint": "&PowerPoint"
        },
        "desc_intro": "BOA je výkonná sada vylepšení přístupnosti pro Microsoft Office, výsledek vývoje za podpory umělé inteligence; navržená tak, aby výrazně zlepšila zážitek uživatelů odečítače obrazovky NVDA. Celá kódová základna byla vygenerována pomocí Anti Gravity 2.0. Přímo opravuje nepřístupné komponenty uživatelského rozhraní a přináší nástroje pro rychlou navigaci v Excelu a PowerPointu.",
        "cl_201": """### Version 2.0.1
#### Vylepšení UX/UI
* **Nastavení s kartami:** Panel nastavení BOA byl reorganizován do přístupných karet (&Excel, &Word a &PowerPoint) pomocí `wx.Notebook`, což výrazně zlepšuje navigaci s odečítačem obrazovky a eliminuje dlouhé seznamy.
* **Kompatibilita s NVDA 2026.2:** Otestováno a certifikováno pro NVDA 2026.2.""",
        "readme_201": """### Version 2.0.1
#### Vylepšení UX/UI
* **Nastavení s kartami:** Panel nastavení BOA byl reorganizován do přístupných karet (&Excel, &Word a &PowerPoint) pomocí `wx.Notebook`, což výrazně zlepšuje navigaci s odečítačem obrazovky a eliminuje dlouhé seznamy. Mezi kartami můžete rychle přepínat pomocí `Alt+E`, `Alt+W`, `Alt+P` nebo standardních klávesových zkratek `Ctrl+PageDown`/`Ctrl+PageUp`.
* **Kompatibilita s NVDA 2026.2:** Otestováno a certifikováno pro NVDA 2026.2."""
    },
    "ar": {
        "ui": {
            "Enable &Bulk Sheet Organizer and Quick Sheet Mover": "تمكين منظِّم أوراق العمل بالجملة والنقل السريع للورقة (&B)",
            "Explicitly announce mer&ged cells": "الإعلان بوضوح عن الخلايا الم&دمجة",
            "&Excel": "&Excel",
            "Prevent NVDA &crashes in Word text fields (SafeRichEdit)": "منع ت&عطل NVDA في حقول نصوص Word (SafeRichEdit)",
            "&Word": "&Word",
            "Read hidden &hex codes when navigating the Standard Color hexagon grid": "قراءة رموز &Hex المخفية عند التنقل في شبكة الألوان القياسية السداسية",
            "Ensure the Hex Color edit &field is properly labeled": "التأكد من تسمية &حقل تحرير لون Hex بشكل صحيح",
            "Enable &Slide Layout Analyzer (Prefix + L)": "تمكين محلل تخطيط ال&شريحة (Prefix + L)",
            "Shape Movement &Audio Mode:": "الوضع ال&صوتي لحركة الأشكال:",
            "&PowerPoint": "&PowerPoint"
        },
        "desc_intro": "BOA هي مجموعة قوية من تحسينات إمكانية الوصول لـ Microsoft Office، وهي نتاج التطوير بمساعدة الذكاء الاصطناعي؛ تم تصميمها لتحسين تجربة قارئ الشاشة لمستخدمي NVDA بشكل كبير. تم إنشاء قاعدة التعليمات البرمجية بالكامل باستخدام Anti Gravity 2.0. وهي تعالج مكونات واجهة المستخدم غير القابلة للوصول وتوفر أدوات تنقل سريعة في Excel وPowerPoint.",
        "cl_201": """### Version 2.0.1
#### تحسينات واجهة وتجربة المستخدم
* **مربع حوار الإعدادات المبوبة:** تمت إعادة تنظيم لوحة إعدادات BOA في علامات تبويب يمكن الوصول إليها (&Excel و&Word و&PowerPoint) باستخدام `wx.Notebook`، مما يسهل التنقل بشكل كبير باستخدام قارئ الشاشة ويزيل القوائم الطويلة.
* **التوافق مع NVDA 2026.2:** تم اختباره واعتماد توافقه مع NVDA 2026.2.""",
        "readme_201": """### Version 2.0.1
#### تحسينات واجهة وتجربة المستخدم
* **مربع حوار الإعدادات المبوبة:** تمت إعادة تنظيم لوحة إعدادات BOA في علامات تبويب يمكن الوصول إليها (&Excel و&Word و&PowerPoint) باستخدام `wx.Notebook`، مما يسهل التنقل بشكل كبير باستخدام قارئ الشاشة ويزيل القوائم الطويلة. يمكنك التبديل بسرعة بين علامات التبويب باستخدام `Alt+E` أو `Alt+W` أو `Alt+P` أو الاختصارات القياسية `Ctrl+PageDown`/`Ctrl+PageUp`.
* **التوافق مع NVDA 2026.2:** تم اختباره واعتماد توافقه مع NVDA 2026.2."""
    },
    "ur": {
        "ui": {
            "Enable &Bulk Sheet Organizer and Quick Sheet Mover": "بلک شیٹ آرگنائزر اور فوری شیٹ موور فعال کریں (&B)",
            "Explicitly announce mer&ged cells": "ضم شدہ (&G) سیلز کا واضح اعلان کریں",
            "&Excel": "&Excel",
            "Prevent NVDA &crashes in Word text fields (SafeRichEdit)": "Word کے متنی خانوں میں NVDA کریش کو روکیں (&C) (SafeRichEdit)",
            "&Word": "&Word",
            "Read hidden &hex codes when navigating the Standard Color hexagon grid": "معیاری رنگوں کی مسدس گرڈ پر نیویگیٹ کرتے وقت چھپے ہوئے &Hex کوڈز پڑھیں",
            "Ensure the Hex Color edit &field is properly labeled": "یقینی بنائیں کہ Hex رنگین ترمیمی فیلڈ (&F) پر مناسب لیبل موجود ہے",
            "Enable &Slide Layout Analyzer (Prefix + L)": "سلائیڈ لے آؤٹ اینالائزر کو فعال کریں (&S) (Prefix + L)",
            "Shape Movement &Audio Mode:": "شکل کی نقل و حرکت کا آڈیو موڈ (&A):",
            "&PowerPoint": "&PowerPoint"
        },
        "desc_intro": "BOA مائیکروسافٹ آفس کے لیے رسائی میں بہتری کا ایک طاقتور سوٹ ہے، جو AI کی مدد سے کی گئی ترقی کا نتیجہ ہے؛ یہ NVDA صارفین کے اسکرین ریڈر کے تجربے کو نمایاں طور پر بہتر بنانے کے لیے ڈیزائن کیا گیا ہے۔ پورا کوڈ بیس Anti Gravity 2.0 کے ذریعے تیار کیا گیا ہے۔ یہ غیر رسائی والے UI اجزاء کو براہ راست درست کرتا ہے اور Excel اور PowerPoint کے لیے تیز رفتار نیویگیشن ٹولز متعارف کراتا ہے۔",
        "cl_201": """### Version 2.0.1
#### UX/UI میں بہتری
* **ٹیب شدہ ترتیبات کا ڈائیلاگ:** `wx.Notebook` کا استعمال کرتے ہوئے BOA ترتیبات کے پینل کو قابل رسائی ٹیبز (&Excel، &Word، اور &PowerPoint) میں دوبارہ ترتیب دیا گیا ہے، جس سے اسکرین ریڈر کی نیویگیشن میں زبردست بہتری آئی ہے اور لمبی اسکرولنگ لسٹ ختم ہو گئی ہے۔
* **NVDA 2026.2 مطابقت:** NVDA 2026.2 کے لیے جانچا اور تصدیق شدہ۔""",
        "readme_201": """### Version 2.0.1
#### UX/UI میں بہتری
* **ٹیب شدہ ترتیبات کا ڈائیلاگ:** `wx.Notebook` کا استعمال کرتے ہوئے BOA ترتیبات کے پینل کو قابل رسائی ٹیبز (&Excel، &Word، اور &PowerPoint) میں دوبارہ ترتیب دیا گیا ہے، جس سے اسکرین ریڈر کی نیویگیشن میں زبردست بہتری آئی ہے اور لمبی اسکرولنگ لسٹ ختم ہو گئی ہے۔ آپ `Alt+E`، `Alt+W`، `Alt+P` یا معیاری شارٹ کٹس `Ctrl+PageDown`/`Ctrl+PageUp` کا استعمال کر کے ٹیبز کے درمیان فوری سوئچ کر سکتے ہیں۔
* **NVDA 2026.2 مطابقت:** NVDA 2026.2 کے لیے جانچا اور تصدیق شدہ۔"""
    },
    "pa": {
        "ui": {
            "Enable &Bulk Sheet Organizer and Quick Sheet Mover": "ਬਲਕ ਸ਼ੀਟ ਆਰਗੇਨਾਈਜ਼ਰ ਅਤੇ ਤਤਕਾਲ ਸ਼ੀਟ ਮੂਵਰ ਨੂੰ ਸਮਰੱਥ ਕਰੋ (&B)",
            "Explicitly announce mer&ged cells": "ਮਿਲਾਏ ਗਏ ਸੈੱਲਾਂ ਦਾ ਸਪਸ਼ਟ ਐਲਾਨ ਕਰੋ (&G)",
            "&Excel": "&Excel",
            "Prevent NVDA &crashes in Word text fields (SafeRichEdit)": "Word ਟੈਕਸਟ ਖੇਤਰਾਂ ਵਿੱਚ NVDA ਕ੍ਰੈਸ਼ ਹੋਣ ਤੋਂ ਰੋਕੋ (&C) (SafeRichEdit)",
            "&Word": "&Word",
            "Read hidden &hex codes when navigating the Standard Color hexagon grid": "ਸਟੈਂਡਰਡ ਕਲਰ ਹੈਕਸਾਗਨ ਗਰਿੱਡ 'ਤੇ ਨੈਵੀਗੇਟ ਕਰਦੇ ਸਮੇਂ ਲੁਕੇ ਹੋਏ &Hex ਕੋਡ ਪੜ੍ਹੋ",
            "Ensure the Hex Color edit &field is properly labeled": "ਯਕੀਨੀ ਬਣਾਓ ਕਿ Hex ਰੰਗ ਸੰਪਾਦਨ ਖੇਤਰ (&F) ਸਹੀ ਢੰਗ ਨਾਲ ਲੇਬਲ ਕੀਤਾ ਗਿਆ ਹੈ",
            "Enable &Slide Layout Analyzer (Prefix + L)": "ਸਲਾਈਡ ਲੇਆਉਟ ਐਨਾਲਾਈਜ਼ਰ ਨੂੰ ਸਮਰੱਥ ਕਰੋ (&S) (Prefix + L)",
            "Shape Movement &Audio Mode:": "ਆਕਾਰ ਅੰਦੋਲਨ ਆਡੀਓ ਮੋਡ (&A):",
            "&PowerPoint": "&PowerPoint"
        },
        "desc_intro": "BOA ਮਾਈਕ੍ਰੋਸਾਫਟ ਆਫਿਸ ਲਈ ਪਹੁੰਚਯੋਗਤਾ ਸੁਧਾਰਾਂ ਦਾ ਇੱਕ ਸ਼ਕਤੀਸ਼ਾਲੀ ਸੂਟ ਹੈ, ਜੋ AI ਸਹਾਇਤਾ ਪ੍ਰਾਪਤ ਵਿਕਾਸ ਦਾ ਨਤੀਜਾ ਹੈ; NVDA ਉਪਭੋਗਤਾਵਾਂ ਲਈ ਸਕ੍ਰੀਨ ਰੀਡਰ ਅਨੁਭਵ ਨੂੰ ਬਹੁਤ ਬਿਹਤਰ ਬਣਾਉਣ ਲਈ ਤਿਆਰ ਕੀਤਾ ਗਿਆ ਹੈ। ਸਮੁੱਚਾ ਕੋਡਬੇਸ Anti Gravity 2.0 ਦੀ ਵਰਤੋਂ ਕਰਕੇ ਤਿਆਰ ਕੀਤਾ ਗਿਆ ਹੈ। ਇਹ ਪਹੁੰਚਯੋਗ ਨਾ ਹੋਣ ਵਾਲੇ UI ਹਿੱਸਿਆਂ ਨੂੰ ਸਿੱਧਾ ਪੈਚ ਕਰਦਾ ਹੈ ਅਤੇ Excel ਅਤੇ PowerPoint ਲਈ ਤੇਜ਼ ਨੇਵੀਗੇਸ਼ਨ ਟੂਲ ਪੇਸ਼ ਕਰਦਾ ਹੈ।",
        "cl_201": """### Version 2.0.1
#### UX/UI ਸੁਧਾਰ
* **ਟੈਬ ਕੀਤੀਆਂ ਸੈਟਿੰਗਾਂ ਡਾਇਲਾਗ:** `wx.Notebook` ਦੀ ਵਰਤੋਂ ਕਰਦੇ ਹੋਏ BOA ਸੈਟਿੰਗਾਂ ਪੈਨਲ ਨੂੰ ਪਹੁੰਚਯੋਗ ਟੈਬਾਂ (&Excel, &Word, ਅਤੇ &PowerPoint) ਵਿੱਚ ਮੁੜ ਵਿਵਸਥਿਤ ਕੀਤਾ ਗਿਆ ਹੈ, ਜਿਸ ਨਾਲ ਸਕ੍ਰੀਨ ਰੀਡਰ ਨੈਵੀਗੇਸ਼ਨ ਵਿੱਚ ਬਹੁਤ ਸੁਧਾਰ ਹੋਇਆ ਹੈ।
* **NVDA 2026.2 ਅਨੁਕੂਲਤਾ:** NVDA 2026.2 ਲਈ ਪਰਖਿਆ ਅਤੇ ਪ੍ਰਮਾਣਿਤ ਕੀਤਾ ਗਿਆ ਹੈ।""",
        "readme_201": """### Version 2.0.1
#### UX/UI ਸੁਧਾਰ
* **ਟੈਬ ਕੀਤੀਆਂ ਸੈਟਿੰਗਾਂ ਡਾਇਲਾਗ:** `wx.Notebook` ਦੀ ਵਰਤੋਂ ਕਰਦੇ ਹੋਏ BOA ਸੈਟਿੰਗਾਂ ਪੈਨਲ ਨੂੰ ਪਹੁੰਚਯੋਗ ਟੈਬਾਂ (&Excel, &Word, ਅਤੇ &PowerPoint) ਵਿੱਚ ਮੁੜ ਵਿਵਸਥਿਤ ਕੀਤਾ ਗਿਆ ਹੈ, ਜਿਸ ਨਾਲ ਸਕ੍ਰੀਨ ਰੀਡਰ ਨੈਵੀਗੇਸ਼ਨ ਵਿੱਚ ਬਹੁਤ ਸੁਧਾਰ ਹੋਇਆ ਹੈ। ਤੁਸੀਂ `Alt+E`, `Alt+W`, `Alt+P` ਜਾਂ ਸਟੈਂਡਰਡ ਸ਼ਾਰਟਕੱਟ `Ctrl+PageDown`/`Ctrl+PageUp` ਦੀ ਵਰਤੋਂ ਕਰਕੇ ਟੈਬਾਂ ਵਿਚਕਾਰ ਤੇਜ਼ੀ ਨਾਲ ਸਵਿੱਚ ਕਰ ਸਕਦੇ ਹੋ।
* **NVDA 2026.2 ਅਨੁਕੂਲਤਾ:** NVDA 2026.2 ਲਈ ਪਰਖਿਆ ਅਤੇ ਪ੍ਰਮਾਣਿਤ ਕੀਤਾ ਗਿਆ ਹੈ।"""
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
            for marker in ["\nKey Features:", "\n    Key Features:", "\n    主要功能：", "\n    主要機能:", "\n    주요 기능:", "\n    Temel Özellikler:", "\n    Klíčové funkce:", "\n    الميزات الرئيسية:", "\n    اہم خصوصیات:", "\n    ਮੁੱਖ ਵਿਸ਼ੇਸ਼ਤਾਵਾਂ:"]:
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
