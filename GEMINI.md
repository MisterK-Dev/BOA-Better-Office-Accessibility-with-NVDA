# BOA (Better Office Accessibility) for NVDA

Welcome! This `GEMINI.md` file serves as the master onboarding blueprint for the BOA repository. If you are an AI agent, read this document to orient yourself before proceeding.

## Mission Statement
BOA is a critical NVDA add-on designed to drastically improve the accessibility of Microsoft Office (Excel, Word, PowerPoint) by providing intelligent semantic reporting, bulk organizers, standard color grids, and enhanced COM object interaction without freezing the single-threaded NVDA engine.

## Agent Directives
1. **Always read `.agent/rules/rules.md` first**. It contains non-negotiable strict boundaries regarding NVDA safety, Excel COM handling, and our global translation function `_()`.
2. Do not arbitrarily rename files or guess API methods. Use `.agent/REFS_Sources/` for official documentation.
3. If you ever update the directory structure, you MUST trigger the `architecture_mapper` skill to regenerate this document.

## Architectural Guidelines & Safeguards

### 1. Approved Tech Stack
- **GUI:** Use `wxPython`.
- **Localization:** Use GNU `gettext` via `addonHandler.initTranslation()`.

### 2. Top 3 Anti-Patterns
1. **The Freeze Death:** Using `time.sleep()` or blocking `while` loops on the main thread instantly kills NVDA speech. Always use `core.callLater` or `wx.CallAfter`.
2. **The Translation Nuke:** Assigning data to `_` (e.g., `_, val = data`) overwrites the global localization function, causing instant `TypeError` crashes whenever text is spoken. Always use `_unused` instead.
3. **The COM Overload:** Looping over massive Excel grids or blindly reading `.UsedRange` without strict safety limits (e.g., `if count > 2000: break`) causes catastrophic COM bridge freezes.

### 3. Handling Ambiguity
If you encounter a bug, face an architectural doubt, or are unsure of an NVDA API, **do not guess or hallucinate**. 
1. Consult `.agent/rules/rules.md`.
2. Review the previous conversation history and transcripts.
3. If the path forward is still unclear, pause immediately and ask the Developer for clarification.

### 4. AI Action Triggers
- **Verify Compilation:** Run `scons`.
- **Update Translations:** Run `scons pot` immediately after altering UI/Speech text.
- **Git Operations:** Do not stage or commit code (`git add`, `git commit`) unless the Developer explicitly instructs you to do so. Never push without permission.

## Architecture Map
<!-- MAP_START -->
```
├── .agent/
│   ├── REFS_Sources/
│   │   ├── NVDA 2026.1.1 Developer Guide.mhtml
│   │   └── post-mortem.md
│   ├── rules/
│   │   └── rules.md
│   └── skills/
│       ├── architecture_mapper/
│       │   ├── scripts/
│       │   └── SKILL.md
│       ├── build_fixer/
│       │   ├── build_fixer.md
│       │   └── fix_po_newlines.py
│       ├── comment_analyzer/
│       │   └── comment_analyzer.md
│       ├── compatibility_guard/
│       │   ├── scripts/
│       │   ├── api_matrix.json
│       │   └── compatibility_guard.md
│       ├── gesture_auditor/
│       │   ├── scripts/
│       │   └── gesture_auditor.md
│       ├── release_validation/
│       │   └── release_validation.md
│       └── translation_pipeline/
│           ├── scripts/
│           └── translation_pipeline.md
├── addon/
│   ├── appModules/
│   │   ├── boa_enhancements/
│   │   │   ├── excel_enhancements/
│   │   │   ├── powerpoint_enhancements/
│   │   │   ├── word_enhancements/
│   │   │   ├── __init__.py
│   │   │   ├── boa_config.py
│   │   │   ├── boa_gui.py
│   │   │   └── safe_rich_edit.py
│   │   ├── excel.py
│   │   ├── powerpnt.py
│   │   └── winword.py
│   ├── doc/
│   │   ├── ar/
│   │   │   ├── readme.html
│   │   │   └── readme.md
│   │   ├── bn/
│   │   │   ├── readme.html
│   │   │   └── readme.md
│   │   ├── cs/
│   │   │   ├── readme.html
│   │   │   └── readme.md
│   │   ├── de/
│   │   │   ├── readme.html
│   │   │   └── readme.md
│   │   ├── en/
│   │   │   ├── readme.html
│   │   │   └── readme.md
│   │   ├── es/
│   │   │   ├── readme.html
│   │   │   └── readme.md
│   │   ├── fr/
│   │   │   ├── readme.html
│   │   │   └── readme.md
│   │   ├── gu/
│   │   │   ├── readme.html
│   │   │   └── readme.md
│   │   ├── hi/
│   │   │   ├── readme.html
│   │   │   └── readme.md
│   │   ├── it/
│   │   │   ├── readme.html
│   │   │   └── readme.md
│   │   ├── ja/
│   │   │   ├── readme.html
│   │   │   └── readme.md
│   │   ├── kn/
│   │   │   ├── readme.html
│   │   │   └── readme.md
│   │   ├── ko/
│   │   │   ├── readme.html
│   │   │   └── readme.md
│   │   ├── ml/
│   │   │   ├── readme.html
│   │   │   └── readme.md
│   │   ├── mr/
│   │   │   ├── readme.html
│   │   │   └── readme.md
│   │   ├── pa/
│   │   │   ├── readme.html
│   │   │   └── readme.md
│   │   ├── pl/
│   │   │   ├── readme.html
│   │   │   └── readme.md
│   │   ├── pt/
│   │   │   ├── readme.html
│   │   │   └── readme.md
│   │   ├── ru/
│   │   │   ├── readme.html
│   │   │   └── readme.md
│   │   ├── ta/
│   │   │   ├── readme.html
│   │   │   └── readme.md
│   │   ├── te/
│   │   │   ├── readme.html
│   │   │   └── readme.md
│   │   ├── tr/
│   │   │   ├── readme.html
│   │   │   └── readme.md
│   │   ├── uk/
│   │   │   ├── readme.html
│   │   │   └── readme.md
│   │   ├── ur/
│   │   │   ├── readme.html
│   │   │   └── readme.md
│   │   ├── zh_CN/
│   │   │   ├── readme.html
│   │   │   └── readme.md
│   │   └── style.css
│   ├── globalPlugins/
│   │   └── boa_settings.py
│   ├── locale/
│   │   ├── ar/
│   │   │   ├── LC_MESSAGES/
│   │   │   └── manifest.ini
│   │   ├── bn/
│   │   │   ├── LC_MESSAGES/
│   │   │   └── manifest.ini
│   │   ├── cs/
│   │   │   ├── LC_MESSAGES/
│   │   │   └── manifest.ini
│   │   ├── de/
│   │   │   ├── LC_MESSAGES/
│   │   │   └── manifest.ini
│   │   ├── es/
│   │   │   ├── LC_MESSAGES/
│   │   │   └── manifest.ini
│   │   ├── fr/
│   │   │   ├── LC_MESSAGES/
│   │   │   └── manifest.ini
│   │   ├── gu/
│   │   │   ├── LC_MESSAGES/
│   │   │   └── manifest.ini
│   │   ├── hi/
│   │   │   ├── LC_MESSAGES/
│   │   │   └── manifest.ini
│   │   ├── it/
│   │   │   ├── LC_MESSAGES/
│   │   │   └── manifest.ini
│   │   ├── ja/
│   │   │   ├── LC_MESSAGES/
│   │   │   └── manifest.ini
│   │   ├── kn/
│   │   │   ├── LC_MESSAGES/
│   │   │   └── manifest.ini
│   │   ├── ko/
│   │   │   ├── LC_MESSAGES/
│   │   │   └── manifest.ini
│   │   ├── ml/
│   │   │   ├── LC_MESSAGES/
│   │   │   └── manifest.ini
│   │   ├── mr/
│   │   │   ├── LC_MESSAGES/
│   │   │   └── manifest.ini
│   │   ├── pa/
│   │   │   ├── LC_MESSAGES/
│   │   │   └── manifest.ini
│   │   ├── pl/
│   │   │   ├── LC_MESSAGES/
│   │   │   └── manifest.ini
│   │   ├── pt/
│   │   │   ├── LC_MESSAGES/
│   │   │   └── manifest.ini
│   │   ├── ru/
│   │   │   ├── LC_MESSAGES/
│   │   │   └── manifest.ini
│   │   ├── ta/
│   │   │   ├── LC_MESSAGES/
│   │   │   └── manifest.ini
│   │   ├── te/
│   │   │   ├── LC_MESSAGES/
│   │   │   └── manifest.ini
│   │   ├── tr/
│   │   │   ├── LC_MESSAGES/
│   │   │   └── manifest.ini
│   │   ├── uk/
│   │   │   ├── LC_MESSAGES/
│   │   │   └── manifest.ini
│   │   ├── ur/
│   │   │   ├── LC_MESSAGES/
│   │   │   └── manifest.ini
│   │   └── zh_CN/
│   │       ├── LC_MESSAGES/
│   │       └── manifest.ini
│   └── manifest.ini
├── .gitignore
├── .sconsign.dblite
├── CHANGELOG.md
├── COPYING.txt
├── GEMINI.md
├── README.md
├── buildVars.py
├── manifest-translated.ini.tpl
├── manifest.ini.tpl
├── sconstruct
└── style.css
```

<!-- MAP_END -->
