# -*- coding: UTF-8 -*-
# Copyright (C) 2026 KIRAN G T {MisterK} and Antigravity 2
# This file is covered by the GNU General Public License (GPL), version 2.
# See the file COPYING.txt for more details.

# Build customizations
# Change this file instead of sconstruct or manifest files, whenever possible.

_ = lambda x: x

# Full add-on dictionary, to build manifest and add-on plural tags
addon_info = {
    # add-on Name, internal for nvda
    "addon_name": "BOA",
    # Add-on summary, usually the user visible name of the addon.
    "addon_summary": "BOA: Better Office Accessibility",
    # Add-on description
    "addon_description": _("""A powerful suite of accessibility enhancements for Microsoft Office, designed to vastly improve the screen reader experience for NVDA users. 
    
    Key Features:
    - Cell Monitor: Save and continuously track dynamic Excel cells.
    - Sheet Layout Analyzer: Instantly scan worksheets for hidden tabs, protected states, and data blocks.
    - Bulk Sheet Organizer & Quick Mover: Instantly move or reorder multiple sheets using accessible dialogs.
    - PowerPoint Color Pickers: Read hidden Hex and RGB codes inside standard color grids.
    
    Please refer to the add-on help documentation for a complete list of features and shortcuts."""),
    "addon_version": "2.0.0",
    # Author(s)
    "addon_author": "MisterK with Anti Gravity2 <ravana.krn@gmail.com>",
    # URL for the add-on documentation support
    "addon_url": "https://github.com/MisterK-Dev/BOA-Better-Office-Accessibility-with-NVDA",
    # URL for the add-on repository where the source code is hosted
    "addon_sourceURL": "https://github.com/MisterK-Dev/BOA-Better-Office-Accessibility-with-NVDA",
    # Documentation file name
    "addon_docFileName": "readme.html",
    # Minimum NVDA version supported
    "addon_minimumNVDAVersion": "2026.1.0",
    # Last NVDA version supported/tested
    "addon_lastTestedNVDAVersion": "2026.1.1",
    # True if the addon requires an update to the manifest format.
    "addon_updateChannel": None,
    # Changelog URL or None
    "addon_changelog": _("""### Version 2.0.0
#### New Features
* **PowerPoint: Complete Document Analyzer (Experimental) (`NVDA+E`, then `D`):** A highly advanced, background-processed accessibility tool that maps out an entire presentation without freezing NVDA's speech engine. It provides a deeply navigable Virtual Table of Contents, detects Reading Order Mismatches (Visual Order vs. Z-Order), flags "Wall of Text" slides, and maps complex objects like SmartArt and Data Tables.
* **PowerPoint: Slide Layout Analyzer (Experimental) (`NVDA+E`, then `L`):** Instantly scans your currently active slide to understand its spatial layout and accessibility constraints, ensuring a completely smooth and responsive screen reader experience. means, here you will get details about current slide similar to Excel's sheet lay out analyzer.
* **PowerPoint: Bulk Slide Organizer (Experimental) (`NVDA+E`, then `X`):** Similar to the Excel feature, you can now instantly reorder, move, and arrange multiple PowerPoint slides at once using a fully accessible dialog.
* **PowerPoint: Shape Movement Audio Mode (Experimental):** Introduces 3D Spatial Audio cues to the PowerPoint canvas. Provides auditory feedback indicating the direction and boundary limits of an object as you move it, vastly improving spatial awareness. as mentioned this is experemental and waiting for feedbacks to improve or remove.
* **Word: Formatting Auditor (`NVDA+E`, then `F`):** Scans your Word document for formatting inconsistencies to ensure visual standards.
* **Word: Document Analyzer (`NVDA+E`, then `D`):** Instantly pull up a structural overview of your Word document. *(A special note of credit and thanks to Paul: This feature was directly inspired by his brilliant "Word Access" add-on. We are deeply grateful for his foundational work in this space!)*
* **Word: Automated Footnote Announcer:** Footnotes will now be automatically announced inline as you read, depending on your custom BOA settings. *(Note: Support for endnotes and comments is planned for a future release).*
* **Excel: The Power Editor (Accessible Formula Editor):** An absolute game-changer for modifying massive formulas.
  - **Single-Tap `NVDA+E`, then `F2`:** Instantly announces the raw formula string of the active cell (or announces "No formula").
  - **Double-Tap `NVDA+E`, then `F2`:** Opens a fully accessible, multi-line editor to safely modify massive, nested formulas. Native `Enter` adds line breaks for easy reading, and `Ctrl+Enter` saves it back to Excel.
  - *Safety Checks:* Safely traps syntax errors before they corrupt your sheet, and detects post-calculation errors (like `#NAME?` or `#DIV/0!`) to instantly warn you if a formula broke.
* **Excel: Formula Auditing & Evaluation:** Added custom shortcuts (`NVDA+E`, then `Shift+P` and `NVDA+E`, then `Shift+D`) to reliably trace Precedents and Dependents. Furthermore, Excel's native "Evaluate Formula" dialog is now fully accessible; NVDA automatically reads the evaluated results as you step through the calculation!
* **Excel: Cell Monitor Pro Upgrades:** 
  - **Slot Manager Dialog (`NVDA+E`, then `Alt+M`):** Opens a dialog listing all your actively monitored cells. Press `Enter` to instantly jump to one.
  - **Warp Back (`NVDA+E`, then `\\`):** Instantly teleports you back to your previous working cell after checking a slot.
  - **Direct Slot Jump (`Alt` + `Slot Number`):** Bypass the prefix entirely and instantly jump to an assigned cell slot.
* **Input Gestures Customization:** All features across all Office apps have been explicitly exposed to the native NVDA Input Gestures dialog, granting you complete freedom to customize every keyboard shortcut.

#### UX/UI Enhancements
* **Unified Browseable Reports:** We have adopted a unified HTML reporting system across the add-on. Features like the Excel Conditional Formatting Announcer, Layout Analyzers, and Document Analyzers no longer just speak massive blocks of text; their results now open in a native, navigable HTML window, allowing you to review the data at your own pace.
* **Excel: Enhanced Dependents/Precedents Tracking:** Vastly enhanced the speech output for Excel's native formula tracing shortcuts (`Ctrl+[` for Direct Precedents, and `Ctrl+]` for Direct Dependents). NVDA will now explicitly announce exactly what cells were selected.
* **Excel: Merge Cell Support:** Merged cells are now correctly detected and explicitly announced by the gap-skipping cell tracker.

#### Bug Fixes
* **Word: List Item Double Reading:** Implemented a temporary patch to fix the bug where NVDA double-reads paragraph list items in certain Word views.
* **Excel: Cell Monitor Localization Bug:** Resolved underlying tracking bugs caused by the recent translation localization updates."""),
}

# Define the python files that are the sources of your add-on.
# You can either list every file (using ""/") as a path separator,
# or use glob expressions.
pythonSources = [
    "addon/globalPlugins/*.py", 
    "addon/appModules/*.py",
    "addon/appModules/boa_enhancements/*.py",
    "addon/appModules/boa_enhancements/excel_enhancements/*.py",
    "addon/appModules/boa_enhancements/powerpoint_enhancements/*.py",
    "addon/appModules/boa_enhancements/word_enhancements/*.py",
    "addon/appModules/boa_enhancements/ui/*.py"
]

# Files that contain strings for translation. Usually your python sources
i18nSources = pythonSources + ["buildVars.py"]

# Files that will be ignored when building the nvda-addon file
# Paths are relative to the addon directory, not to the root directory of your addon sources.
excludedFiles = []

# Base language for the NVDA add-on
# If your add-on is written in a language other than english, modify this variable.
# For example, set baseLanguage to "es" if your add-on is primarily written in spanish.
baseLanguage = "en"

# Markdown extensions for add-on documentation
# Most add-ons do not require additional Markdown extensions.
# If you need to add support for markup such as tables, fill out the below list.
# Extensions string must be of the format "markdown.extensions.extensionName"
# e.g. "markdown.extensions.tables" to add tables.
markdownExtensions = ["markdown.extensions.tables"]

# Added for SCons compatibility
brailleTables = []
symbolDictionaries = []
speechDictionaries = []
