# BOA: Better Office Accessibility

BOA is a powerful suite of accessibility enhancements for Microsoft Office, outcome of AI assisted development; designed to vastly improve the screen reader experience for NVDA users. The Entire code base has been generated using Anti Gravity 2.0. It directly patches inaccessible UI components and introduces rapid navigation tools for Excel and PowerPoint.

---

## ⌨️ Hotkey Reference
* Prefix refers to NVDA plus E, default, you can change it under input gestures.

| Feature | Key Combination | Context / Notes |
| :--- | :--- | :--- |
| **Enter Command Mode** | `[Prefix]` (Default: `NVDA+E`) | Activates Command Prefix Mode (triggers a high-pitched beep) |
| **Cancel Command Mode** | `Escape` | Exits Command Prefix Mode |
| **EXCEL ENHANCEMENTS** | | |
| **Analyze Sheet Layout** | `[Prefix]`, then `L` | Run within Excel before navigating data blocks |
| **Jump to Nearest Data Block** | `[Prefix]`, then `J` | Requires Layout Analysis first |
| **Open Bulk Sheet Organizer** | `[Prefix]`, then `X` | Opens the accessible sheet reordering dialog |
| **Raw Formula Announcer** | `[Prefix]`, then `F2` | Single tap to hear the raw formula string |
| **Power Formula Editor** | `[Prefix]`, then `F2` twice | Double tap to open the accessible multi-line formula editor |
| **Trace Precedents** | `[Prefix]`, then `Shift+P` | Trace Precedents same feature in accessible manner.|
| **Trace Dependents** | `[Prefix]`, then `Shift+D` | Trace Dependents same feature in acesible manner, pressing enter on a cell will teleport you there.|
| **Detailed Conditional Formatting**| `[Prefix]`, then `F` | Announces complete formatting details of focused cell |
| **Move Active Sheet Left** | `NVDA+Shift+LeftArrow` | Shifts the active sheet one position up |
| **Move Active Sheet Right** | `NVDA+Shift+RightArrow` | Shifts the active worksheet one position down |
| **Move Sheet to Start/End** | `NVDA+Shift+Home` / `End` | Sends worksheet to the absolute boundaries |
| **Hide / Unhide Row** | `Ctrl+9` / `Ctrl+Shift+9` | Native shortcut; BOA explicitly announces the visibility change |
| **Hide / Unhide Column** | `Ctrl+0` / `Ctrl+Shift+0` | Native shortcut; BOA explicitly announces the visibility change |
| **Unhide Column (Fallback)** | `NVDA+Ctrl+Shift+0` | Bypasses Windows input language hotkey conflicts |
| **Map Cell to Memory Slot** | `[Prefix]`, then `Shift+1` to `Shift+9` | Assigns current cell to a background monitor slot |
| **Read Monitored Cell Slot** | `[Prefix]`, then `1` to `9` | Recalls and reads the value of the assigned slot |
| **Direct Slot Jump** | `Alt` + `1` to `9` | Instantly jump your cursor to a monitored slot |
| **Warp Back to Previous Cell** | `[Prefix]`, then `\` | Instantly teleports you back after checking a slot |
| **Slot Manager Dialog** | `[Prefix]`, then `Alt+M` | Opens a dialog to view and manage all active monitors |
| **Toggle Background Monitoring** | `[Prefix]`, then `M` | Manually toggles background calculation tracking |
| **Clear All Memory Slots** | `[Prefix]`, then `Backspace` | Purges all saved background cell monitors |
| **POWERPOINT ENHANCEMENTS** | | |
| **Slide Layout Analyzer** | `[Prefix]`, then `L` | Analyzes and announces the spatial layout of the current slide |
| **Document Analyzer** | `[Prefix]`, then `D` | Generates a comprehensive Table of Contents and health report |
| **Bulk Slide Organizer** | `[Prefix]`, then `X` | Opens the accessible dialog to reorder multiple slides |
| **WORD ENHANCEMENTS** | | |
| **Formatting Auditor** | `[Prefix]`, then `F` | Audits the current document for formatting inconsistencies |
| **Document Analyzer** | `[Prefix]`, then `D` | Analyzes the current Word document layout and structure |

---

## 🚀 Features

### Excel Enhancements

#### 1. Sheet Layout Analyzer & Caching
Instantly scan any Excel worksheet to understand its structure, hidden elements, and data blocks.
* **How it works:** BOA quickly scans the sheet and announces active data blocks. It also warns you about **Hidden Worksheet Tabs**, active **Filters**, **Protected Modes**, and **Hidden Outer Boundaries** (e.g., if columns near the right edge of the sheet are hidden, preventing you from missing off-screen data).
* **Data Navigation:** After scanning, you can use the data block jump hotkeys to instantly warp your cursor between discovered data blocks, effortlessly bypassing thousands of empty cells.

#### 2. Bulk Sheet Organizer
Instantly reorder and arrange multiple sheets at once using a fully accessible dialog .
* **How it works:** Opens a dialog where you can select a sheet and map it to a new position. Scheduled moves are listed in a data table (press `Del` to remove a mistake). Click `OK` and your workbook is rearranged instantly.

#### 3. Quick Sheet Mover
Move the active sheet left, right, to the very beginning, or to the very end instantly using your keyboard shortcuts.

#### 4. Accessible Sheet Renaming
* When renaming a sheet, NVDA natively struggles to read the characters you are typing.
* BOA injects a custom `ExcelSheetRenameEdit` class that uses the `SafeRichEdit` engine, meaning you can precisely read by character, word, or line while renaming. This serves as an enhancement to the existing default renaming behavior.

#### 5. Hidden Row/Column Tracker
* Proactively tracks your movement across the grid to prevent you from missing hidden or filtered data.
* **Crossed Fragmented Cells:** If you jump across a heavily fragmented or hidden section of the grid (e.g., moving from Row 3 to Row 10 because Rows 4–9 are hidden), BOA explicitly announces "Rows 4 through 9 hidden". This ensures you always know when data has been skipped in the structure.

#### 6. Conditional Formatting Announcer
* Automatically reads the color, font style, and background shade of cells that have been dynamically changed by Excel's Conditional Formatting rules.
* Gives you the true visual state of the cell rather than just the raw underlying value. Initially, when focusing on the cell, it announces "has conditional formatting, and some other minor details". For comprehensive info, use the detailed hotkey configuration which is NVDA E and F.

#### 7. Better selection announcement
reads if cell or range selected or unselected.

#### 8 Cell monitor:
* **Cell Monitor:** Use command paths to map specific cells to memory slots. You can jump back and read them anytime using the assigned numerical slot.
* **Continuous Monitoring:** Slotted cells are automatically monitored in the background. If Excel triggers a recalculation or cell edit, BOA instantly announces the new value. Toggle manually or clear all via command slots.
* **Excel: Cell Monitor Pro Upgrades:** 
  - **Slot Manager Dialog (`NVDA+E`, then `Alt+M`):** Opens a dialog listing all your actively monitored cells. Press `Enter` to instantly jump to one.
  - **Warp Back (`NVDA+E`, then `\`):** Instantly teleports you back to your previous working cell after checking a slot.
  - **Direct Slot Jump (`Prefix + Alt` + `Slot Number`):** Bypass the instantly jump to an assigned cell slot.

#### 9 Power editor
* **Excel: The Power Editor (Accessible Formula Editor):** An absolute game-changer for modifying massive formulas.
  - **Single-Tap `NVDA+E`, then `F2`:** Instantly announces the raw formula string of the active cell (or announces "No formula").
  - **Double-Tap `NVDA+E`, then `F2`:** Opens a fully accessible, multi-line editor to safely modify massive, nested formulas. Native `Enter` adds line breaks for easy reading, and `Ctrl+Enter` saves it back to Excel.
  - *Safety Checks:* Safely traps syntax errors before they corrupt your sheet, and detects post-calculation errors (like `#NAME?` or `#DIV/0!`) to instantly warn you if a formula broke.

#### 10 Formula auditing and evaluation enhancements:
* **Excel: Formula Auditing & Evaluation:** Added custom shortcuts (`NVDA+E`, then `Shift+P` and `NVDA+E`, then `Shift+D`) to reliably trace Precedents and Dependents. Furthermore, Excel's native "Evaluate Formula" dialog is now fully accessible; NVDA automatically reads the evaluated results as you step through the calculation!

### PowerPoint Enhancements

#### 1. Accessible Color Pickers
* Unlocks the Custom Color dialog in PowerPoint.
* Identifies and explicitly reads out the "Red", "Green", and "Blue" edit boxes correctly (by overriding `PowerPointRGBEdit`).
* Maps the previously invisible Hex input field so NVDA can read the full Hex color value cleanly.

#### 2. Standard Color Grid Support
* Navigating the PowerPoint "Standard" color hexagon grid normally reads as "Graphic" or silence.
* BOA tracks your arrow keys across the hexagon and silently fetches the hidden color value, announcing it to you in real-time (e.g., "Color #FF0000").

#### 3 Bulk Slide Organizer:
* **PowerPoint: Bulk Slide Organizer (Experimental) (`NVDA+E`, then `X`):** Similar to the Excel feature, you can now instantly reorder, move, and arrange multiple PowerPoint slides at once using a fully accessible dialog.

#### 4 Slide lay out analyzer
* **PowerPoint: Slide Layout Analyzer (Experimental) (`NVDA+E`, then `L`):** Instantly scans your currently active slide to understand its spatial layout and accessibility constraints, ensuring a completely smooth and responsive screen reader experience. means, here you will get details about current slide similar to Excel's sheet lay out analyzer.


#### 5 Complete Document [PPT] analyzer
* **PowerPoint: Complete Document Analyzer (Experimental) (`NVDA+E`, then `D`):** A highly advanced, background-processed accessibility tool that maps out an entire presentation without freezing NVDA's speech engine. It provides a deeply navigable Virtual Table of Contents, detects Reading Order Mismatches (Visual Order vs. Z-Order), flags "Wall of Text" slides, and maps complex objects like SmartArt and Data Tables.

#### 6 shape movement [adjustment] enhancements:
* **PowerPoint: Shape Movement Audio Mode (Experimental):** Introduces 3D Spatial Audio cues to the PowerPoint canvas. Provides auditory feedback indicating the direction and boundary limits of an object as you move it, vastly improving spatial awareness.

### Word Enhancements:
#### 1. Document Analyzer inspired and derived from Paul's word access addon:
* **Word: Document Analyzer (`NVDA+E`, then `D`):** Instantly pull up a structural overview of your Word document. *(A special note of credit and thanks to Paul: This feature was directly inspired by his brilliant "Word Access" add-on. We are deeply grateful for his foundational work in this space!)*

#### 2 Formatting Auditor
* **Word: Formatting Auditor (`NVDA+E`, then `F`):** Scans your Word document for formatting inconsistencies to ensure visual standards.

#### 3 Foot note reader:
* **Word: Automated Footnote Announcer:** Footnotes will now be automatically announced inline as you read, depending on your custom BOA settings. *(Note: Support for endnotes and comments is planned for a future release).*

### Infrastructure & Technical Mechanisms

#### The Command Prefix Mode
To prevent keystroke conflicts with other NVDA plugins, BOA uses a **Command Prefix Mode**:
1. Press the activation hotkey to enter Command Mode. You will hear a high-pitched beep. Default is NVDA plus E.
2. Press a secondary key to trigger a specific feature.
3. If you press an invalid key, you will hear an error beep.

#### Customization & Settings Panel
* BOA features are fully modular and can be enabled or disabled at any time. Go to `NVDA Menu -> Preferences -> Settings -> BOA Office Enhancements` to toggle individual features on or off.
* **Intelligent Accelerator Keys:** Every single setting features a mathematically unique `Alt+Key` accelerator shortcut within the panel. For example, press `Alt+E` to instantly jump to the Excel group, `Alt+P` for PowerPoint, and `Alt+W` for Word.
* Settings are saved securely to a standalone JSON file (`boa_settings.json`), ensuring your core NVDA configuration is never corrupted.
* If Microsoft Office officially fixes an accessibility bug in the future, you can safely disable BOA's specific override hook without losing the rest of the addon's functionality.
* **Input Gestures Customization:** All features across all Office apps have been explicitly exposed to the native NVDA Input Gestures dialog, granting you complete freedom to customize every keyboard shortcut.

#### Security & Integration Boundaries
* Clipboard injections strictly verify window foreground process IDs to prevent leakage of data into other applications.
* some Custom hotkeys are fully exposed in NVDA's Input Gestures dialog under the "Better Office Accessibility" category.

---

## 📋 Requirements

* **NVDA:** Version 2026.1.0 or later.
* **Applications:** Microsoft Excel & Microsoft PowerPoint.

---

## 💾 Installation

1. Download the latest `.nvda-addon` release file, or locate it within the native NVDA Add-on Store.
2. if installing from file, Open the file or use `NVDA's Add-on Store -> Install from external file`.
3. Restart NVDA.

---

## 🛠️ Changelog
### Version 2.0.1
#### UX/UI Enhancements
* **Tabbed Settings Dialog: [inspired by Vision Assistant Pro]** Reorganized the BOA Settings Panel into accessible tabs (&Excel, &Word, and &PowerPoint) using `wx.Notebook`, vastly improving screen reader navigation and eliminating long scrolling lists.
* **NVDA 2026.2 Compatibility:** Tested and ensured compatibility for NVDA 2026.2.

### Version 2.0.0
#### New Features
* **PowerPoint: Complete Document Analyzer (Experimental) (`NVDA+E`, then `D`):** A highly advanced, background-processed accessibility tool that maps out an entire presentation without freezing NVDA's speech engine. It provides a deeply navigable Virtual Table of Contents, detects Reading Order Mismatches (Visual Order vs. Z-Order), flags "Wall of Text" slides, and maps complex objects like SmartArt and Data Tables.
* **PowerPoint: Slide Layout Analyzer (Experimental) (`NVDA+E`, then `L`):** Instantly scans your currently active slide to understand its spatial layout and accessibility constraints, ensuring a completely smooth and responsive screen reader experience. means, here you will get details about current slide similar to Excel's sheet lay out analyzer.
* **PowerPoint: Bulk Slide Organizer (Experimental) (`NVDA+E`, then `X`):** Similar to the Excel feature, you can now instantly reorder, move, and arrange multiple PowerPoint slides at once using a fully accessible dialog.
* **PowerPoint: Shape Movement Audio Mode (Experimental):** Introduces 3D Spatial Audio cues to the PowerPoint canvas. Provides auditory feedback indicating the direction and boundary limits of an object as you move it, vastly improving spatial awareness. well as mentioned this is experimental, waiting for feed backs to improve it.
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
  - **Warp Back (`NVDA+E`, then `\`):** Instantly teleports you back to your previous working cell after checking a slot.
  - **Direct Slot Jump (`Alt` + `Slot Number`):** Bypass the prefix entirely and instantly jump to an assigned cell slot.
* **Input Gestures Customization:** All features across all Office apps have been explicitly exposed to the native NVDA Input Gestures dialog, granting you complete freedom to customize every keyboard shortcut.

#### UX/UI Enhancements
* **Unified Browseable Reports:** We have adopted a unified HTML reporting system across the add-on. Features like the Excel Conditional Formatting Announcer, Layout Analyzers, and Document Analyzers no longer just speak massive blocks of text; their results now open in a native, navigable HTML window, allowing you to review the data at your own pace.
* **Excel: Enhanced Dependents/Precedents Tracking:** Vastly enhanced the speech output for Excel's native formula tracing shortcuts (`Ctrl+[` for Direct Precedents, and `Ctrl+]` for Direct Dependents). NVDA will now explicitly announce exactly what cells were selected.
* **Excel: Merge Cell Support:** Merged cells are now correctly detected and explicitly announced by the gap-skipping cell tracker.

#### Bug Fixes
* **Word: List Item Double Reading:** Implemented a temporary patch to fix the bug where NVDA double-reads paragraph list items in certain Word views.
* **Excel: Cell Monitor Localization Bug:** Resolved underlying tracking bugs caused by the recent translation localization updates.

### What's New in v1.6.1
* **Deep File Localization**: Fixed missing string translations deep within the Excel enhancement modules (such as the Sheet Layout Analyzer and Quick Sheet Mover) to ensure 100% localization coverage.
* **Expanded Translation Support**: Added 7 new languages to the system (Turkish, Polish, Korean, Ukrainian, Czech, Urdu, and Punjabi). 
  *(Note: These translations were generated by AI, so some minor translation errors or inaccuracies may be present.)*

### v1.6.0
* **Comprehensive Translation Support**: The add-on is now fully localized with support for 17 global languages. 
  *(Note: These translations were generated by AI, so some minor translation errors or inaccuracies may be present.)*
* **Strict Code Governance**: Applied GPL-2.0 copyright headers across the entire codebase."""),

### Version 1.5.0 
#### New Features
##### End of Data Radar
When navigating through large spreadsheets, it can be difficult to tell if an empty cell means you've reached the end of a list, or if there is simply a gap in the data. The **End of Data Radar** acts as a smart perimeter check to save you from blindly arrowing through empty space.
Whenever you navigate into an empty cell, BOA instantly scans the remaining cells in your direction of travel. If there is absolutely no data left, it will proactively announce:
* *"No more data below"*
* *"No more data above"*
* *"No more data to the right"*
* *"No more data to the left"*
**Configuration Options:**
You can configure this feature via `NVDA Preferences -> Settings -> BOA Office Enhancements`. Because spreadsheets can contain hidden complexities (like invisible formulas or collapsed rows), the radar provides three operating modes:
1. **Off**: Disables the radar entirely.
2. **Strict Memory Check (CountA) [Default]**: The safest and fastest approach. It checks the raw memory of the spreadsheet. If it detects *anything* below you (including hidden rows, text, numbers, or invisible formulas), it stays completely silent to prevent false alarms. It only announces "No more data" when the remainder of the sheet is 100% mathematically blank.
3. **Visible Data Only (Math Engine)**: A highly advanced engine designed for complex sheets. It intelligently filters out hidden rows and invisible formulas (e.g., `=""`). It will only stay silent if there are actual, visible numbers or text left in your path.

### Version 1.4 - 2026-06-12
#### New Features
* **Cell Monitor:** Use command paths to map specific cells to memory slots. You can jump back and read them anytime using the assigned numerical slot.
* **Continuous Monitoring:** Slotted cells are automatically monitored in the background. If Excel triggers a recalculation or cell edit, BOA instantly announces the new value. Toggle manually or clear all via command slots.

#### Bug Fixes

### Version 1.3.0 — 2026-06-05
*Final release.*

#### New Features
* **Sheet Layout Analyzer:** Added powerful layout scanning infrastructure. Instantly detects Worksheet Protection, active Column Filters, Hidden Worksheet Tabs, and hidden absolute borders while caching discovered data blocks.
* **Guided Data Block Navigation:** Post-analysis navigation allows immediate cursor warps between major clusters of data, bypassing empty cells seamlessly.
* **Conditional Formatting Announcer:** Automatically detects and reads the dynamic color, font style, and background shade of cells altered by Excel's Conditional Formatting rules.
* **Explicit Settings Accelerators:** Completely overhauled the BOA Settings GUI to strictly comply with NVDA architecture. Every feature checkbox now possesses a globally unique `Alt+Letter` shortcut, preventing keyboard cycling and eliminating first-letter navigation failures.

#### Bug Fixes
* **Absolute Edge Boundary Detection:** Replaced native COM `UsedRange` edge checks with absolute 1D mathematical boundary checks (`Row 1048576` and `Column 16384`) to guarantee detection of hidden rows/columns even if they lie far outside the active data block.
* **Lazy COM Property Safe Bailouts:** Hardened COM property loops to prevent NVDA thread freezes when evaluating millions of contiguous hidden structures.

### Version 1.2.0 — 2026-06-03
*Final release.*

#### New Features
* **App-Launch Caching:** Major architectural overhaul. Core modules are now lazy-loaded exactly when you focus on Office applications, eliminating boot lag, completely solving the 'unknown' object focus glitch on rename dialogs, and preserving multi-file codebase structure.
* **Enhanced Cell Tracker (1D COM Math):** Rewrote the hidden cell gap detection logic to only evaluate one-dimensional cross-sections (`current_col` or `current_row`). This reduces the COM calculation payload by over 16 million cells, instantly eliminating navigation freezes when jumping hidden ranges.
* **Process Memory Wiping:** Implemented Excel Window Handle (`Hwnd`) tracking to detect when the user closes and reopens Excel. This actively wipes out stale global state memory and completely solves the false "Sheet hidden" announcement when opening a fresh "Book1".

#### Bug Fixes
* **Double Selection Announcement:** Migrated away from unreliable asynchronous `winUser.getKeyState` and implemented `api.getLastInputGesture()` to perfectly suppress double announcements when using Shift+Arrow keys.
* **Boundary Detector Deactivation:** The Proactive Boundary Detector has been deactivated to protect NVDA native navigation stability, falling back entirely to the gap-skipping tracker.

### Version 1.1.0 — 2026-05-30
*Final release.*

#### New Features
* **Settings GUI:** Added a native BOA Office Enhancements panel inside `NVDA -> Preferences -> Settings` to easily toggle features on or off.
* **SafeRichEdit Hook:** Prevents silent NVDA crashes when interacting with RichEdit controls in Office 2024.
* **Customizable Hotkeys:** All BOA hotkeys are now fully exposed in NVDA's Input Gestures dialog under the "Better Office Accessibility" category.
* **Excel: Hidden Row/Column Skip Detection:** Proactively announces when navigating past hidden rows or columns, ensuring you never miss filtered data. Can be toggled in settings.

#### Bug Fixes
* **Thread Safety:** Removed all blocking delays (`time.sleep`) and replaced them with non-blocking NVDA asynchronous callbacks to ensure the screen reader never stutters during background operations.

### Version 1.0.0 — 2026-05-24
*Initial public release.*

#### New Features
* **Excel: Bulk Sheet Organizer:** Instantly reorder multiple sheets at once using a fully accessible dialog.
* **Excel: Quick Sheet Mover:** Move the active sheet left, right, to start, or to end via keyboard commands.
* **Excel: Accessible Sheet Renaming:** Intercepts the inaccessible native rename field and replaces it with a reliable accessible dialog.
* **Excel: Smart Selection Tracking:** Accurately announces multi-cell range selections and deselections.
* **PowerPoint: Accessible Color Pickers:** Enables NVDA to accurately read RGB and Hex values inside the Custom Color dialog.
* **PowerPoint: Standard Color Grid Support:** Intercepts arrow key navigation to read hidden Hex codes from the inaccessible color hexagon grid.