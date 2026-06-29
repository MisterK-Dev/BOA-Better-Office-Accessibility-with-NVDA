# Changelog

### Version 2.0.0
#### New Features
* **PowerPoint: Complete Document Analyzer (Experimental) (`NVDA+E`, then `D`):** A highly advanced, background-processed accessibility tool that maps out an entire presentation without freezing NVDA's speech engine. It provides a deeply navigable Virtual Table of Contents, detects Reading Order Mismatches (Visual Order vs. Z-Order), flags "Wall of Text" slides, and maps complex objects like SmartArt and Data Tables.
* **PowerPoint: Slide Layout Analyzer (Experimental) (`NVDA+E`, then `L`):** Instantly scans your currently active slide to understand its spatial layout and accessibility constraints, ensuring a completely smooth and responsive screen reader experience. means, here you will get details about current slide similar to Excel's sheet lay out analyzer.
* **PowerPoint: Bulk Slide Organizer (Experimental) (`NVDA+E`, then `X`):** Similar to the Excel feature, you can now instantly reorder, move, and arrange multiple PowerPoint slides at once using a fully accessible dialog.
* **PowerPoint: Shape Movement Audio Mode (Experimental):** Introduces 3D Spatial Audio cues to the PowerPoint canvas. Provides auditory feedback indicating the direction and boundary limits of an object as you move it, vastly improving spatial awareness.
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
* **Word: List Item Double Reading:** Implemented a temporary patch to fix the bug where NVDA double-reads paragraph list items.
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

### Version 1.4.0 - 2026-06-12
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
