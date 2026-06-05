# BOA: Better Office Accessibility

BOA is a powerful suite of accessibility enhancements for Microsoft Office, designed to vastly improve the screen reader experience for NVDA users. It directly patches inaccessible UI components and introduces rapid navigation tools for Excel and PowerPoint.

## Excel Features

### Bulk Sheet Organizer
Instantly reorder and arrange multiple sheets at once using a fully accessible dialog table.
* **Shortcut:** `NVDA + Alt + C`
* **How it works:** Opens a dialog where you can select a sheet and map it to a new position. Scheduled moves are listed in a data table (press `Del` to remove a mistake). Click `OK` and your workbook is rearranged instantly.

### Quick Sheet Mover
Move the active sheet left, right, to the very beginning, or to the very end instantly using your keyboard.
* **Shortcuts:**
#### 1. Better Focus Loss Prevention
- Excel's native UI automation frequently loses focus when unselecting cells. BOA automatically restores and reads the correct cell focus to prevent NVDA from going silent.

#### 2. The Bulk Sheet Organizer
- Easily sort, delete, and rename sheets in massive Excel workbooks via an accessible dialog box.
- Press `NVDA+E` followed by `X` to instantly open the Grid Mover Dialog. 
- You can rapidly delete multiple sheets with the `Del` key and move them around using standard arrow keys.

#### 3. Quick Sheet Movers
- Swiftly swap sheet positions without opening a menu. 
- Press `NVDA+Shift+LeftArrow` to move the current sheet to the left.
- Press `NVDA+Shift+RightArrow` to move the current sheet to the right.

#### 4. Accessible Sheet Renaming
- When renaming a sheet, NVDA natively struggles to read the characters you are typing.
- BOA injects a custom `ExcelSheetRenameEdit` class that uses the `SafeRichEdit` engine, meaning you can precisely read by character, word, or line while renaming.

#### 5. Hidden Row/Column Tracker
- Proactively tracks your movement across the grid to prevent you from missing hidden or filtered data.
- **Crossed Fragmented Cells:** If you jump across a heavily fragmented or hidden section of the grid (e.g. moving from Row 3 to Row 10 because Rows 4-9 are hidden), BOA explicitly announces "Rows 4 through 9 hidden". This ensures you always know when data has been skipped in the structure.

#### 6. Sheet Layout Analyzer & Proactive Guidance
- Excel sheets with scattered data can be extremely disorienting. BOA scans the sheet and proactively guides you when you get lost in empty space.
- **Cache the Layout:** Press `NVDA+E, L` at any time to scan the active sheet and memorize all data blocks.
- **Guided Announcement:** While navigating, if you land on an empty cell, BOA instantly calculates the Manhattan distance to the nearest cached data block and tells you exactly where it is (e.g., "Nearest data at B4").
- **One-time Announcement:** Automatically calculates and announces the nearest data block the very first time you land on a new sheet, then goes dormant to keep quiet.
- **Manual Jump Teleportation:** At any time, if you are lost in empty cells, press `NVDA+E, J`. BOA will instantly teleport your cursor directly to the nearest data block.

---

### PowerPoint Features

#### 1. Accessible Color Pickers
- Unlocks the Custom Color dialog in PowerPoint.
- Identifies and explicitly reads out the "Red", "Green", and "Blue" edit boxes correctly (by overriding `PowerPointRGBEdit`).
- Maps the previously invisible Hex input field so NVDA can read the full Hex color value cleanly.

#### 2. Standard Color Grid Support
- Navigating the PowerPoint "Standard" color hexagon grid normally reads as "Graphic" or silence.
- BOA tracks your arrow keys (`Up, Down, Left, Right`) across the hexagon and silently fetches the hidden color value, announcing it to you in real-time (e.g. "Color #FF0000").

---

## The `NVDA+E` Command Prefix

To prevent keystroke conflicts with other NVDA plugins, BOA uses a **Command Prefix Mode**:
1. Press `NVDA+E` to enter Command Mode. You will hear a high-pitched beep.
2. Press a secondary key to trigger a feature.
   - Example: `NVDA+E` then `X` = Open the Excel Bulk Sheet Organizer.
   - Example: `NVDA+E` then `Escape` = Cancel command mode.

If you press an invalid key, you will hear an error beep.

## Customization

BOA features are fully modular and can be enabled or disabled at any time. 
Go to `NVDA Menu -> Preferences -> Settings -> Better Office Accessibility` to toggle individual features on or off without restarting NVDA.

## ⚙️ Settings & Configuration

As of v1.1.0, BOA includes a fully accessible native NVDA Settings Panel. 
You can find it by navigating to **NVDA Menu -> Preferences -> Settings -> BOA Office Enhancements**.

**Features of the Configuration Manager:**
- All features are enabled by default, but can be individually toggled off.
- The settings are grouped securely by application (Excel, PowerPoint, Word).
- Settings are saved securely to a standalone JSON file (`boa_settings.json`), ensuring your core NVDA configuration is never corrupted.
- If Microsoft Office officially fixes an accessibility bug in the future, you can safely disable BOA's specific override hook without losing the rest of the addon's functionality!

## Security
BOA is built with strict security boundaries. Background COM manipulations are executed within safely initialized sandboxes, and clipboard injections strictly verify window foreground process IDs to prevent leakage of data into other applications.

## Requirements
- NVDA 2026.1.0 or later.
- Microsoft Excel & PowerPoint.

## Installation
1. Download the latest `.nvda-addon` release file.
2. Open the file or use NVDA's Add-on Store -> Install from external file.
3. Restart NVDA.

## Changelog

---

### Version 1.3.0-dev2 — 2026-06-05
**Development build.**

#### New Features
- **Conditional Formatting Detection** — Added a lightweight background scanner that passively detects conditional formatting as you navigate.
- **Conditional Formatting Deep Dive (`NVDA+E, F`)** — Added a dedicated shortcut to extract the complete formatting rules (Color Scales, Data Bars, Icon Sets, formulas) and display them in the accessible BOA dialog, including the "Applies To" range and active visual outcomes.

#### Technical Learnings & Architecture Fixes
- **Utilizing NVDA Infrastructure (Colors)** — We learned a valuable lesson in not reinventing the wheel. We initially built a custom math algorithm to translate COM integer colors to text, which produced results like "Light Pink" when NVDA natively said "Light Pale Orange". We removed the custom dictionary and successfully hooked directly into NVDA's native `colors` module, guaranteeing 100% parity with NVDA's built-in formatting announcements.
- **COM Named Properties (`AppliesTo.Address`)** — We discovered that in Python's `comtypes`, fetching the `Address` property on a `FormatCondition` range returns a `<comtypes.client.lazybind.NamedProperty>` object instead of the text string because it can accept arguments (e.g. absolute vs relative). We fixed this by explicitly invoking it with parentheses `Address()` to evaluate the default property fetch.
- **Deep Rule Parsing** — A rule type of `xlColorScale` or `xlDatabar` is just the surface. We learned to drill down into the hidden `ColorScaleCriteria` and `MinPoint`/`MaxPoint` sub-objects to extract the exact `ConditionValue` type (e.g., "Lowest Value", "Percentile", "Number") to provide a 1:1 recreation of the visual Rules Manager logic.
- **Visual Hiding Extraction** — We mapped the `ShowValue` and `ShowIconOnly` boolean properties so the add-on can explicitly warn blind users when a sighted user has visually hidden the underlying cell text in favor of a Data Bar or Icon.

---

### Version 1.3.0-dev1 — 2026-06-05
**Development build.**

#### New Features
- **Sheet Layout Analyzer** — Added `NVDA+E, L` to scan and memorize scattered data blocks on a sheet.
- **Proactive Guidance (Auto-Announce)** — Added Guided and One-time announcement modes that proactively tell the user the coordinates of the nearest data block when navigating into empty space.
- **Manual Jump (`NVDA+E, J`)** — Added a teleportation shortcut that calculates the Manhattan distance to all known blocks and natively moves the Excel selection to the closest one.

#### Technical Learnings & Architecture Fixes
- **Robust Empty Cell Detection** — Excel COM `cell.Value` is notoriously unreliable for detecting "empty" cells (it can return `VT_EMPTY`, `VT_NULL`, or evaluate formulas to `""` or `0`). We built a highly robust checker that evaluates both the underlying COM `Value` and the visual `cell.Text` to guarantee the cell actually appears empty to the user before triggering guidance.
- **Garbage Collection of NVDA Timers** — We discovered that scheduling delayed speech in NVDA via `core.callLater(1000, lambda: ui.message(msg))` can fail silently because Python's garbage collector destroys the weak lambda closure before the 1-second timer fires. This was fixed by passing the explicit function and arguments `core.callLater(1000, ui.message, msg)` to lock them in memory.
- **UIA COM Object Resolution** — We discovered that under modern UI Automation, relying on `AccessibleObjectFromWindow` using the `EXCEL7` window class fails entirely because the window class is now `NetUIHWND`. We fortified the object manager to natively fallback to `GetActiveObject("Excel.Application")` to ensure COM connectivity remains stable regardless of the UI framework.
- **Eliminating Speech Delays** — Removed intentional 1000ms speech delays from guidance announcements. They were originally added to prevent NVDA's native selection reading from interrupting the add-on, but users preferred instant feedback over staggered speech queues.

### Version 1.2.0 — 2026-06-03
**Final release.**

#### New Features
- **App-Launch Caching** — Major architectural overhaul. Core modules are now lazy-loaded exactly when you focus on Office applications, eliminating boot lag, completely solving the 'unknown' object focus glitch on rename dialogs, and preserving multi-file codebase structure.
- **Enhanced Cell Tracker (1D COM Math)** — Rewrote the hidden cell gap detection logic to only evaluate one-dimensional cross-sections (`current_col` or `current_row`). This reduces the COM calculation payload by over 16 million cells, instantly eliminating navigation freezes when jumping hidden ranges.
- **Process Memory Wiping** — Implemented Excel Window Handle (`Hwnd`) tracking to detect when the user closes and reopens Excel. This actively wipes out stale global state memory and completely solves the false "Sheet hidden" announcement when opening a fresh "Book1".
- **Intelligent API Codebase Comments** — Completely documented the codebase using Automated Docstring injections. Every function, method, and NVDA UIA interaction is now fully documented for future maintainability.

#### Bug Fixes
- **Double Selection Announcement** — Migrated away from unreliable asynchronous `winUser.getKeyState` and implemented `api.getLastInputGesture()` to perfectly suppress double announcements when using Shift+Arrow keys.
- **Boundary Detector Deactivation** — The Proactive Boundary Detector has been deactivated to protect NVDA native navigation stability, falling back entirely to the gap-skipping tracker.

---

### Version 1.1.0 — 2026-05-30
**Final release.**

#### New Features
- **Settings GUI** — Added a native BOA Office Enhancements panel inside NVDA -> Preferences -> Settings to easily toggle features on or off.
- **SafeRichEdit Hook** — Prevents silent NVDA crashes when interacting with RichEdit controls in Office 2024.
- **Customizable Hotkeys** — All BOA hotkeys are now fully exposed in NVDA's Input Gestures dialog under the "Better Office Accessibility" category.
- **Excel: Hidden Row/Column Skip Detection** — Proactively announces when navigating past hidden rows or columns, ensuring you never miss filtered data. Can be toggled in settings.

#### Bug Fixes
- **Thread Safety** — Removed all blocking delays (`time.sleep`) and replaced them with non-blocking NVDA asynchronous callbacks to ensure the screen reader never stutters during background operations.
- **Excel: Selection Double-Announcement** — Fixed an issue where BOA redundantly announced multi-cell selections that NVDA natively handles (e.g., via Shift+Arrow keys).

---

### Version 1.0.0 — 2026-05-24
**Initial public release.**

#### New Features
- **Excel: Bulk Sheet Organizer** (`NVDA+Alt+C`) — Instantly reorder multiple sheets at once using a fully accessible dialog.
- **Excel: Quick Sheet Mover** (`NVDA+Shift+Arrows`) — Move the active sheet left, right, to start, or to end via keyboard.
- **Excel: Accessible Sheet Renaming** — Intercepts the inaccessible native rename field and replaces it with a reliable accessible dialog.
- **Excel: Smart Selection Tracking** — Accurately announces multi-cell range selections and deselections.
- **PowerPoint: Accessible Color Pickers** — Enables NVDA to accurately read RGB and Hex values inside the Custom Color dialog.
- **PowerPoint: Standard Color Grid Support** — Intercepts arrow key navigation to read hidden Hex codes from the inaccessible color hexagon grid.
