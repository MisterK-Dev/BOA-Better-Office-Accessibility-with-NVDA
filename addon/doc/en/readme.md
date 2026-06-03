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

### Version 1.2.0-dev26 — 2026-06-03
**Development build.**

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
