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
  * Move Left: `NVDA + Shift + Left Arrow` (or PageUp)
  * Move Right: `NVDA + Shift + Right Arrow` (or PageDown)
  * Move to Start: `NVDA + Shift + Home`
  * Move to End: `NVDA + Shift + End`

### Accessible Sheet Renaming
Bypasses the notoriously inaccessible Excel "Rename Sheet" edit field.
* **How it works:** When you trigger the native Excel sheet rename command (e.g., via the ribbon or context menu), BOA intercepts it and opens a reliable, 100% accessible standard text dialog. Type your new name, press Enter, and BOA handles the renaming safely in the background.

### Smart Selection Tracking
Accurately announces multi-cell range selections and deselections.
* **How it works:** Fixes the silence when selecting ranges via the `F5` Go To dialog or when unexpectedly deselecting a range.

## PowerPoint Features

### Accessible Color Pickers
Enables NVDA to accurately read RGB and Hex values inside the Custom Color dialog.
* **How it works:** Seamlessly binds to the Red, Green, Blue, and Hex text boxes in the PowerPoint color picker, properly labeling them for NVDA object navigation and focus.

### Standard Color Grid Support
Read the exact color you are selecting in the inaccessible "Standard" color hexagon grid.
* **How it works:** When you navigate the visual color hexagon using the arrow keys, BOA intercepts the movement and reads the hidden Hex code from PowerPoint's internal memory in real-time.

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
