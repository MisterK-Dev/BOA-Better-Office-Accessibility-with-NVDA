# BOA: Better Office Accessibility

BOA is a powerful suite of accessibility enhancements for Microsoft Office, designed to vastly improve the screen reader experience for NVDA users. It directly patches inaccessible UI components and introduces rapid navigation tools for Excel and PowerPoint.

---

## ⌨️ Hotkey Reference

| Feature | Key Combination | Context / Notes |
| :--- | :--- | :--- |
| **Enter Command Mode** | `NVDA+E` | Activates Command Prefix Mode (triggers a high-pitched beep) |
| **Analyze Sheet Layout** | `NVDA+E`, then `L` | Run within Excel before navigating data blocks |
| **Jump to Nearest Data Block** | `NVDA+E`, then `J` /  | Requires Layout Analysis first |
| **Open Bulk Sheet Organizer** | `NVDA+E`, then `X` | Opens the accessible sheet reordering dialog |
| **Move Active Sheet Left** | `NVDA+Shift+LeftArrow` | Shifts the active sheet one position up|
| **Move Active Sheet Right** | `NVDA+Shift+RightArrow` | Shifts the active worksheet one position down|
| **Move Sheet to Start/End** | `NVDA+Shift+Home` / `End` | Sends worksheet to the absolute absolute boundaries |
| **Detailed Conditional Formatting**| `NVDA+E`, then `F` | Announces complete formatting details of focused cell |
| **Map Cell to Memory Slot** | `NVDA+E`, then `Shift+1` to `Shift+9` | Assigns current cell to a background monitor slot |
| **Read Monitored Cell Slot** | `NVDA+E`, then `1` to `9` | Recalls and reads the value of the assigned slot |
| **Toggle Background Monitoring** | `NVDA+E`, then `M` | Manually toggles background calculation tracking |
| **Clear All Memory Slots** | `NVDA+E`, then `Backspace` | Purges all saved background cell monitors |
| **Cancel Command Mode** | `Escape` | Exits Command Prefix Mode |

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

### PowerPoint Enhancements

#### 1. Accessible Color Pickers
* Unlocks the Custom Color dialog in PowerPoint.
* Identifies and explicitly reads out the "Red", "Green", and "Blue" edit boxes correctly (by overriding `PowerPointRGBEdit`).
* Maps the previously invisible Hex input field so NVDA can read the full Hex color value cleanly.

#### 2. Standard Color Grid Support
* Navigating the PowerPoint "Standard" color hexagon grid normally reads as "Graphic" or silence.
* BOA tracks your arrow keys across the hexagon and silently fetches the hidden color value, announcing it to you in real-time (e.g., "Color #FF0000").

### Infrastructure & Technical Mechanisms

#### The Command Prefix Mode
To prevent keystroke conflicts with other NVDA plugins, BOA uses a **Command Prefix Mode**:
1. Press the activation hotkey to enter Command Mode. You will hear a high-pitched beep.
2. Press a secondary key to trigger a specific feature.
3. If you press an invalid key, you will hear an error beep.

#### Customization & Settings Panel
* BOA features are fully modular and can be enabled or disabled at any time. Go to `NVDA Menu -> Preferences -> Settings -> BOA Office Enhancements` to toggle individual features on or off.
* **Intelligent Accelerator Keys:** Every single setting features a mathematically unique `Alt+Key` accelerator shortcut within the panel. For example, press `Alt+E` to instantly jump to the Excel group, `Alt+P` for PowerPoint, and `Alt+W` for Word.
* Settings are saved securely to a standalone JSON file (`boa_settings.json`), ensuring your core NVDA configuration is never corrupted.
* If Microsoft Office officially fixes an accessibility bug in the future, you can safely disable BOA's specific override hook without losing the rest of the addon's functionality.

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