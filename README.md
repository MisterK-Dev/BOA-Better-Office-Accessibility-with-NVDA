# BOA: Blind Office Accessibility

BOA is a powerful suite of accessibility enhancements for Microsoft Office, designed to vastly improve the screen reader experience for NVDA users. It directly patches inaccessible UI components and introduces rapid navigation tools for Excel and PowerPoint.

## Excel Features

### 1. Bulk Sheet Organizer
Instantly reorder and arrange multiple sheets at once using a fully accessible dialog table.
* **Shortcut:** `NVDA + Alt + C`
* **How it works:** Opens a dialog where you can select a sheet and map it to a new position. Scheduled moves are listed in a data table (press `Del` to remove a mistake). Click `OK` and your workbook is rearranged instantly.

### 2. Quick Sheet Mover
Move the active sheet left, right, to the very beginning, or to the very end instantly using your keyboard.
* **Shortcuts:**
  * Move Left: `NVDA + Shift + Left Arrow` (or PageUp)
  * Move Right: `NVDA + Shift + Right Arrow` (or PageDown)
  * Move to Start: `NVDA + Shift + Home`
  * Move to End: `NVDA + Shift + End`

### 3. Accessible Sheet Renaming
Bypasses the notoriously inaccessible Excel "Rename Sheet" edit field.
* **How it works:** When you trigger the native Excel sheet rename command (e.g., via the ribbon or context menu), BOA intercepts it and opens a reliable, 100% accessible standard text dialog. Type your new name, press Enter, and BOA handles the renaming safely in the background.

### 4. Smart Selection Tracking
Accurately announces multi-cell range selections and deselections.
* **How it works:** Fixes the silence when selecting ranges via the `F5` Go To dialog or when unexpectedly deselecting a range.

## PowerPoint Features

### 1. Accessible Color Pickers
Enables NVDA to accurately read RGB and Hex values inside the Custom Color dialog.
* **How it works:** Seamlessly binds to the Red, Green, Blue, and Hex text boxes in the PowerPoint color picker, properly labeling them for NVDA object navigation and focus.

### 2. Standard Color Grid Support
Read the exact color you are selecting in the inaccessible "Standard" color hexagon grid.
* **How it works:** When you navigate the visual color hexagon using the arrow keys, BOA intercepts the movement and reads the hidden Hex code from PowerPoint's internal memory in real-time.

## Security
BOA is built with strict security boundaries. Background COM manipulations are executed within safely initialized sandboxes, and clipboard injections strictly verify window foreground process IDs to prevent leakage of data into other applications.

## Requirements
- NVDA 2026.1.0 or later.
- Microsoft Excel & PowerPoint.

## Installation
1. Download the latest `.nvda-addon` release file.
2. Open the file or use NVDA's Add-on Store -> Install from external file.
3. Restart NVDA.
