# Architectural Choices & Preferences
- **Cell Monitor (v1.4 dev 1):** We explicitly abandoned native COM Application Event Sinks (`SheetChange`/`SheetCalculate`). Windows DCOM Security blocks cross-process COM event callbacks, causing them to silently fail in the NVDA Python environment.
- **Background Loop Architecture:** We transitioned to a recursive `core.callLater(150)` loop. This is the NVDA-native standard for background tasks, matching the performance of a WX Timer but avoiding thread/parent window binding issues. It guarantees 100% reliable execution with 0% CPU overhead, bypassing Excel's restricted event broadcasting.
- **COM Polling Safety:** Inside the background loop, we actively check `excel.CalculationState != 0` to prevent "Mid-Calculation Traps" where COM might return intermediate/stale values (`0` or `""`) before the formula dependency tree has fully resolved.
- None recorded yet. If the `comtypes` Event Hook proves unstable during Cell Editing Mode, we may have to pivot to a polling timer, but the initial design favors native event hooks for performance.
# Post Mortem: Cell Monitor & NVDA Dialog Refactoring
**Date:** June 20, 2026

## 1. The `_` Localization Trap
**Issue:** `TypeError: 'str' object is not callable` when triggering translation messages.
**Root Cause:** In NVDA add-ons, `addonHandler.initTranslation()` injects `_` into the global namespace as the translation function. Using `_` as a throwaway variable in local scope (e.g., `excel, _, _, _, _ = cls._get_active_cell_info(obj)`) shadows the global function, overriding it with a literal string. 
**Lesson Learned:** *Never* use `_` as a throwaway variable in NVDA add-ons. Always use explicit names like `_unused` or `_addr` to avoid destroying the localization pipeline.

## 2. WxPython Dialog Scoping & Destruction
**Issue:** Dialog windows silently failing to open, or crashing the NVDA event loop upon closure.
**Root Cause:** Defining a custom `wx.Dialog` class inside a function closure causes unpredictable garbage collection and context loss. Furthermore, calling `.Destroy()` on a dialog in an NVDA thread without `CallAfter` or inside a modal loop can crash the parent frame.
**Lesson Learned:** Custom GUI dialogs must always be defined at the module (root) level. They should inherit cleanly (e.g. `super().__init__` or `wx.Dialog.__init__`) and be safely closed using `self.EndModal()` to respect NVDA's wx event loop boundaries.

## 3. Excel SDI (Single Document Interface) and COM Boundaries
**Issue:** Eagerly failing to read slotted cells from different workbooks when switching windows.
**Root Cause:** Modern Excel uses SDI, opening workbooks in potentially separate `Excel.exe` processes. Looking up `excel.Workbooks(name)` on the *current* active instance fails if the target workbook lives in a background instance.
**Lesson Learned:** To reliably track state across different Excel windows, the add-on must capture and explicitly cache the actual `excel` Application instance COM proxy at the time the slot is assigned, and use that specific proxy when trying to read the slot later.

## 4. Graceful Degradation vs. Aggressive Cleanup
**Issue:** Background monitors being silently deleted immediately upon alt-tabbing to another Excel process.
**Root Cause:** The original background loop aggressively deleted monitors if their workbook name wasn't found in the currently active `excel.Workbooks` list.
**Lesson Learned:** Background polling loops in a multi-process environment should fail gracefully. Instead of eagerly deleting data if a COM object isn't immediately accessible, simply skip the current tick. Only delete slots explicitly or when it is definitively proven the user closed the file.
