# -*- coding: UTF-8 -*-
# Copyright (C) 2026 KIRAN G T {MisterK} and Antigravity 2
# This file is covered by the GNU General Public License (GPL), version 2.
# See the file COPYING.txt for more details.

import addonHandler
addonHandler.initTranslation()

import ui
import queueHandler
from logHandler import log

class CellMonitorManager:
    """
    Manages the 9 slots and continuous monitoring for Excel cells.
    
    Architectural Intent & Considerations:
    Windows DCOM security natively prevents standard COM event sinks for Excel cell changes 
    across out-of-process boundaries. Because we cannot "listen" for an event when a background 
    cell changes, we MUST use a non-blocking polling architecture (`core.callLater`). 
    This class manages both manual memory slots (1-9) and the active continuous background polling.
    """
    _slots = {}  # Format: { "1": {"wb": "Book1", "sheet": "Sheet1", "cell": "$A$1", "val": "100"} }
    _monitors = {} # Format: { "Book1|Sheet1|$A$1": "100" }
    
    _monitoring_active = False

    @classmethod
    def _start_timer(cls, excelApp):
        """
        Initiates the background polling loop for cell monitoring.
        
        Architectural Intent & Considerations:
        We use NVDA's native `core.callLater(150, ...)` instead of `wx.Timer` or standard `time.sleep`. 
        `time.sleep` would catastrophically block NVDA's single thread, freezing the entire screen reader. 
        `core.callLater` safely schedules the execution on NVDA's main thread loop every 150ms.
        """
        cls._active_excel = excelApp
        if not cls._monitoring_active:
            cls._monitoring_active = True
            import core
            # 150ms feels instantaneous. We use NVDA's native core.callLater instead of wx.Timer for guaranteed execution.
            core.callLater(150, cls._check_all_monitors)
            log.info("BOA: Started Cell Monitor Loop.")

    @classmethod
    def _stop_timer(cls):
        """
        Safely halts the continuous monitoring loop.
        
        Architectural Intent & Considerations:
        To prevent memory leaks and unnecessary CPU polling when no cells are being monitored, 
        we toggle the `_monitoring_active` flag. The loop function respects this flag and gracefully terminates.
        """
        cls._monitoring_active = False
        cls._active_excel = None
        log.info("BOA: Stopped Cell Monitor Loop.")

    @classmethod
    def _get_active_cell_info(cls, obj):
        """
        Retrieves the exact workbook, sheet, address, and value of the currently focused cell.
        
        Architectural Intent & Considerations:
        Because NVDA objects don't always expose complete structural metadata natively, we must 
        query Excel's COM model directly to guarantee we have the absolute $A$1 address and parent 
        workbook name required for strict slot tracking.
        """
        try:
            import comtypes.client
            import ctypes
            import comtypes.automation
            excel = None
            
            # Try getting active object first
            try:
                excel = comtypes.client.GetActiveObject("Excel.Application")
            except Exception:
                pass
                
            # Fallback Consideration: If GetActiveObject fails (common if Excel is in edit mode or blocked by security boundaries),
            # we must manually dig for the EXCEL7 window class handle to force a back-door connection.
            if not excel:
                hwnd7 = obj.windowHandle if getattr(obj, "windowClassName", "") == "EXCEL7" else None
                if hwnd7:
                    # Dynamically load the oleacc library.
                    oleacc = ctypes.windll.oleacc if hasattr(ctypes.windll, 'oleacc') else ctypes.windll.user32.oleacc
                    ptr = ctypes.POINTER(comtypes.automation.IDispatch)()
                    # OBJID_NATIVEOM (-16) retrieves the native COM object underneath the window.
                    res = oleacc.AccessibleObjectFromWindow(hwnd7, -16, ctypes.byref(comtypes.automation.IDispatch._iid_), ctypes.byref(ptr))
                    if res == 0 and ptr:
                        # Safely cast the raw COM pointer back into a usable Python Excel Application object.
                        excel = comtypes.client.dynamic.Dispatch(ptr).Application

            if excel:
                cell = excel.ActiveCell
                sheet = excel.ActiveSheet.Name
                wb = excel.ActiveWorkbook.Name
                address = cell.Address(False, False) # relative A1 without verbose $ symbols
                val = str(cell.Text) if cell.Text is not None else ""
                return excel, wb, sheet, address, val
        except Exception as e:
            log.debugWarning(f"BOA CellMonitor: Failed to get cell info: {e}")
        return None, None, None, None, None

    @classmethod
    def assign_slot(cls, slot_str, obj):
        """
        Assigns the active cell to a memory slot (1-9).
        
        Architectural Intent & Considerations:
        Users need a way to quickly check specific cells without physically navigating to them.
        By assigning a cell to a dictionary slot, we cache its exact coordinate path. We also automatically 
        enable continuous monitoring for this cell so that if its value changes while the user is elsewhere, 
        they are immediately notified.
        """
        excel, wb, sheet, address, val = cls._get_active_cell_info(obj)
        if not excel:
            ui.message(_("Error: Could not read Excel cell."))
            return

        is_replace = slot_str in cls._slots
        old_info = cls._slots.get(slot_str)

        cls._slots[slot_str] = {
            "wb": wb,
            "sheet": sheet,
            "cell": address,
            "val": val
        }

        # Auto-enable continuous monitoring for slotted cells
        monitor_key = f"{wb}|{sheet}|{address}"
        cls._monitors[monitor_key] = val
        cls._start_timer(excel)

        # Remove old cell from monitors if it was only monitored via this slot
        if is_replace:
            old_monitor_key = f"{old_info['wb']}|{old_info['sheet']}|{old_info['cell']}"
            if old_monitor_key != monitor_key:
                # Keep it in monitors if it's assigned to another slot or manually monitored,
                # but to keep things simple for now, we leave it in the monitor list unless cleared.
                pass
            ui.message(_("{old_cell} has been replaced by {address} for slot {slot_str}").format(
                old_cell=old_info['cell'], address=address, slot_str=slot_str))
        else:
            ui.message(_("{address} set to slot {slot_str}").format(address=address, slot_str=slot_str))

    @classmethod
    def read_slot(cls, slot_str, obj):
        """
        Reads the current value of the mapped slot via COM.
        
        Architectural Intent & Considerations:
        When the user presses the slot key, we must query the live COM model rather than returning 
        a cached value. This ensures the reading is 100% accurate even if continuous polling was disabled.
        """
        if slot_str not in cls._slots:
            ui.message(_("Slot {slot_str} is empty.").format(slot_str=slot_str))
            return

        info = cls._slots[slot_str]
        excel, current_wb, current_sheet, _, _ = cls._get_active_cell_info(obj)
        if not excel:
            # Try to grab excel instance if we lost it
            if cls._active_excel:
                excel = cls._active_excel
            else:
                ui.message(_("Error: Excel not accessible."))
                return

        try:
            # First verify the workbook still exists (hasn't been renamed or closed)
            wb_names = [wb.Name for wb in excel.Workbooks]
            if info["wb"] not in wb_names:
                ui.message(_("Slot {slot_str} lost. Workbook '{wb}' was renamed or closed.").format(
                    slot_str=slot_str, wb=info['wb']))
                del cls._slots[slot_str]
                monitor_key = f"{info['wb']}|{info['sheet']}|{info['cell']}"
                if monitor_key in cls._monitors:
                    del cls._monitors[monitor_key]
                return
                
            target_wb = excel.Workbooks(info["wb"])
            
            # Verify the sheet still exists (hasn't been renamed or deleted)
            sheet_names = [s.Name for s in target_wb.Sheets]
            if info["sheet"] not in sheet_names:
                ui.message(_("Slot {slot_str} lost. Sheet '{sheet}' was renamed or deleted.").format(
                    slot_str=slot_str, sheet=info['sheet']))
                del cls._slots[slot_str]
                monitor_key = f"{info['wb']}|{info['sheet']}|{info['cell']}"
                if monitor_key in cls._monitors:
                    del cls._monitors[monitor_key]
                return
                
            target_sheet = target_wb.Sheets(info["sheet"])
            target_cell = target_sheet.Range(info["cell"])
            val = str(target_cell.Text) if target_cell.Text is not None else ""
            
            # Update cache
            info["val"] = val
            monitor_key = f"{info['wb']}|{info['sheet']}|{info['cell']}"
            if monitor_key in cls._monitors:
                cls._monitors[monitor_key] = val
                
            if current_wb == info['wb'] and current_sheet == info['sheet']:
                location_str = ""
            elif current_wb == info['wb']:
                location_str = f" in {info['sheet']}"
            else:
                location_str = f" in {info['sheet']} of {info['wb']}"
                
            ui.message(_("{val} - {cell}{location_str}").format(
                val=val, cell=info['cell'], location_str=location_str))
        except Exception:
            ui.message(_("Cannot read slot {slot_str}. Excel may be busy.").format(slot_str=slot_str))

    @classmethod
    def toggle_monitor(cls, obj):
        """
        Toggles continuous background monitoring for the active cell.
        
        Architectural Intent & Considerations:
        Allows users to selectively monitor a single cell (like a total sum) without assigning it 
        to a specific 1-9 slot. If all monitors are cleared, it proactively shuts down the background 
        polling timer to conserve NVDA system resources.
        """
        excel, wb, sheet, address, val = cls._get_active_cell_info(obj)
        if not excel:
            ui.message(_("Error: Could not read Excel cell."))
            return

        monitor_key = f"{wb}|{sheet}|{address}"
        
        if monitor_key in cls._monitors:
            del cls._monitors[monitor_key]
            ui.message(_("Continuous monitor OFF for {address}").format(address=address))
            if not cls._monitors:
                cls._stop_timer()
        else:
            cls._monitors[monitor_key] = val
            cls._start_timer(excel)
            ui.message(_("Continuous monitor ON for {address}").format(address=address))

    @classmethod
    def clear_all(cls, obj):
        """
        Wipes all slots and active monitors from memory.
        
        Architectural Intent & Considerations:
        Provides a necessary reset switch for the user. Crucially, it must also call `_stop_timer()` 
        to ensure the background polling loop is fully terminated.
        """
        cls._slots.clear()
        cls._monitors.clear()
        cls._stop_timer()
        ui.message(_("All monitored and slotted cells cleared."))

    @classmethod
    def _check_all_monitors(cls):
        """
        The core background polling loop that checks all registered cells for value changes.
        
        Architectural Intent & Considerations:
        This function recursively calls itself via `core.callLater`. It iterates through the `_monitors` 
        dictionary, queries the live COM cell value, and compares it against the cached value. 
        It includes strict safety gates to prevent crashing if Excel is mid-calculation or a workbook is closed.
        """
        if not cls._monitoring_active:
            return

        # Reschedule first to guarantee continuous loop
        import core
        core.callLater(150, cls._check_all_monitors)

        if not cls._monitors:
            return

        try:
            import comtypes.client
            try:
                # Fetch fresh excel instance dynamically to avoid stale COM proxies
                excel = comtypes.client.GetActiveObject("Excel.Application")
            except Exception:
                excel = cls._active_excel
                
            if not excel:
                return

            # Safety Gate 1: Mid-Calculation Trap
            try:
                # xlDone = 0. If Excel is calculating (1) or pending (2), wait.
                if excel.CalculationState != 0:
                    return
            except Exception:
                pass

            # Safety Gate 2: Ghost Workbook Cleanup & Rename Detection
            open_wbs = None
            try:
                # Fetch all open workbook names. If this fails, Excel is busy (Edit Mode).
                open_wbs = [wb.Name for wb in excel.Workbooks]
            except Exception:
                pass

            to_remove = []

            for key, last_val in cls._monitors.items():
                wb_name, sheet_name, cell_addr = key.split("|")
                
                # Active detection of renamed/lost Workbooks and Sheets
                if open_wbs is not None:
                    if wb_name not in open_wbs:
                        to_remove.append(key)
                        continue
                        
                    # If workbook exists, safely verify sheet exists.
                    try:
                        target_wb = excel.Workbooks(wb_name)
                        sheet_names = [s.Name for s in target_wb.Sheets]
                        if sheet_name not in sheet_names:
                            to_remove.append(key)
                            continue
                    except Exception:
                        # If checking sheets fails due to Edit Mode, safely ignore for this tick.
                        pass

                try:
                    target_wb = excel.Workbooks(wb_name)
                    target_sheet = target_wb.Sheets(sheet_name)
                    target_cell = target_sheet.Range(cell_addr)
                    
                    # Safety Gate 3: Text vs Value display trap
                    current_val = str(target_cell.Text) if target_cell.Text is not None else ""
                    # If column is too narrow, Excel returns ######. Fallback to raw value.
                    if current_val.startswith("###"):
                        raw_val = target_cell.Value
                        current_val = str(raw_val) if raw_val is not None else ""
                    
                    if current_val != last_val:
                        cls._monitors[key] = current_val
                        for s_key, s_info in cls._slots.items():
                            if s_info["wb"] == wb_name and s_info["sheet"] == sheet_name and s_info["cell"] == cell_addr:
                                cls._slots[s_key]["val"] = current_val
                                
                        import ui
                        try:
                            active_wb = excel.ActiveWorkbook.Name if excel.ActiveWorkbook else None
                            active_sheet = excel.ActiveSheet.Name if excel.ActiveSheet else None
                        except Exception:
                            active_wb, active_sheet = None, None
                            
                        if active_wb == wb_name and active_sheet == sheet_name:
                            location_str = ""
                        elif active_wb == wb_name:
                            location_str = f" in {sheet_name}"
                        else:
                            location_str = f" in {sheet_name} of {wb_name}"
                            
                        ui.message(_("{cell_addr} updated: {current_val}{location_str}").format(
                            cell_addr=cell_addr, current_val=current_val, location_str=location_str))
                except Exception:
                    pass

            for key in to_remove:
                del cls._monitors[key]
                wb_closed, sheet_closed, cell_addr = key.split("|")
                
                # Check if it belongs to a slot and clear it
                slot_cleared = None
                for s_key, s_info in list(cls._slots.items()):
                    if s_info["wb"] == wb_closed and s_info["sheet"] == sheet_closed and s_info["cell"] == cell_addr:
                        slot_cleared = s_key
                        del cls._slots[s_key]
                
                import ui
                if slot_cleared:
                    ui.message(_("Monitor for Slot {slot_cleared} lost due to name change or closure.").format(
                        slot_cleared=slot_cleared))
                else:
                    ui.message(_("Monitor cleared: {sheet_closed} in {wb_closed} lost.").format(
                        sheet_closed=sheet_closed, wb_closed=wb_closed))

        except Exception:
            pass
