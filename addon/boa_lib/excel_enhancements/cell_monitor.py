import ui
import queueHandler
from logHandler import log

class CellMonitorManager:
    """
    Manages the 9 slots and continuous monitoring for Excel cells.
    """
    _slots = {}  # Format: { "1": {"wb": "Book1", "sheet": "Sheet1", "cell": "$A$1", "val": "100"} }
    _monitors = {} # Format: { "Book1|Sheet1|$A$1": "100" }
    
    _timer = None

    @classmethod
    def _start_timer(cls, excelApp):
        if cls._timer is None:
            import wx
            cls._timer = wx.Timer()
            cls._timer.Bind(wx.EVT_TIMER, cls._on_timer_tick)
        
        if not cls._timer.IsRunning():
            cls._active_excel = excelApp
            # 750ms is highly performant, indistinguishable from instant for speech, and costs 0% CPU.
            cls._timer.Start(750)
            log.info("BOA: Started Cell Monitor Timer.")

    @classmethod
    def _stop_timer(cls):
        if cls._timer and cls._timer.IsRunning():
            cls._timer.Stop()
            cls._active_excel = None
            log.info("BOA: Stopped Cell Monitor Timer.")

    @classmethod
    def _on_timer_tick(cls, evt):
        # Safely route the check through the queue handler
        import queueHandler
        queueHandler.queueFunction(cls._check_all_monitors)

    @classmethod
    def _get_active_cell_info(cls, obj):
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
                
            # Fallback to HWND digging
            if not excel:
                hwnd7 = obj.windowHandle if getattr(obj, "windowClassName", "") == "EXCEL7" else None
                if hwnd7:
                    oleacc = ctypes.windll.oleacc if hasattr(ctypes.windll, 'oleacc') else ctypes.windll.user32.oleacc
                    ptr = ctypes.POINTER(comtypes.automation.IDispatch)()
                    res = oleacc.AccessibleObjectFromWindow(hwnd7, -16, ctypes.byref(comtypes.automation.IDispatch._iid_), ctypes.byref(ptr))
                    if res == 0 and ptr:
                        excel = comtypes.client.dynamic.Dispatch(ptr).Application

            if excel:
                cell = excel.ActiveCell
                sheet = excel.ActiveSheet.Name
                wb = excel.ActiveWorkbook.Name
                address = cell.Address(True, True) # absolute $A$1
                val = str(cell.Text) if cell.Text is not None else ""
                return excel, wb, sheet, address, val
        except Exception as e:
            log.debugWarning(f"BOA CellMonitor: Failed to get cell info: {e}")
        return None, None, None, None, None

    @classmethod
    def assign_slot(cls, slot_str, obj):
        """ Assigns the active cell to slot 1-9 """
        excel, wb, sheet, address, val = cls._get_active_cell_info(obj)
        if not excel:
            ui.message("Error: Could not read Excel cell.")
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
            ui.message(f"{old_info['cell']} has been replaced by {address} for slot {slot_str}")
        else:
            ui.message(f"{address} set to slot {slot_str}")

    @classmethod
    def read_slot(cls, slot_str, obj):
        """ Reads the value of the mapped slot via COM """
        if slot_str not in cls._slots:
            ui.message(f"Slot {slot_str} is empty.")
            return

        info = cls._slots[slot_str]
        excel, _, _, _, _ = cls._get_active_cell_info(obj)
        if not excel:
            # Try to grab excel instance if we lost it
            if cls._active_excel:
                excel = cls._active_excel
            else:
                ui.message("Error: Excel not accessible.")
                return

        try:
            # We access the specific workbook, sheet, and range directly
            target_wb = excel.Workbooks(info["wb"])
            target_sheet = target_wb.Sheets(info["sheet"])
            target_cell = target_sheet.Range(info["cell"])
            val = str(target_cell.Text) if target_cell.Text is not None else ""
            
            # Update cache
            info["val"] = val
            monitor_key = f"{info['wb']}|{info['sheet']}|{info['cell']}"
            if monitor_key in cls._monitors:
                cls._monitors[monitor_key] = val
                
            ui.message(val)
        except Exception:
            ui.message(f"Cannot read slot {slot_str}. Excel may be busy or workbook closed.")

    @classmethod
    def toggle_monitor(cls, obj):
        """ Toggles continuous monitoring for the active cell """
        excel, wb, sheet, address, val = cls._get_active_cell_info(obj)
        if not excel:
            ui.message("Error: Could not read Excel cell.")
            return

        monitor_key = f"{wb}|{sheet}|{address}"
        
        if monitor_key in cls._monitors:
            del cls._monitors[monitor_key]
            ui.message(f"Continuous monitor OFF for {address}")
            if not cls._monitors:
                cls._stop_timer()
        else:
            cls._monitors[monitor_key] = val
            cls._start_timer(excel)
            ui.message(f"Continuous monitor ON for {address}")

    @classmethod
    def clear_all(cls, obj):
        """ Clears all slots and monitors """
        cls._slots.clear()
        cls._monitors.clear()
        cls._stop_timer()
        ui.message("All monitored and slotted cells cleared.")

    @classmethod
    def _check_all_monitors(cls):
        """ Called safely by NVDA queueHandler when Excel fires a change event """
        if not cls._monitors or not cls._active_excel:
            return

        excel = cls._active_excel
        to_remove = []

        try:
            for key, last_val in cls._monitors.items():
                wb_name, sheet_name, cell_addr = key.split("|")
                try:
                    # Attempt to read current value
                    target_wb = excel.Workbooks(wb_name)
                    target_sheet = target_wb.Sheets(sheet_name)
                    target_cell = target_sheet.Range(cell_addr)
                    current_val = str(target_cell.Text) if target_cell.Text is not None else ""
                    
                    if current_val != last_val:
                        # Value changed!
                        cls._monitors[key] = current_val
                        # Update slot cache if it exists
                        for s_key, s_info in cls._slots.items():
                            if s_info["wb"] == wb_name and s_info["sheet"] == sheet_name and s_info["cell"] == cell_addr:
                                cls._slots[s_key]["val"] = current_val
                                
                        ui.message(f"{cell_addr} updated: {current_val}")
                except Exception:
                    # Target might be closed or invalid, we ignore and potentially clean up later
                    pass
        except Exception as e:
            # Excel is busy (e.g. Cell Editing Mode), graceful failure
            pass
