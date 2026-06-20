# -*- coding: UTF-8 -*-
# Copyright (C) 2026 KIRAN G T {MisterK} and Antigravity 2
# This file is covered by the GNU General Public License (GPL), version 2.
# See the file COPYING.txt for more details.

import addonHandler
addonHandler.initTranslation()

import wx
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
            "val": val,
            "excel": excel
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
        if slot_str not in cls._slots:
            ui.message(_("Slot {slot_str} is empty.").format(slot_str=slot_str))
            return
            
        info = cls._slots[slot_str]
        
        excel = info.get("excel")
        active_excel, current_wb, current_sheet, _unused_addr, _unused_val = cls._get_active_cell_info(obj)
        
        if not excel:
            excel = active_excel
            
        if not excel:
            ui.message(_("Error: Excel not accessible."))
            return

        try:
            target_wb = None
            try:
                target_wb = excel.Workbooks(info["wb"])
            except Exception:
                if active_excel:
                    try:
                        target_wb = active_excel.Workbooks(info["wb"])
                    except Exception:
                        pass
                        
            if not target_wb:
                ui.message(_("Slot {slot_str} lost. Workbook '{wb}' is closed or inaccessible.").format(
                    slot_str=slot_str, wb=info['wb']))
                return
            
            sheet_names = [s.Name for s in target_wb.Sheets]
            if info["sheet"] not in sheet_names:
                ui.message(_("Slot {slot_str} lost. Sheet '{sheet}' was renamed or deleted.").format(
                    slot_str=slot_str, sheet=info['sheet']))
                return
                
            target_sheet = target_wb.Sheets(info["sheet"])
            target_cell = target_sheet.Range(info["cell"])
            try:
                val = str(target_cell.Text)
            except Exception:
                val = ""
            if not val or val.startswith("###"):
                raw_val = target_cell.Value
                val = str(raw_val) if raw_val is not None else ""
            
            info["val"] = val
            monitor_key = f"{info['wb']}|{info['sheet']}|{info['cell']}"
            if monitor_key in cls._monitors:
                cls._monitors[monitor_key] = val
                
            if current_wb == info['wb'] and current_sheet == info['sheet']:
                location_str = ""
            elif current_wb == info['wb']:
                location_str = _(" in {sheet}").format(sheet=info['sheet'])
            else:
                location_str = _(" in {sheet} of {wb}").format(sheet=info['sheet'], wb=info['wb'])
                
            ui.message("{val} - {cell}{location_str}".format(
                val=val, cell=info['cell'], location_str=location_str))
        except Exception as e:
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

            for key, last_val in cls._monitors.items():
                wb_name, sheet_name, cell_addr = key.split("|")
                
                try:
                    # Robustly find workbook object without proxy wrapping
                    target_wb = None
                    for w in excel.Workbooks:
                        w_name = w.Name
                        if w_name == wb_name or w_name.split('.')[0] == wb_name.split('.')[0]:
                            target_wb = w
                            break
                    
                    if not target_wb:
                        continue # Skip this tick, might be in a background Excel instance
                        
                    # Safely verify sheet exists
                    sheet_names = [s.Name for s in target_wb.Sheets]
                    if sheet_name not in sheet_names:
                        continue # Skip this tick
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
                            location_str = _(" in {sheet}").format(sheet=sheet_name)
                        else:
                            location_str = _(" in {sheet} of {wb}").format(sheet=sheet_name, wb=wb_name)
                            
                        ui.message(_("{cell_addr} updated: {current_val}{location_str}").format(
                            cell_addr=cell_addr, current_val=current_val, location_str=location_str))
                except Exception:
                    pass

        except Exception:
            pass

    _last_working_cell = None

    @classmethod
    def _jump_to_address(cls, excel, wb_name, sheet_name, cell_addr):
        try:
            # Cache current location before jumping
            try:
                if excel.ActiveWorkbook and excel.ActiveSheet and excel.ActiveCell:
                    cls._last_working_cell = {
                        "wb": excel.ActiveWorkbook.Name,
                        "sheet": excel.ActiveSheet.Name,
                        "cell": excel.ActiveCell.Address()
                    }
            except Exception:
                pass
                
            wb = None
            # Robustly resolve workbook, ignoring extension mismatch issues
            for w in excel.Workbooks:
                w_name = w.Name
                if w_name == wb_name or w_name.split('.')[0] == wb_name.split('.')[0]:
                    wb = w
                    break
                    
            if not wb:
                ui.message(_("Cannot jump. The workbook '{wb}' is closed.").format(wb=wb_name))
                return
                
            try:
                sheet = wb.Sheets(sheet_name)
            except Exception:
                ui.message(_("Cannot jump. The sheet '{sheet}' was renamed or deleted.").format(sheet=sheet_name))
                return
                
            wb.Activate()
            sheet.Activate()
            sheet.Range(cell_addr).Select()
        except Exception:
            ui.message(_("Cannot jump. The cell address is invalid or Excel is busy."))

    @classmethod
    def jump_to_slot(cls, slot_num, obj):
        if slot_num not in cls._slots:
            ui.message(_("No cell assigned to slot {slot_num}").format(slot_num=slot_num))
            return
            
        info = cls._slots[slot_num]
        excel, _unused_wb, _unused_sheet, _unused_addr, _unused_val = cls._get_active_cell_info(obj)
        if not excel:
            if cls._active_excel:
                excel = cls._active_excel
            else:
                ui.message(_("Error: Excel not accessible."))
                return
                
        cls._jump_to_address(excel, info["wb"], info["sheet"], info["cell"])

    @classmethod
    def jump_back(cls, obj):
        if not cls._last_working_cell:
            ui.message(_("No previous cell to jump back to."))
            return
            
        info = cls._last_working_cell
        excel, _unused_wb, _unused_sheet, _unused_addr, _unused_val = cls._get_active_cell_info(obj)
        if not excel:
            if cls._active_excel:
                excel = cls._active_excel
            else:
                ui.message(_("Error: Excel not accessible."))
                return
                
        cls._jump_to_address(excel, info["wb"], info["sheet"], info["cell"])
        # Clear it so we don't jump back and forth infinitely
        cls._last_working_cell = None

class ActiveMonitorsDialog(wx.Dialog):
    def __init__(self, parent, slots, monitors, excel_app):
        super().__init__(parent, title=_("Active Cell Monitors"), style=wx.DEFAULT_DIALOG_STYLE | wx.RESIZE_BORDER)
        self.slots = slots
        self.monitors = monitors
        self.excel = excel_app
        self.mapping = [] 
        
        mainSizer = wx.BoxSizer(wx.VERTICAL)
        helpLabel = wx.StaticText(self, label=_("Select a cell to jump to it. Press Delete to remove it from monitors."))
        mainSizer.Add(helpLabel, 0, wx.ALL, 5)
        
        self.listBox = wx.ListBox(self, size=(500, 300))
        self.populate_list()
        self.listBox.Bind(wx.EVT_LISTBOX_DCLICK, self.onJump)
        self.listBox.Bind(wx.EVT_CHAR_HOOK, self.onCharHook)
        mainSizer.Add(self.listBox, 1, wx.EXPAND | wx.ALL, 5)
        
        btnSizer = wx.StdDialogButtonSizer()
        jumpBtn = wx.Button(self, wx.ID_OK, label=_("&Jump"))
        jumpBtn.Bind(wx.EVT_BUTTON, self.onJump)
        btnSizer.AddButton(jumpBtn)
        closeBtn = wx.Button(self, wx.ID_CANCEL, label=_("&Close"))
        closeBtn.Bind(wx.EVT_BUTTON, self.onClose)
        btnSizer.AddButton(closeBtn)
        btnSizer.Realize()
        
        mainSizer.Add(btnSizer, 0, wx.EXPAND | wx.ALL, 5)
        self.SetSizer(mainSizer)
        self.CenterOnParent()
        
        if self.listBox.GetCount() > 0:
            self.listBox.SetSelection(0)
            
    def populate_list(self):
        self.listBox.Clear()
        self.mapping = []
        for slot_num in sorted(self.slots.keys()):
            info = self.slots[slot_num]
            val = info["val"]
            display_text = f"Slot {slot_num}: {info['sheet']}!{info['cell']} ({val})"
            self.listBox.Append(display_text)
            self.mapping.append((True, slot_num))
            
        for key, val in self.monitors.items():
            wb, sheet, cell = key.split("|")
            is_slotted = False
            for s_info in self.slots.values():
                if s_info["wb"] == wb and s_info["sheet"] == sheet and s_info["cell"] == cell:
                    is_slotted = True
                    break
            if not is_slotted:
                display_text = f"Monitor: {sheet}!{cell} ({val})"
                self.listBox.Append(display_text)
                self.mapping.append((False, key))

    def onCharHook(self, evt):
        if evt.GetKeyCode() == wx.WXK_DELETE:
            idx = self.listBox.GetSelection()
            if idx != wx.NOT_FOUND:
                is_slot, key = self.mapping[idx]
                if is_slot:
                    info = self.slots[key]
                    monitor_key = f"{info['wb']}|{info['sheet']}|{info['cell']}"
                    del self.slots[key]
                    if monitor_key in self.monitors:
                        del self.monitors[monitor_key]
                    import ui
                    ui.message(_("Slot {slot} deleted").format(slot=key))
                else:
                    del self.monitors[key]
                    import ui
                    ui.message(_("Monitor deleted"))
                self.populate_list()
                if self.listBox.GetCount() > 0:
                    self.listBox.SetSelection(min(idx, self.listBox.GetCount() - 1))
        elif evt.GetKeyCode() == wx.WXK_ESCAPE:
            self.onClose(evt)
        elif evt.GetKeyCode() == wx.WXK_RETURN:
            self.onJump(evt)
        else:
            evt.Skip()

    def onJump(self, evt):
        idx = self.listBox.GetSelection()
        if idx != wx.NOT_FOUND:
            is_slot, key = self.mapping[idx]
            if is_slot:
                info = self.slots[key]
                CellMonitorManager._jump_to_address(self.excel, info["wb"], info["sheet"], info["cell"])
            else:
                wb, sheet, cell = key.split("|")
                CellMonitorManager._jump_to_address(self.excel, wb, sheet, cell)
            self.EndModal(wx.ID_OK)
            
    def onClose(self, evt):
        self.EndModal(wx.ID_CANCEL)

CellMonitorManager.open_monitor_dialog = classmethod(lambda cls, obj: _open_monitor_dialog(cls, obj))

def _open_monitor_dialog(cls, obj):
    import gui
    excel, _unused_wb, _unused_sheet, _unused_addr, _unused_val = cls._get_active_cell_info(obj)
    if not excel:
        if cls._active_excel:
            excel = cls._active_excel
        else:
            import ui
            ui.message(_("Error: Excel not accessible."))
            return

    def _show():
        try:
            gui.mainFrame.prePopup()
            dlg = ActiveMonitorsDialog(gui.mainFrame, cls._slots, cls._monitors, excel)
            dlg.ShowModal()
        except Exception as e:
            import ui
            ui.message(f"Dialog failed: {str(e)}")
            from logHandler import log
            log.error(f"BOA Monitor Dialog Error: {e}", exc_info=True)
        finally:
            try:
                dlg.Destroy()
            except Exception:
                pass
            gui.mainFrame.postPopup()
            
    wx.CallAfter(_show)
