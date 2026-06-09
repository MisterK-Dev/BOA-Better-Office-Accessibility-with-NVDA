import controlTypes
from logHandler import log
import UIAHandler
import NVDAObjects.UIA
import NVDAObjects.IAccessible
import NVDAObjects.window.edit
import wx
import gui
import threading
import time
import winUser
import keyboardHandler
import core
from scriptHandler import script
import queueHandler

class BulkSheetOrganizer(object):
    """
    Handles launching and applying bulk sheet rearrangement operations.
    ARCHITECTURAL INTENT: Reordering multiple sheets in Excel natively requires dragging tabs 
    with a mouse or repeatedly opening the "Move or Copy" dialog. This class provides an
    accessible, programmatic way to queue multiple sheet moves and execute them via COM
    all at once.
    """
    @staticmethod
    def _show_bulk_dialog(sheet_names, hwnd):
        """
        Creates and displays the wx.Dialog. Parses the result and triggers COM execution.
        """
        import gui
        gui.mainFrame.prePopup()
        dlg = ExcelBulkSheetOrganizerDialog(gui.mainFrame, sheet_names)
        dlg.Raise()
        res = dlg.ShowModal()
        planned_moves = dict(dlg.planned_moves) if res == wx.ID_OK else None
        dlg.Destroy()
        gui.mainFrame.postPopup()
        
        if planned_moves:
            def _do_com_moves():
                """
                Executes the queued sheet moves via Excel COM automation.
                """
                import comtypes.client
                import comtypes.automation
                import ctypes
                import ui
                import logHandler
                
                # Fetch Excel via the raw HWND to bypass COM unavailability
                oleacc = ctypes.windll.user32.oleacc if hasattr(ctypes.windll.user32, 'oleacc') else ctypes.windll.oleacc
                ptr = ctypes.POINTER(comtypes.automation.IDispatch)()
                res = oleacc.AccessibleObjectFromWindow(hwnd, -16, ctypes.byref(comtypes.automation.IDispatch._iid_), ctypes.byref(ptr))
                if res == 0 and ptr:
                    try:
                        excel = comtypes.client.dynamic.Dispatch(ptr).Application
                        wb = excel.ActiveWorkbook
                        
                        # Calculate the final order of sheets before moving anything
                        unmoved = [s for s in sheet_names if s not in planned_moves]
                        moved = [(s, planned_moves[s]) for s in planned_moves]
                        moved.sort(key=lambda x: x[1])
                        
                        final_list = unmoved[:]
                        for s, pos in moved:
                            insert_idx = min(pos - 1, len(final_list))
                            final_list.insert(insert_idx, s)
                            
                        total_sheets = wb.Sheets.Count
                        if total_sheets > 1:
                            # Step 1: Secure the absolute last sheet first
                            last_sheet_name = final_list[-1]
                            sheet = wb.Sheets(last_sheet_name)
                            if sheet.Index < total_sheets:
                                if sheet.Index < total_sheets - 1:
                                    sheet.Move(wb.Sheets(total_sheets))
                                wb.Sheets(total_sheets).Move(sheet)
                                
                            # Step 2: Working backwards, move every other sheet directly BEFORE the one to its right.
                            for i in range(len(final_list) - 2, -1, -1):
                                sheet_name = final_list[i]
                                sheet_after_name = final_list[i + 1]
                                
                                # We must re-fetch the sheet objects by name on every iteration 
                                # to avoid stale COM references caused by the previous move operations.
                                sheet = wb.Sheets(sheet_name)
                                sheet_after = wb.Sheets(sheet_after_name)
                                
                                if sheet.Index != sheet_after.Index - 1:
                                    sheet.Move(sheet_after)
                                    
                        ui.message("Bulk arrangement complete")
                    except Exception as e:
                        ui.message(f"Error during bulk move: {e}")
                        logHandler.log.error(f"BOA bulk bg error: {e}")
                        
            def _apply_moves():
                """
                Ensures Excel is in the foreground before starting COM moves.
                """
                import winUser
                winUser.setForegroundWindow(hwnd)
                import core
                core.callLater(200, _do_com_moves)
                
            import core
            core.callLater(100, _apply_moves)

    @staticmethod
    def launch_dialog(obj):
        """
        Connects to Excel, grabs a list of all current sheet names, 
        and then opens the custom Bulk Sheet Organizer WX dialog.
        ARCHITECTURAL INTENT: We perform the heavy COM lifting here before launching the UI
        so the dialog appears instantly and is prepopulated with correct data.
        """
        import comtypes.client
        import comtypes.automation
        import ctypes
        import ui
        import wx
        
        try:
            # Safely hook into Excel grid via HWND traversal
            hwnd7 = None
            if getattr(obj, "windowClassName", "") == "EXCEL7":
                hwnd7 = obj.windowHandle
            else:
                hwnd = ctypes.windll.user32.FindWindowW("XLMAIN", None)
                if hwnd:
                    xldesk = ctypes.windll.user32.FindWindowExW(hwnd, 0, "XLDESK", None)
                    if xldesk:
                        hwnd7 = ctypes.windll.user32.FindWindowExW(xldesk, 0, "EXCEL7", None)
            
            if not hwnd7:
                ui.message("Could not find Excel grid.")
                return
                
            oleacc = ctypes.windll.user32.oleacc if hasattr(ctypes.windll.user32, 'oleacc') else ctypes.windll.oleacc
            ptr = ctypes.POINTER(comtypes.automation.IDispatch)()
            res = oleacc.AccessibleObjectFromWindow(hwnd7, -16, ctypes.byref(comtypes.automation.IDispatch._iid_), ctypes.byref(ptr))
            
            if res != 0 or not ptr:
                ui.message("Failed to hook Excel.")
                return
                
            win = comtypes.client.dynamic.Dispatch(ptr)
            excel = win.Application
            wb = excel.ActiveWorkbook
            if not wb:
                ui.message("No active workbook.")
                return
                
            # Extract the names of every sheet currently in the workbook to populate the dialog.
            total_sheets = wb.Sheets.Count
            sheet_names = [wb.Sheets(i).Name for i in range(1, total_sheets + 1)]
            
            # Use wx.CallAfter to safely push the dialog creation onto NVDA's main GUI thread.
            wx.CallAfter(BulkSheetOrganizer._show_bulk_dialog, sheet_names, hwnd7)
        except Exception as e:
            ui.message("Error opening organizer")
            import logHandler
            logHandler.log.error(f"ExcelGridMover bulk error: {e}")

class ExcelBulkSheetOrganizerDialog(wx.Dialog):
    """
    A custom wxPython Dialog that provides a fully accessible interface for bulk moving sheets.
    wxPython is the GUI framework used by NVDA.
    """
    def __init__(self, parent, sheet_names):
        super().__init__(parent, title="Bulk Sheet Organizer")
        self.sheet_names = sheet_names
        # Dictionary to track the user's requested moves before they press OK.
        # Format: {"Sheet1": 3, "Sheet2": 1}
        self.planned_moves = {} 
        
        # main_sizer is the master layout container. 
        # BoxSizers automatically stack elements vertically or horizontally, making the dialog accessible and scalable.
        main_sizer = wx.BoxSizer(wx.VERTICAL)
        
        # --- Combo 1: Sheet Name Selection ---
        row1 = wx.BoxSizer(wx.HORIZONTAL)
        row1.Add(wx.StaticText(self, label="Sheet Name:"), 0, wx.ALL|wx.ALIGN_CENTER_VERTICAL, 5)
        self.cb_sheet = wx.ComboBox(self, choices=self.sheet_names, style=wx.CB_READONLY)
        if self.sheet_names:
            self.cb_sheet.SetSelection(0)
        # Bind the combobox change event to our custom handler so we can update the Target Position box dynamically.
        self.cb_sheet.Bind(wx.EVT_COMBOBOX, self.on_sheet_change)
        row1.Add(self.cb_sheet, 1, wx.ALL, 5)
        
        # --- Combo 2: Target Position Selection ---
        row2 = wx.BoxSizer(wx.HORIZONTAL)
        row2.Add(wx.StaticText(self, label="Target Position:"), 0, wx.ALL|wx.ALIGN_CENTER_VERTICAL, 5)
        positions = [str(i) for i in range(1, len(self.sheet_names) + 1)]
        self.cb_pos = wx.ComboBox(self, choices=positions, style=wx.CB_READONLY)
        if positions:
            self.cb_pos.SetSelection(0)
        # When a position is selected, it immediately saves the move to `self.planned_moves`
        self.cb_pos.Bind(wx.EVT_COMBOBOX, self.on_pos_change)
        row2.Add(self.cb_pos, 1, wx.ALL, 5)
        
        main_sizer.Add(row1, 0, wx.EXPAND)
        main_sizer.Add(row2, 0, wx.EXPAND)
        
        # --- List of Scheduled Moves ---
        main_sizer.Add(wx.StaticText(self, label="Scheduled Moves (Press Del to remove):"), 0, wx.LEFT|wx.TOP, 5)
        # LC_REPORT style creates a standard data table which is highly accessible to NVDA.
        self.list_moves = wx.ListCtrl(self, size=(-1, 150), style=wx.LC_REPORT | wx.LC_SINGLE_SEL)
        self.list_moves.InsertColumn(0, "Sheet", width=150)
        self.list_moves.InsertColumn(1, "Target Position", width=100)
        # Bind the Delete key so users can easily remove mistakes.
        self.list_moves.Bind(wx.EVT_LIST_KEY_DOWN, self.on_list_key_down)
        main_sizer.Add(self.list_moves, 1, wx.ALL|wx.EXPAND, 5)
        
        # --- OK / Cancel Buttons ---
        btn_sizer = wx.BoxSizer(wx.HORIZONTAL)
        btn_ok = wx.Button(self, wx.ID_OK, label="&OK")
        btn_cancel = wx.Button(self, wx.ID_CANCEL, label="&Cancel")
        btn_sizer.Add(btn_ok, 0, wx.ALL, 5)
        btn_sizer.Add(btn_cancel, 0, wx.ALL, 5)
        main_sizer.Add(btn_sizer, 0, wx.ALL|wx.ALIGN_RIGHT, 5)
        
        self.SetSizerAndFit(main_sizer)
        
        # Force NVDA focus to the first combo box when the dialog opens so the user can start typing immediately.
        self.cb_sheet.SetFocus()

    def on_sheet_change(self, event):
        """
        Triggered when the user changes the Sheet Name combobox.
        It updates the Target Position combobox to either show the already-scheduled move, 
        or the current position of the sheet.
        """
        sheet = self.cb_sheet.GetValue()
        if sheet in self.planned_moves:
            self.cb_pos.SetSelection(self.planned_moves[sheet] - 1)
        else:
            if sheet in self.sheet_names:
                self.cb_pos.SetSelection(self.sheet_names.index(sheet))

    def on_pos_change(self, event):
        """
        Triggered when the user selects a new Target Position.
        It registers the move and updates the data table visually.
        """
        sheet = self.cb_sheet.GetValue()
        pos_str = self.cb_pos.GetValue()
        if sheet and pos_str:
            pos = int(pos_str)
            self.planned_moves[sheet] = pos
            self._refresh_list()
            import ui
            ui.message(f"Scheduled: {sheet} to position {pos}")

    def on_list_key_down(self, event):
        """
        Listens for the Delete key being pressed while focused on the Scheduled Moves list.
        """
        if event.GetKeyCode() == wx.WXK_DELETE:
            self.on_remove()
        event.Skip()

    def on_remove(self):
        sel = self.list_moves.GetFirstSelected()
        if sel != -1:
            sheet = self.list_moves.GetItemText(sel, 0)
            if sheet in self.planned_moves:
                del self.planned_moves[sheet]
            self._refresh_list()
            import ui
            ui.message("Move removed")
            # Update combo box to original position
            self.on_sheet_change(None)
            
    def _refresh_list(self):
        self.list_moves.DeleteAllItems()
        for sheet, pos in self.planned_moves.items():
            self.list_moves.Append([str(sheet), str(pos)])
