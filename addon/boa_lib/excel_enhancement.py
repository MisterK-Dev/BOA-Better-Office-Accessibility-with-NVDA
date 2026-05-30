import NVDAObjects.UIA
from logHandler import log
import NVDAObjects.IAccessible
import NVDAObjects.window.edit
import UIAHandler
import wx
import gui
import threading
import time
import winUser
import keyboardHandler
import core

_is_renaming_sheet = False
_last_selection_count = 1

# Tracks the last multi-cell address BOA announced aloud.
# This prevents BOA from re-announcing the same range that NVDA already spoke natively.
_last_announced_address = None

# Track states for structural Excel changes that lack native COM/UIA events
_last_freeze_panes_state = None
_last_visible_sheet_count = None

# Track last focused cell coordinates to detect jumps over hidden rows/cols
_last_focused_row = None
_last_focused_col = None
_last_focused_sheet = None
_last_focused_wb = None

def check_unselect(obj):
    """
    Called whenever Excel selection or focus changes.
    Its primary purpose is to detect when a user has a multi-cell selection (e.g., A1 through D1)
    and suddenly presses an arrow key, causing the selection to drop down to a single cell.
    Standard NVDA does not announce "unselected" in Excel, leaving the user unaware they lost their range.
    """
    global _last_selection_count
    try:
        className = getattr(obj, "windowClassName", "")
        # Filter purely to Excel's grid/cell classes to avoid unnecessary and expensive COM calls on other UI elements.
        if className not in ("EXCEL7", "NetUIHWND", "XLDESK"):
            return
            
        import comtypes.client
        import comtypes.automation
        import ctypes
        
        # Attempt to hook into the active Excel instance via standard COM.
        excel = None
        try:
            excel = comtypes.client.GetActiveObject("Excel.Application")
        except Exception:
            # Fallback: If GetActiveObject fails (common due to Windows security boundaries or multiple instances),
            # we manually crawl the window tree to find the raw EXCEL7 grid handle.
            hwnd7 = ctypes.windll.user32.FindWindowW("XLMAIN", None)
            if hwnd7:
                xldesk = ctypes.windll.user32.FindWindowExW(hwnd7, 0, "XLDESK", None)
                if xldesk:
                    hwnd7 = ctypes.windll.user32.FindWindowExW(xldesk, 0, "EXCEL7", None)
            if hwnd7:
                # Use AccessibleObjectFromWindow to force a back-door COM connection directly from the HWND.
                oleacc = ctypes.windll.user32.oleacc if hasattr(ctypes.windll.user32, 'oleacc') else ctypes.windll.oleacc
                ptr = ctypes.POINTER(comtypes.automation.IDispatch)()
                res = oleacc.AccessibleObjectFromWindow(hwnd7, -16, ctypes.byref(comtypes.automation.IDispatch._iid_), ctypes.byref(ptr))
                if res == 0 and ptr:
                    excel = comtypes.client.dynamic.Dispatch(ptr).Application

        if excel:
            sel = excel.Selection
            if getattr(sel, 'Cells', None):
                count = sel.Cells.Count
                # If we previously had >1 cells selected, and now we only have 1, the user dropped the selection.
                if count == 1 and _last_selection_count > 1:
                    import speech
                    speech.speakMessage("unselected")
                _last_selection_count = count
    except Exception:
        pass

class ExcelSheetRenameEdit(NVDAObjects.IAccessible.IAccessible):
    """
    Override for the Excel 'Rename sheet' edit box.
    Instead of trying to force Excel's broken edit box to speak,
    we intercept it and show a real, 100% accessible wx.Dialog.
    """
    def event_gainFocus(self):
        super().event_gainFocus()
        global _is_renaming_sheet
        if not _is_renaming_sheet:
            _is_renaming_sheet = True
            initial_name = self._fetch_sheet_name()
            wx.CallAfter(self._show_rename_dialog, initial_name, self.windowHandle)

    def _show_rename_dialog(self, initial_name, hwnd):
        """
        Creates the custom WX dialog to capture the new sheet name from the user.
        """
        gui.mainFrame.prePopup()
        dlg = wx.TextEntryDialog(gui.mainFrame, "Enter new sheet name:", "Rename Sheet", initial_name)
        dlg.Raise()
        res = dlg.ShowModal()
        new_name = dlg.GetValue() if res == wx.ID_OK else None
        dlg.Destroy()
        gui.mainFrame.postPopup()
        
        def bg_task():
            """
            Background thread that securely injects the typed sheet name back into Excel's native edit field.
            Using a background thread prevents blocking NVDA's main core loop.
            """
            global _is_renaming_sheet
            try:
                import time
                time.sleep(0.1)
                # Force Excel back to the foreground so it can receive keystrokes.
                winUser.setForegroundWindow(hwnd)
                time.sleep(0.2)
                
                clean_name = new_name.strip() if new_name else ""
                if not clean_name:
                    # If the user cancelled or typed nothing, send Escape to abort Excel's native rename mode.
                    core.callLater(10, lambda: keyboardHandler.KeyboardInputGesture.fromName("escape").send())
                    return
                    
                # SECURITY CHECK: Verify the foreground process ID matches Excel.
                # This prevents the addon from accidentally pasting the sheet name into a password field
                # if the user aggressively Alt-Tabbed to their browser during the delay.
                import ctypes
                fg_hwnd = winUser.getForegroundWindow()
                fg_pid = ctypes.c_ulong()
                ctypes.windll.user32.GetWindowThreadProcessId(fg_hwnd, ctypes.byref(fg_pid))
                
                target_pid = ctypes.c_ulong()
                ctypes.windll.user32.GetWindowThreadProcessId(hwnd, ctypes.byref(target_pid))
                
                if fg_pid.value != target_pid.value:
                    log.warning("BOA: Foreground window mismatch! Aborting keystroke injection to prevent pasting into wrong app.")
                    return
                
                # Safely back up the user's existing clipboard contents.
                import api
                old_clip = ""
                try:
                    old_clip = api.getClipData()
                except Exception:
                    pass
                
                try:
                    # Inject the string via the clipboard and simulate Ctrl+V.
                    # This is vastly more reliable and instantaneous than simulating individual keystrokes.
                    api.copyToClip(clean_name)
                    time.sleep(0.2)
                    
                    if winUser.getForegroundWindow() == fg_hwnd:
                        core.callLater(10, lambda: keyboardHandler.KeyboardInputGesture.fromName("control+v").send())
                        time.sleep(0.2)
                        core.callLater(10, lambda: keyboardHandler.KeyboardInputGesture.fromName("enter").send())
                        import speech
                        core.callLater(10, lambda: speech.speakMessage(f"Renaming to {clean_name}"))
                finally:
                    # Restore the user's original clipboard content once finished.
                    time.sleep(0.5)
                    if old_clip:
                        api.copyToClip(old_clip)
            finally:
                import time
                time.sleep(1.5)
                _is_renaming_sheet = False

        # Launch the background thread immediately after the dialog closes.
        threading.Thread(target=bg_task).start()

    def _get_name(self):
        return "Rename sheet"

    def _fetch_sheet_name(self):
        """
        Since the native edit field does not correctly expose its initial text to NVDA,
        we must manually hunt down the selected sheet tab in the Excel UI tree
        to find out what the sheet's current name is before renaming it.
        """
        try:
            # Step 1: Traverse up the object tree to find the main Excel window (XLMAIN).
            p = getattr(self, 'parent', None)
            xlmain_hwnd = None
            while p:
                if getattr(p, 'windowClassName', '') == 'XLMAIN':
                    xlmain_hwnd = p.windowHandle
                    break
                p = getattr(p, 'parent', None)
            
            if xlmain_hwnd:
                xlmain_uia = UIAHandler.handler.clientObject.ElementFromHandle(xlmain_hwnd)
                if xlmain_uia:
                    log.info("BOA: XLMAIN UIA found! Searching for TabItems...")
                    # Step 2: Search the UI Automation tree for TabItems (standard sheet tabs).
                    condition = UIAHandler.handler.clientObject.CreatePropertyCondition(
                        UIAHandler.UIA_ControlTypePropertyId, 
                        UIAHandler.UIA_TabItemControlTypeId
                    )
                    tabs = xlmain_uia.FindAll(UIAHandler.TreeScope_Descendants, condition)
                    if tabs:
                        log.info(f"BOA: Found {tabs.length} TabItems in XLMAIN")
                        for i in range(tabs.length):
                            tab = tabs.GetElement(i)
                            try:
                                # Property 30079 is UIA_SelectionItemIsSelectedPropertyId
                                is_sel = tab.GetCurrentPropertyValue(30079)
                                name = tab.CurrentName
                                log.info(f"BOA: TabItem {i}: name='{name}', selected={is_sel}")
                                if is_sel:
                                    return name
                            except Exception as e:
                                log.info(f"BOA: Error checking selection for tab {i}: {e}")
                    else:
                        # Step 3: Fallback for newer Excel 365 builds where sheet tabs are rendered as ListItems!
                        log.info("BOA: No TabItems found. Searching for ListItem controls instead...")
                        condition2 = UIAHandler.handler.clientObject.CreatePropertyCondition(
                            UIAHandler.UIA_ControlTypePropertyId, 
                            UIAHandler.UIA_ListItemControlTypeId
                        )
                        list_items = xlmain_uia.FindAll(UIAHandler.TreeScope_Descendants, condition2)
                        if list_items:
                            for i in range(list_items.length):
                                item = list_items.GetElement(i)
                                try:
                                    is_sel = item.GetCurrentPropertyValue(30079)
                                    name = item.CurrentName
                                    if is_sel and name and ("sheet" in name.lower() or name != ""):
                                        log.info(f"BOA: ListItem {i}: name='{name}', selected={is_sel}")
                                        return name
                                except Exception:
                                    pass
        except Exception as e:
            log.info(f"BOA: UIA Sheet Tab search failed: {e}")

        return ""

class SafeRichEdit(NVDAObjects.window.edit.Edit):
    """
    Override for Office RichEdit20W/RichEdit50W controls that crash 
    due to ITextDocument failing with OSError. 
    By bypassing ITextDocumentTextInfo and falling back to standard EditTextInfo, we avoid the crash.
    """
    TextInfo = NVDAObjects.window.edit.EditTextInfo
    
    def _get_ITextDocumentObject(self):
        return None

class ExcelGridMover(NVDAObjects.window.Window):
    def _check_boundary_bump(self, direction):
        """
        Detects if the user pressed an arrow key but focus did not move because they
        are stuck against an absolute hidden boundary (e.g. Row 1 is hidden).
        """
        try:
            import comtypes.client
            import ctypes
            import comtypes.automation
            
            hwnd7 = self.windowHandle if getattr(self, "windowClassName", "") == "EXCEL7" else None
            if not hwnd7:
                return
                
            oleacc = ctypes.windll.oleacc if hasattr(ctypes.windll, 'oleacc') else ctypes.windll.user32.oleacc
            ptr = ctypes.POINTER(comtypes.automation.IDispatch)()
            res = oleacc.AccessibleObjectFromWindow(hwnd7, -16, ctypes.byref(comtypes.automation.IDispatch._iid_), ctypes.byref(ptr))
            
            if res == 0 and ptr:
                excel = comtypes.client.dynamic.Dispatch(ptr).Application
                active_cell = excel.ActiveCell
                r = active_cell.Row
                c = active_cell.Column
                import ui
                
                if direction == "up" and r > 1:
                    # Check if ALL rows above are hidden (meaning we are stuck against the ceiling)
                    if excel.Range(excel.Cells(1, 1), excel.Cells(r - 1, 1)).EntireRow.Hidden in (True, -1):
                        if r - 1 == 1:
                            ui.message("Row 1 hidden")
                        else:
                            ui.message(f"Rows 1 through {r - 1} hidden")
                            
                elif direction == "left" and c > 1:
                    if excel.Range(excel.Cells(1, 1), excel.Cells(1, c - 1)).EntireColumn.Hidden in (True, -1):
                        def col_num_to_letter(n):
                            s = ""
                            while n > 0:
                                n, remainder = divmod(n - 1, 26)
                                s = chr(65 + remainder) + s
                            return s
                            
                        if c - 1 == 1:
                            ui.message("Column A hidden")
                        else:
                            ui.message(f"Columns A through {col_num_to_letter(c - 1)} hidden")
                            
                elif direction == "down" and r < 1048576:
                    if excel.Range(excel.Cells(r + 1, 1), excel.Cells(1048576, 1)).EntireRow.Hidden in (True, -1):
                        if r + 1 == 1048576:
                            ui.message("Row 1048576 hidden")
                        else:
                            ui.message(f"Rows {r + 1} through 1048576 hidden")
                            
                elif direction == "right" and c < 16384:
                    if excel.Range(excel.Cells(1, c + 1), excel.Cells(1, 16384)).EntireColumn.Hidden in (True, -1):
                        def col_num_to_letter(n):
                            s = ""
                            while n > 0:
                                n, remainder = divmod(n - 1, 26)
                                s = chr(65 + remainder) + s
                            return s
                            
                        if c + 1 == 16384:
                            ui.message("Column XFD hidden")
                        else:
                            ui.message(f"Columns {col_num_to_letter(c + 1)} through XFD hidden")
        except Exception:
            pass

    @script(
        description="Checks top boundary bump.",
        category="Better Office Accessibility"
    )
    def script_moveUp(self, gesture):
        self._check_boundary_bump("up")
        # Let NVDA's native Excel appModule process the keystroke natively
        raise NotImplementedError

    @script(
        description="Checks left boundary bump.",
        category="Better Office Accessibility"
    )
    def script_moveLeft(self, gesture):
        self._check_boundary_bump("left")
        # Let NVDA's native Excel appModule process the keystroke natively
        raise NotImplementedError

    @script(
        description="Checks bottom boundary bump.",
        category="Better Office Accessibility"
    )
    def script_moveDown(self, gesture):
        self._check_boundary_bump("down")
        # Let NVDA's native Excel appModule process the keystroke natively
        raise NotImplementedError

    @script(
        description="Checks right boundary bump.",
        category="Better Office Accessibility"
    )
    def script_moveRight(self, gesture):
        self._check_boundary_bump("right")
        # Let NVDA's native Excel appModule process the keystroke natively
        raise NotImplementedError

    def event_gainFocus(self):
        super(ExcelGridMover, self).event_gainFocus()
        import core
        core.callLater(50, self._check_multi_selection)

    def _check_multi_selection(self):
        """
        Manually checks the Excel COM model to announce the currently selected cell range.
        This fires after focus lands on the Excel grid (e.g., after closing the Go To dialog,
        pressing F5, or using the Name Box to jump to a range).

        DOUBLE-ANNOUNCEMENT PREVENTION:
        NVDA natively handles manual selection expansions (e.g., Shift+Arrow, Ctrl+Space).
        If the user is physically holding Shift or Control when this check fires, it means
        they are manually extending the selection via keyboard. We abort and let NVDA natively
        announce it. We only speak if no modifiers are held (which implies a programmatic jump
        like hitting Enter in the Go To box).
        """
        global _last_announced_address
        import winUser
        
        # GUARD 1: If Shift or Control is held down, the user is manually selecting.
        # NVDA will natively announce this. Do not double-announce.
        if (winUser.getKeyState(winUser.VK_SHIFT) & 32768) or (winUser.getKeyState(winUser.VK_CONTROL) & 32768):
            return

        import comtypes.client
        import comtypes.automation
        import ctypes
        import speech
        try:
            hwnd7 = self.windowHandle if getattr(self, "windowClassName", "") == "EXCEL7" else None
            if not hwnd7:
                return

            oleacc = ctypes.windll.oleacc if hasattr(ctypes.windll, 'oleacc') else ctypes.windll.user32.oleacc
            OBJID_NATIVEOM = -16
            ptr = ctypes.POINTER(comtypes.automation.IDispatch)()
            res = oleacc.AccessibleObjectFromWindow(hwnd7, OBJID_NATIVEOM, ctypes.byref(comtypes.automation.IDispatch._iid_), ctypes.byref(ptr))

            if res == 0 and ptr:
                win = comtypes.client.dynamic.Dispatch(ptr)
                excel = win.Application
                sel = excel.Selection

                if getattr(sel, 'Cells', None):
                    # ALWAYS check structural changes and hidden row skips on focus change
                    self._check_structural_changes(excel)
                    self._check_hidden_skip(excel)
                    
                    # Verify it is a Range object (has Cells property) and has more than 1 cell selected.
                    if sel.Cells.Count > 1:
                        address = sel.Address(False, False)  # Returns string like "A1:D1"

                        # GUARD: Only announce if this is a NEW address BOA hasn't spoken yet.
                        # This is the core fix for the double-announcement problem.
                        # When the user selects a range via Shift+Arrow, NVDA fires its own UIA
                        # event_selectionChange and speaks natively. Then event_gainFocus fires here
                        # and would speak again. By comparing against our last announced address,
                        # we skip the duplicate and stay silent.
                        if address == _last_announced_address:
                            return

                        _last_announced_address = address
                        spoken_address = address.replace(":", " through ")
                        speech.speakMessage(f"{spoken_address} selected")
                    else:
                        # Selection collapsed to single cell — reset tracker so
                        # the next multi-cell jump is announced cleanly.
                        _last_announced_address = None
        except Exception:
            pass

    def _check_hidden_skip(self, excel):
        """
        Detects if the user's focus jumped over completely hidden rows or columns.
        Uses COM bulk checking and SpecialCells to avoid false positives.
        Announces mixed scenarios (e.g. Ctrl+DownArrow) accurately.
        """
        global _last_focused_row, _last_focused_col, _last_focused_sheet, _last_focused_wb
        import ui
        
        try:
            active_cell = excel.ActiveCell
            current_row = active_cell.Row
            current_col = active_cell.Column
            
            try:
                current_sheet = excel.ActiveSheet.Name
                current_wb = excel.ActiveWorkbook.Name
            except Exception:
                current_sheet = None
                current_wb = None

            # Reset tracker if sheet/workbook changed to avoid ghost skips
            if current_sheet != _last_focused_sheet or current_wb != _last_focused_wb:
                _last_focused_row = None
                _last_focused_col = None
                _last_focused_sheet = current_sheet
                _last_focused_wb = current_wb
            
            if _last_focused_row is not None and _last_focused_col is not None:
                # ---------------- ROW JUMPS ----------------
                if current_col == _last_focused_col and current_row != _last_focused_row:
                    min_r = min(_last_focused_row, current_row)
                    max_r = max(_last_focused_row, current_row)
                    
                    if max_r - min_r > 1:
                        gap_range = excel.Range(excel.Cells(min_r + 1, 1), excel.Cells(max_r - 1, 1)).EntireRow
                        hidden_val = gap_range.Hidden
                        
                        if hidden_val == True or hidden_val == -1:
                            if max_r - min_r == 2:
                                ui.message(f"Row {min_r + 1} hidden")
                            else:
                                ui.message(f"Rows {min_r + 1} through {max_r - 1} hidden")
                        elif hidden_val is None:
                            # COM returns None (Mixed). Validate it's actually mixed and not a COM glitch.
                            has_hidden = False
                            if max_r - min_r < 100:
                                for r in range(min_r + 1, max_r):
                                    if excel.Rows(r).Hidden in (True, -1):
                                        has_hidden = True
                                        break
                            else:
                                try:
                                    visible_cells = gap_range.SpecialCells(12)
                                    visible_count = sum(area.Rows.Count for area in visible_cells.Areas)
                                    if visible_count < (max_r - min_r - 1):
                                        has_hidden = True
                                except Exception:
                                    has_hidden = True
                                    
                            if has_hidden:
                                ui.message("Some hidden rows")

                    # Row Boundary Check (Navigating UP into Row 1)
                    if current_row < _last_focused_row and current_row > 1:
                        if excel.Rows(current_row - 1).Hidden in (True, -1):
                            top = current_row - 1
                            while top > 1 and excel.Rows(top - 1).Hidden in (True, -1):
                                top -= 1
                            if top == current_row - 1:
                                ui.message(f"Row {top} hidden")
                            else:
                                ui.message(f"Rows {top} through {current_row - 1} hidden")

                # ---------------- COLUMN JUMPS ----------------
                if current_row == _last_focused_row and current_col != _last_focused_col:
                    min_c = min(_last_focused_col, current_col)
                    max_c = max(_last_focused_col, current_col)
                    
                    def col_num_to_letter(n):
                        s = ""
                        while n > 0:
                            n, remainder = divmod(n - 1, 26)
                            s = chr(65 + remainder) + s
                        return s
                        
                    if max_c - min_c > 1:
                        gap_range = excel.Range(excel.Cells(1, min_c + 1), excel.Cells(1, max_c - 1)).EntireColumn
                        hidden_val = gap_range.Hidden
                        
                        start_letter = col_num_to_letter(min_c + 1)
                        if hidden_val == True or hidden_val == -1:
                            if max_c - min_c == 2:
                                ui.message(f"Column {start_letter} hidden")
                            else:
                                end_letter = col_num_to_letter(max_c - 1)
                                ui.message(f"Columns {start_letter} through {end_letter} hidden")
                        elif hidden_val is None:
                            has_hidden = False
                            if max_c - min_c < 100:
                                for c in range(min_c + 1, max_c):
                                    if excel.Columns(c).Hidden in (True, -1):
                                        has_hidden = True
                                        break
                            else:
                                try:
                                    visible_cells = gap_range.SpecialCells(12)
                                    visible_count = sum(area.Columns.Count for area in visible_cells.Areas)
                                    if visible_count < (max_c - min_c - 1):
                                        has_hidden = True
                                except Exception:
                                    has_hidden = True
                            
                            if has_hidden:
                                ui.message("Some hidden columns")

                    # Column Boundary Check (Navigating LEFT into Col A)
                    if current_col < _last_focused_col and current_col > 1:
                        if excel.Columns(current_col - 1).Hidden in (True, -1):
                            left = current_col - 1
                            while left > 1 and excel.Columns(left - 1).Hidden in (True, -1):
                                left -= 1
                            start_letter = col_num_to_letter(left)
                            if left == current_col - 1:
                                ui.message(f"Column {start_letter} hidden")
                            else:
                                end_letter = col_num_to_letter(current_col - 1)
                                ui.message(f"Columns {start_letter} through {end_letter} hidden")
                        
            _last_focused_row = current_row
            _last_focused_col = current_col
        except Exception:
            pass

    def _check_structural_changes(self, excel):
        """
        Monitors for changes in Excel's structural layout that do not natively fire events,
        such as toggling Freeze Panes or Hiding/Unhiding a worksheet.
        By caching the previous state, we can detect if a change occurred while the user
        was interacting with the Ribbon or right-click menus, and announce it upon returning
        focus to the grid. Uses ui.message for safe output routing.
        """
        global _last_freeze_panes_state, _last_visible_sheet_count, _last_focused_sheet, _last_focused_wb
        import ui
        
        try:
            # Check Freeze Panes
            # ActiveWindow.FreezePanes returns a boolean indicating if panes are frozen.
            current_freeze = excel.ActiveWindow.FreezePanes
            if _last_freeze_panes_state is not None and current_freeze != _last_freeze_panes_state:
                if current_freeze:
                    ui.message("Panes frozen")
                else:
                    ui.message("Panes unfrozen")
            _last_freeze_panes_state = current_freeze
            
            # Check sheet counts and hidden sheets
            try:
                current_visible = 0
                total_sheets = excel.ActiveWorkbook.Sheets.Count
                
                # Check for skipped hidden sheets if navigated via Ctrl+PageDown
                if _last_focused_sheet is not None and _last_focused_wb == excel.ActiveWorkbook.Name:
                    current_idx = excel.ActiveSheet.Index
                    try:
                        last_idx = excel.ActiveWorkbook.Sheets(_last_focused_sheet).Index
                        if abs(current_idx - last_idx) > 1:
                            min_s = min(last_idx, current_idx)
                            max_s = max(last_idx, current_idx)
                            for i in range(min_s + 1, max_s):
                                sheet = excel.ActiveWorkbook.Sheets(i)
                                if sheet.Visible != -1:  # -1 is xlSheetVisible
                                    ui.message(f"{sheet.Name} hidden")
                    except Exception:
                        pass

                for i in range(1, total_sheets + 1):
                    sheet = excel.ActiveWorkbook.Sheets(i)
                    # Visible property returns -1 for visible, 0 for hidden, 2 for very hidden
                    if sheet.Visible == -1:
                        current_visible += 1
                
                if _last_visible_sheet_count is not None and current_visible != _last_visible_sheet_count:
                    if current_visible < _last_visible_sheet_count:
                        ui.message("Sheet hidden")
                    else:
                        ui.message("Sheet unhidden")
                _last_visible_sheet_count = current_visible
            except Exception:
                pass
                
        except Exception:
            pass

    __gestures = {
        "kb:upArrow": "moveUp",
        "kb:leftArrow": "moveLeft",
        "kb:downArrow": "moveDown",
        "kb:rightArrow": "moveRight"
    }

    def _move_sheet(self, direction):
        import comtypes.client
        import comtypes.automation
        import ctypes
        import speech
        try:
            # Safely get the Excel object bypassing GetActiveObject to prevent MK_E_UNAVAILABLE errors
            hwnd7 = None
            if getattr(self, "windowClassName", "") == "EXCEL7":
                hwnd7 = self.windowHandle
            else:
                hwnd = ctypes.windll.user32.FindWindowW("XLMAIN", None)
                if hwnd:
                    xldesk = ctypes.windll.user32.FindWindowExW(hwnd, 0, "XLDESK", None)
                    if xldesk:
                        hwnd7 = ctypes.windll.user32.FindWindowExW(xldesk, 0, "EXCEL7", None)
            
            if not hwnd7:
                speech.speakMessage("Could not find Excel grid.")
                return
                
            oleacc = ctypes.windll.oleacc
            OBJID_NATIVEOM = -16
            ptr = ctypes.POINTER(comtypes.automation.IDispatch)()
            
            res = oleacc.AccessibleObjectFromWindow(
                hwnd7, OBJID_NATIVEOM, 
                ctypes.byref(comtypes.automation.IDispatch._iid_), 
                ctypes.byref(ptr)
            )
            
            if res != 0 or not ptr:
                speech.speakMessage("Failed to hook Excel.")
                return
                
            win = comtypes.client.dynamic.Dispatch(ptr)
            excel = win.Application
            wb = excel.ActiveWorkbook
            sheet = excel.ActiveSheet
            
            current_index = sheet.Index
            total_sheets = wb.Sheets.Count
            
            if total_sheets <= 1:
                speech.speakMessage("Only one sheet in workbook")
                return
                
            if direction == "left":
                if current_index == 1:
                    speech.speakMessage("Already at beginning")
                    return
                sheet.Move(wb.Sheets(current_index - 1))
            elif direction == "right":
                if current_index == total_sheets:
                    speech.speakMessage("Already at end")
                    return
                # To move right, we just move the right neighbor before us!
                wb.Sheets(current_index + 1).Move(sheet)
            elif direction == "start":
                if current_index == 1:
                    speech.speakMessage("Already at beginning")
                    return
                sheet.Move(wb.Sheets(1))
            elif direction == "end":
                if current_index == total_sheets:
                    speech.speakMessage("Already at end")
                    return
                # Two-step COM trick to move to the very end without the broken 'After' parameter:
                # 1. Move our sheet BEFORE the very last sheet (if we aren't already just before it)
                if current_index < total_sheets - 1:
                    sheet.Move(wb.Sheets(total_sheets))
                # 2. Move the last sheet BEFORE our sheet! This puts our sheet at the very end!
                wb.Sheets(total_sheets).Move(sheet)
                
            # Moving sheets can sometimes change the active sheet (especially the 2-step end trick).
            # Force our target sheet to be the active one.
            sheet.Activate()
            
            new_index = excel.ActiveSheet.Index
            sheet_name = excel.ActiveSheet.Name
            speech.speakMessage(f"Moved {sheet_name} to position {new_index} of {total_sheets}")
        except Exception as e:
            speech.speakMessage("Failed to move sheet")
            import logHandler
            logHandler.log.error(f"ExcelGridMover error: {e}")

    from scriptHandler import script

    @script(
        description="Moves the active Excel sheet to the left.",
        category="Better Office Accessibility"
    )
    def script_moveSheetLeft(self, gesture):
        self._move_sheet("left")

    @script(
        description="Moves the active Excel sheet to the right.",
        category="Better Office Accessibility"
    )
    def script_moveSheetRight(self, gesture):
        self._move_sheet("right")

    @script(
        description="Moves the active Excel sheet to the very beginning of the workbook.",
        category="Better Office Accessibility"
    )
    def script_moveSheetStart(self, gesture):
        self._move_sheet("start")

    @script(
        description="Moves the active Excel sheet to the very end of the workbook.",
        category="Better Office Accessibility"
    )
    def script_moveSheetEnd(self, gesture):
        self._move_sheet("end")
        
    def _show_bulk_dialog(self, sheet_names, hwnd):
        import gui
        gui.mainFrame.prePopup()
        dlg = ExcelBulkSheetOrganizerDialog(gui.mainFrame, sheet_names)
        dlg.Raise()
        res = dlg.ShowModal()
        planned_moves = dict(dlg.planned_moves) if res == wx.ID_OK else None
        dlg.Destroy()
        gui.mainFrame.postPopup()
        
        if planned_moves:
            def bg_task():
                """
                Background thread for executing the bulk sheet arrangements via COM.
                We execute this in the background to prevent the Excel COM calls from freezing NVDA.
                """
                import time
                import winUser
                import comtypes.client
                import comtypes.automation
                import ctypes
                import speech
                import logHandler
                
                # CRITICAL: Any background thread interacting with Office COM MUST call CoInitialize
                # before making API calls, otherwise comtypes will crash or disconnect randomly.
                ctypes.windll.ole32.CoInitialize(None)
                try:
                    time.sleep(0.1)
                    winUser.setForegroundWindow(hwnd)
                    time.sleep(0.2)
                    
                    oleacc = ctypes.windll.user32.oleacc if hasattr(ctypes.windll.user32, 'oleacc') else ctypes.windll.oleacc
                    ptr = ctypes.POINTER(comtypes.automation.IDispatch)()
                    res = oleacc.AccessibleObjectFromWindow(hwnd, -16, ctypes.byref(comtypes.automation.IDispatch._iid_), ctypes.byref(ptr))
                    if res == 0 and ptr:
                        excel = comtypes.client.dynamic.Dispatch(ptr).Application
                        wb = excel.ActiveWorkbook
                        
                        # Step 1: Calculate the final desired mathematical order of the sheets.
                        unmoved = [s for s in sheet_names if s not in planned_moves]
                        moved = [(s, planned_moves[s]) for s in planned_moves]
                        moved.sort(key=lambda x: x[1])
                        
                        final_list = unmoved[:]
                        for s, pos in moved:
                            insert_idx = min(pos - 1, len(final_list))
                            final_list.insert(insert_idx, s)
                            
                        # Step 2: Apply the final computed order to Excel.
                        # Because Excel's COM Move() command only supports 'Before' and 'After', 
                        # moving elements iteratively from left to right causes index shifting bugs.
                        # We reconstruct the array safely by moving from Right to Left!
                        total_sheets = wb.Sheets.Count
                        if total_sheets > 1:
                            # Secure the absolute last sheet first
                            last_sheet_name = final_list[-1]
                            sheet = wb.Sheets(last_sheet_name)
                            if sheet.Index < total_sheets:
                                if sheet.Index < total_sheets - 1:
                                    sheet.Move(wb.Sheets(total_sheets))
                                wb.Sheets(total_sheets).Move(sheet)
                                
                            # Working backwards, move every other sheet directly BEFORE the one to its right.
                            for i in range(len(final_list) - 2, -1, -1):
                                sheet_name = final_list[i]
                                sheet_after_name = final_list[i + 1]
                                
                                # We must re-fetch the sheet objects by name on every iteration 
                                # to avoid stale COM references caused by the previous move operations.
                                sheet = wb.Sheets(sheet_name)
                                sheet_after = wb.Sheets(sheet_after_name)
                                
                                if sheet.Index != sheet_after.Index - 1:
                                    sheet.Move(sheet_after)
                                    
                        speech.speakMessage("Bulk arrangement complete")
                except Exception as e:
                    import speech
                    speech.speakMessage(f"Error during bulk move: {e}")
                    logHandler.log.error(f"BOA bulk bg error: {e}")
                finally:
                    # CRITICAL: Always uninitialize the COM apartment when the thread finishes to prevent memory leaks.
                    ctypes.windll.ole32.CoUninitialize()
                    
            import threading
            threading.Thread(target=bg_task).start()

    @script(
        description="Opens the BOA Bulk Sheet Organizer dialog.",
        category="Better Office Accessibility"
    )
    def script_openBulkSheetOrganizer(self, gesture):
        """
        NVDA script triggered by the user (NVDA+Alt+C).
        It connects to Excel, grabs a list of all current sheet names, 
        and then opens the custom Bulk Sheet Organizer WX dialog.
        """
        import comtypes.client
        import comtypes.automation
        import ctypes
        import speech
        import wx
        
        try:
            hwnd7 = None
            if getattr(self, "windowClassName", "") == "EXCEL7":
                hwnd7 = self.windowHandle
            else:
                hwnd = ctypes.windll.user32.FindWindowW("XLMAIN", None)
                if hwnd:
                    xldesk = ctypes.windll.user32.FindWindowExW(hwnd, 0, "XLDESK", None)
                    if xldesk:
                        hwnd7 = ctypes.windll.user32.FindWindowExW(xldesk, 0, "EXCEL7", None)
            
            if not hwnd7:
                speech.speakMessage("Could not find Excel grid.")
                return
                
            oleacc = ctypes.windll.user32.oleacc if hasattr(ctypes.windll.user32, 'oleacc') else ctypes.windll.oleacc
            ptr = ctypes.POINTER(comtypes.automation.IDispatch)()
            res = oleacc.AccessibleObjectFromWindow(hwnd7, -16, ctypes.byref(comtypes.automation.IDispatch._iid_), ctypes.byref(ptr))
            
            if res != 0 or not ptr:
                speech.speakMessage("Failed to hook Excel.")
                return
                
            win = comtypes.client.dynamic.Dispatch(ptr)
            excel = win.Application
            wb = excel.ActiveWorkbook
            if not wb:
                speech.speakMessage("No active workbook.")
                return
                
            # Extract the names of every sheet currently in the workbook to populate the dialog.
            total_sheets = wb.Sheets.Count
            sheet_names = [wb.Sheets(i).Name for i in range(1, total_sheets + 1)]
            
            # Use wx.CallAfter to safely push the dialog creation onto NVDA's main GUI thread.
            wx.CallAfter(self._show_bulk_dialog, sheet_names, hwnd7)
        except Exception as e:
            speech.speakMessage("Error opening organizer")
            import logHandler
            logHandler.log.error(f"ExcelGridMover bulk error: {e}")

    @script(
        description="Hides the currently selected row.",
        category="Better Office Accessibility"
    )
    def script_hideRow(self, gesture):
        """
        Intercepts Ctrl+9 to hide the selected row and announces the change.
        Passes the keystroke to Excel natively first.
        """
        self._execute_and_verify_visibility_change(gesture, "row", True)

    @script(
        description="Unhides the currently selected row.",
        category="Better Office Accessibility"
    )
    def script_unhideRow(self, gesture):
        """
        Intercepts Ctrl+Shift+9 to unhide the selected row and announces the change.
        """
        self._execute_and_verify_visibility_change(gesture, "row", False)

    @script(
        description="Hides the currently selected column.",
        category="Better Office Accessibility"
    )
    def script_hideColumn(self, gesture):
        """
        Intercepts Ctrl+0 to hide the selected column and announces the change.
        """
        self._execute_and_verify_visibility_change(gesture, "column", True)

    @script(
        description="Unhides the currently selected column.",
        category="Better Office Accessibility"
    )
    def script_unhideColumn(self, gesture):
        """
        Intercepts Ctrl+Shift+0 to unhide the selected column and announces the change.
        """
        self._execute_and_verify_visibility_change(gesture, "column", False)

    @script(
        description="Unhides the currently selected column natively via COM (Fallback for when Windows blocks Ctrl+Shift+0).",
        category="Better Office Accessibility"
    )
    def script_unhideColumnFallback(self, gesture):
        """
        Forces the column to unhide using Excel COM.
        """
        self._execute_and_verify_visibility_change(gesture, "column", False, force_com=True)

    def _execute_and_verify_visibility_change(self, gesture, element_type, is_hiding, force_com=False):
        """
        Fetches the initial visibility state, passes the native keystroke to Excel (or uses COM if forced),
        waits briefly, and then checks if the COM model state actually changed.
        Uses core.callLater to safely delay without blocking NVDA's single-threaded core.
        """
        import comtypes.client
        import comtypes.automation
        import ctypes
        
        initial_state = None
        try:
            hwnd7 = self.windowHandle if getattr(self, "windowClassName", "") == "EXCEL7" else None
            if hwnd7:
                oleacc = ctypes.windll.oleacc if hasattr(ctypes.windll, 'oleacc') else ctypes.windll.user32.oleacc
                ptr = ctypes.POINTER(comtypes.automation.IDispatch)()
                res = oleacc.AccessibleObjectFromWindow(hwnd7, -16, ctypes.byref(comtypes.automation.IDispatch._iid_), ctypes.byref(ptr))
                if res == 0 and ptr:
                    excel = comtypes.client.dynamic.Dispatch(ptr).Application
                    if element_type == "row":
                        initial_state = excel.Selection.EntireRow.Hidden
                        if force_com:
                            excel.Selection.EntireRow.Hidden = is_hiding
                    elif element_type == "column":
                        initial_state = excel.Selection.EntireColumn.Hidden
                        if force_com:
                            excel.Selection.EntireColumn.Hidden = is_hiding
        except Exception:
            pass

        if not force_com:
            # Send the original gesture through to Excel natively
            gesture.send()
        
        import core
        # 200ms delay gives Excel enough time to process the keystroke and update its COM model
        core.callLater(200, self._verify_visibility_change_callback, element_type, is_hiding, initial_state)
        
    def _verify_visibility_change_callback(self, element_type, is_hiding, initial_state):
        """
        Callback executed by core.callLater to verify the visibility state.
        Announces the change via ui.message if successful and state actually changed.
        """
        import comtypes.client
        import comtypes.automation
        import ctypes
        import ui
        import logHandler
        
        try:
            hwnd7 = self.windowHandle if getattr(self, "windowClassName", "") == "EXCEL7" else None
            if not hwnd7:
                return

            oleacc = ctypes.windll.oleacc if hasattr(ctypes.windll, 'oleacc') else ctypes.windll.user32.oleacc
            OBJID_NATIVEOM = -16
            ptr = ctypes.POINTER(comtypes.automation.IDispatch)()
            res = oleacc.AccessibleObjectFromWindow(hwnd7, OBJID_NATIVEOM, ctypes.byref(comtypes.automation.IDispatch._iid_), ctypes.byref(ptr))

            if res == 0 and ptr:
                win = comtypes.client.dynamic.Dispatch(ptr)
                excel = win.Application
                sel = excel.Selection
                
                is_hidden = None
                address_str = ""
                if element_type == "row":
                    is_hidden = sel.EntireRow.Hidden
                    start_row = sel.Row
                    end_row = sel.Row + sel.Rows.Count - 1
                    if start_row == end_row:
                        address_str = f"Row {start_row}"
                    else:
                        address_str = f"Rows {start_row} through {end_row}"
                elif element_type == "column":
                    is_hidden = sel.EntireColumn.Hidden
                    start_col = sel.Column
                    end_col = sel.Column + sel.Columns.Count - 1
                    
                    def col_num_to_letter(n):
                        s = ""
                        while n > 0:
                            n, remainder = divmod(n - 1, 26)
                            s = chr(65 + remainder) + s
                        return s
                        
                    start_letter = col_num_to_letter(start_col)
                    if start_col == end_col:
                        address_str = f"Column {start_letter}"
                    else:
                        end_letter = col_num_to_letter(end_col)
                        address_str = f"Columns {start_letter} through {end_letter}"
                    
                # Only announce if the state successfully changed to avoid false positives
                if is_hidden is not None and is_hidden != initial_state:
                    state_str = "hidden" if is_hidden else "unhidden"
                    ui.message(f"{address_str} {state_str}")
        except Exception as e:
            logHandler.log.debugWarning(f"BOA: Failed to verify {element_type} visibility change. {e}")

    __gestures = {
        "kb:NVDA+shift+leftArrow": "moveSheetLeft",
        "kb:NVDA+shift+rightArrow": "moveSheetRight",
        "kb:NVDA+shift+left": "moveSheetLeft",
        "kb:NVDA+shift+right": "moveSheetRight",
        "kb:NVDA+shift+pageUp": "moveSheetLeft",
        "kb:NVDA+shift+pageDown": "moveSheetRight",
        "kb:NVDA+shift+home": "moveSheetStart",
        "kb:NVDA+shift+end": "moveSheetEnd",
        "kb:NVDA+alt+leftArrow": "moveSheetLeft",
        "kb:NVDA+alt+rightArrow": "moveSheetRight",
        "kb:NVDA+alt+c": "openBulkSheetOrganizer",
        "kb:control+9": "hideRow",
        "kb:control+shift+9": "unhideRow",
        "kb:control+0": "hideColumn",
        "kb:control+shift+0": "unhideColumn",
        "kb:NVDA+control+shift+0": "unhideColumnFallback",
    }

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
            import speech
            speech.speakMessage(f"Scheduled: {sheet} to position {pos}")

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
            import speech
            speech.speakMessage("Move removed")
            # Update combo box to original position
            self.on_sheet_change(None)
            
    def _refresh_list(self):
        self.list_moves.DeleteAllItems()
        for sheet, pos in self.planned_moves.items():
            self.list_moves.Append([str(sheet), str(pos)])
