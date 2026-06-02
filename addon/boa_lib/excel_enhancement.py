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
from scriptHandler import script

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
    try:
        import controlTypes
        className = getattr(obj, "windowClassName", "")
        role = getattr(obj, "role", None)
        # Filter purely to Excel's grid/cell classes to avoid unnecessary and expensive COM calls on other UI elements.
        if className not in ("EXCEL7", "NetUIHWND", "XLDESK") and role != controlTypes.Role.TABLECELL:
            return
            
        import core
        core.callLater(50, _do_check_unselect)
    except Exception:
        pass

def _do_check_unselect():
    global _last_selection_count
    try:
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
                try:
                    # ALWAYS check structural changes and hidden row skips on focus/selection change
                    import boa_config
                    if boa_config.get_feature_state("excel", "hidden_row_skip"):
                        ExcelGridMover.check_structural_changes(excel)
                        ExcelGridMover.check_hidden_skip(excel)
                except Exception:
                    pass
                count = sel.Cells.Count
                # If we previously had >1 cells selected, and now we only have 1, the user dropped the selection.
                if count == 1 and _last_selection_count > 1:
                    import ui
                    ui.message("unselected")
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
        
        def _restore_clip_and_reset(old_clip):
            if old_clip:
                import api
                try:
                    api.copyToClip(old_clip)
                except Exception:
                    pass
            global _is_renaming_sheet
            _is_renaming_sheet = False

        def _do_enter(clean_name, old_clip):
            keyboardHandler.KeyboardInputGesture.fromName("enter").send()
            import ui
            ui.message(f"Renaming to {clean_name}")
            core.callLater(1500, lambda: _restore_clip_and_reset(old_clip))

        def _do_inject(old_clip, clean_name, fg_hwnd):
            if winUser.getForegroundWindow() == fg_hwnd:
                keyboardHandler.KeyboardInputGesture.fromName("control+v").send()
                core.callLater(200, lambda: _do_enter(clean_name, old_clip))
            else:
                _restore_clip_and_reset(old_clip)

        def _do_clipboard(clean_name, fg_hwnd):
            import api
            old_clip = ""
            try:
                old_clip = api.getClipData()
            except Exception:
                pass
            try:
                api.copyToClip(clean_name)
                core.callLater(200, lambda: _do_inject(old_clip, clean_name, fg_hwnd))
            except Exception:
                _restore_clip_and_reset(old_clip)

        def _check_security():
            import ctypes
            fg_hwnd = winUser.getForegroundWindow()
            fg_pid = ctypes.c_ulong()
            ctypes.windll.user32.GetWindowThreadProcessId(fg_hwnd, ctypes.byref(fg_pid))
            
            target_pid = ctypes.c_ulong()
            ctypes.windll.user32.GetWindowThreadProcessId(hwnd, ctypes.byref(target_pid))
            
            if fg_pid.value != target_pid.value:
                log.warning("BOA: Foreground window mismatch! Aborting keystroke injection to prevent pasting into wrong app.")
                global _is_renaming_sheet
                _is_renaming_sheet = False
                return
            
            clean_name = new_name.strip() if new_name else ""
            if not clean_name:
                keyboardHandler.KeyboardInputGesture.fromName("escape").send()
                _is_renaming_sheet = False
                return
                
            _do_clipboard(clean_name, fg_hwnd)

        def _set_foreground():
            winUser.setForegroundWindow(hwnd)
            core.callLater(200, _check_security)
            
        core.callLater(100, _set_foreground)

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
    """
    Excel Grid Keystroke Interceptor and Event Monitor.
    WHY THIS EXISTS (Architecture intent):
    Native NVDA UIA relies on Excel exposing accurate bounds and states. However, Excel often
    hides rows/columns physically but still exposes them in UIA, causing NVDA to jump silently.
    This class intercepts arrow keys and structural changes to manually calculate and announce
    boundaries, hidden skips, and sheet movements, bypassing the unreliable UIA layer entirely.
    """





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
        import comtypes.client
        import comtypes.automation
        import ctypes
        import ui
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
                    import boa_config
                    if boa_config.get_feature_state("excel", "hidden_row_skip"):
                        ExcelGridMover.check_structural_changes(excel)
                        ExcelGridMover.check_hidden_skip(excel)
                    
                    # Verify it is a Range object (has Cells property) and has more than 1 cell selected.
                    if sel.Cells.Count > 1:
                        # GUARD 1: If Shift or Control is held down, the user is manually selecting.
                        # NVDA will natively announce this. Do not double-announce multi-cell selections.
                        if (winUser.getKeyState(winUser.VK_SHIFT) & 32768) or (winUser.getKeyState(winUser.VK_CONTROL) & 32768):
                            return
                            
                        address = sel.Address(False, False)  # Returns string like "A1:D1"

                        # GUARD: Only announce if this is a NEW address BOA hasn't spoken yet.
                        if address == _last_announced_address:
                            return

                        _last_announced_address = address
                        spoken_address = address.replace(":", " through ")
                        ui.message(f"{spoken_address} selected")
                    else:
                        # Selection collapsed to single cell — reset tracker
                        _last_announced_address = None
        except Exception:
            pass

    @staticmethod
    def check_hidden_skip(excel):
        """
        Detects if the user's focus jumped over completely hidden rows or columns.
        Uses COM bulk checking and SpecialCells to avoid false positives and loops.
        Announces mixed scenarios accurately and includes proactive boundary detection.
        """
        global _last_focused_row, _last_focused_col, _last_focused_sheet, _last_focused_wb
        import ui
        import logHandler
        
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
                row_changed = current_row != _last_focused_row
                col_changed = current_col != _last_focused_col
                
                def col_num_to_letter(n):
                    s = ""
                    while n > 0:
                        n, remainder = divmod(n - 1, 26)
                        s = chr(65 + remainder) + s
                    return s
                
                skip_announced_top = False
                skip_announced_bottom = False
                skip_announced_left = False
                skip_announced_right = False
                
                # ---------------- ROW JUMPS ----------------
                if col_changed == False and row_changed:
                    min_r = min(_last_focused_row, current_row)
                    max_r = max(_last_focused_row, current_row)
                    
                    if max_r - min_r > 1:
                        gap_range = excel.Range(excel.Cells(min_r + 1, 1), excel.Cells(max_r - 1, 1)).EntireRow
                        hidden_val = gap_range.Hidden
                        
                        # Architecture intent: COM Range.Hidden returns True if all rows are hidden, 
                        # False if all are visible, and None/Null if it's a mix.
                        # By checking 'is not False', we instantly catch both fully hidden and mixed 
                        # ranges without using slow and error-prone SpecialCells iterations.
                        if hidden_val is not False:
                            if hidden_val in (True, -1):
                                if max_r - min_r == 2:
                                    ui.message(f"Row {min_r + 1} hidden")
                                else:
                                    ui.message(f"Rows {min_r + 1} through {max_r - 1} hidden")
                            else:
                                ui.message("Crossed hidden rows")
                                
                            if current_row > _last_focused_row:
                                skip_announced_top = True
                            else:
                                skip_announced_bottom = True

                # ---------------- COLUMN JUMPS ----------------
                if row_changed == False and col_changed:
                    min_c = min(_last_focused_col, current_col)
                    max_c = max(_last_focused_col, current_col)
                        
                    if max_c - min_c > 1:
                        gap_range = excel.Range(excel.Cells(1, min_c + 1), excel.Cells(1, max_c - 1)).EntireColumn
                        hidden_val = gap_range.Hidden
                        
                        start_letter = col_num_to_letter(min_c + 1)
                        # Architecture intent: Similar to rows, we check if the column hidden value
                        # is not False to reliably detect mixed or fully hidden columns during bulk navigation.
                        if hidden_val is not False:
                            if hidden_val in (True, -1):
                                if max_c - min_c == 2:
                                    ui.message(f"Column {start_letter} hidden")
                                else:
                                    end_letter = col_num_to_letter(max_c - 1)
                                    ui.message(f"Columns {start_letter} through {end_letter} hidden")
                            else:
                                ui.message("Crossed hidden columns")
                                
                            if current_col > _last_focused_col:
                                skip_announced_left = True
                            else:
                                skip_announced_right = True

                # ---------------- PROACTIVE BOUNDARY DETECTION ----------------
                # Check top boundary
                if row_changed and current_row > 1 and not skip_announced_top:
                    try:
                        if excel.Rows(current_row - 1).Hidden in (True, -1) and excel.Rows(1).Hidden in (True, -1):
                            if excel.Range(excel.Cells(1, 1), excel.Cells(current_row - 1, 1)).EntireRow.Hidden in (True, -1):
                                if current_row == 2:
                                    ui.message("Top boundary. Row 1 hidden")
                                else:
                                    ui.message(f"Top boundary. Rows 1 to {current_row - 1} hidden")
                    except Exception:
                        pass
                
                # Check bottom boundary
                if row_changed and current_row < 1048576 and not skip_announced_bottom:
                    try:
                        if excel.Rows(current_row + 1).Hidden in (True, -1) and excel.Rows(1048576).Hidden in (True, -1):
                            if excel.Range(excel.Cells(current_row + 1, 1), excel.Cells(1048576, 1)).EntireRow.Hidden in (True, -1):
                                if current_row == 1048575:
                                    ui.message("Bottom boundary. Row 1048576 hidden")
                                else:
                                    ui.message(f"Bottom boundary. Rows {current_row + 1} to 1048576 hidden")
                    except Exception:
                        pass
                        
                # Check left boundary
                if col_changed and current_col > 1 and not skip_announced_left:
                    try:
                        if excel.Columns(current_col - 1).Hidden in (True, -1) and excel.Columns(1).Hidden in (True, -1):
                            if excel.Range(excel.Cells(1, 1), excel.Cells(1, current_col - 1)).EntireColumn.Hidden in (True, -1):
                                if current_col == 2:
                                    ui.message("Left boundary. Column A hidden")
                                else:
                                    ui.message(f"Left boundary. Columns A to {col_num_to_letter(current_col - 1)} hidden")
                    except Exception:
                        pass

                # Check right boundary
                if col_changed and current_col < 16384 and not skip_announced_right:
                    try:
                        if excel.Columns(current_col + 1).Hidden in (True, -1) and excel.Columns(16384).Hidden in (True, -1):
                            if excel.Range(excel.Cells(1, current_col + 1), excel.Cells(1, 16384)).EntireColumn.Hidden in (True, -1):
                                if current_col == 16383:
                                    ui.message("Right boundary. Column XFD hidden")
                                else:
                                    ui.message(f"Right boundary. Columns {col_num_to_letter(current_col + 1)} to XFD hidden")
                    except Exception:
                        pass

            _last_focused_row = current_row
            _last_focused_col = current_col
        except Exception as e:
            try:
                logHandler.log.debug(f"BOA: Failed to check hidden skip: {e}")
            except Exception:
                pass
    @staticmethod
    def check_structural_changes(excel):
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

    # Gestures merged to the bottom __gestures block

    def _move_sheet(self, direction):
        import comtypes.client
        import comtypes.automation
        import ctypes
        import ui
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
                ui.message("Could not find Excel grid.")
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
                ui.message("Failed to hook Excel.")
                return
                
            win = comtypes.client.dynamic.Dispatch(ptr)
            excel = win.Application
            wb = excel.ActiveWorkbook
            sheet = excel.ActiveSheet
            
            current_index = sheet.Index
            total_sheets = wb.Sheets.Count
            
            if total_sheets <= 1:
                ui.message("Only one sheet in workbook")
                return
                
            if direction == "left":
                if current_index == 1:
                    ui.message("Already at beginning")
                    return
                sheet.Move(wb.Sheets(current_index - 1))
            elif direction == "right":
                if current_index == total_sheets:
                    ui.message("Already at end")
                    return
                # To move right, we just move the right neighbor before us!
                wb.Sheets(current_index + 1).Move(sheet)
            elif direction == "start":
                if current_index == 1:
                    ui.message("Already at beginning")
                    return
                sheet.Move(wb.Sheets(1))
            elif direction == "end":
                if current_index == total_sheets:
                    ui.message("Already at end")
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
            ui.message(f"Moved {sheet_name} to position {new_index} of {total_sheets}")
        except Exception as e:
            ui.message("Failed to move sheet")
            import logHandler
            logHandler.log.error(f"ExcelGridMover error: {e}")

    # (script import moved to top of file)

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
            def _do_com_moves():
                import comtypes.client
                import comtypes.automation
                import ctypes
                import ui
                import logHandler
                
                oleacc = ctypes.windll.user32.oleacc if hasattr(ctypes.windll.user32, 'oleacc') else ctypes.windll.oleacc
                ptr = ctypes.POINTER(comtypes.automation.IDispatch)()
                res = oleacc.AccessibleObjectFromWindow(hwnd, -16, ctypes.byref(comtypes.automation.IDispatch._iid_), ctypes.byref(ptr))
                if res == 0 and ptr:
                    try:
                        excel = comtypes.client.dynamic.Dispatch(ptr).Application
                        wb = excel.ActiveWorkbook
                        
                        unmoved = [s for s in sheet_names if s not in planned_moves]
                        moved = [(s, planned_moves[s]) for s in planned_moves]
                        moved.sort(key=lambda x: x[1])
                        
                        final_list = unmoved[:]
                        for s, pos in moved:
                            insert_idx = min(pos - 1, len(final_list))
                            final_list.insert(insert_idx, s)
                            
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
                                    
                        ui.message("Bulk arrangement complete")
                    except Exception as e:
                        ui.message(f"Error during bulk move: {e}")
                        logHandler.log.error(f"BOA bulk bg error: {e}")
                        
            def _apply_moves():
                import winUser
                winUser.setForegroundWindow(hwnd)
                import core
                core.callLater(200, _do_com_moves)
                
            import core
            core.callLater(100, _apply_moves)

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
        import ui
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
            wx.CallAfter(self._show_bulk_dialog, sheet_names, hwnd7)
        except Exception as e:
            ui.message("Error opening organizer")
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
