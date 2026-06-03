import NVDAObjects.UIA
import NVDAObjects.IAccessible
import NVDAObjects.window.edit
import controlTypes
from logHandler import log
import UIAHandler
import wx
import gui
import threading
import time
import winUser
import keyboardHandler
import core
from scriptHandler import script
import queueHandler
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
_last_structural_wb = None
_last_excel_hwnd = None
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
                    from boa_lib import boa_config
                    if boa_config.get_feature_state("excel", "hidden_row_skip"):
                        CellNavigationTracker.check_structural_changes(excel)
                        CellNavigationTracker.check_hidden_skip(excel)
                except Exception:
                    pass
                count = sel.Cells.Count
                # If we previously had >1 cells selected, and now we only have 1, the user dropped the selection.
                if count == 1 and _last_selection_count > 1:
                    if boa_config.get_feature_state("excel", "unselect_tracking"):
                        import ui
                        ui.message("unselected")
                _last_selection_count = count
    except Exception:
        pass

class CellNavigationTracker(object):
    """
    Excel Grid Keystroke Interceptor and Event Monitor.
    WHY THIS EXISTS (Architecture intent):
    Native NVDA UIA relies on Excel exposing accurate bounds and states. However, Excel often
    hides rows/columns physically but still exposes them in UIA, causing NVDA to jump silently.
    This class intercepts arrow keys and structural changes to manually calculate and announce
    boundaries, hidden skips, and sheet movements, bypassing the unreliable UIA layer entirely.
    """





    def event_gainFocus(self):
        super(CellNavigationTracker, self).event_gainFocus()
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
                    from boa_lib import boa_config
                    if boa_config.get_feature_state("excel", "hidden_row_skip"):
                        CellNavigationTracker.check_structural_changes(excel)
                        CellNavigationTracker.check_hidden_skip(excel)
                    
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
                        
                        if boa_config.get_feature_state("excel", "unselect_tracking"):
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
                        hidden_count = 0
                        fragmented = False
                        first_hidden = None
                        last_hidden = None
                        
                        sheet = excel.ActiveSheet
                        
                        if max_r - min_r == 2:
                            if excel.Rows(min_r + 1).Hidden in (True, -1):
                                first_hidden = min_r + 1
                                last_hidden = min_r + 1
                                hidden_count = 1
                        else:
                            gap_range = sheet.Range(sheet.Cells(min_r + 1, current_col), sheet.Cells(max_r - 1, current_col))
                            try:
                                visible_range = gap_range.SpecialCells(12)
                                areas_count = visible_range.Areas.Count
                                
                                if areas_count == 1:
                                    vis_start = visible_range.Row
                                    vis_end = visible_range.Row + visible_range.Rows.Count - 1
                                    int_start = gap_range.Row
                                    int_end = gap_range.Row + gap_range.Rows.Count - 1
                                    
                                    if vis_start > int_start:
                                        first_hidden = int_start
                                        last_hidden = vis_start - 1
                                        hidden_count = last_hidden - first_hidden + 1
                                    elif vis_end < int_end:
                                        first_hidden = vis_end + 1
                                        last_hidden = int_end
                                        hidden_count = last_hidden - first_hidden + 1
                                elif areas_count == 2:
                                    area1 = visible_range.Areas.Item(1)
                                    area2 = visible_range.Areas.Item(2)
                                    first_hidden = area1.Row + area1.Rows.Count
                                    last_hidden = area2.Row - 1
                                    hidden_count = last_hidden - first_hidden + 1
                                else:
                                    fragmented = True
                                    hidden_count = 3  # Triggers plural
                            except Exception:
                                # "No cells were found" -> Entire gap is hidden
                                first_hidden = gap_range.Row
                                last_hidden = gap_range.Row + gap_range.Rows.Count - 1
                                hidden_count = last_hidden - first_hidden + 1
                                
                        if hidden_count > 0:
                            if not fragmented:
                                if hidden_count == 1:
                                    ui.message(f"Row {first_hidden} hidden")
                                else:
                                    if first_hidden == min_r + 1 and last_hidden == max_r - 1 and (max_r - min_r - 1) >= 500:
                                        ui.message(f"Rows {min_r + 1} through {max_r - 1} hidden")
                                    else:
                                        ui.message(f"Rows {first_hidden} through {last_hidden} hidden")
                            else:
                                ui.message("Crossed heavily fragmented hidden rows")
                                
                            if current_row > _last_focused_row:
                                skip_announced_top = True
                            else:
                                skip_announced_bottom = True

                # ---------------- COLUMN JUMPS ----------------
                if row_changed == False and col_changed:
                    min_c = min(_last_focused_col, current_col)
                    max_c = max(_last_focused_col, current_col)
                        
                    if max_c - min_c > 1:
                        hidden_count = 0
                        fragmented = False
                        first_hidden = None
                        last_hidden = None
                        
                        sheet = excel.ActiveSheet
                        
                        if max_c - min_c == 2:
                            if excel.Columns(min_c + 1).Hidden in (True, -1):
                                first_hidden = min_c + 1
                                last_hidden = min_c + 1
                                hidden_count = 1
                        else:
                            gap_range = sheet.Range(sheet.Cells(current_row, min_c + 1), sheet.Cells(current_row, max_c - 1))
                            try:
                                visible_range = gap_range.SpecialCells(12)
                                areas_count = visible_range.Areas.Count
                                
                                if areas_count == 1:
                                    vis_start = visible_range.Column
                                    vis_end = visible_range.Column + visible_range.Columns.Count - 1
                                    int_start = gap_range.Column
                                    int_end = gap_range.Column + gap_range.Columns.Count - 1
                                    
                                    if vis_start > int_start:
                                        first_hidden = int_start
                                        last_hidden = vis_start - 1
                                        hidden_count = last_hidden - first_hidden + 1
                                    elif vis_end < int_end:
                                        first_hidden = vis_end + 1
                                        last_hidden = int_end
                                        hidden_count = last_hidden - first_hidden + 1
                                elif areas_count == 2:
                                    area1 = visible_range.Areas.Item(1)
                                    area2 = visible_range.Areas.Item(2)
                                    first_hidden = area1.Column + area1.Columns.Count
                                    last_hidden = area2.Column - 1
                                    hidden_count = last_hidden - first_hidden + 1
                                else:
                                    fragmented = True
                                    hidden_count = 3  # Triggers plural
                            except Exception:
                                # "No cells were found" -> Entire intersection is hidden
                                first_hidden = gap_range.Column
                                last_hidden = gap_range.Column + gap_range.Columns.Count - 1
                                hidden_count = last_hidden - first_hidden + 1
                                
                        if hidden_count > 0:
                            if not fragmented:
                                if hidden_count == 1:
                                    ui.message(f"Column {col_num_to_letter(first_hidden)} hidden")
                                else:
                                    if first_hidden == min_c + 1 and last_hidden == max_c - 1 and (max_c - min_c - 1) >= 500:
                                        ui.message(f"Columns {col_num_to_letter(min_c + 1)} through {col_num_to_letter(max_c - 1)} hidden")
                                    else:
                                        ui.message(f"Columns {col_num_to_letter(first_hidden)} through {col_num_to_letter(last_hidden)} hidden")
                            else:
                                ui.message("Crossed heavily fragmented hidden columns")
                                
                            if current_col > _last_focused_col:
                                skip_announced_left = True
                            else:
                                skip_announced_right = True

                # ---------------- PROACTIVE BOUNDARY DETECTION ----------------
                # Check top boundary
                if row_changed and current_row > 1 and not skip_announced_top:
                    try:
                        if excel.Rows(current_row - 1).Hidden in (True, -1):
                            gap_range = excel.Range(excel.Cells(1, current_col), excel.Cells(current_row - 1, current_col))
                            try:
                                visible_cells = gap_range.SpecialCells(12)
                                last_area = visible_cells.Areas.Item(visible_cells.Areas.Count)
                                hidden_start = last_area.Row + last_area.Rows.Count
                            except Exception:
                                hidden_start = 1
                                
                            if hidden_start == 1:
                                ui.message(f"Top boundary. Rows 1 to {current_row - 1} hidden")
                    except Exception:
                        pass
                
                # Check bottom boundary
                if row_changed and current_row < 1048576 and not skip_announced_bottom:
                    try:
                        if excel.Rows(current_row + 1).Hidden in (True, -1):
                            gap_range = excel.Range(excel.Cells(current_row + 1, current_col), excel.Cells(1048576, current_col))
                            try:
                                visible_cells = gap_range.SpecialCells(12)
                                first_area = visible_cells.Areas.Item(1)
                                hidden_end = first_area.Row - 1
                            except Exception:
                                hidden_end = 1048576
                            
                            if hidden_end == 1048576:
                                ui.message(f"Bottom boundary. Rows {current_row + 1} to 1048576 hidden")
                    except Exception:
                        pass
                        
                # Check left boundary
                if col_changed and current_col > 1 and not skip_announced_left:
                    try:
                        if excel.Columns(current_col - 1).Hidden in (True, -1):
                            gap_range = excel.Range(excel.Cells(current_row, 1), excel.Cells(current_row, current_col - 1))
                            try:
                                visible_cells = gap_range.SpecialCells(12)
                                last_area = visible_cells.Areas.Item(visible_cells.Areas.Count)
                                hidden_start = last_area.Column + last_area.Columns.Count
                            except Exception:
                                hidden_start = 1
                            
                            if hidden_start == 1:
                                ui.message(f"Left boundary. Columns A to {col_num_to_letter(current_col - 1)} hidden")
                    except Exception:
                        pass

                # Check right boundary
                if col_changed and current_col < 16384 and not skip_announced_right:
                    try:
                        if excel.Columns(current_col + 1).Hidden in (True, -1):
                            gap_range = excel.Range(excel.Cells(current_row, current_col + 1), excel.Cells(current_row, 16384))
                            try:
                                visible_cells = gap_range.SpecialCells(12)
                                first_area = visible_cells.Areas.Item(1)
                                hidden_end = first_area.Column - 1
                            except Exception:
                                hidden_end = 16384
                                
                            if hidden_end == 16384:
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
        global _last_freeze_panes_state, _last_visible_sheet_count, _last_focused_sheet, _last_focused_wb, _last_structural_wb, _last_excel_hwnd
        import ui
        
        try:
            current_wb_name = excel.ActiveWorkbook.Name
            current_hwnd = excel.Application.Hwnd
            
            # Reset trackers if the user closed and reopened Excel (new Process Window Handle)
            if _last_excel_hwnd != current_hwnd:
                _last_structural_wb = None
                _last_focused_wb = None
                _last_visible_sheet_count = None
                _last_freeze_panes_state = None
                _last_excel_hwnd = current_hwnd
            
            # Reset trackers if the user switched to a different workbook
            if _last_structural_wb != current_wb_name:
                _last_visible_sheet_count = None
                _last_freeze_panes_state = None
                _last_structural_wb = current_wb_name
                
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

        from boa_lib import boa_config
        if not force_com:
            # Send the original gesture through to Excel natively
            gesture.send()
        
        if not boa_config.get_feature_state("excel", "hidden_row_skip"):
            return
            
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
        "kb:control+9": "hideRow",
        "kb:control+shift+9": "unhideRow",
        "kb:control+0": "hideColumn",
        "kb:control+shift+0": "unhideColumn",
        "kb:NVDA+control+shift+0": "unhideColumnFallback",
    }
