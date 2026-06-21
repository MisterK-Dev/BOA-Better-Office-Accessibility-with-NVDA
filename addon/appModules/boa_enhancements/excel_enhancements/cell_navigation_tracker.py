# -*- coding: UTF-8 -*-
# Copyright (C) 2026 KIRAN G T {MisterK} and Antigravity 2
# This file is covered by the GNU General Public License (GPL), version 2.
# See the file COPYING.txt for more details.

import addonHandler
addonHandler.initTranslation()

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
_cached_excel_app = None
_cached_excel_pid = None
_excel_event_connection = None
_last_force_synced_address = None

_drift_timer_running = False
_pending_injection_address = None

def _drift_poll_loop():
    global _drift_timer_running
    if not _drift_timer_running: return
    try:
        global _cached_excel_app
        if not _cached_excel_app: 
            import core
            core.callLater(100, _drift_poll_loop)
            return
            
        active_cell = _cached_excel_app.ActiveCell
        if not active_cell:
            import core
            core.callLater(100, _drift_poll_loop)
            return
            
        global _last_focused_row, _last_focused_col, _last_focused_sheet, _last_focused_wb
        
        current_row = active_cell.Row
        current_col = active_cell.Column
        try:
            current_sheet = _cached_excel_app.ActiveSheet.Name
            current_wb = _cached_excel_app.ActiveWorkbook.Name
        except Exception:
            current_sheet = None
            current_wb = None
        
        is_match = False
        if _last_focused_row == current_row and _last_focused_col == current_col:
            if _last_focused_sheet == current_sheet and _last_focused_wb == current_wb:
                is_match = True
                
        global _last_force_synced_address
        global _pending_injection_address
        current_address = f"[{current_wb}]{current_sheet}!{active_cell.Address(False, False)}"
        
        if not is_match:
            if current_address != _last_force_synced_address and current_address != _pending_injection_address:
                _pending_injection_address = current_address
                
                def _verify_and_inject(expected_row, expected_col, expected_sheet, expected_wb, expected_addr, expected_cell):
                    global _pending_injection_address
                    _pending_injection_address = None
                    
                    # Check if NVDA cache STILL hasn't updated to match expected
                    global _last_focused_row, _last_focused_col, _last_focused_sheet, _last_focused_wb
                    still_mismatched = False
                    if _last_focused_row != expected_row or _last_focused_col != expected_col:
                        still_mismatched = True
                    elif _last_focused_sheet != expected_sheet or _last_focused_wb != expected_wb:
                        still_mismatched = True
                        
                    if still_mismatched:
                        import logHandler
                        logHandler.log.debug(f"BOA: Drift confirmed after 100ms. NVDA totally missed jump to {expected_addr}. Forcing COM sync.")
                        global _last_force_synced_address
                        _last_force_synced_address = expected_addr
                        try:
                            import api
                            import eventHandler
                            import controlTypes
                            
                            # 1. Update our cache manually to ensure the tracker doesn't break on the next arrow key
                            _last_focused_row = expected_row
                            _last_focused_col = expected_col
                            _last_focused_sheet = expected_sheet
                            _last_focused_wb = expected_wb
                            
                            native_success = False
                            
                            # Attempt 1: Try querying OS focus
                            focus_obj = api.getDesktopObject().objectWithFocus()
                            if focus_obj and getattr(focus_obj, "role", None) == controlTypes.Role.TABLECELL:
                                eventHandler.executeEvent("gainFocus", focus_obj)
                                native_success = True
                            
                            # Attempt 2: Try building the NVDAObject manually (Legacy MSAA)
                            if not native_success:
                                try:
                                    from NVDAObjects.window.excel import ExcelCell
                                    fg = api.getForegroundObject()
                                    excelWindow = fg
                                    while excelWindow and getattr(excelWindow, "windowClassName", "") != "EXCEL7":
                                        excelWindow = excelWindow.parent
                                        
                                    if excelWindow:
                                        obj = ExcelCell(windowHandle=excelWindow.windowHandle, excelWindowObject=excelWindow, excelCellObject=expected_cell)
                                        eventHandler.executeEvent("gainFocus", obj)
                                        native_success = True
                                except Exception:
                                    pass
                                    
                            # Fallback: Perfect UI Message Synthesis (Identical to native speech/Braille flash)
                            if not native_success:
                                import ui
                                val = expected_cell.Value
                                if val is None: val = ""
                                elif type(val) is float and val.is_integer(): val = int(val)
                                
                                addr = expected_cell.Address(False, False)
                                ui.message(f"{val} {addr}".strip())
                                
                            # ALWAYS Handle Multi-cell selection ranges (Ctrl+Shift+[)
                            try:
                                if _cached_excel_app.Selection.Cells.Count > 1:
                                    sel_addr = _cached_excel_app.Selection.Address(False, False)
                                    global _last_announced_address
                                    if sel_addr != _last_announced_address:
                                        _last_announced_address = sel_addr
                                        spoken_address = sel_addr.replace(":", " through ")
                                        import ui
                                        ui.message(_("{address} selected").format(address=spoken_address))
                            except Exception:
                                pass
                                
                        except Exception as e:
                            import logHandler
                            logHandler.log.debug(f"BOA: Complete failure in force focus sync: {e}")
                
                import core
                core.callLater(100, _verify_and_inject, current_row, current_col, current_sheet, current_wb, current_address, active_cell)
        else:
            # Successfully matched natively! Clear the force sync tracker.
            _last_force_synced_address = None
            _pending_injection_address = None
    except Exception:
        pass
        
    import core
    core.callLater(100, _drift_poll_loop)

def _get_excel_app():
    """
    Retrieves and caches the active Excel Application COM object.
    Uses the cached connection if it is still alive to avoid expensive Running Object Table lookups.
    """
    global _cached_excel_app, _cached_excel_pid
    if _cached_excel_app is not None:
        try:
            # Query a simple, fast property to verify the COM object is alive and responding
            _unused_version = _cached_excel_app.Version
            return _cached_excel_app
        except Exception:
            _cached_excel_app = None
            _cached_excel_pid = None

    import comtypes.client
    import comtypes.automation
    import ctypes
    import winUser
    
    app = None
    try:
        # Standard fast retrieval
        app = comtypes.client.GetActiveObject("Excel.Application")
    except Exception:
        pass

    if not app:
        try:
            # Fallback raw grid tree crawl
            hwnd7 = ctypes.windll.user32.FindWindowW("XLMAIN", None)
            if hwnd7:
                xldesk = ctypes.windll.user32.FindWindowExW(hwnd7, 0, "XLDESK", None)
                if xldesk:
                    hwnd7 = ctypes.windll.user32.FindWindowExW(xldesk, 0, "EXCEL7", None)
            if hwnd7:
                oleacc = ctypes.windll.oleacc if hasattr(ctypes.windll, 'oleacc') else ctypes.windll.user32.oleacc
                ptr = ctypes.POINTER(comtypes.automation.IDispatch)()
                res = oleacc.AccessibleObjectFromWindow(hwnd7, -16, ctypes.byref(comtypes.automation.IDispatch._iid_), ctypes.byref(ptr))
                if res == 0 and ptr:
                    app = comtypes.client.dynamic.Dispatch(ptr).Application
        except Exception:
            pass

    if app:
        try:
            global _drift_timer_running
            _cached_excel_app = app
            # Get the exact PID from the COM process
            import ctypes
            pid = ctypes.c_ulong()
            hwnd = app.Hwnd
            ctypes.windll.user32.GetWindowThreadProcessId(hwnd, ctypes.byref(pid))
            _cached_excel_pid = pid.value
            
            # Start background poller instead of relying on COM Event Sinks
            if not _drift_timer_running:
                _drift_timer_running = True
                import core
                core.callLater(100, _drift_poll_loop)
                
            return _cached_excel_app
        except Exception:
            pass

    return None

def suspend_tracker_and_release():
    """
    Instantly pauses the background 100ms COM polling loop and checks for Excel closure.
    
    Architectural Intent & Considerations:
    If this loop runs continuously while Excel is minimized, it causes a global performance 
    drain on NVDA by firing COM requests across process boundaries endlessly. By setting 
    _drift_timer_running to False when Excel loses focus, we preserve CPU. The tracker will 
    automatically revive itself via _get_excel_app() the instant the user returns to the grid.
    """
    global _drift_timer_running
    _drift_timer_running = False
    release_if_closed()

def release_if_closed():
    """
    Checks if there are any EXCEL.EXE windows remaining for the cached PID.
    If none exist, releases the COM cache so the background process can terminate.
    """
    global _cached_excel_app, _cached_excel_pid
    if _cached_excel_app is None or _cached_excel_pid is None:
        return
        
    import winUser
    import ctypes
    import ctypes.wintypes as wintypes
    import gc
    
    found_window = False
    
    # EnumWindows callback
    def enum_windows_proc(hwnd, lParam):
        nonlocal found_window
        if found_window:
            return False # Stop enumeration
            
        # Check class name
        class_name = winUser.getClassName(hwnd)
        if class_name == "XLMAIN":
            _unused_thread_id, pid = winUser.getWindowThreadProcessID(hwnd)
            if pid == _cached_excel_pid:
                found_window = True
                return False # Stop enumeration
        return True
        
    CMPFUNC = ctypes.WINFUNCTYPE(ctypes.c_bool, wintypes.HWND, wintypes.LPARAM)
    enum_func = CMPFUNC(enum_windows_proc)
    
    ctypes.windll.user32.EnumWindows(enum_func, 0)
    
    if not found_window:
        # No windows belonging to this Excel process are open. Release COM!
        global _drift_timer_running
        _drift_timer_running = False
        _cached_excel_app = None
        _cached_excel_pid = None
        # Explicit garbage collection ensures COM proxy is deleted, reducing ref count
        gc.collect()

# Tracks the last multi-cell address BOA announced aloud.
# This prevents BOA from re-announcing the same range that NVDA already spoke natively.
_last_announced_address = None

# Track states for structural Excel changes that lack native COM/UIA events
_last_freeze_panes_state = None
_last_visible_sheet_count = None
_last_total_sheet_count = None

# Track last focused cell coordinates to detect jumps over hidden rows/cols
_last_focused_row = None
_last_focused_col = None
_last_focused_sheet = None
_last_focused_wb = None
_last_structural_wb = None
_last_structural_sheet = None
_last_excel_hwnd = None
def check_unselect(obj):
    """
    Called whenever Excel selection or focus changes.
    ARCHITECTURAL INTENT: Detects when a user has a multi-cell selection (e.g., A1 through D1)
    and suddenly presses an arrow key, causing the selection to drop down to a single cell.
    Standard NVDA does not announce "unselected" in Excel, leaving the user unaware they lost their range.
    By hooking into global focus events, we can detect this regression and provide feedback.
    """
    try:
        import controlTypes
        className = getattr(obj, "windowClassName", "")
        role = getattr(obj, "role", None)
        # Filter purely to Excel's grid/cell classes to avoid unnecessary and expensive COM calls on other UI elements.
        if className not in ("EXCEL7", "NetUIHWND", "XLDESK") and role != controlTypes.Role.TABLECELL:
            return
            
        import core
        # Defer execution by 50ms so Excel's internal selection state has time to update
        core.callLater(50, _do_check_unselect)
    except Exception:
        pass

def _do_check_unselect():
    """
    Performs the actual COM check to determine if the selection has collapsed.
    ARCHITECTURAL INTENT: Runs asynchronously to avoid blocking NVDA's event queue. Uses COM
    to read the Selection.Count and compares it to the previously cached count.
    """
    global _last_selection_count
    try:
        excel = _get_excel_app()
        if excel:
            sel = excel.Selection
            # Ensure the selection is actually a Range (has Cells) and not a Shape/Chart
            if getattr(sel, 'Cells', None):
                try:
                    # ALWAYS check structural changes and hidden row skips on focus/selection change
                    from appModules.boa_enhancements import boa_config
                    if boa_config.get_feature_state("excel", "hidden_row_skip"):
                        CellNavigationTracker.check_structural_changes(excel)
                        
                    if (boa_config.get_feature_state("excel", "hidden_row_skip") or 
                        boa_config.get_feature_state("excel", "auto_announce_first_block") in ["one_time", "guided"]):
                        CellNavigationTracker.check_hidden_skip(excel)
                except Exception:
                    pass
                count = sel.Cells.Count
                # If we previously had >1 cells selected, and now we only have 1, the user dropped the selection.
                if count == 1 and _last_selection_count > 1:
                    if boa_config.get_feature_state("excel", "unselect_tracking"):
                        import ui
                        ui.message(_("unselected"))
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
        """
        Triggered natively by NVDA when the Excel grid gains focus.
        
        Architectural Intent & Considerations:
        We MUST call `super()` immediately to ensure NVDA's base UIA handler processes the focus event 
        properly. We then schedule our custom multi-selection check with a 50ms delay, allowing Excel's 
        internal COM engine enough time to register the new focus state before we query it.
        """
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
            excel = _get_excel_app()
            if excel:
                sel = excel.Selection

                if getattr(sel, 'Cells', None):
                    # ALWAYS check structural changes and hidden row skips on focus change
                    from appModules.boa_enhancements import boa_config
                    
                    if boa_config.get_feature_state("excel", "hidden_row_skip"):
                        CellNavigationTracker.check_structural_changes(excel)
                        
                    if (boa_config.get_feature_state("excel", "hidden_row_skip") or 
                        boa_config.get_feature_state("excel", "auto_announce_first_block") in ["one_time", "guided"]):
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
                            ui.message(_("{address} selected").format(address=spoken_address))
                    else:
                        # Selection collapsed to single cell — reset tracker
                        _last_announced_address = None
        except Exception:
            pass

    @staticmethod
    def check_hidden_skip(excel):
        """
        Detects if the user's focus jumped over completely hidden rows or columns.
        ARCHITECTURAL INTENT: NVDA relies on UI Automation (UIA) events to track focus. 
        However, Excel exposes physically hidden cells to UIA as if they were visible, causing 
        NVDA to silently skip over them without warning the user. We use COM bulk checking 
        (SpecialCells) here to manually calculate hidden gaps whenever the focus coordinates 
        jump by more than 1 unit.
        """
        global _last_focused_row, _last_focused_col, _last_focused_sheet, _last_focused_wb
        import ui
        import logHandler
        from appModules.boa_enhancements import boa_config
        
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
                # One-time mode triggers only on sheet change/open
                if boa_config.get_feature_state("excel", "auto_announce_first_block") == "one_time":
                    from .sheet_layout_analyzer import SheetLayoutAnalyzer
                    SheetLayoutAnalyzer.auto_announce_one_time(excel)
                    
                _last_focused_row = None
                _last_focused_col = None
                _last_focused_sheet = current_sheet
                _last_focused_wb = current_wb
            
            row_changed = True
            col_changed = True
            
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
                
                if boa_config.get_feature_state("excel", "hidden_row_skip"):
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
                                # Consideration: Manually looping through hundreds of rows to check their .Hidden property 
                                # via COM is disastrously slow and will freeze NVDA. Instead, we define the gap and 
                                # ask Excel to return ONLY the visible cells using SpecialCells(12) (xlCellTypeVisible). 
                                # By comparing the visible chunks against the total gap, we mathematically deduce what is hidden instantly.
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
                                        ui.message(_("Row {row_num} hidden").format(row_num=first_hidden))
                                    else:
                                        if first_hidden == min_r + 1 and last_hidden == max_r - 1 and (max_r - min_r - 1) >= 500:
                                            ui.message(_("Rows {start} through {end} hidden").format(start=min_r + 1, end=max_r - 1))
                                        else:
                                            ui.message(_("Rows {start} through {end} hidden").format(start=first_hidden, end=last_hidden))
                                else:
                                    ui.message(_("Crossed heavily fragmented hidden rows"))
                                    
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
                                        ui.message(_("Column {col_letter} hidden").format(col_letter=col_num_to_letter(first_hidden)))
                                    else:
                                        if first_hidden == min_c + 1 and last_hidden == max_c - 1 and (max_c - min_c - 1) >= 500:
                                            ui.message(_("Columns {start} through {end} hidden").format(start=col_num_to_letter(min_c + 1), end=col_num_to_letter(max_c - 1)))
                                        else:
                                            ui.message(_("Columns {start} through {end} hidden").format(start=col_num_to_letter(first_hidden), end=col_num_to_letter(last_hidden)))
                                else:
                                    ui.message(_("Crossed heavily fragmented hidden columns"))
                                    
                                if current_col > _last_focused_col:
                                    skip_announced_left = True
                                else:
                                    skip_announced_right = True
    
                # Check right boundary

            # --- End of Data Radar (Empty Cell Tracker) ---
            # Architectural Intent: Acts as a radar to inform the user if there is no more data left 
            # in the direction they are traveling, preventing them from blindly arrowing through empty space.
            radar_mode = boa_config.get_feature_state("excel", "end_of_data_radar")
            if radar_mode and radar_mode != "off":
                if (row_changed or col_changed) and _last_focused_row is not None and _last_focused_col is not None:
                    is_empty = False
                    try:
                        val = active_cell.Value
                        text = active_cell.Text
                        if val is None or str(text).strip() == "":
                            is_empty = True
                    except Exception:
                        pass
                        
                    if is_empty:
                        sheet = excel.ActiveSheet
                        max_row = sheet.Rows.Count
                        max_col = sheet.Columns.Count
                        
                        def has_data(rng):
                            if radar_mode == "visible":
                                # Visible Math Engine: Filters out hidden rows via SpecialCells, 
                                # then subtracts CountBlank from Total Cells to detect any visible text/numbers
                                # while safely ignoring formulas that evaluate to "".
                                try:
                                    visible_rng = rng.SpecialCells(12) # 12 = xlCellTypeVisible
                                    for area in visible_rng.Areas:
                                        total = area.Cells.Count
                                        blanks = excel.WorksheetFunction.CountBlank(area)
                                        if total - blanks > 0:
                                            return True
                                    return False
                                except Exception:
                                    # If the entire range is hidden, SpecialCells throws an exception
                                    return False
                            else:
                                # Default CountA Engine
                                return excel.WorksheetFunction.CountA(rng) > 0
                        
                        try:
                            # Moving Down
                            if current_row > _last_focused_row and col_changed == False:
                                if current_row < max_row:
                                    rng = sheet.Range(sheet.Cells(current_row + 1, current_col), sheet.Cells(max_row, current_col))
                                    if not has_data(rng):
                                        ui.message(_("No more data below"))
                                        
                            # Moving Up
                            elif current_row < _last_focused_row and col_changed == False:
                                if current_row > 1:
                                    rng = sheet.Range(sheet.Cells(1, current_col), sheet.Cells(current_row - 1, current_col))
                                    if not has_data(rng):
                                        ui.message(_("No more data above"))
                                        
                            # Moving Right
                            elif current_col > _last_focused_col and row_changed == False:
                                if current_col < max_col:
                                    rng = sheet.Range(sheet.Cells(current_row, current_col + 1), sheet.Cells(current_row, max_col))
                                    if not has_data(rng):
                                        ui.message(_("No more data to the right"))
                                        
                            # Moving Left
                            elif current_col < _last_focused_col and row_changed == False:
                                if current_col > 1:
                                    rng = sheet.Range(sheet.Cells(current_row, 1), sheet.Cells(current_row, current_col - 1))
                                    if not has_data(rng):
                                        ui.message(_("No more data to the left"))
                        except Exception as e:
                            try:
                                import logHandler
                                logHandler.log.debug(f"BOA: End of data radar failed: {e}")
                            except Exception:
                                pass

            _last_focused_row = current_row
            _last_focused_col = current_col
            
            # Guided mode triggers on cell movement
            if boa_config.get_feature_state("excel", "auto_announce_first_block") == "guided":
                if row_changed or col_changed:
                    from .sheet_layout_analyzer import SheetLayoutAnalyzer
                    SheetLayoutAnalyzer.auto_announce_guided(excel)
                    
            if boa_config.get_feature_state("excel", "conditional_formatting"):
                if row_changed or col_changed:
                    from .conditional_formatting import ConditionalFormattingTracker
                    msg = ConditionalFormattingTracker.check_quick_format(excel)
                    if msg:
                        ui.message(msg)
        except Exception as e:
            try:
                logHandler.log.debug(f"BOA: Failed to check hidden skip: {e}")
            except Exception:
                pass
    @staticmethod
    def check_structural_changes(excel):
        """
        Monitors for changes in Excel's structural layout that do not natively fire events.
        ARCHITECTURAL INTENT: Actions like toggling Freeze Panes or Hiding/Unhiding a worksheet
        do not emit standard UIA/MSAA property change events. By caching the previous state and 
        polling it on focus return, we can detect if a structural change occurred while the user
        was interacting with the Ribbon or right-click menus, and announce it proactively.
        """
        global _last_freeze_panes_state, _last_visible_sheet_count, _last_total_sheet_count, _last_focused_sheet, _last_focused_wb, _last_structural_wb, _last_structural_sheet, _last_excel_hwnd
        import ui
        
        try:
            active_wb = excel.ActiveWorkbook
            if not active_wb:
                return
            current_wb_name = active_wb.Name
            current_hwnd = excel.Hwnd
            
            # Reset trackers if the user closed and reopened Excel (new Process Window Handle)
            if _last_excel_hwnd != current_hwnd:
                _last_structural_wb = None
                _last_structural_sheet = None
                _last_focused_wb = None
                _last_visible_sheet_count = None
                _last_total_sheet_count = None
                _last_freeze_panes_state = None
                _last_excel_hwnd = current_hwnd
            
            # Reset trackers if the user switched to a different workbook
            if _last_structural_wb != current_wb_name:
                _last_visible_sheet_count = None
                _last_total_sheet_count = None
                _last_freeze_panes_state = None
                _last_structural_wb = current_wb_name
                _last_structural_sheet = None
                
            # Check Freeze Panes
            # ActiveWindow.FreezePanes returns a boolean indicating if panes are frozen.
            active_win = excel.ActiveWindow
            if active_win:
                current_freeze = active_win.FreezePanes
                if _last_freeze_panes_state is not None and current_freeze != _last_freeze_panes_state:
                    if current_freeze:
                        ui.message(_("Panes frozen"))
                    else:
                        ui.message(_("Panes unfrozen"))
                _last_freeze_panes_state = current_freeze
            
            # Check sheet counts and hidden sheets
            try:
                sheets = active_wb.Sheets
                total_sheets = sheets.Count
                active_sheet = excel.ActiveSheet
                if not active_sheet:
                    return
                current_sheet = active_sheet.Name
                
                # Throttle execution: only execute heavy sheet checking loop if:
                # - The active sheet has changed
                # - The active workbook has changed
                # - The total sheet count has changed
                if (_last_structural_sheet == current_sheet and 
                    _last_structural_wb == current_wb_name and 
                    _last_total_sheet_count == total_sheets):
                    return
                
                sheet_was_deleted = False
                # Check for sheet deletion within the same workbook by iterating Sheets
                if (_last_structural_sheet is not None and 
                    _last_structural_wb == current_wb_name and 
                    _last_total_sheet_count is not None and 
                    total_sheets < _last_total_sheet_count):
                    found = False
                    for s in sheets:
                        try:
                            if s.Name == _last_structural_sheet:
                                found = True
                                break
                        except Exception:
                            pass
                    if not found:
                        sheet_was_deleted = True
                        ui.message(_("{sheet_name} deleted").format(sheet_name=_last_structural_sheet))
                
                # Check for skipped hidden sheets if navigated via Ctrl+PageDown
                # Only valid if the sheet count hasn't changed (meaning no deletion/insertion shifted indices)
                last_idx = None
                if (not sheet_was_deleted and 
                    _last_total_sheet_count == total_sheets and 
                    _last_focused_sheet is not None and 
                    _last_focused_wb == current_wb_name):
                    
                    # Find last index safely via loop to avoid failing comtypes key index lookups
                    for s in sheets:
                        try:
                            if s.Name == _last_focused_sheet:
                                last_idx = s.Index
                                break
                        except Exception:
                            pass
                            
                if last_idx is not None:
                    current_idx = active_sheet.Index
                    if abs(current_idx - last_idx) > 1:
                        min_s = min(last_idx, current_idx)
                        max_s = max(last_idx, current_idx)
                        for i in range(min_s + 1, max_s):
                            try:
                                sheet = sheets(i)
                                if sheet.Visible != -1:  # -1 is xlSheetVisible
                                    ui.message(_("{sheet} hidden").format(sheet=sheet.Name))
                            except Exception:
                                pass

                current_visible = 0
                for sheet in sheets:
                    try:
                        # Visible property returns -1 for visible, 0 for hidden, 2 for very hidden
                        if sheet.Visible == -1:
                            current_visible += 1
                    except Exception:
                        pass
                
                if _last_visible_sheet_count is not None and current_visible != _last_visible_sheet_count:
                    total_changed = _last_total_sheet_count is not None and total_sheets != _last_total_sheet_count
                    if current_visible < _last_visible_sheet_count:
                        if not sheet_was_deleted:
                            ui.message(_("Sheet hidden"))
                    else:
                        if not total_changed:
                            ui.message(_("Sheet unhidden"))
                
                _last_visible_sheet_count = current_visible
                _last_total_sheet_count = total_sheets
                _last_structural_sheet = current_sheet
                _last_structural_wb = current_wb_name
            except Exception:
                pass
                
        except Exception:
            pass
    @script(
        description="Hides the currently selected row.",
        category=_("BOA (Better Office Accessibility)")
    )
    def script_hideRow(self, gesture):
        """
        Intercepts Ctrl+9 to hide the selected row and announces the change.
        
        Architectural Intent & Considerations:
        Native Excel completely lacks auditory feedback when a row is hidden via shortcut. We intercept 
        the shortcut, forward it to the OS so Excel executes it natively, and then asynchronously query 
        the COM model to verify the change actually occurred before speaking it.
        """
        self._execute_and_verify_visibility_change(gesture, "row", True)

    @script(
        description="Unhides the currently selected row.",
        category=_("BOA (Better Office Accessibility)")
    )
    def script_unhideRow(self, gesture):
        """
        Intercepts Ctrl+Shift+9 to unhide the selected row and announces the change.
        
        Architectural Intent & Considerations:
        Native Excel completely lacks auditory feedback when a row is unhidden via shortcut. We intercept 
        the shortcut, forward it to the OS, and asynchronously query the COM model to verify the change.
        """
        self._execute_and_verify_visibility_change(gesture, "row", False)

    @script(
        description="Hides the currently selected column.",
        category=_("BOA (Better Office Accessibility)")
    )
    def script_hideColumn(self, gesture):
        """
        Intercepts Ctrl+0 to hide the selected column and announces the change.
        
        Architectural Intent & Considerations:
        Native Excel completely lacks auditory feedback when a column is hidden via shortcut. We intercept 
        the shortcut, forward it to the OS, and asynchronously query the COM model to verify the change.
        """
        self._execute_and_verify_visibility_change(gesture, "column", True)

    @script(
        description="Unhides the currently selected column.",
        category=_("BOA (Better Office Accessibility)")
    )
    def script_unhideColumn(self, gesture):
        """
        Intercepts Ctrl+Shift+0 to unhide the selected column and announces the change.
        
        Architectural Intent & Considerations:
        Native Excel completely lacks auditory feedback when a column is unhidden via shortcut. We intercept 
        the shortcut, forward it to the OS, and asynchronously query the COM model to verify the change.
        """
        self._execute_and_verify_visibility_change(gesture, "column", False)

    @script(
        description="Unhides the currently selected column natively via COM (Fallback for when Windows blocks Ctrl+Shift+0).",
        category=_("BOA (Better Office Accessibility)")
    )
    def script_unhideColumnFallback(self, gesture):
        """
        Forces the column to unhide using Excel COM.
        
        Architectural Intent & Considerations:
        Windows 10/11 frequently hijacks Ctrl+Shift+0 for changing keyboard layouts, completely breaking 
        Excel's native unhide shortcut. This fallback bypasses the OS keyboard hook entirely and forces 
        the column visible via direct COM manipulation (`force_com=True`).
        """
        self._execute_and_verify_visibility_change(gesture, "column", False, force_com=True)

    def _execute_and_verify_visibility_change(self, gesture, element_type, is_hiding, force_com=False):
        """
        Fetches the initial visibility state, passes the native keystroke to Excel (or uses COM if forced),
        waits briefly, and then checks if the COM model state actually changed.
        ARCHITECTURAL INTENT: Excel's keyboard shortcuts for hiding rows/columns (Ctrl+9, Ctrl+0) 
        do not provide auditory feedback natively. Furthermore, Windows sometimes hijacks Ctrl+Shift+0.
        This wrapper verifies the state change asynchronously via COM and speaks the result.
        """
        initial_state = None
        try:
            excel = _get_excel_app()
            if excel:
                sel = excel.Selection
                if element_type == "row":
                    initial_state = sel.EntireRow.Hidden
                    if force_com:
                        sel.EntireRow.Hidden = is_hiding
                elif element_type == "column":
                    initial_state = sel.EntireColumn.Hidden
                    if force_com:
                        sel.EntireColumn.Hidden = is_hiding
        except Exception:
            pass

        from appModules.boa_enhancements import boa_config
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
        ARCHITECTURAL INTENT: By checking the state after a 200ms delay, we give Excel's 
        internal engine time to process the keystroke and update the COM model, ensuring 
        accurate feedback without blocking NVDA's single-threaded core.
        """
        import ui
        import logHandler
        
        try:
            excel = _get_excel_app()
            if excel:
                sel = excel.Selection
                
                is_hidden = None
                address_str = ""
                if element_type == "row":
                    is_hidden = sel.EntireRow.Hidden
                    start_row = sel.Row
                    end_row = sel.Row + sel.Rows.Count - 1
                    if start_row == end_row:
                        address_str = _("Row {num}").format(num=start_row)
                    else:
                        address_str = _("Rows {start} through {end}").format(start=start_row, end=end_row)
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
                        address_str = _("Column {col}").format(col=start_letter)
                    else:
                        end_letter = col_num_to_letter(end_col)
                        address_str = _("Columns {start} through {end}").format(start=start_letter, end=end_letter)
                    
                # Only announce if the state successfully changed to avoid false positives
                if is_hidden is not None and is_hidden != initial_state:
                    state_str = _("hidden") if is_hidden else _("unhidden")
                    ui.message(_("{address} {state}").format(address=address_str, state=state_str))
        except Exception as e:
            logHandler.log.debugWarning(f"BOA: Failed to verify {element_type} visibility change. {e}")

    __gestures = {
        "kb:control+9": "hideRow",
        "kb:control+shift+9": "unhideRow",
        "kb:control+0": "hideColumn",
        "kb:control+shift+0": "unhideColumn",
        "kb:NVDA+control+shift+0": "unhideColumnFallback",
    }
