# -*- coding: UTF-8 -*-
# Copyright (C) 2026 KIRAN G T {MisterK} and Antigravity 2
# This file is covered by the GNU General Public License (GPL), version 2.
# See the file COPYING.txt for more details.

import addonHandler
addonHandler.initTranslation()

"""
Excel Enhancement Manager

This module acts as the central dispatcher for all Excel-specific accessibility enhancements.
Architectural Intent:
Instead of dumping all Excel logic into a single monolithic file (like `globalPlugins`), 
this manager cleanly injects independent, feature-specific classes into the NVDA object resolution chain.
This strictly adheres to `rules.mdc` modular isolation constraints.
"""

from appModules.boa_enhancements import boa_config
from .cell_navigation_tracker import CellNavigationTracker
from .quick_sheet_mover import QuickSheetMover
from .bulk_sheet_organizer import BulkSheetOrganizer
from .accessible_rename import ExcelSheetRenameEdit
from logHandler import log

def inject_excel_grid_classes(clsList):
    """
    Injects modular Excel grid enhancements based on user config.
    ARCHITECTURAL INTENT: Instead of hardcoding feature classes directly into the NVDA event loop,
    this function dynamically injects independent mixin classes into the NVDA object's inheritance
    chain (clsList) at runtime. This allows features to be independently toggled without polluting
    the base Excel object class.
    """
    # Bulk Sheet Organizer (GUI dialog) for reordering sheets via an accessible menu
    if boa_config.get_feature_state("excel", "grid_mover"):
        clsList.insert(0, BulkSheetOrganizer)
        
    # Quick Sheet Movers (NVDA+Shift+Arrows) for fast keyboard-driven sheet reordering
    if boa_config.get_feature_state("excel", "grid_mover"):
        clsList.insert(0, QuickSheetMover)
        
    # Cell Navigation Tracker (Gap/Boundary detection) for spatial awareness
    if (boa_config.get_feature_state("excel", "unselect_tracking") or 
        boa_config.get_feature_state("excel", "hidden_row_skip") or 
        boa_config.get_feature_state("excel", "auto_announce_first_block") in ["one_time", "guided"] or
        boa_config.get_feature_state("excel", "conditional_formatting")):
        clsList.insert(0, CellNavigationTracker)


def inject_excel_rename_class(clsList):
    """
    Injects the accessible Sheet Rename field override.
    ARCHITECTURAL INTENT: Excel's native "Rename Sheet" edit field is poorly exposed to UIA.
    By injecting ExcelSheetRenameEdit into the class list for the 'EXCEL=' window class,
    we intercept the focus event and seamlessly substitute our own fully accessible wx.Dialog.
    """
    if boa_config.get_feature_state("excel", "sheet_rename"):
        clsList.insert(0, ExcelSheetRenameEdit)


def inject_evaluate_formula_dialog_class(clsList):
    """
    Injects the accessible Evaluate Formula evaluation hook into Buttons of legacy Excel dialogs.
    ARCHITECTURAL INTENT: Tied strictly to the formula_auditing config. It allows tracking 
    of dynamic formula evaluation text changes without completely overriding button logic.
    """
    if boa_config.get_feature_state("excel", "formula_auditing_announcements"):
        from .formula_auditor import ExcelLegacyDialogButtonMixin
        clsList.insert(0, ExcelLegacyDialogButtonMixin)

def show_bulk_sheet_organizer(obj):
    if boa_config.get_feature_state("excel", "grid_mover"):
        import tones
        tones.beep(600, 50)
        from .bulk_sheet_organizer import BulkSheetOrganizer
        BulkSheetOrganizer.launch_dialog(obj)
        return True
    return False

def announce_sheet_layout(obj):
    if boa_config.get_feature_state("excel", "sheet_layout_analyzer"):
        import tones
        tones.beep(600, 50)
        from .sheet_layout_analyzer import SheetLayoutAnalyzer
        try:
            import comtypes.client
            excel = comtypes.client.GetActiveObject("Excel.Application")
            if excel:
                SheetLayoutAnalyzer.announce_layout(excel)
        except Exception:
            try:
                hwnd7 = obj.windowHandle if getattr(obj, "windowClassName", "") == "EXCEL7" else None
                if hwnd7:
                    import ctypes
                    oleacc = ctypes.windll.oleacc if hasattr(ctypes.windll, 'oleacc') else ctypes.windll.user32.oleacc
                    ptr = ctypes.POINTER(comtypes.automation.IDispatch)()
                    res = oleacc.AccessibleObjectFromWindow(hwnd7, -16, ctypes.byref(comtypes.automation.IDispatch._iid_), ctypes.byref(ptr))
                    if res == 0 and ptr:
                        excel = comtypes.client.dynamic.Dispatch(ptr).Application
                        SheetLayoutAnalyzer.announce_layout(excel)
            except Exception:
                pass
        return True
    return False

def jump_to_nearest_block(obj):
    if boa_config.get_feature_state("excel", "sheet_layout_analyzer"):
        import tones
        tones.beep(600, 50)
        from .sheet_layout_analyzer import SheetLayoutAnalyzer
        try:
            import comtypes.client
            excel = comtypes.client.GetActiveObject("Excel.Application")
            if excel:
                SheetLayoutAnalyzer.jump_to_nearest_block(excel)
        except Exception:
            try:
                hwnd7 = obj.windowHandle if getattr(obj, "windowClassName", "") == "EXCEL7" else None
                if hwnd7:
                    import ctypes
                    oleacc = ctypes.windll.oleacc if hasattr(ctypes.windll, 'oleacc') else ctypes.windll.user32.oleacc
                    ptr = ctypes.POINTER(comtypes.automation.IDispatch)()
                    res = oleacc.AccessibleObjectFromWindow(hwnd7, -16, ctypes.byref(comtypes.automation.IDispatch._iid_), ctypes.byref(ptr))
                    if res == 0 and ptr:
                        excel = comtypes.client.dynamic.Dispatch(ptr).Application
                        SheetLayoutAnalyzer.jump_to_nearest_block(excel)
            except Exception:
                pass
        return True
    return False

def announce_conditional_formatting(obj):
    if boa_config.get_feature_state("excel", "conditional_formatting"):
        import tones
        tones.beep(600, 50)
        from .conditional_formatting import ConditionalFormattingTracker
        try:
            import comtypes.client
            excel = comtypes.client.GetActiveObject("Excel.Application")
            if excel:
                ConditionalFormattingTracker.announce_deep_dive(excel)
        except Exception:
            try:
                hwnd7 = obj.windowHandle if getattr(obj, "windowClassName", "") == "EXCEL7" else None
                if hwnd7:
                    import ctypes
                    oleacc = ctypes.windll.oleacc if hasattr(ctypes.windll, 'oleacc') else ctypes.windll.user32.oleacc
                    ptr = ctypes.POINTER(comtypes.automation.IDispatch)()
                    res = oleacc.AccessibleObjectFromWindow(hwnd7, -16, ctypes.byref(comtypes.automation.IDispatch._iid_), ctypes.byref(ptr))
                    if res == 0 and ptr:
                        excel = comtypes.client.dynamic.Dispatch(ptr).Application
                        ConditionalFormattingTracker.announce_deep_dive(excel)
            except Exception:
                pass
        return True
    return False

def toggle_cell_monitor(obj):
    if boa_config.get_feature_state("excel", "cell_monitor"):
        from .cell_monitor import CellMonitorManager
        CellMonitorManager.toggle_monitor(obj)
        return True
    return False

def clear_all_cell_monitors(obj):
    if boa_config.get_feature_state("excel", "cell_monitor"):
        from .cell_monitor import CellMonitorManager
        CellMonitorManager.clear_all(obj)
        return True
    return False

def handle_prefix_command(command_key, obj):
    """
    Routes a secondary key (pressed after NVDA+E) to the appropriate feature.
    ARCHITECTURAL INTENT: Implementing a prefix key system (NVDA+E, then <key>) conserves
    global hotkeys and logically groups Excel-specific commands under one namespace.
    """
    if command_key == 'x':
        return show_bulk_sheet_organizer(obj)
            
    if command_key == 'f2':
        import scriptHandler
        from . import cell_editor
        if scriptHandler.getLastScriptRepeatCount() == 0:
            cell_editor.speak_formula()
            return "keep_alive"
        else:
            cell_editor.open_power_editor()
            return True
            
    if command_key == 'l':
        return announce_sheet_layout(obj)
            
    if command_key == 'j':
        return jump_to_nearest_block(obj)

    if command_key == 'f':
        return announce_conditional_formatting(obj)

    if command_key == 'shift+p':
        from .formula_auditor import show_precedents_dialog
        show_precedents_dialog(obj)
        return True

    if command_key == 'shift+d':
        from .formula_auditor import show_dependents_dialog
        show_dependents_dialog(obj)
        return True

    # --- Cell Monitor Commands ---
    if command_key == 'backspace':
        return clear_all_cell_monitors(obj)
        
    if command_key == 'm':
        return toggle_cell_monitor(obj)
        
    # Check for slots 1-9
    if len(command_key) == 1 and command_key.isdigit() and command_key != '0':
        if boa_config.get_feature_state("excel", "cell_monitor"):
            from .cell_monitor import CellMonitorManager
            CellMonitorManager.read_slot(command_key, obj)
            return True
        
    # Check for assigning slots (shift+1 to shift+9)
    if command_key.startswith("shift+") and len(command_key) == 7 and command_key[-1].isdigit() and command_key[-1] != '0':
        if boa_config.get_feature_state("excel", "cell_monitor"):
            from .cell_monitor import CellMonitorManager
            CellMonitorManager.assign_slot(command_key[-1], obj)
            return True
    # Check for jumping to slots (alt+1 to alt+9)
    if command_key.startswith("alt+") and len(command_key) == 5 and command_key[-1].isdigit() and command_key[-1] != '0':
        if boa_config.get_feature_state("excel", "cell_monitor"):
            from .cell_monitor import CellMonitorManager
            CellMonitorManager.jump_to_slot(command_key[-1], obj)
            return True

    # Jump Back
    if command_key == '\\':
        if boa_config.get_feature_state("excel", "cell_monitor"):
            from .cell_monitor import CellMonitorManager
            CellMonitorManager.jump_back(obj)
            return True

    # Open Monitor Dialog
    if command_key == 'alt+m':
        if boa_config.get_feature_state("excel", "cell_monitor"):
            from .cell_monitor import CellMonitorManager
            CellMonitorManager.open_monitor_dialog(obj)
            return True

    return False
