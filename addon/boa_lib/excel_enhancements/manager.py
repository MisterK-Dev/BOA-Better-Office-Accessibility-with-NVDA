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

from boa_lib import boa_config
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


def handle_prefix_command(command_key, obj):
    """
    Routes a secondary key (pressed after NVDA+E) to the appropriate feature.
    ARCHITECTURAL INTENT: Implementing a prefix key system (NVDA+E, then <key>) conserves
    global hotkeys and logically groups Excel-specific commands under one namespace.
    """
    if command_key == 'x':
        # Feature check ensures we only execute if the user has enabled grid_mover
        if boa_config.get_feature_state("excel", "grid_mover"):
            import tones
            # Provide auditory feedback that the prefix command was accepted
            tones.beep(600, 50)
            from .bulk_sheet_organizer import BulkSheetOrganizer
            # Launch the organizer dialog, passing the active Excel object
            BulkSheetOrganizer.launch_dialog(obj)
            return True
            
    if command_key == 'l':
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
                    # Fallback Consideration: If GetActiveObject fails (e.g. if Excel is busy), 
                    # we must manually dig for the EXCEL7 window class handle from the NVDA object to force a connection.
                    hwnd7 = obj.windowHandle if getattr(obj, "windowClassName", "") == "EXCEL7" else None
                    if hwnd7:
                        import ctypes
                        # Dynamically load the oleacc library.
                        oleacc = ctypes.windll.oleacc if hasattr(ctypes.windll, 'oleacc') else ctypes.windll.user32.oleacc
                        ptr = ctypes.POINTER(comtypes.automation.IDispatch)()
                        # Use AccessibleObjectFromWindow (-16 is OBJID_NATIVEOM) to force a back-door COM connection directly from the HWND.
                        res = oleacc.AccessibleObjectFromWindow(hwnd7, -16, ctypes.byref(comtypes.automation.IDispatch._iid_), ctypes.byref(ptr))
                        if res == 0 and ptr:
                            # Safely cast the raw COM pointer back into a usable Python Excel Application object.
                            excel = comtypes.client.dynamic.Dispatch(ptr).Application
                            SheetLayoutAnalyzer.announce_layout(excel)
                except Exception:
                    pass
            return True
            
    if command_key == 'j':
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
                    # Fallback Consideration: If GetActiveObject fails (e.g. if Excel is busy), 
                    # we must manually dig for the EXCEL7 window class handle from the NVDA object to force a connection.
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

    if command_key == 'f':
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
                    # Fallback Consideration: If GetActiveObject fails (e.g. if Excel is busy), 
                    # we must manually dig for the EXCEL7 window class handle from the NVDA object to force a connection.
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

    # --- Cell Monitor Commands ---
    if command_key == 'backspace':
        if boa_config.get_feature_state("excel", "cell_monitor"):
            from .cell_monitor import CellMonitorManager
            CellMonitorManager.clear_all(obj)
            return True
        
    if command_key == 'm':
        if boa_config.get_feature_state("excel", "cell_monitor"):
            from .cell_monitor import CellMonitorManager
            CellMonitorManager.toggle_monitor(obj)
            return True
        
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

    # Add future commands here (e.g. 'r' for rename)
    return False
