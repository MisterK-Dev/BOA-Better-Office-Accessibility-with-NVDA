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
    if boa_config.get_feature_state("excel", "unselect_tracking") or boa_config.get_feature_state("excel", "hidden_row_skip"):
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
            
    # Add future commands here (e.g. 'r' for rename)
    return False
