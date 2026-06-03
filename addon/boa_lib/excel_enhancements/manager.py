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
    Called by office_bridge when an 'EXCEL7' or similar grid class receives focus.
    Each feature is completely isolated in its own file and can be toggled independently.
    """
    # Bulk Sheet Organizer (GUI dialog)
    if boa_config.get_feature_state("excel", "grid_mover"): # Currently sharing the same toggle
        clsList.insert(0, BulkSheetOrganizer)
        
    # Quick Sheet Movers (NVDA+Shift+Arrows)
    if boa_config.get_feature_state("excel", "grid_mover"):
        clsList.insert(0, QuickSheetMover)
        
    # Cell Navigation Tracker (Gap/Boundary detection)
    if boa_config.get_feature_state("excel", "grid_mover"):
        clsList.insert(0, CellNavigationTracker)

def inject_excel_rename_class(clsList):
    """
    Injects the accessible Sheet Rename field override.
    Called by office_bridge when the 'EXCEL=' class receives focus.
    """
    if boa_config.get_feature_state("excel", "sheet_rename"):
        clsList.insert(0, ExcelSheetRenameEdit)

def handle_prefix_command(command_key, obj):
    """
    Routes a secondary key (pressed after NVDA+E) to the appropriate feature.
    Returns True if handled, False if invalid.
    """
    if command_key == 'x':
        if boa_config.get_feature_state("excel", "grid_mover"):
            import tones
            tones.beep(600, 50)
            from .bulk_sheet_organizer import BulkSheetOrganizer
            BulkSheetOrganizer.launch_dialog(obj)
            return True
            
    # Add future commands here (e.g. 'r' for rename)
    return False
