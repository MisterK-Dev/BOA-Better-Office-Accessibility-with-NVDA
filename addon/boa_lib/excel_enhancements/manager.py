import boa_config
from .cell_navigation_tracker import CellNavigationTracker
from .quick_sheet_mover import QuickSheetMover
from .bulk_sheet_organizer import BulkSheetOrganizer
from .accessible_rename import ExcelSheetRenameEdit
from logHandler import log

def inject_excel_grid_classes(clsList):
    """
    Injects modular Excel grid enhancements based on user config.
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
            tones.beep(600, 50) # Success beep
            # To open the dialog, we just instantiate it if we have the active obj
            import wx
            from .bulk_sheet_organizer import ExcelBulkSheetOrganizerDialog
            import gui
            gui.mainFrame.prePopup()
            d = ExcelBulkSheetOrganizerDialog(gui.mainFrame, obj)
            d.Show()
            d.Raise()
            gui.mainFrame.postPopup()
            return True
            
    # Add future commands here (e.g. 'r' for rename)
    return False
