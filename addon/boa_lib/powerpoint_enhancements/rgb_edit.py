import NVDAObjects.UIA
import controlTypes
import UIAHandler
from logHandler import log

# --- Swappable PowerPoint UIA Automation IDs ---
# These IDs correspond to the UIAAutomationId property of the edit fields in the Custom Color dialog.
# If NVDA Object Navigator shows different UIAAutomationIds for newer Office versions, update these variables.
PPT_RED_ID = "101"
PPT_GREEN_ID = "102"
PPT_BLUE_ID = "103"

class PowerPointRGBEdit(NVDAObjects.UIA.UIA):
    """
    STRICT UIA Override for PowerPoint Custom Colors RGB edit boxes.
    The RGB text boxes use UI Automation. Since their accessible names are often missing or confusing,
    we identify them by their UIAAutomationId and explicitly provide "Red", "Green", or "Blue" to NVDA.
    """
    def _get_name(self):
        """
        Overrides the native name resolution for RGB edits.
        Checks the UIAAutomationId and returns 'Red', 'Green', or 'Blue' respectively.
        """
        # UIA uses UIAAutomationId, not windowControlID, so we fetch it safely.
        auto_id = getattr(self, 'UIAAutomationId', '')
        if auto_id == PPT_RED_ID:
            return "Red"
        elif auto_id == PPT_GREEN_ID:
            return "Green"
        elif auto_id == PPT_BLUE_ID:
            return "Blue"
        
        # Fallback to the default name if it doesn't match our specific IDs.
        return super(PowerPointRGBEdit, self).name
