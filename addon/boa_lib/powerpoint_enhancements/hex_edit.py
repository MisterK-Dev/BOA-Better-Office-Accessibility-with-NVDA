import NVDAObjects.UIA
from logHandler import log
import NVDAObjects.IAccessible
import controlTypes
import UIAHandler

from excel_enhancement import SafeRichEdit

# --- Swappable PowerPoint UIA Automation IDs ---
# These IDs correspond to the UIAAutomationId property of the edit fields in the Custom Color dialog.
# If NVDA Object Navigator shows different UIAAutomationIds for newer Office versions, update these variables.
PPT_RED_ID = "101"
PPT_GREEN_ID = "102"
PPT_BLUE_ID = "103"

class PowerPointHexEdit(SafeRichEdit):
    """
    Specific override for the PowerPoint Color Hex Edit box.
    By default, this edit box does not expose an accessible name.
    We identify it using its fixed windowControlID (1637) and explicitly return "Hex Color".
    """
    def _get_name(self):
        # 1637 is the standard Win32 control ID for the Hex field in the Office color picker.
        if hasattr(self, 'windowControlID') and self.windowControlID == 1637:
            return "Hex Color"
        return super(PowerPointHexEdit, self).name
