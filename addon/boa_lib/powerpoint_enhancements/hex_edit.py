import NVDAObjects.UIA
import controlTypes
import UIAHandler
from logHandler import log
from boa_lib.safe_rich_edit import SafeRichEdit

class PowerPointHexEdit(SafeRichEdit):
    """
    Specific override for the PowerPoint Color Hex Edit box.
    By default, this edit box does not expose an accessible name.
    We identify it using its fixed windowControlID (1637) and explicitly return "Hex Color".
    """
    def _get_name(self):
        """
        Overrides the native name resolution.
        Checks if the control ID matches the known Hex input field (1637).
        If it matches, returns the hardcoded string "Hex Color", otherwise falls back to native behavior.
        """
        # 1637 is the standard Win32 control ID for the Hex field in the Office color picker.
        if hasattr(self, 'windowControlID') and self.windowControlID == 1637:
            return "Hex Color"
        return super(PowerPointHexEdit, self).name
