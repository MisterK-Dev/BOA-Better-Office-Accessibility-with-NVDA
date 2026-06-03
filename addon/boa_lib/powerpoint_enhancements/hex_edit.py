import NVDAObjects.UIA
import controlTypes
import UIAHandler
from logHandler import log
from boa_lib.safe_rich_edit import SafeRichEdit

class PowerPointHexEdit(SafeRichEdit):
    """
    Specific override for the PowerPoint Color Hex Edit box.
    
    Architectural Intent:
    By default, this edit box does not expose an accessible name to assistive technologies, 
    making it read as "edit blank". We inherit from `SafeRichEdit` to safely hook into the 
    NVDA object lifecycle. We identify the control using its fixed windowControlID (1637) 
    and explicitly return the name "Hex Color", ensuring screen reader users have immediate 
    context when they focus on this field.
    """
    def _get_name(self):
        """
        Overrides the native NVDA name resolution property getter (`_get_name`).
        
        Architectural Intent:
        Intercepts the NVDA API call that asks "What is the name of this object?".
        Checks if the underlying Win32 control ID matches the known Hex input field (1637).
        If it matches, it short-circuits the standard UIA/Win32 accessible name retrieval
        and returns the hardcoded string "Hex Color". Otherwise, it falls back to the native 
        behavior provided by `SafeRichEdit`.
        """
        # Check if the NVDA object has a 'windowControlID' attribute (standard for Win32 windows).
        # 1637 is the standard Win32 control ID for the Hex field in the Office color picker.
        if hasattr(self, 'windowControlID') and self.windowControlID == 1637:
            # Return our custom accessible name, overriding the missing default name.
            return "Hex Color"
        
        # If it's not the Hex field (e.g. another RichEdit control), delegate back to the parent class.
        return super(PowerPointHexEdit, self).name
