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

class PowerPointRGBEdit(NVDAObjects.UIA.UIA):
    """
    STRICT UIA Override for PowerPoint Custom Colors RGB edit boxes.
    The RGB text boxes use UI Automation. Since their accessible names are often missing or confusing,
    we identify them by their UIAAutomationId and explicitly provide "Red", "Green", or "Blue" to NVDA.
    """
    def _get_name(self):
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

class PowerPointStandardColorGrid(NVDAObjects.window.Window):
    """
    Intercepts arrow keys when the user is navigating the 'Standard' color hexagon grid in PowerPoint.
    Since the hexagon itself is visually drawn and completely inaccessible via standard APIs, 
    we silently search for the hidden 'Custom' tab edit controls in the background, extract the Hex value, 
    and read it out loud so the user knows exactly what color they are currently focused on.
    """
    def _read_hidden_hex_color(self):
        import ctypes
        import winUser
        import speech
        
        hwnd = self.windowHandle
        children = []
        
        # Callback to collect all child windows of the current dialog.
        def callback(child, param):
            children.append(child)
            return True
            
        EnumChildProc = ctypes.WINFUNCTYPE(ctypes.c_bool, ctypes.c_int, ctypes.c_void_p)
        ctypes.windll.user32.EnumChildWindows(hwnd, EnumChildProc(callback), 0)
        
        hex_val = None
        # Iterate through all discovered child windows to find the Edit control containing the Hex string.
        for child in children:
            buf = ctypes.create_unicode_buffer(256)
            ctypes.windll.user32.GetClassNameW(child, buf, 256)
            cls = buf.value
            
            # The Hex field is an Edit or RichEdit control.
            if "Edit" in cls or "RichEdit" in cls:
                # Query the text length first to prevent buffer overflows from maliciously crafted files.
                length = ctypes.windll.user32.SendMessageW(child, winUser.WM_GETTEXTLENGTH, 0, 0)
                if 0 < length < 100:
                    text_buf = ctypes.create_unicode_buffer(length + 1)
                    ctypes.windll.user32.SendMessageW(child, winUser.WM_GETTEXT, length + 1, text_buf)
                    text = text_buf.value.strip()
                    
                    # Verify if the text exactly matches a standard Hex format (e.g., "#FFFFFF").
                    if text.startswith("#") and len(text) == 7:
                        hex_val = text
                        break
                        
        if hex_val:
            speech.speakMessage(f"Color {hex_val}")

    # The following scripts intercept standard navigation keys.
    # They first pass the original key press to PowerPoint (gesture.send()),
    # and then schedule a slight delay (50ms) to read the new hidden hex color,
    # giving the PowerPoint UI enough time to internally update the hidden edit boxes.

    from scriptHandler import script

    @script(
        description="Reads the hidden hex color while navigating UP in the PowerPoint Standard color grid.",
        category="Better Office Accessibility"
    )
    def script_arrowUp(self, gesture):
        gesture.send()
        import core
        core.callLater(50, self._read_hidden_hex_color)

    @script(
        description="Reads the hidden hex color while navigating DOWN in the PowerPoint Standard color grid.",
        category="Better Office Accessibility"
    )
    def script_arrowDown(self, gesture):
        gesture.send()
        import core
        core.callLater(50, self._read_hidden_hex_color)

    @script(
        description="Reads the hidden hex color while navigating LEFT in the PowerPoint Standard color grid.",
        category="Better Office Accessibility"
    )
    def script_arrowLeft(self, gesture):
        gesture.send()
        import core
        core.callLater(50, self._read_hidden_hex_color)

    @script(
        description="Reads the hidden hex color while navigating RIGHT in the PowerPoint Standard color grid.",
        category="Better Office Accessibility"
    )
    def script_arrowRight(self, gesture):
        gesture.send()
        import core
        core.callLater(50, self._read_hidden_hex_color)

    # Bind the interceptor scripts to the standard arrow keys.
    __gestures = {
        "kb:upArrow": "arrowUp",
        "kb:downArrow": "arrowDown",
        "kb:leftArrow": "arrowLeft",
        "kb:rightArrow": "arrowRight",
    }
