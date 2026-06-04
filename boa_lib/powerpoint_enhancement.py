import NVDAObjects.UIA
from logHandler import log
import NVDAObjects.IAccessible
import controlTypes
import UIAHandler

from excel_enhancement import SafeRichEdit

# --- Swappable PowerPoint UIA Automation IDs ---
# If NVDA Object Navigator shows different UIAAutomationIds for Office 2024, change these variables.
PPT_RED_ID = "101"
PPT_GREEN_ID = "102"
PPT_BLUE_ID = "103"

class PowerPointHexEdit(SafeRichEdit):
    """
    Specific override for PowerPoint Color Hex Edit box.
    """
    def _get_name(self):
        if hasattr(self, 'windowControlID') and self.windowControlID == 1637:
            return "Hex Color"
        return super(PowerPointHexEdit, self).name

class PowerPointRGBEdit(NVDAObjects.UIA.UIA):
    """
    STRICT UIA Override for PowerPoint Custom Colors RGB edit boxes.
    """
    def _get_name(self):
        # UIA uses UIAAutomationId, not windowControlID
        auto_id = getattr(self, 'UIAAutomationId', '')
        if auto_id == PPT_RED_ID:
            return "Red"
        elif auto_id == PPT_GREEN_ID:
            return "Green"
        elif auto_id == PPT_BLUE_ID:
            return "Blue"
        
        return super(PowerPointRGBEdit, self).name

class PowerPointStandardColorGrid(NVDAObjects.window.Window):
    """
    Intercepts arrow keys when the user is navigating the 'Standard' color hexagon in PowerPoint.
    Since the hexagon itself is totally inaccessible, we silently read the RGB Hex value from the hidden 'Custom' tab edits!
    """
    def _read_hidden_hex_color(self):
        import ctypes
        import winUser
        import speech
        
        hwnd = self.windowHandle
        children = []
        def callback(child, param):
            children.append(child)
            return True
        EnumChildProc = ctypes.WINFUNCTYPE(ctypes.c_bool, ctypes.c_int, ctypes.c_void_p)
        ctypes.windll.user32.EnumChildWindows(hwnd, EnumChildProc(callback), 0)
        
        hex_val = None
        for child in children:
            buf = ctypes.create_unicode_buffer(256)
            ctypes.windll.user32.GetClassNameW(child, buf, 256)
            cls = buf.value
            if "Edit" in cls or "RichEdit" in cls:
                length = ctypes.windll.user32.SendMessageW(child, winUser.WM_GETTEXTLENGTH, 0, 0)
                if 0 < length < 100:
                    text_buf = ctypes.create_unicode_buffer(length + 1)
                    ctypes.windll.user32.SendMessageW(child, winUser.WM_GETTEXT, length + 1, text_buf)
                    text = text_buf.value.strip()
                    if text.startswith("#") and len(text) == 7:
                        hex_val = text
                        break
        if hex_val:
            speech.speakMessage(f"Color {hex_val}")

    def script_arrowUp(self, gesture):
        gesture.send()
        import core
        core.callLater(50, self._read_hidden_hex_color)

    def script_arrowDown(self, gesture):
        gesture.send()
        import core
        core.callLater(50, self._read_hidden_hex_color)

    def script_arrowLeft(self, gesture):
        gesture.send()
        import core
        core.callLater(50, self._read_hidden_hex_color)

    def script_arrowRight(self, gesture):
        gesture.send()
        import core
        core.callLater(50, self._read_hidden_hex_color)

    __gestures = {
        "kb:upArrow": "arrowUp",
        "kb:downArrow": "arrowDown",
        "kb:leftArrow": "arrowLeft",
        "kb:rightArrow": "arrowRight",
    }
