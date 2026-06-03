import ctypes
import wx
import winUser
import ui
import queueHandler
from scriptHandler import script
from logHandler import log

class PowerPointStandardColorGrid(object):
    """
    Intercepts arrow keys when the user is navigating the 'Standard' color hexagon grid in PowerPoint.
    Since the hexagon itself is visually drawn and completely inaccessible via standard APIs, 
    we silently search for the hidden 'Custom' tab edit controls in the background, extract the Hex value, 
    and read it out loud so the user knows exactly what color they are currently focused on.
    """
    
    # Cache the HWND of the hex edit box to prevent expensive recursive searches on every keystroke
    _cached_hex_hwnd = None

    def _read_hidden_hex_color(self):
        """
        Scans child windows to find the hidden Hex edit box.
        Reads the text and announces the Hex color cleanly using ui.message.
        Uses a cached HWND if available to minimize CTypes API overhead.
        """
        # Fast path: Check if cached handle is still valid and has a hex string
        if self._cached_hex_hwnd and winUser.isWindow(self._cached_hex_hwnd):
            text = self._get_text_from_hwnd(self._cached_hex_hwnd)
            if text and text.startswith("#") and len(text) == 7:
                ui.message(f"Color {text}")
                return

        # Slow path: Enum child windows
        hwnd = self.windowHandle
        children = []
        
        # Callback to collect all child windows of the current dialog.
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
                text = self._get_text_from_hwnd(child)
                if text and text.startswith("#") and len(text) == 7:
                    hex_val = text
                    self._cached_hex_hwnd = child
                    break
                        
        if hex_val:
            ui.message(f"Color {hex_val}")

    def _get_text_from_hwnd(self, hwnd):
        """
        Helper method to safely extract text from a Win32 HWND.
        """
        length = ctypes.windll.user32.SendMessageW(hwnd, winUser.WM_GETTEXTLENGTH, 0, 0)
        if 0 < length < 100:
            text_buf = ctypes.create_unicode_buffer(length + 1)
            ctypes.windll.user32.SendMessageW(hwnd, winUser.WM_GETTEXT, length + 1, text_buf)
            return text_buf.value.strip()
        return None

    def _trigger_delayed_hex_read(self):
        """
        Schedules a thread-safe UI update read.
        """
        # Using wx.CallLater directly as it is safe and doesn't violate single-threaded constraints.
        wx.CallLater(50, queueHandler.queueFunction, queueHandler.eventQueue, self._read_hidden_hex_color)

    @script(
        description="Reads the hidden hex color while navigating UP in the PowerPoint Standard color grid.",
        category="Better Office Accessibility"
    )
    def script_arrowUp(self, gesture):
        gesture.send()
        self._trigger_delayed_hex_read()

    @script(
        description="Reads the hidden hex color while navigating DOWN in the PowerPoint Standard color grid.",
        category="Better Office Accessibility"
    )
    def script_arrowDown(self, gesture):
        gesture.send()
        self._trigger_delayed_hex_read()

    @script(
        description="Reads the hidden hex color while navigating LEFT in the PowerPoint Standard color grid.",
        category="Better Office Accessibility"
    )
    def script_arrowLeft(self, gesture):
        gesture.send()
        self._trigger_delayed_hex_read()

    @script(
        description="Reads the hidden hex color while navigating RIGHT in the PowerPoint Standard color grid.",
        category="Better Office Accessibility"
    )
    def script_arrowRight(self, gesture):
        gesture.send()
        self._trigger_delayed_hex_read()

    __gestures = {
        "kb:upArrow": "arrowUp",
        "kb:downArrow": "arrowDown",
        "kb:leftArrow": "arrowLeft",
        "kb:rightArrow": "arrowRight",
    }
