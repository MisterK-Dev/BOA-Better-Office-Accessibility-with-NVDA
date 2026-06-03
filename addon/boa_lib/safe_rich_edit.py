from logHandler import log
import UIAHandler
import wx
import gui
import threading
import time
import winUser
import keyboardHandler
import core
from scriptHandler import script
import queueHandler

class SafeRichEdit(object):
    """
    Override for Office RichEdit20W/RichEdit50W controls that crash 
    due to ITextDocument failing with OSError. 
    By bypassing ITextDocumentTextInfo and falling back to standard EditTextInfo, we avoid the crash.
    """
    @property
    def TextInfo(self):
        import NVDAObjects.window.edit
        return NVDAObjects.window.edit.EditTextInfo
    
    def _get_ITextDocumentObject(self):
        return None
