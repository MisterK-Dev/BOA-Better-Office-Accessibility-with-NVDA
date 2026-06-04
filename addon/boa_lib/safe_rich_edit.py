from logHandler import log
import NVDAObjects.UIA
import NVDAObjects.IAccessible
import NVDAObjects.window.edit
import controlTypes
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
    Override for Office RichEdit20W/RichEdit50W controls that notoriously crash NVDA.
    Architectural Why: NVDA normally tries to read advanced text formatting in Office fields by 
    querying the ITextDocument interface via COM. However, in certain fields (like the Excel Name Box 
    or PowerPoint Hex Box), this COM call hangs and eventually throws an OSError, crashing the NVDA core.
    By forcibly stripping out the ITextDocumentTextInfo behavior and falling back to a plain EditTextInfo, 
    we lose advanced formatting announcements (bold/italics) but completely eliminate the fatal crash.
    """
    # Override the TextInfo class to use standard, safe edit field logic rather than COM-based logic
    TextInfo = NVDAObjects.window.edit.EditTextInfo
    
    def _get_ITextDocumentObject(self):
        """
        Architectural Why: NVDA internally checks for the presence of this object. 
        By explicitly returning None, we trick NVDA into thinking the rich text COM interface 
        is completely unavailable, forcing it onto the safe fallback path.
        """
        return None
