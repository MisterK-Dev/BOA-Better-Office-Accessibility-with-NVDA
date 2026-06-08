import NVDAObjects.UIA
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

_is_renaming_sheet = False

class ExcelSheetRenameEdit(object):
    """
    Override for the Excel 'Rename sheet' edit box.
    ARCHITECTURAL INTENT: Excel's native sheet rename field is deeply inaccessible. It fails
    to expose its initial text or cursor position to UIA correctly. Instead of fighting
    Excel's broken implementation, we intercept the focus event, hide the broken UI behind
    a fully accessible standard wx.Dialog, and then use secure clipboard injection to push 
    the result back into Excel.
    """
    def event_gainFocus(self):
        """
        Intercepts focus when the user triggers 'Rename Sheet'.
        """
        super(ExcelSheetRenameEdit, self).event_gainFocus()
        global _is_renaming_sheet
        if not _is_renaming_sheet:
            _is_renaming_sheet = True
            # Fetch the actual sheet name BEFORE opening the dialog because the native edit box is empty in UIA
            initial_name = self._fetch_sheet_name()
            # Safely push the dialog creation onto NVDA's main GUI thread to prevent deadlocks
            wx.CallAfter(self._show_rename_dialog, initial_name, self.windowHandle)

    def _show_rename_dialog(self, initial_name, hwnd):
        """
        Creates the custom WX dialog to capture the new sheet name from the user.
        """
        gui.mainFrame.prePopup()
        dlg = wx.TextEntryDialog(gui.mainFrame, "Enter new sheet name:", "Rename Sheet", initial_name)
        dlg.Raise()
        res = dlg.ShowModal()
        new_name = dlg.GetValue() if res == wx.ID_OK else None
        dlg.Destroy()
        gui.mainFrame.postPopup()
        
        def _restore_clip_and_reset(old_clip):
            """
            Restores the user's original clipboard content after we are done injecting the new name.
            """
            if old_clip:
                import api
                try:
                    api.copyToClip(old_clip)
                except Exception:
                    pass
            global _is_renaming_sheet
            _is_renaming_sheet = False

        def _do_enter(clean_name, old_clip):
            """
            Simulates pressing Enter to finalize the rename in Excel's native edit field.
            """
            keyboardHandler.KeyboardInputGesture.fromName("enter").send()
            import ui
            ui.message(f"Renaming to {clean_name}")
            # Delay restoring the clipboard to ensure Excel has time to process the Enter key
            core.callLater(1500, lambda: _restore_clip_and_reset(old_clip))

        def _do_inject(old_clip, clean_name, fg_hwnd):
            """
            Injects the new name into Excel by pasting it from the clipboard.
            """
            if winUser.getForegroundWindow() == fg_hwnd:
                # Ctrl+V is much faster and more reliable than typing characters individually
                keyboardHandler.KeyboardInputGesture.fromName("control+v").send()
                core.callLater(200, lambda: _do_enter(clean_name, old_clip))
            else:
                # If Excel lost focus, abort to prevent pasting over the user's data in another app
                _restore_clip_and_reset(old_clip)

        def _do_clipboard(clean_name, fg_hwnd):
            """
            Backs up the clipboard and copies the new sheet name.
            """
            import api
            old_clip = ""
            try:
                old_clip = api.getClipData()
            except Exception:
                pass
            try:
                api.copyToClip(clean_name)
                core.callLater(200, lambda: _do_inject(old_clip, clean_name, fg_hwnd))
            except Exception:
                _restore_clip_and_reset(old_clip)

        def _check_security():
            """
            Security check to prevent errant keystroke injection.
            Ensures that Excel is still the active window before we start pasting and pressing Enter.
            """
            global _is_renaming_sheet
            import ctypes
            fg_hwnd = winUser.getForegroundWindow()
            fg_pid = ctypes.c_ulong()
            ctypes.windll.user32.GetWindowThreadProcessId(fg_hwnd, ctypes.byref(fg_pid))
            
            target_pid = ctypes.c_ulong()
            ctypes.windll.user32.GetWindowThreadProcessId(hwnd, ctypes.byref(target_pid))
            
            # GUARD: Abort if the foreground process does not match the Excel process
            if fg_pid.value != target_pid.value:
                log.warning("BOA: Foreground window mismatch! Aborting keystroke injection to prevent pasting into wrong app.")
                _is_renaming_sheet = False
                return
            
            clean_name = new_name.strip() if new_name else ""
            if not clean_name:
                # User cancelled or entered empty string; escape out of the native rename edit box
                keyboardHandler.KeyboardInputGesture.fromName("escape").send()
                _is_renaming_sheet = False
                return
                
            _do_clipboard(clean_name, fg_hwnd)

        def _set_foreground():
            """
            Forces Excel back to the foreground before injecting keystrokes.
            """
            winUser.setForegroundWindow(hwnd)
            core.callLater(200, _check_security)
            
        core.callLater(100, _set_foreground)

    def _get_name(self):
        """
        Override the UIA name property to ensure it announces logically.
        """
        return "Rename sheet"

    def _fetch_sheet_name(self):
        """
        Hunts down the selected sheet tab in the Excel UI Automation tree.
        ARCHITECTURAL INTENT: Since the native edit field does not expose its initial text 
        correctly to NVDA, we must manually traverse the UIA tree, find the active TabItem 
        (or ListItem in newer builds), and extract its name.
        """
        try:
            # Step 1: Traverse up the object tree to find the main Excel window (XLMAIN).
            p = getattr(self, 'parent', None)
            xlmain_hwnd = None
            while p:
                if getattr(p, 'windowClassName', '') == 'XLMAIN':
                    xlmain_hwnd = p.windowHandle
                    break
                p = getattr(p, 'parent', None)
            
            if xlmain_hwnd:
                xlmain_uia = UIAHandler.handler.clientObject.ElementFromHandle(xlmain_hwnd)
                if xlmain_uia:
                    log.info("BOA: XLMAIN UIA found! Searching for TabItems...")
                    # Step 2: Search the UI Automation tree for TabItems (standard sheet tabs).
                    condition = UIAHandler.handler.clientObject.CreatePropertyCondition(
                        UIAHandler.UIA_ControlTypePropertyId, 
                        UIAHandler.UIA_TabItemControlTypeId
                    )
                    tabs = xlmain_uia.FindAll(UIAHandler.TreeScope_Descendants, condition)
                    if tabs:
                        log.info(f"BOA: Found {tabs.length} TabItems in XLMAIN")
                        for i in range(tabs.length):
                            tab = tabs.GetElement(i)
                            try:
                                # Property 30079 is UIA_SelectionItemIsSelectedPropertyId
                                is_sel = tab.GetCurrentPropertyValue(30079)
                                name = tab.CurrentName
                                log.info(f"BOA: TabItem {i}: name='{name}', selected={is_sel}")
                                if is_sel:
                                    return name
                            except Exception as e:
                                log.info(f"BOA: Error checking selection for tab {i}: {e}")
                    else:
                        # Step 3: Fallback for newer Excel 365 builds where sheet tabs are rendered as ListItems!
                        log.info("BOA: No TabItems found. Searching for ListItem controls instead...")
                        condition2 = UIAHandler.handler.clientObject.CreatePropertyCondition(
                            UIAHandler.UIA_ControlTypePropertyId, 
                            UIAHandler.UIA_ListItemControlTypeId
                        )
                        list_items = xlmain_uia.FindAll(UIAHandler.TreeScope_Descendants, condition2)
                        if list_items:
                            for i in range(list_items.length):
                                item = list_items.GetElement(i)
                                try:
                                    is_sel = item.GetCurrentPropertyValue(30079)
                                    name = item.CurrentName
                                    # Ensure it's actually a sheet tab by checking text/selection
                                    if is_sel and name and ("sheet" in name.lower() or name != ""):
                                        log.info(f"BOA: ListItem {i}: name='{name}', selected={is_sel}")
                                        return name
                                except Exception:
                                    pass
        except Exception as e:
            log.info(f"BOA: UIA Sheet Tab search failed: {e}")

        return ""
