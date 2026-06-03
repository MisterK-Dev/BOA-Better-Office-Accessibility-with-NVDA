import controlTypes
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

class QuickSheetMover(object):
    # Gestures merged to the bottom __gestures block

    def _move_sheet(self, direction):
        import comtypes.client
        import comtypes.automation
        import ctypes
        import ui
        try:
            # Safely get the Excel object bypassing GetActiveObject to prevent MK_E_UNAVAILABLE errors
            hwnd7 = None
            if getattr(self, "windowClassName", "") == "EXCEL7":
                hwnd7 = self.windowHandle
            else:
                hwnd = ctypes.windll.user32.FindWindowW("XLMAIN", None)
                if hwnd:
                    xldesk = ctypes.windll.user32.FindWindowExW(hwnd, 0, "XLDESK", None)
                    if xldesk:
                        hwnd7 = ctypes.windll.user32.FindWindowExW(xldesk, 0, "EXCEL7", None)
            
            if not hwnd7:
                ui.message("Could not find Excel grid.")
                return
                
            oleacc = ctypes.windll.oleacc
            OBJID_NATIVEOM = -16
            ptr = ctypes.POINTER(comtypes.automation.IDispatch)()
            
            res = oleacc.AccessibleObjectFromWindow(
                hwnd7, OBJID_NATIVEOM, 
                ctypes.byref(comtypes.automation.IDispatch._iid_), 
                ctypes.byref(ptr)
            )
            
            if res != 0 or not ptr:
                ui.message("Failed to hook Excel.")
                return
                
            win = comtypes.client.dynamic.Dispatch(ptr)
            excel = win.Application
            wb = excel.ActiveWorkbook
            sheet = excel.ActiveSheet
            
            current_index = sheet.Index
            total_sheets = wb.Sheets.Count
            
            if total_sheets <= 1:
                ui.message("Only one sheet in workbook")
                return
                
            if direction == "left":
                if current_index == 1:
                    ui.message("Already at beginning")
                    return
                sheet.Move(wb.Sheets(current_index - 1))
            elif direction == "right":
                if current_index == total_sheets:
                    ui.message("Already at end")
                    return
                # To move right, we just move the right neighbor before us!
                wb.Sheets(current_index + 1).Move(sheet)
            elif direction == "start":
                if current_index == 1:
                    ui.message("Already at beginning")
                    return
                sheet.Move(wb.Sheets(1))
            elif direction == "end":
                if current_index == total_sheets:
                    ui.message("Already at end")
                    return
                # Two-step COM trick to move to the very end without the broken 'After' parameter:
                # 1. Move our sheet BEFORE the very last sheet (if we aren't already just before it)
                if current_index < total_sheets - 1:
                    sheet.Move(wb.Sheets(total_sheets))
                # 2. Move the last sheet BEFORE our sheet! This puts our sheet at the very end!
                wb.Sheets(total_sheets).Move(sheet)
                
            # Moving sheets can sometimes change the active sheet (especially the 2-step end trick).
            # Force our target sheet to be the active one.
            sheet.Activate()
            
            new_index = excel.ActiveSheet.Index
            sheet_name = excel.ActiveSheet.Name
            ui.message(f"Moved {sheet_name} to position {new_index} of {total_sheets}")
        except Exception as e:
            ui.message("Failed to move sheet")
            import logHandler
            logHandler.log.error(f"ExcelGridMover error: {e}")

    # (script import moved to top of file)

    @script(
        description="Moves the active Excel sheet to the left.",
        category="Better Office Accessibility"
    )
    def script_moveSheetLeft(self, gesture):
        self._move_sheet("left")

    @script(
        description="Moves the active Excel sheet to the right.",
        category="Better Office Accessibility"
    )
    def script_moveSheetRight(self, gesture):
        self._move_sheet("right")

    @script(
        description="Moves the active Excel sheet to the very beginning of the workbook.",
        category="Better Office Accessibility"
    )
    def script_moveSheetStart(self, gesture):
        self._move_sheet("start")

    @script(
        description="Moves the active Excel sheet to the very end of the workbook.",
        category="Better Office Accessibility"
    )
    def script_moveSheetEnd(self, gesture):
        self._move_sheet("end")

    __gestures = {
        "kb:NVDA+shift+leftArrow": "moveSheetLeft",
        "kb:NVDA+shift+rightArrow": "moveSheetRight",
        "kb:NVDA+shift+left": "moveSheetLeft",
        "kb:NVDA+shift+right": "moveSheetRight",
        "kb:NVDA+shift+pageUp": "moveSheetLeft",
        "kb:NVDA+shift+pageDown": "moveSheetRight",
        "kb:NVDA+shift+home": "moveSheetStart",
        "kb:NVDA+shift+end": "moveSheetEnd",
        "kb:NVDA+alt+leftArrow": "moveSheetLeft",
        "kb:NVDA+alt+rightArrow": "moveSheetRight",
    }
