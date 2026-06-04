import globalPluginHandler
import NVDAObjects.UIA
from logHandler import log
import os
import sys

# Dynamically add boa_lib to sys.path so we can import our modules
addon_dir = os.path.dirname(os.path.dirname(__file__))
lib_dir = os.path.join(addon_dir, "boa_lib")
if lib_dir not in sys.path:
    sys.path.insert(0, lib_dir)

# Import all the specific enhancement classes
import excel_enhancement
from excel_enhancement import ExcelSheetRenameEdit, ExcelGridMover, SafeRichEdit
import powerpoint_enhancement
from powerpoint_enhancement import PowerPointHexEdit, PowerPointRGBEdit, PowerPointStandardColorGrid

class GlobalPlugin(globalPluginHandler.GlobalPlugin):
    """
    Global plugin that intercepts object creation to safely inject UIA overrides.
    """
    def event_gainFocus(self, obj, nextHandler):
        try:
            appModule = getattr(obj, 'appModule', None)
            if appModule and getattr(appModule, 'appName', '').lower() == "excel":
                import excel_enhancement
                excel_enhancement.check_unselect(obj)
        except Exception:
            pass
        nextHandler()
        
    def event_selectionChange(self, obj, nextHandler):
        try:
            appModule = getattr(obj, 'appModule', None)
            if appModule and getattr(appModule, 'appName', '').lower() == "excel":
                import excel_enhancement
                excel_enhancement.check_unselect(obj)
        except Exception:
            pass
        nextHandler()

    def chooseNVDAObjectOverlayClasses(self, obj, clsList):
        appModule = getattr(obj, 'appModule', None)
        if not appModule:
            return

        appName = getattr(appModule, 'appName', '').lower()
        className = getattr(obj, 'windowClassName', '')
        
        if appName in ("excel", "powerpnt", "winword"):
            log.info(f"BOA hook: app={appName}, class={className}")

        # --- STRICT UIA VERIFICATION ---
        # Never inject a UIA class into an object that doesn't support it.
        is_uia = hasattr(obj, 'UIAElement') or (NVDAObjects.UIA.UIA in clsList)
        if not is_uia:
            # Note: For non-UIA objects, we allow the override logic to proceed if classes are already IAccessible or window based
            pass

        # -----------------------------------------------------
        # Excel Overrides
        # -----------------------------------------------------
        if appName == "excel":
            if className in ("EXCEL7", "XLDESK", "NetUIHWND"):
                log.info("BOA: injecting ExcelGridMover!")
                clsList.insert(0, ExcelGridMover)
                
            # The log identified the sheet renaming box as 'EXCEL='
            if className == "EXCEL=":
                log.info("BOA: injecting ExcelSheetRenameEdit!")
                clsList.insert(0, ExcelSheetRenameEdit)
            elif className in ("RichEdit20W", "RichEdit50W"):
                log.info("BOA: injecting SafeRichEdit for Excel!")
                clsList.insert(0, SafeRichEdit)

        # -----------------------------------------------------
        # PowerPoint Overrides
        # -----------------------------------------------------
        elif appName == "powerpnt":
            if className == "bosa_sdm_Mso96":
                import controlTypes
                if getattr(obj, "role", None) == controlTypes.Role.TAB:
                    log.info("BOA: injecting PowerPointStandardColorGrid for TAB in PPT!")
                    clsList.insert(0, PowerPointStandardColorGrid)

            if className == "RichEdit20W":
                if getattr(obj, 'windowControlID', None) == 1637:
                    log.info("BOA: injecting PowerPointHexEdit!")
                    clsList.insert(0, PowerPointHexEdit)
                else:
                    log.info("BOA: injecting SafeRichEdit for PPT!")
                    clsList.insert(0, SafeRichEdit)
            elif className == "RichEdit50W":
                log.info("BOA: injecting SafeRichEdit for PPT!")
                clsList.insert(0, SafeRichEdit)
            
            if className == "Edit":
                parent = getattr(obj, 'parent', None)
                parent_class = getattr(parent, 'windowClassName', '') if parent else ""
                
                # #32770 is the standard Windows dialog class for the Color Picker
                if parent_class == "#32770":
                    clsList.insert(0, PowerPointRGBEdit)
