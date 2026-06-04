import globalPluginHandler
import NVDAObjects.UIA
from logHandler import log
import os
import sys

# Dynamically add the 'boa_lib' directory to sys.path.
# This ensures that NVDA can import our custom external logic files (excel_enhancement, powerpoint_enhancement)
# without cluttering the globalPlugins directory.
addon_dir = os.path.dirname(os.path.dirname(__file__))
lib_dir = os.path.join(addon_dir, "boa_lib")
if lib_dir not in sys.path:
    sys.path.insert(0, lib_dir)

# Import all specific enhancement classes from our custom library directory.
import excel_enhancement
from excel_enhancement import ExcelSheetRenameEdit, ExcelGridMover, SafeRichEdit
import powerpoint_enhancement
from powerpoint_enhancement import PowerPointHexEdit, PowerPointRGBEdit, PowerPointStandardColorGrid

class GlobalPlugin(globalPluginHandler.GlobalPlugin):
    """
    The GlobalPlugin acts as the main entry point for the BOA Add-on.
    It intercepts NVDA events (like focus and selection changes) and injects custom Python classes 
    over standard Microsoft Office UI elements to override their default, often inaccessible behavior.
    """
    
    def __init__(self, *args, **kwargs):
        super(GlobalPlugin, self).__init__(*args, **kwargs)
        import boa_gui
        import gui.settingsDialogs
        try:
            gui.settingsDialogs.NVDASettingsDialog.categoryClasses.append(boa_gui.BOASettingsPanel)
            log.info("BOA SettingsPanel registered.")
        except Exception as e:
            log.error(f"BOA: Failed to register settings panel: {e}")

    def terminate(self):
        super(GlobalPlugin, self).terminate()
        import boa_gui
        import gui.settingsDialogs
        try:
            if boa_gui.BOASettingsPanel in gui.settingsDialogs.NVDASettingsDialog.categoryClasses:
                gui.settingsDialogs.NVDASettingsDialog.categoryClasses.remove(boa_gui.BOASettingsPanel)
                log.info("BOA SettingsPanel unregistered.")
        except Exception as e:
            log.error(f"BOA: Failed to unregister settings panel: {e}")
    
    def event_gainFocus(self, obj, nextHandler):
        """
        Triggered every time a new object gains focus in the operating system.
        We use this to track selection state changes specifically within Excel.
        """
        try:
            appModule = getattr(obj, 'appModule', None)
            if appModule and getattr(appModule, 'appName', '').lower() == "excel":
                import boa_config
                if boa_config.get_feature_state("excel", "unselect_tracking"):
                    import excel_enhancement
                    # Call our custom selection tracking logic before allowing NVDA to handle the focus event.
                    excel_enhancement.check_unselect(obj)
        except Exception:
            pass
        nextHandler()
        
    def event_selectionChange(self, obj, nextHandler):
        """
        Triggered when a selection changes (e.g., highlighting a different group of cells).
        """
        try:
            appModule = getattr(obj, 'appModule', None)
            if appModule and getattr(appModule, 'appName', '').lower() == "excel":
                import boa_config
                if boa_config.get_feature_state("excel", "unselect_tracking"):
                    import excel_enhancement
                    # Notify the user if a multi-cell selection was unexpectedly deselected.
                    excel_enhancement.check_unselect(obj)
        except Exception:
            pass
        nextHandler()

    def chooseNVDAObjectOverlayClasses(self, obj, clsList):
        """
        This is the core injection method. When NVDA discovers a new UI element, 
        it asks plugins if they want to apply any custom class overrides to that object.
        We check the application name and window class to selectively inject our enhancements.
        """
        appModule = getattr(obj, 'appModule', None)
        if not appModule:
            return

        appName = getattr(appModule, 'appName', '').lower()
        className = getattr(obj, 'windowClassName', '')
        
        if appName in ("excel", "powerpnt", "winword"):
            log.info(f"BOA hook: app={appName}, class={className}")

        # --- STRICT UIA VERIFICATION ---
        # Never inject a UIA-specific class into an object that doesn't natively support UI Automation.
        # This prevents catastrophic crashes within NVDA's C++ core.
        is_uia = hasattr(obj, 'UIAElement') or (NVDAObjects.UIA.UIA in clsList)
        if not is_uia:
            pass

        # -----------------------------------------------------
        # Excel Overrides
        # -----------------------------------------------------
        import boa_config
        if appName == "excel":
            # The main Excel spreadsheet grid classes.
            if className in ("EXCEL7", "XLDESK", "NetUIHWND") and boa_config.get_feature_state("excel", "grid_mover"):
                log.info("BOA: injecting ExcelGridMover!")
                # Insert our custom Grid Mover at the top of the class hierarchy so it intercepts keystrokes first.
                clsList.insert(0, ExcelGridMover)
                
            # The 'EXCEL=' class specifically represents the native "Rename Sheet" edit box.
            if className == "EXCEL=" and boa_config.get_feature_state("excel", "sheet_rename"):
                log.info("BOA: injecting ExcelSheetRenameEdit!")
                clsList.insert(0, ExcelSheetRenameEdit)
            # Standard RichEdit controls in Office sometimes crash when ITextDocument is accessed.
            elif className in ("RichEdit20W", "RichEdit50W") and boa_config.get_feature_state("excel", "safe_rich_edit"):
                log.info("BOA: injecting SafeRichEdit for Excel!")
                clsList.insert(0, SafeRichEdit)

        # -----------------------------------------------------
        # PowerPoint Overrides
        # -----------------------------------------------------
        elif appName == "powerpnt":
            # 'bosa_sdm_Mso96' is a legacy Office dialog class, used here for the 'Standard' color hexagon.
            if className == "bosa_sdm_Mso96" and boa_config.get_feature_state("powerpoint", "standard_color_grid"):
                import controlTypes
                if getattr(obj, "role", None) == controlTypes.Role.TAB:
                    log.info("BOA: injecting PowerPointStandardColorGrid for TAB in PPT!")
                    clsList.insert(0, PowerPointStandardColorGrid)

            # RichEdit20W and RichEdit50W are used extensively in PowerPoint dialogs.
            if className == "RichEdit20W":
                # windowControlID 1637 uniquely identifies the Hex input field in the Color picker.
                if getattr(obj, 'windowControlID', None) == 1637 and boa_config.get_feature_state("powerpoint", "hex_edit"):
                    log.info("BOA: injecting PowerPointHexEdit!")
                    clsList.insert(0, PowerPointHexEdit)
                elif boa_config.get_feature_state("powerpoint", "safe_rich_edit"):
                    log.info("BOA: injecting SafeRichEdit for PPT!")
                    clsList.insert(0, SafeRichEdit)
            elif className == "RichEdit50W" and boa_config.get_feature_state("powerpoint", "safe_rich_edit"):
                log.info("BOA: injecting SafeRichEdit for PPT!")
                clsList.insert(0, SafeRichEdit)
            
            # The standard 'Edit' class is used for the RGB input fields.
            if className == "Edit" and boa_config.get_feature_state("powerpoint", "rgb_edit"):
                parent = getattr(obj, 'parent', None)
                parent_class = getattr(parent, 'windowClassName', '') if parent else ""
                
                # Verify that the parent dialog is the standard Windows '#32770' dialog box 
                # to avoid injecting this into random Edit fields throughout PowerPoint.
                if parent_class == "#32770":
                    clsList.insert(0, PowerPointRGBEdit)

        # -----------------------------------------------------
        # Word Overrides
        # -----------------------------------------------------
        elif appName == "winword":
            if className in ("RichEdit20W", "RichEdit50W") and boa_config.get_feature_state("word", "safe_rich_edit"):
                log.info("BOA: injecting SafeRichEdit for Word!")
                clsList.insert(0, SafeRichEdit)
