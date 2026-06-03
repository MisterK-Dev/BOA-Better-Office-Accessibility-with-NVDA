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

import boa_config
from excel_enhancements import manager as excel_manager
from safe_rich_edit import SafeRichEdit
import powerpoint_enhancement
from powerpoint_enhancements import manager as ppt_manager

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
                if boa_config.get_feature_state("excel", "unselect_tracking") or boa_config.get_feature_state("excel", "hidden_row_skip"):
                    from excel_enhancements.cell_navigation_tracker import check_unselect
                    check_unselect(obj)
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
                if boa_config.get_feature_state("excel", "unselect_tracking") or boa_config.get_feature_state("excel", "hidden_row_skip"):
                    from excel_enhancements.cell_navigation_tracker import check_unselect
                    check_unselect(obj)
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
        if appName == "excel":
            if className in ("EXCEL7", "XLDESK", "NetUIHWND"):
                excel_manager.inject_excel_grid_classes(clsList)
                
            if className == "EXCEL=":
                excel_manager.inject_excel_rename_class(clsList)
                
            elif className in ("RichEdit20W", "RichEdit50W") and boa_config.get_feature_state("excel", "safe_rich_edit"):
                clsList.insert(0, SafeRichEdit)

        # -----------------------------------------------------
        # PowerPoint Overrides
        # -----------------------------------------------------
        elif appName == "powerpnt":
            if className == "bosa_sdm_Mso96":
                import controlTypes
                if getattr(obj, "role", None) == controlTypes.Role.TAB:
                    ppt_manager.inject_ppt_color_grid(clsList)

            if className == "RichEdit20W":
                if getattr(obj, 'windowControlID', None) == 1637:
                    ppt_manager.inject_ppt_hex_edit(clsList)
                elif boa_config.get_feature_state("powerpoint", "safe_rich_edit"):
                    clsList.insert(0, SafeRichEdit)
            elif className == "RichEdit50W" and boa_config.get_feature_state("powerpoint", "safe_rich_edit"):
                clsList.insert(0, SafeRichEdit)
            
            if className == "Edit":
                parent = getattr(obj, 'parent', None)
                parent_class = getattr(parent, 'windowClassName', '') if parent else ""
                if parent_class == "#32770":
                    ppt_manager.inject_ppt_rgb_edit(clsList)

        elif appName == "winword":
            if className in ("RichEdit20W", "RichEdit50W") and boa_config.get_feature_state("word", "safe_rich_edit"):
                log.info("BOA: injecting SafeRichEdit for Word!")
                clsList.insert(0, SafeRichEdit)

    def script_triggerCommandPrefix(self, gesture):
        import tones
        tones.beep(800, 50)
        
        # Bind interceptor keys dynamically
        self.bindGesture("kb:escape", "cancelCommandPrefix")
        # Bind the entire alphabet and numbers to intercept valid commands
        for char in "abcdefghijklmnopqrstuvwxyz0123456789":
            self.bindGesture(f"kb:{char}", "handleCommandKey")
            
    script_triggerCommandPrefix.__doc__ = "Triggers the BOA command prefix mode."

    def script_cancelCommandPrefix(self, gesture):
        import tones
        tones.beep(300, 50)
        self._clear_command_bindings()
        
    def script_handleCommandKey(self, gesture):
        import tones
        key = gesture.displayName.lower()
        self._clear_command_bindings()
        
        # Get active object and app
        import api
        obj = api.getFocusObject()
        appModule = getattr(obj, 'appModule', None)
        appName = getattr(appModule, 'appName', '').lower() if appModule else ""
        
        handled = False
        if appName == "excel":
            handled = excel_manager.handle_prefix_command(key, obj)
        elif appName == "powerpnt":
            handled = ppt_manager.handle_prefix_command(key, obj)
            
        if not handled:
            tones.beep(150, 50) # Error beep
            
    def _clear_command_bindings(self):
        try:
            self.removeGestureBinding("kb:escape")
            for char in "abcdefghijklmnopqrstuvwxyz0123456789":
                self.removeGestureBinding(f"kb:{char}")
        except Exception:
            pass

    __gestures = {
        "kb:NVDA+e": "triggerCommandPrefix"
    }
