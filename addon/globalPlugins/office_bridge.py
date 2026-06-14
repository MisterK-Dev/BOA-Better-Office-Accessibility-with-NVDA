# -*- coding: UTF-8 -*-
# Copyright (C) 2026 KIRAN G T {MisterK} and Antigravity 2
# This file is covered by the GNU General Public License (GPL), version 2.
# See the file COPYING.txt for more details.

import globalPluginHandler

from logHandler import log
import os
import sys
import addonHandler

# Initialize gettext translation for the addon domain
addonHandler.initTranslation()

# Dynamically add the 'boa_lib' directory to sys.path.
# This ensures that NVDA can import our custom external logic files (excel_enhancement, powerpoint_enhancement)
# without cluttering the globalPlugins directory.
addon_dir = os.path.dirname(os.path.dirname(__file__))
if addon_dir not in sys.path:
    sys.path.insert(0, addon_dir)

from boa_lib import boa_config

# App-Launch Caching Variables
excel_manager = None
ppt_manager = None
word_manager = None
from boa_lib.safe_rich_edit import SafeRichEdit

class GlobalPlugin(globalPluginHandler.GlobalPlugin):
    """
    The GlobalPlugin acts as the main entry point for the BOA Add-on.
    Architecturally, it intercepts NVDA events (like focus and selection changes) at a global level
    and injects custom Python classes over standard Microsoft Office UI elements.
    This allows us to override their default, often inaccessible behavior without modifying NVDA's core.
    """
    
    def __init__(self, *args, **kwargs):
        """
        Initializes the global plugin.
        Architectural Why: We register the BOA Settings panel here so it becomes available
        in the NVDA settings dialog immediately upon add-on initialization.
        """
        super(GlobalPlugin, self).__init__(*args, **kwargs)
        from boa_lib import boa_gui
        import gui.settingsDialogs
        try:
            # Register our custom GUI panel into NVDA's main settings dialog categories
            gui.settingsDialogs.NVDASettingsDialog.categoryClasses.append(boa_gui.BOASettingsPanel)
            log.info("BOA SettingsPanel registered.")
        except Exception as e:
            log.error(f"BOA: Failed to register settings panel: {e}")

    def terminate(self):
        """
        Cleans up when the add-on is unloaded or NVDA exits.
        Architectural Why: It is critical to unregister our custom settings panel to prevent
        memory leaks or crashes in NVDA's GUI when the add-on is no longer active.
        """
        super(GlobalPlugin, self).terminate()
        from boa_lib import boa_gui
        import gui.settingsDialogs
        try:
            # Check if our panel is still registered before attempting removal
            if boa_gui.BOASettingsPanel in gui.settingsDialogs.NVDASettingsDialog.categoryClasses:
                gui.settingsDialogs.NVDASettingsDialog.categoryClasses.remove(boa_gui.BOASettingsPanel)
                log.info("BOA SettingsPanel unregistered.")
        except Exception as e:
            log.error(f"BOA: Failed to unregister settings panel: {e}")
    
    def event_gainFocus(self, obj, nextHandler):
        """
        Triggered every time a new object gains focus in the operating system.
        Architectural Why: We intercept focus events globally to monitor state changes,
        specifically for tracking Excel cell selection when navigating.
        """
        try:
            # Retrieve the application module associated with the focused object
            appModule = getattr(obj, 'appModule', None)
            # Check if the application is Microsoft Excel
            if appModule and getattr(appModule, 'appName', '').lower() == "excel":
                # Verify if tracking features are enabled in BOA's configuration
                if (boa_config.get_feature_state("excel", "unselect_tracking") or 
                    boa_config.get_feature_state("excel", "hidden_row_skip") or 
                    boa_config.get_feature_state("excel", "auto_announce_first_block") in ["one_time", "guided"] or
                    boa_config.get_feature_state("excel", "conditional_formatting")):
                    from boa_lib.excel_enhancements.cell_navigation_tracker import check_unselect
                    # Delegate the check to the external tracker to keep this global hook lean
                    check_unselect(obj)
        except Exception:
            pass
        # Always call the nextHandler to ensure NVDA continues normal focus processing
        nextHandler()
        
    def event_selectionChange(self, obj, nextHandler):
        """
        Triggered when a selection changes (e.g., highlighting a different group of cells).
        Architectural Why: Similar to focus events, we track selection changes to maintain
        accurate context of user navigation in complex grids like Excel.
        """
        try:
            # Retrieve the application module
            appModule = getattr(obj, 'appModule', None)
            if appModule and getattr(appModule, 'appName', '').lower() == "excel":
                if (boa_config.get_feature_state("excel", "unselect_tracking") or 
                    boa_config.get_feature_state("excel", "hidden_row_skip") or 
                    boa_config.get_feature_state("excel", "auto_announce_first_block") in ["one_time", "guided"] or
                    boa_config.get_feature_state("excel", "conditional_formatting")):
                    from boa_lib.excel_enhancements.cell_navigation_tracker import check_unselect
                    check_unselect(obj)
        except Exception:
            pass
        # Pass control back to NVDA's core handlers
        nextHandler()

    def chooseNVDAObjectOverlayClasses(self, obj, clsList):
        """
        This is the core injection method. 
        Architectural Why: When NVDA discovers a new UI element, it asks plugins if they want 
        to apply any custom class overrides to that object. We check the application name and 
        window class to selectively inject our enhancements, overriding default behavior without modifying NVDA core.
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
        is_uia = hasattr(obj, 'UIAElement') or any(c.__name__ == 'UIA' for c in clsList)
        if not is_uia:
            pass

        # -----------------------------------------------------
        # Excel Overrides
        # -----------------------------------------------------
        if appName == "excel":
            # Use global cached reference to avoid repeated expensive imports on every UI event
            global excel_manager
            if excel_manager is None:
                from boa_lib.excel_enhancements import manager as excel_manager

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
            # Cache PowerPoint manager to improve injection performance
            global ppt_manager
            if ppt_manager is None:
                from boa_lib.powerpoint_enhancements import manager as ppt_manager

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
            # Cache Word manager for performance
            global word_manager
            if word_manager is None:
                from boa_lib.word_enhancements import manager as word_manager

            if className in ("RichEdit20W", "RichEdit50W"):
                word_manager.inject_word_safe_rich_edit(clsList)

    def script_triggerCommandPrefix(self, gesture):
        """
        Triggers the BOA command prefix mode.
        Architectural Why: We use a multi-key command architecture (like screen readers often do)
        to avoid consuming too many global shortcut keys. Pressing the prefix key captures subsequent keystrokes.
        """
        import tones
        # Emit a high-pitched beep to indicate the prefix mode is active
        tones.beep(800, 50)
        
        # Bind interceptor keys dynamically to capture the next keystroke
        self.bindGesture("kb:escape", "cancelCommandPrefix")
        self.bindGesture("kb:backspace", "handleCommandKey")
        # Bind the entire alphabet and numbers to intercept valid commands
        for char in "abcdefghijklmnopqrstuvwxyz0123456789":
            self.bindGesture(f"kb:{char}", "handleCommandKey")
        # Bind shifted numbers for assigning slots
        for i in range(1, 10):
            self.bindGesture(f"kb:shift+{i}", "handleCommandKey")

    def script_cancelCommandPrefix(self, gesture):
        """
        Cancels the active BOA command prefix mode.
        Architectural Why: Provides a fallback (escape hatch) for the user if they accidentally
        triggered the prefix or changed their mind, restoring normal keyboard functionality.
        """
        import tones
        # Emit a lower-pitched beep to indicate cancellation
        tones.beep(300, 50)
        self._clear_command_bindings()
        
    def script_handleCommandKey(self, gesture):
        """
        Handles the keystroke immediately following the command prefix.
        Architectural Why: Routes the intercepted keystroke to the appropriate application-specific
        manager (Excel, PowerPoint, Word) based on the currently active application.
        """
        import tones
        key = gesture.displayName.lower()
        # Immediately clear bindings so subsequent keys act normally
        self._clear_command_bindings()
        
        # Get active object and app to determine routing
        import api
        obj = api.getFocusObject()
        appModule = getattr(obj, 'appModule', None)
        appName = getattr(appModule, 'appName', '').lower() if appModule else ""
        
        handled = False
        # Route the command to the respective manager
        if appName == "excel":
            global excel_manager
            if excel_manager is None:
                from boa_lib.excel_enhancements import manager as excel_manager
            handled = excel_manager.handle_prefix_command(key, obj)
        elif appName == "powerpnt":
            global ppt_manager
            if ppt_manager is None:
                from boa_lib.powerpoint_enhancements import manager as ppt_manager
            handled = ppt_manager.handle_prefix_command(key, obj)
        elif appName == "winword":
            global word_manager
            if word_manager is None:
                from boa_lib.word_enhancements import manager as word_manager
            handled = word_manager.handle_prefix_command(key, obj)
            
        if not handled:
            tones.beep(150, 50) # Error beep for unhandled or invalid commands
            
    def _clear_command_bindings(self):
        """
        Removes dynamic keyboard bindings used during the prefix mode.
        Architectural Why: Ensures that our temporary interception of alphanumeric keys
        does not interfere with normal typing once the command sequence is complete or cancelled.
        """
        try:
            self.removeGestureBinding("kb:escape")
            self.removeGestureBinding("kb:backspace")
            for char in "abcdefghijklmnopqrstuvwxyz0123456789":
                self.removeGestureBinding(f"kb:{char}")
            for i in range(1, 10):
                self.removeGestureBinding(f"kb:shift+{i}")
        except Exception:
            pass

    __gestures = {
        "kb:NVDA+e": "triggerCommandPrefix"
    }
