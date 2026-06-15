# -*- coding: UTF-8 -*-
# Copyright (C) 2026 KIRAN G T {MisterK} and Antigravity 2
# This file is covered by the GNU General Public License (GPL), version 2.

import appModuleHandler
import api
import addonHandler

addonHandler.initTranslation()

from nvdaBuiltin.appModules.excel import AppModule as CoreExcelAppModule
from appModules.boa_enhancements.excel_enhancements import manager as excel_manager
from appModules.boa_enhancements.excel_enhancements.cell_navigation_tracker import check_unselect
from appModules.boa_enhancements.safe_rich_edit import SafeRichEdit

class AppModule(CoreExcelAppModule):
    """
    BOA's Excel AppModule.
    This subclasses NVDA's built-in Excel appModule to safely add our custom enhancements
    without breaking core functionality.
    """

    def __init__(self, *args, **kwargs):
        super(AppModule, self).__init__(*args, **kwargs)

    def chooseNVDAObjectOverlayClasses(self, obj, clsList):
        """
        Inject custom classes for specific UI elements in Excel.
        CRITICAL: super() must be called FIRST so NVDA's core Excel classes are placed,
        then BOA inserts its enhancements at position 0 (highest priority) on top.
        """
        super(AppModule, self).chooseNVDAObjectOverlayClasses(obj, clsList)
        from appModules.boa_enhancements import boa_config
        className = getattr(obj, "windowClassName", "")
        
        if className in ("EXCEL7", "XLDESK", "NetUIHWND"):
            excel_manager.inject_excel_grid_classes(clsList)
            
        if className == "EXCEL=":
            excel_manager.inject_excel_rename_class(clsList)
            
        elif className in ("RichEdit20W", "RichEdit50W") and boa_config.get_feature_state("excel", "safe_rich_edit"):
            clsList.insert(0, SafeRichEdit)

    def event_gainFocus(self, obj, nextHandler):
        try:
            from appModules.boa_enhancements import boa_config
            if (boa_config.get_feature_state("excel", "unselect_tracking") or 
                boa_config.get_feature_state("excel", "hidden_row_skip") or 
                boa_config.get_feature_state("excel", "auto_announce_first_block") in ["one_time", "guided"] or
                boa_config.get_feature_state("excel", "conditional_formatting")):
                check_unselect(obj)
        except Exception:
            pass
        nextHandler()

    def event_selectionChange(self, obj, nextHandler):
        try:
            from appModules.boa_enhancements import boa_config
            if (boa_config.get_feature_state("excel", "unselect_tracking") or 
                boa_config.get_feature_state("excel", "hidden_row_skip") or 
                boa_config.get_feature_state("excel", "auto_announce_first_block") in ["one_time", "guided"] or
                boa_config.get_feature_state("excel", "conditional_formatting")):
                check_unselect(obj)
        except Exception:
            pass
        nextHandler()

    def _clear_command_bindings(self):
        try:
            self.removeGestureBinding("kb:escape")
            self.removeGestureBinding("kb:backspace")
            for char in "abcdefghijklmnopqrstuvwxyz0123456789":
                self.removeGestureBinding(f"kb:{char}")
            for i in range(1, 10):
                self.removeGestureBinding(f"kb:shift+{i}")
        except Exception:
            pass

    def script_triggerCommandPrefix(self, gesture):
        """
        Catches the initial NVDA+E prefix for Excel.
        """
        import tones
        tones.beep(800, 50)
        
        self.bindGesture("kb:escape", "cancelCommandPrefix")
        self.bindGesture("kb:backspace", "handleCommandKey")
        for char in "abcdefghijklmnopqrstuvwxyz0123456789":
            self.bindGesture(f"kb:{char}", "handleCommandKey")
        for i in range(1, 10):
            self.bindGesture(f"kb:shift+{i}", "handleCommandKey")
    
    script_triggerCommandPrefix.__doc__ = _("Triggers the BOA command prefix. Press this, followed by a specific command key.")

    def script_cancelCommandPrefix(self, gesture):
        import tones
        tones.beep(300, 50)
        self._clear_command_bindings()
        
    def script_handleCommandKey(self, gesture):
        """
        Catches the subsequent key pressed after NVDA+E.
        Routes to excel_manager.handle_prefix_command(key, obj).
        """
        import tones
        try:
            kb_id = list(gesture.identifiers)[-1]
            key = kb_id.split(":", 1)[1].lower() if ":" in kb_id else gesture.displayName.lower()
        except Exception:
            key = gesture.displayName.lower()
            
        self._clear_command_bindings()
        
        obj = api.getFocusObject()
        handled = excel_manager.handle_prefix_command(key, obj)
        if not handled:
            tones.beep(150, 50)

    __gestures = {
        "kb:NVDA+e": "triggerCommandPrefix"
    }
