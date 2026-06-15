# -*- coding: UTF-8 -*-
# Copyright (C) 2026 KIRAN G T {MisterK} and Antigravity 2
# This file is covered by the GNU General Public License (GPL), version 2.

import appModuleHandler
import api
import controlTypes
import addonHandler

addonHandler.initTranslation()

from nvdaBuiltin.appModules.powerpnt import AppModule as CorePowerpntAppModule
from appModules.boa_enhancements.powerpoint_enhancements import manager as ppt_manager
from appModules.boa_enhancements.safe_rich_edit import SafeRichEdit

class AppModule(CorePowerpntAppModule):
    """
    BOA's PowerPoint AppModule.
    This subclasses NVDA's built-in PowerPoint appModule to safely add our custom enhancements.
    """

    def __init__(self, *args, **kwargs):
        super(AppModule, self).__init__(*args, **kwargs)

    def chooseNVDAObjectOverlayClasses(self, obj, clsList):
        super(AppModule, self).chooseNVDAObjectOverlayClasses(obj, clsList)
        from appModules.boa_enhancements import boa_config
        className = getattr(obj, "windowClassName", "")
        
        if className == "bosa_sdm_Mso96":
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
        import tones
        tones.beep(800, 50)
        
        self.bindGesture("kb:escape", "cancelCommandPrefix")
        self.bindGesture("kb:backspace", "handleCommandKey")
        for char in "abcdefghijklmnopqrstuvwxyz0123456789":
            self.bindGesture(f"kb:{char}", "handleCommandKey")
        for i in range(1, 10):
            self.bindGesture(f"kb:shift+{i}", "handleCommandKey")
    
    script_triggerCommandPrefix.__doc__ = _("Triggers the BOA command prefix for PowerPoint. Press this, followed by a specific command key.")

    def script_cancelCommandPrefix(self, gesture):
        import tones
        tones.beep(300, 50)
        self._clear_command_bindings()
        
    def script_handleCommandKey(self, gesture):
        import tones
        try:
            kb_id = list(gesture.identifiers)[-1]
            key = kb_id.split(":", 1)[1].lower() if ":" in kb_id else gesture.displayName.lower()
        except Exception:
            key = gesture.displayName.lower()
            
        self._clear_command_bindings()
        
        obj = api.getFocusObject()
        handled = ppt_manager.handle_prefix_command(key, obj)
        if not handled:
            tones.beep(150, 50)

    __gestures = {
        "kb:NVDA+e": "triggerCommandPrefix"
    }
