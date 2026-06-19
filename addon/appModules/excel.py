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
    scriptCategory = _("BOA (Better Office Accessibility)")

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
            # Safely grab the exact Process ID of the currently focused Excel instance without COM!
            from appModules.boa_enhancements.excel_enhancements import cell_navigation_tracker
            cell_navigation_tracker._cached_excel_pid = obj.processID
            
            from appModules.boa_enhancements import boa_config
            if (boa_config.get_feature_state("excel", "unselect_tracking") or 
                boa_config.get_feature_state("excel", "hidden_row_skip") or 
                boa_config.get_feature_state("excel", "auto_announce_first_block") in ["one_time", "guided"] or
                boa_config.get_feature_state("excel", "conditional_formatting")):
                check_unselect(obj)
        except Exception:
            pass
        nextHandler()

    def event_appModule_loseFocus(self, *args, **kwargs):
        try:
            import core
            from appModules.boa_enhancements.excel_enhancements.cell_navigation_tracker import release_if_closed
            # Defer by 200ms to allow Windows to update the focused HWND state before checking
            core.callLater(200, release_if_closed)
        except Exception:
            pass
            
        # Safely call nextHandler if it exists, or fallback to super() for NVDA's base method
        if args and len(args) >= 2 and callable(args[1]):
            args[1]()
        elif 'nextHandler' in kwargs and callable(kwargs['nextHandler']):
            kwargs['nextHandler']()
        else:
            try:
                super(AppModule, self).event_appModule_loseFocus(*args, **kwargs)
            except Exception:
                pass

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

    def event_UIA_invoke(self, obj, nextHandler):
        # Placeholder for future UIA invoke enhancements
        nextHandler()

    def _clear_command_bindings(self):
        try:
            self.removeGestureBinding("kb:escape")
            self.removeGestureBinding("kb:backspace")
            self.removeGestureBinding("kb:f2")
            for char in "abcdefghijklmnopqrstuvwxyz0123456789":
                self.removeGestureBinding(f"kb:{char}")
                self.removeGestureBinding(f"kb:shift+{char}")
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
        self.bindGesture("kb:f2", "handleCommandKey")
        for char in "abcdefghijklmnopqrstuvwxyz0123456789":
            self.bindGesture(f"kb:{char}", "handleCommandKey")
            self.bindGesture(f"kb:shift+{char}", "handleCommandKey")
    
    script_triggerCommandPrefix.__doc__ = _("Triggers the BOA command prefix. Press this, followed by a specific command key.")

    def script_cancelCommandPrefix(self, gesture):
        import tones
        tones.beep(300, 50)
        self._clear_command_bindings()
        
    def script_handleCommandKey(self, gesture):
        # Catches the subsequent key pressed after NVDA+E.
        # Routes to excel_manager.handle_prefix_command(key, obj).
        import tones
        try:
            kb_id = list(gesture.identifiers)[-1]
            key = kb_id.split(":", 1)[1].lower() if ":" in kb_id else gesture.displayName.lower()
        except Exception:
            key = gesture.displayName.lower()
            
        obj = api.getFocusObject()
        result = excel_manager.handle_prefix_command(key, obj)
        
        if result == "keep_alive":
            # Do NOT clear bindings yet. Give the user time to press a second key (like double-tapping F2).
            # Schedule a 500ms cleanup to ensure they don't get stuck in prefix mode.
            import core
            def delayed_cleanup():
                import scriptHandler
                # If they didn't trigger another prefix script, clean up.
                if scriptHandler.getLastScriptRepeatCount() == 0:
                    self._clear_command_bindings()
            core.callLater(500, delayed_cleanup)
        else:
            self._clear_command_bindings()
            if not result:
                tones.beep(150, 50)

    def script_bulkSheetOrganizer(self, gesture):
        obj = api.getFocusObject()
        excel_manager.show_bulk_sheet_organizer(obj)
    script_bulkSheetOrganizer.__doc__ = _("Launches the Bulk Sheet Organizer dialog to reorder sheets.")

    def script_announceSheetLayout(self, gesture):
        obj = api.getFocusObject()
        excel_manager.announce_sheet_layout(obj)
    script_announceSheetLayout.__doc__ = _("Analyzes and announces the layout structure of the current sheet.")

    def script_jumpToNearestBlock(self, gesture):
        obj = api.getFocusObject()
        excel_manager.jump_to_nearest_block(obj)
    script_jumpToNearestBlock.__doc__ = _("Jumps the focus to the nearest cell block.")

    def script_announceConditionalFormatting(self, gesture):
        obj = api.getFocusObject()
        excel_manager.announce_conditional_formatting(obj)
    script_announceConditionalFormatting.__doc__ = _("Announces conditional formatting details for the current cell.")

    def script_toggleCellMonitor(self, gesture):
        obj = api.getFocusObject()
        excel_manager.toggle_cell_monitor(obj)
    script_toggleCellMonitor.__doc__ = _("Toggles monitoring on the active cell.")

    def script_clearAllCellMonitors(self, gesture):
        obj = api.getFocusObject()
        excel_manager.clear_all_cell_monitors(obj)
    script_clearAllCellMonitors.__doc__ = _("Clears all monitored cells.")

    def script_tracePrecedents(self, gesture):
        from appModules.boa_enhancements.excel_enhancements import formula_auditor
        formula_auditor.trace_precedents()
    script_tracePrecedents.__doc__ = _("Traces precedents and opens a dialog to jump to cells that provide data to the active cell.")

    def script_traceDependents(self, gesture):
        from appModules.boa_enhancements.excel_enhancements import formula_auditor
        formula_auditor.trace_dependents()
    script_traceDependents.__doc__ = _("Traces dependents and opens a dialog to jump to cells that depend on the active cell.")

    __gestures = {
        "kb:NVDA+e": "triggerCommandPrefix"
    }
