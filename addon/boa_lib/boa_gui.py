"""
BOA Settings GUI Module

This module integrates BOA configuration into the native NVDA settings dialog.
WHY THIS EXISTS (Architecture intent):
We want a seamless user experience where BOA settings are located alongside standard NVDA preferences.
By inheriting from gui.settingsDialogs.SettingsPanel, we natively inject our panel into the NVDA
Preferences menu, avoiding clunky standalone WX dialogs that break NVDA UI consistency.
"""
import wx
import gui
from gui.settingsDialogs import SettingsPanel
import boa_config
from logHandler import log

class BOASettingsPanel(SettingsPanel):
    """
    The WX GUI Panel for BOA settings.
    WHY THIS EXISTS: Provides the visual layer for boa_config.json. It renders checkboxes grouped
    by application (Word, Excel, PowerPoint) to allow users to selectively disable specific hooks.
    This granular control is critical for preventing lock-ins if a future Microsoft Office update
    natively fixes a bug and our hook becomes obsolete.
    """
    # Title for the NVDA settings dialog list and the panel title
    title = "BOA Office Enhancements"
    
    def makeSettings(self, settingsSizer):
        """Build the GUI for the settings panel."""
        self.checkboxes = {}
        
        # Ensure config is loaded
        config = boa_config.get_all_config()
        
        # Helper to create a group
        def create_group(app_name, display_name, features):
            group_sizer = wx.StaticBoxSizer(wx.VERTICAL, self, display_name)
            self.checkboxes[app_name] = {}
            for feature_key, label in features.items():
                cb = wx.CheckBox(self, label=label)
                # Set initial value from config
                cb.SetValue(config.get(app_name, {}).get(feature_key, True))
                group_sizer.Add(cb, flag=wx.BOTTOM, border=5)
                self.checkboxes[app_name][feature_key] = cb
            settingsSizer.Add(group_sizer, flag=wx.ALL | wx.EXPAND, border=10)

        # Excel Group
        excel_features = {
            "grid_mover": "ExcelGridMover (Arrow Keystroke Interceptor)",
            "sheet_rename": "ExcelSheetRenameEdit (Sheet Renaming Fix)",
            "safe_rich_edit": "SafeRichEdit (Excel RichEdit Override)",
            "unselect_tracking": "Unselect Tracking (Selection Loss Notification)"
        }
        create_group("excel", "Excel Enhancements", excel_features)
        
        # PowerPoint Group
        ppt_features = {
            "standard_color_grid": "PowerPointStandardColorGrid (Standard Colors TAB Fix)",
            "hex_edit": "PowerPointHexEdit (Hex Color Edit Override)",
            "rgb_edit": "PowerPointRGBEdit (RGB Dialog Override)",
            "safe_rich_edit": "SafeRichEdit (PowerPoint RichEdit Override)"
        }
        create_group("powerpoint", "PowerPoint Enhancements", ppt_features)
        
        # Word Group
        word_features = {
            "safe_rich_edit": "SafeRichEdit (Word RichEdit Override)"
        }
        create_group("word", "Word Enhancements", word_features)

    def onSave(self):
        """Called when the user presses OK or Apply."""
        for app, features in self.checkboxes.items():
            for feature_key, cb in features.items():
                boa_config.set_feature_state(app, feature_key, cb.GetValue())
        boa_config.save_config()
        log.info("BOA settings saved successfully.")
