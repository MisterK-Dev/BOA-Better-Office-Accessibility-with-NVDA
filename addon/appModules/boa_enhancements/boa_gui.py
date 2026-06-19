# -*- coding: UTF-8 -*-
# Copyright (C) 2026 KIRAN G T {MisterK} and Antigravity 2
# This file is covered by the GNU General Public License (GPL), version 2.
# See the file COPYING.txt for more details.

import addonHandler
addonHandler.initTranslation()

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
from appModules.boa_enhancements import boa_config
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
    title = _("BOA Office Enhancements")
    
    def makeSettings(self, settingsSizer):
        """
        Build the GUI for the settings panel.
        Architectural Why: This method is called natively by NVDA when building the preferences dialog.
        We dynamically generate standard WX checkboxes tied to our JSON config to provide
        an accessible, native-feeling configuration experience for the user.
        """
        self.checkboxes = {}
        
        # Ensure config is loaded
        # Fetch the entire configuration dictionary from memory to populate initial UI state
        config = boa_config.get_all_config()
        
        def create_group(app_name, display_name, features):
            """
            Dynamically generates a visually grouped subset of configuration controls (checkboxes/dropdowns).
            
            Architectural Intent & Considerations:
            Instead of manually hardcoding individual wx.CheckBox objects, this helper iterates through a 
            dictionary of features and binds them to the loaded JSON state. This keeps the UI code DRY 
            (Don't Repeat Yourself) and makes adding future hooks trivial. It strictly packs them into a 
            labeled `wx.StaticBoxSizer` to ensure NVDA correctly announces the group boundary.
            """
            group_sizer = wx.StaticBoxSizer(wx.VERTICAL, self, display_name)
            self.checkboxes[app_name] = {}
            for feature_key, label in features.items():
                if feature_key == "auto_announce_first_block":
                    sizer = wx.BoxSizer(wx.HORIZONTAL)
                    lbl = wx.StaticText(self, label=label)
                    sizer.Add(lbl, flag=wx.RIGHT | wx.ALIGN_CENTER_VERTICAL, border=5)
                    choices = [_("Off"), _("One-time Announcement"), _("Guided Announcement")]
                    cb = wx.Choice(self, choices=choices)
                    val = config.get(app_name, {}).get(feature_key, "one_time")
                    
                    # Map the stored string value from the JSON config to the correct zero-indexed UI dropdown selection.
                    # Index 0 = "Off", Index 1 = "One-time Announcement", Index 2 = "Guided Announcement".
                    if val == "off":
                        cb.SetSelection(0)
                    elif val == "guided":
                        cb.SetSelection(2)
                    else:
                        cb.SetSelection(1)
                    cb.SetToolTip(_("Select layout announcement mode. 'Guided Announcement' requires you to press NVDA+E, L at least once per sheet to scan the layout."))
                    sizer.Add(cb)
                    group_sizer.Add(sizer, flag=wx.BOTTOM, border=5)
                    self.checkboxes[app_name][feature_key] = cb
                elif feature_key == "end_of_data_radar":
                    sizer = wx.BoxSizer(wx.HORIZONTAL)
                    lbl = wx.StaticText(self, label=label)
                    sizer.Add(lbl, flag=wx.RIGHT | wx.ALIGN_CENTER_VERTICAL, border=5)
                    choices = [_("Off"), _("Strict Memory Check (CountA)"), _("Visible Data Only (Math Engine)")]
                    cb = wx.Choice(self, choices=choices)
                    val = config.get(app_name, {}).get(feature_key, "counta")
                    
                    if val == "off":
                        cb.SetSelection(0)
                    elif val == "visible":
                        cb.SetSelection(2)
                    else:
                        cb.SetSelection(1)
                    cb.SetToolTip(_("Select the engine used to detect the end of data in a row or column."))
                    sizer.Add(cb)
                    group_sizer.Add(sizer, flag=wx.BOTTOM, border=5)
                    self.checkboxes[app_name][feature_key] = cb
                else:
                    cb = wx.CheckBox(self, label=label)
                    # Set initial value from config
                    cb.SetValue(config.get(app_name, {}).get(feature_key, True))
                    group_sizer.Add(cb, flag=wx.BOTTOM, border=5)
                    self.checkboxes[app_name][feature_key] = cb
            settingsSizer.Add(group_sizer, flag=wx.ALL | wx.EXPAND, border=10)

        # Excel Group
        excel_features = {
            "grid_mover": _("&Excel: Enable Bulk Sheet Organizer and Quick Sheet Mover"),
            "sheet_rename": _("Use accessible Sheet &Rename dialog instead of native edit field"),
            "safe_rich_edit": _("Prevent NVDA &crashes in Excel text fields"),
            "unselect_tracking": _("Announce when a m&ulti-cell selection is unexpectedly lost"),
            "hidden_row_skip": _("Proactively announce when navigating past &hidden rows or columns"),
            "sheet_layout_analyzer": _("Enable Sheet &Layout Analyzer via NVDA+E, L"),
            "auto_announce_first_block": _("Sheet Layout Auto-Announce &Mode:"),
            "conditional_formatting": _("Conditional &Formatting and color"),
            "cell_monitor": _("Enable Cell Moni&tor (slots 1-9 and continuous background monitoring)"),
            "end_of_data_radar": _("Announce when there is no more data in the direction you are &navigating"),
            "formula_auditing_announcements": _("Announce Formula Auditing &Hotkeys (Ctrl+[, etc.)")
        }
        create_group("excel", _("Excel Enhancements"), excel_features)
        
        # PowerPoint Group
        ppt_features = {
            "standard_color_grid": _("&PowerPoint: Read hidden Hex codes when navigating the Standard Color hexagon grid"),
            "hex_edit": _("Ensure the He&x Color edit field is properly labeled"),
            "rgb_edit": _("Ensure the R&GB Color edit fields are properly labeled"),
            "safe_rich_edit": _("Prevent NVDA cr&ashes in PowerPoint text fields")
        }
        create_group("powerpoint", _("PowerPoint Enhancements"), ppt_features)
        
        # Word Group
        word_features = {
            "safe_rich_edit": _("&Word: Prevent NVDA crashes in Word text fields")
        }
        create_group("word", _("Word Enhancements"), word_features)

    def onSave(self):
        """
        Called when the user presses OK or Apply in the NVDA settings dialog.
        Architectural Why: We must persist the in-memory GUI state (checkboxes) back to the 
        file system (JSON) so that configuration persists across NVDA restarts.
        """
        # Iterate over each application (e.g., excel, word) and its respective GUI checkboxes
        for app, features in self.checkboxes.items():
            for feature_key, cb in features.items():
                if feature_key == "auto_announce_first_block":
                    sel = cb.GetSelection()
                    
                    # Reverse map the UI dropdown index back to the exact string expected by boa_config.json.
                    # This consideration ensures the physical JSON configuration file remains human-readable 
                    # strings (e.g., "off" or "guided") rather than cryptic GUI integers.
                    if sel == 0:
                        val = "off"
                    elif sel == 2:
                        val = "guided"
                    else:
                        val = "one_time"
                    boa_config.set_feature_state(app, feature_key, val)
                elif feature_key == "end_of_data_radar":
                    sel = cb.GetSelection()
                    if sel == 0:
                        val = "off"
                    elif sel == 2:
                        val = "visible"
                    else:
                        val = "counta"
                    boa_config.set_feature_state(app, feature_key, val)
                else:
                    boa_config.set_feature_state(app, feature_key, cb.GetValue())
        
        # Flush the updated configuration from memory to the physical JSON file
        boa_config.save_config()
        log.info("BOA settings saved successfully.")
