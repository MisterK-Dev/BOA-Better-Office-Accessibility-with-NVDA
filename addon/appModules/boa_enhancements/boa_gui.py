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
import wx  # noqa: E402
from gui.settingsDialogs import SettingsPanel  # noqa: E402
from appModules.boa_enhancements import boa_config  # noqa: E402
from logHandler import log  # noqa: E402

class BOASettingsPanel(SettingsPanel):
	"""
	The WX GUI Panel for BOA settings.
	WHY THIS EXISTS: Provides the visual layer for boa_config.json. It renders checkboxes grouped
	by application (Word, Excel, PowerPoint) to allow users to selectively disable specific hooks.
	This granular control is critical for preventing lock-ins if a future Microsoft Office update
	natively fixes a bug and our hook becomes obsolete.
	"""
	# Title for the NVDA settings dialog list and the panel title
	# Translators: Title of the BOA general settings dialog.
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
					# Translators: Options for the sheet layout announcement mode.
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
					# Translators: Tooltip explaining the layout announcement mode setting.
					cb.SetToolTip(_("Select layout announcement mode. 'Guided Announcement' requires you to press NVDA+E, L at least once per sheet to scan the layout."))
					sizer.Add(cb)
					group_sizer.Add(sizer, flag=wx.BOTTOM, border=5)
					self.checkboxes[app_name][feature_key] = cb
				elif feature_key == "end_of_data_radar":
					sizer = wx.BoxSizer(wx.HORIZONTAL)
					lbl = wx.StaticText(self, label=label)
					sizer.Add(lbl, flag=wx.RIGHT | wx.ALIGN_CENTER_VERTICAL, border=5)
					# Translators: Options for the engine used to detect the end of data.
					choices = [_("Off"), _("Strict Memory Check (CountA)"), _("Visible Data Only (Math Engine)")]
					cb = wx.Choice(self, choices=choices)
					val = config.get(app_name, {}).get(feature_key, "counta")
					
					if val == "off":
						cb.SetSelection(0)
					elif val == "visible":
						cb.SetSelection(2)
					else:
						cb.SetSelection(1)
					# Translators: Tooltip explaining the end of data radar engine setting.
					cb.SetToolTip(_("Select the engine used to detect the end of data in a row or column."))
					sizer.Add(cb)
					group_sizer.Add(sizer, flag=wx.BOTTOM, border=5)
					self.checkboxes[app_name][feature_key] = cb
				elif feature_key == "read_word_notes_inline":
					sizer = wx.BoxSizer(wx.HORIZONTAL)
					lbl = wx.StaticText(self, label=label)
					sizer.Add(lbl, flag=wx.RIGHT | wx.ALIGN_CENTER_VERTICAL, border=5)
					# Translators: Options for inline reading of Word notes
					choices = [_("Default NVDA Behavior (Reference Only)"), _("Read Note Content Inline")]
					cb = wx.Choice(self, choices=choices)
					val = config.get(app_name, {}).get(feature_key, "native")
					
					if val == "inline":
						cb.SetSelection(1)
					else:
						cb.SetSelection(0)
					# Translators: Tooltip for inline reading of Word notes.
					cb.SetToolTip(_("Choose whether NVDA should automatically read the text of footnotes and endnotes inline with the document text."))
					sizer.Add(cb)
					group_sizer.Add(sizer, flag=wx.BOTTOM, border=5)
					self.checkboxes[app_name][feature_key] = cb
				elif feature_key == "canvas_audio_mode":
					sizer = wx.BoxSizer(wx.HORIZONTAL)
					lbl = wx.StaticText(self, label=label)
					sizer.Add(lbl, flag=wx.RIGHT | wx.ALIGN_CENTER_VERTICAL, border=5)
					# Translators: Options for the PowerPoint Shape Movement Audio mode.
					choices = [
						# Translators: Automatically added by BOA compliance auditor
						_("Default NVDA Behavior"),
						# Translators: Automatically added by BOA compliance auditor
						_("Default Behavior + 3D Spatial Audio (Experimental)")
					]
					cb = wx.Choice(self, choices=choices)
					val = config.get(app_name, {}).get(feature_key, "default")
					
					# Fallback logic for users who had previously selected removed options
					if val in ("default_sound", "canvas", "canvas_sound"):
						cb.SetSelection(1)
					else:
						cb.SetSelection(0)
					
						
					# Translators: Tooltip explaining the Audio Canvas mode setting.
					cb.SetToolTip(_("Controls how NVDA announces shape movement in PowerPoint. Canvas modes translate raw math coordinates into human-readable locations like 'Top-Left Quadrant'."))
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
			# Translators: Label for the Bulk Sheet Organizer setting checkbox.
			"grid_mover": _("&Excel: Enable Bulk Sheet Organizer and Quick Sheet Mover"),
			# Translators: Label for the accessible Sheet Rename setting checkbox.
			"sheet_rename": _("Use accessible Sheet &Rename dialog instead of native edit field"),
			# Translators: Label for the SafeRichEdit in Excel setting checkbox.
			"safe_rich_edit": _("Prevent NVDA &crashes in Excel text fields (SafeRichEdit)"),
			# Translators: Label for the multi-cell selection tracking setting checkbox.
			"unselect_tracking": _("Announce when a m&ulti-cell selection is unexpectedly lost"),
			# Translators: Label for the hidden row skip announcement setting checkbox.
			"hidden_row_skip": _("Proactively announce when navigating past &hidden rows or columns"),
			# Translators: Label for the Sheet Layout Analyzer setting checkbox.
			"sheet_layout_analyzer": _("Enable Sheet &Layout Analyzer via NVDA+E, L"),
			# Translators: Label for the auto-announce mode dropdown setting.
			"auto_announce_first_block": _("Sheet Layout Auto-Announce &Mode:"),
			# Translators: Label for the conditional formatting tracking setting checkbox.
			"conditional_formatting": _("Conditional &Formatting and color"),
			# Translators: Label for the Cell Monitor setting checkbox.
			"cell_monitor": _("Enable Cell Moni&tor (slots 1-9 and continuous background monitoring)"),
			# Translators: Label for the End of Data Radar setting checkbox.
			"end_of_data_radar": _("Announce when there is no more data in the direction you are &navigating"),
			# Translators: Label for the Formula Auditor features setting checkbox.
			"formula_auditing_announcements": _("Enable Formula Auditor &features"),
			# Translators: Label for the merged cell tracking setting checkbox.
			"merged_cell_tracking": _("Explicitly announce m&erged cells")
		}
		# Translators: Title of the group box containing Excel specific settings.
		create_group("excel", _("Excel Enhancements"), excel_features)
		
		# PowerPoint Group
		ppt_features = {
			# Translators: Label for the PowerPoint color grid setting checkbox.
			"standard_color_grid": _("&PowerPoint: Read hidden Hex codes when navigating the Standard Color hexagon grid"),
			# Translators: Label for the PowerPoint hex color edit field setting checkbox.
			"hex_edit": _("Ensure the He&x Color edit field is properly labeled"),
			# Translators: Label for the PowerPoint RGB color edit field setting checkbox.
			"rgb_edit": _("Ensure the R&GB Color edit fields are properly labeled"),
			# Translators: Label for the SafeRichEdit in PowerPoint setting checkbox.
			"safe_rich_edit": _("Prevent NVDA cr&ashes in PowerPoint text fields (SafeRichEdit)"),
			# Translators: Label for the PowerPoint Bulk Slide Organizer setting checkbox.
			"bulk_slide_organizer": _("Enable Bulk Slide Organi&zer dialog (Prefix + X)"),
			# Translators: Label for the PowerPoint Slide Layout Analyzer setting checkbox.
			"slide_layout_analyzer": _("Enable S&lide Layout Analyzer (Prefix + L)"),
			# Translators: Label for the PowerPoint Document Analyzer setting checkbox.
			"document_analyzer": _("Enable PowerPoint &Document Analyzer (Prefix + D)"),
			# Translators: Label for the PowerPoint Canvas Audio Mode setting dropdown.
			"canvas_audio_mode": _("Shape Movement Audio Mode:")
		}
		# Translators: Title of the group box containing PowerPoint specific settings.
		create_group("powerpoint", _("PowerPoint Enhancements"), ppt_features)
		
		word_features = {
			# Translators: Label for the SafeRichEdit in Word setting checkbox.
			"safe_rich_edit": _("&Word: Prevent NVDA crashes in Word text fields (SafeRichEdit)"),
			# Translators: Label for the list double-read fix in Word setting checkbox.
			"fix_list_double_read": _("Prevent &double-reading of list items during paragraph navigation"),
			# Translators: Label for the document analyzer setting checkbox.
			"document_analyzer": _("Enable Document &Analyzer via NVDA+E, D"),
			# Translators: Label for the format auditor setting checkbox.
			"format_auditor": _("Enable Format &Auditor via NVDA+E, F"),
			# Translators: Label for reading word notes inline setting.
			"read_word_notes_inline": _("Footnotes and Endnotes reading &mode:")
		}
		# Translators: Title of the group box containing Word specific settings.
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
				elif feature_key == "read_word_notes_inline":
					sel = cb.GetSelection()
					if sel == 1:
						boa_config.set_feature_state(app, feature_key, "inline")
					else:
						boa_config.set_feature_state(app, feature_key, "native")
				elif feature_key == "canvas_audio_mode":
					sel = cb.GetSelection()
					if sel == 1:
						val = "default_sound"
					elif sel == 2:
						val = "canvas"
					elif sel == 3:
						val = "canvas_sound"
					else:
						val = "default"
					boa_config.set_feature_state(app, feature_key, val)
				else:
					boa_config.set_feature_state(app, feature_key, cb.GetValue())
		
		# Flush the updated configuration from memory to the physical JSON file
		boa_config.save_config()
		log.info("BOA settings saved successfully.")
