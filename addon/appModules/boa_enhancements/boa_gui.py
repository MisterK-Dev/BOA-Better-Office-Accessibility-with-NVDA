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
	
	def makeSettings(self, settingsSizer: wx.BoxSizer) -> None:
		"""
		Build the GUI for the settings panel using an accessible tabbed notebook.

		Architectural Intent:
		Rather than presenting a monolithic, 25+ control vertical list that causes navigation
		fatigue, we organize settings into distinct application tabs (Excel, Word, PowerPoint)
		via wx.Notebook. This lets screen reader users switch between applications instantly
		using Left/Right arrows or Ctrl+PageUp/Down, while preserving clean control parenting
		and full backward compatibility with boa_config.json.
		"""
		self.checkboxes = {}
		
		# Fetch the entire configuration dictionary from memory to populate initial UI state
		config = boa_config.get_all_config()
		
		self.notebook = wx.Notebook(self)

		def create_tab(app_name: str, tab_title: str, features: dict[str, str]) -> None:
			"""
			Dynamically generates a tab page populated with application-specific configuration controls.

			Architectural Considerations:
			- Strict Parenting: All child controls must be parented to `page` (the tab panel),
			  never to `self` (the outer SettingsPanel), to prevent controls from bleeding across inactive tabs.
			- Data Binding: Controls are registered into `self.checkboxes[app_name]` so that
			  `onSave()` can cleanly iterate and persist state without any breaking changes.
			"""
			page = wx.Panel(self.notebook)
			page_sizer = wx.BoxSizer(wx.VERTICAL)
			self.checkboxes[app_name] = {}

			for feature_key, label in features.items():
				if feature_key == "auto_announce_first_block":
					sizer = wx.BoxSizer(wx.HORIZONTAL)
					lbl = wx.StaticText(page, label=label)
					sizer.Add(lbl, flag=wx.RIGHT | wx.ALIGN_CENTER_VERTICAL, border=5)
					# Translators: Options for the sheet layout announcement mode.
					choices = [_("Off"), _("One-time Announcement"), _("Guided Announcement")]
					cb = wx.Choice(page, choices=choices)
					val = config.get(app_name, {}).get(feature_key, "one_time")
					if val == "off":
						cb.SetSelection(0)
					elif val == "guided":
						cb.SetSelection(2)
					else:
						cb.SetSelection(1)
					# Translators: Tooltip explaining the layout announcement mode setting.
					cb.SetToolTip(_("Select layout announcement mode. 'Guided Announcement' requires you to press NVDA+E, L at least once per sheet to scan the layout."))
					sizer.Add(cb)
					page_sizer.Add(sizer, flag=wx.BOTTOM, border=5)
					self.checkboxes[app_name][feature_key] = cb
				elif feature_key == "end_of_data_radar":
					sizer = wx.BoxSizer(wx.HORIZONTAL)
					lbl = wx.StaticText(page, label=label)
					sizer.Add(lbl, flag=wx.RIGHT | wx.ALIGN_CENTER_VERTICAL, border=5)
					# Translators: Options for the engine used to detect the end of data.
					choices = [_("Off"), _("Strict Memory Check (CountA)"), _("Visible Data Only (Math Engine)")]
					cb = wx.Choice(page, choices=choices)
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
					page_sizer.Add(sizer, flag=wx.BOTTOM, border=5)
					self.checkboxes[app_name][feature_key] = cb
				elif feature_key == "read_word_notes_inline":
					sizer = wx.BoxSizer(wx.HORIZONTAL)
					lbl = wx.StaticText(page, label=label)
					sizer.Add(lbl, flag=wx.RIGHT | wx.ALIGN_CENTER_VERTICAL, border=5)
					# Translators: Options for inline reading of Word notes
					choices = [_("Default NVDA Behavior (Reference Only)"), _("Read Note Content Inline")]
					cb = wx.Choice(page, choices=choices)
					val = config.get(app_name, {}).get(feature_key, "native")
					if val == "inline":
						cb.SetSelection(1)
					else:
						cb.SetSelection(0)
					# Translators: Tooltip for inline reading of Word notes.
					cb.SetToolTip(_("Choose whether NVDA should automatically read the text of footnotes and endnotes inline with the document text."))
					sizer.Add(cb)
					page_sizer.Add(sizer, flag=wx.BOTTOM, border=5)
					self.checkboxes[app_name][feature_key] = cb
				elif feature_key == "canvas_audio_mode":
					sizer = wx.BoxSizer(wx.HORIZONTAL)
					lbl = wx.StaticText(page, label=label)
					sizer.Add(lbl, flag=wx.RIGHT | wx.ALIGN_CENTER_VERTICAL, border=5)
					# Translators: Options for the PowerPoint Shape Movement Audio mode.
					choices = [
						# Translators: Automatically added by BOA compliance auditor
						_("Default NVDA Behavior"),
						# Translators: Automatically added by BOA compliance auditor
						_("Default Behavior + 3D Spatial Audio (Experimental)")
					]
					cb = wx.Choice(page, choices=choices)
					val = config.get(app_name, {}).get(feature_key, "default")
					if val in ("default_sound", "canvas", "canvas_sound"):
						cb.SetSelection(1)
					else:
						cb.SetSelection(0)
					# Translators: Tooltip explaining the Audio Canvas mode setting.
					cb.SetToolTip(_("Controls how NVDA announces shape movement in PowerPoint. Canvas modes translate raw math coordinates into human-readable locations like 'Top-Left Quadrant'."))
					sizer.Add(cb)
					page_sizer.Add(sizer, flag=wx.BOTTOM, border=5)
					self.checkboxes[app_name][feature_key] = cb
				else:
					cb = wx.CheckBox(page, label=label)
					cb.SetValue(config.get(app_name, {}).get(feature_key, True))
					page_sizer.Add(cb, flag=wx.BOTTOM, border=5)
					self.checkboxes[app_name][feature_key] = cb

			outer_sizer = wx.BoxSizer(wx.VERTICAL)
			outer_sizer.Add(page_sizer, proportion=1, flag=wx.ALL | wx.EXPAND, border=10)
			page.SetSizer(outer_sizer)
			self.notebook.AddPage(page, tab_title)

		# Excel Features
		excel_features = {
			# Translators: Label for the Bulk Sheet Organizer setting checkbox.
			"grid_mover": _("Enable &Bulk Sheet Organizer and Quick Sheet Mover"),
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
			"merged_cell_tracking": _("Explicitly announce mer&ged cells")
		}
		# Translators: Title of the Excel settings tab.
		create_tab("excel", _("&Excel"), excel_features)

		# Word Features
		word_features = {
			# Translators: Label for the SafeRichEdit in Word setting checkbox.
			"safe_rich_edit": _("Prevent NVDA &crashes in Word text fields (SafeRichEdit)"),
			# Translators: Label for the list double-read fix in Word setting checkbox.
			"fix_list_double_read": _("Prevent &double-reading of list items during paragraph navigation"),
			# Translators: Label for the document analyzer setting checkbox.
			"document_analyzer": _("Enable Document &Analyzer via NVDA+E, D"),
			# Translators: Label for the format auditor setting checkbox.
			"format_auditor": _("Enable Format &Auditor via NVDA+E, F"),
			# Translators: Label for reading word notes inline setting.
			"read_word_notes_inline": _("Footnotes and Endnotes reading &mode:")
		}
		# Translators: Title of the Word settings tab.
		create_tab("word", _("&Word"), word_features)

		# PowerPoint Features
		ppt_features = {
			# Translators: Label for the PowerPoint color grid setting checkbox.
			"standard_color_grid": _("Read hidden &hex codes when navigating the Standard Color hexagon grid"),
			# Translators: Label for the PowerPoint hex color edit field setting checkbox.
			"hex_edit": _("Ensure the Hex Color edit &field is properly labeled"),
			# Translators: Label for the PowerPoint RGB color edit field setting checkbox.
			"rgb_edit": _("Ensure the R&GB Color edit fields are properly labeled"),
			# Translators: Label for the SafeRichEdit in PowerPoint setting checkbox.
			"safe_rich_edit": _("Prevent NVDA cr&ashes in PowerPoint text fields (SafeRichEdit)"),
			# Translators: Label for the PowerPoint Bulk Slide Organizer setting checkbox.
			"bulk_slide_organizer": _("Enable Bulk Slide Organi&zer dialog (Prefix + X)"),
			# Translators: Label for the PowerPoint Slide Layout Analyzer setting checkbox.
			"slide_layout_analyzer": _("Enable &Slide Layout Analyzer (Prefix + L)"),
			# Translators: Label for the PowerPoint Document Analyzer setting checkbox.
			"document_analyzer": _("Enable PowerPoint &Document Analyzer (Prefix + D)"),
			# Translators: Label for the PowerPoint Canvas Audio Mode setting dropdown.
			"canvas_audio_mode": _("Shape Movement &Audio Mode:")
		}
		# Translators: Title of the PowerPoint settings tab.
		create_tab("powerpoint", _("&PowerPoint"), ppt_features)

		settingsSizer.Add(self.notebook, proportion=1, flag=wx.EXPAND | wx.ALL, border=5)
		self.Bind(wx.EVT_CHAR_HOOK, self.onCharHook)

	def onCharHook(self, event: wx.KeyEvent) -> None:
		"""
		Handles Alt+Letter accelerator keys to switch between notebook tabs.

		Architectural Intent:
		Windows SysTabControl32 does not natively register Alt+Letter mnemonics for tab switching.
		This hook catches Alt+E (Excel), Alt+W (Word), and Alt+P (PowerPoint) and programmatically
		selects the corresponding tab and focuses the notebook so NVDA announces the change.
		"""
		if event.AltDown() and not event.ControlDown() and not event.ShiftDown():
			key = event.GetKeyCode()
			if key in (ord('E'), ord('e')):
				self.notebook.SetSelection(0)
				self.notebook.SetFocus()
				return
			elif key in (ord('W'), ord('w')):
				self.notebook.SetSelection(1)
				self.notebook.SetFocus()
				return
			elif key in (ord('P'), ord('p')):
				self.notebook.SetSelection(2)
				self.notebook.SetFocus()
				return
		event.Skip()

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
