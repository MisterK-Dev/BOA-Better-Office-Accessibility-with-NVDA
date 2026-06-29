# -*- coding: UTF-8 -*-
# Copyright (C) 2026 KIRAN G T {MisterK} and Antigravity 2
# This file is covered by the GNU General Public License (GPL), version 2.
# See the file COPYING.txt for more details.

import wx
import ui
import gui
import addonHandler
from logHandler import log

addonHandler.initTranslation()

def get_active_cell():
	from .cell_navigation_tracker import _cached_excel_app
	if not _cached_excel_app:
		return None
	try:
		return _cached_excel_app.ActiveCell
	except Exception:
		return None

def speak_formula():
	"""Reads the formula out loud. If none, says 'No formula'."""
	cell = get_active_cell()
	if not cell:
		return
	try:
		has_formula = cell.HasFormula
		if has_formula:
			ui.message(cell.FormulaLocal)
		else:
			# Translators: Message when trying to edit a formula but the cell has none.
			ui.message(_("No formula"))
	except Exception as e:
		log.debug(f"BOA: Failed to speak formula: {e}")

class PowerEditorDialog(wx.Dialog):
	def __init__(self, parent, cell):
		super(PowerEditorDialog, self).__init__(
			# Translators: Title of the BOA Power Editor dialog.
			parent, title=_("BOA Power Editor"), style=wx.DEFAULT_DIALOG_STYLE | wx.RESIZE_BORDER
		)
		self.cell = cell
		
		mainSizer = wx.BoxSizer(wx.VERTICAL)
		
		# Determine initial content safely (FormulaLocal works for plain text too)
		try:
			val = cell.FormulaLocal
			if isinstance(val, float) and val.is_integer():
				initial_text = str(int(val))
			else:
				initial_text = str(val) if val is not None else ""
		except Exception:
			initial_text = ""
			
		# Translators: Instructions for the Power Editor dialog.
		helpLabel = wx.StaticText(self, label=_("&Edit cell content (press Ctrl+Enter to save, Esc to cancel):"))
		mainSizer.Add(helpLabel, 0, wx.ALL, 5)
		
		# Multiline text control (removed TE_PROCESS_ENTER so native Enter creates a new line)
		self.editor = wx.TextCtrl(self, style=wx.TE_MULTILINE, size=(500, 300))
		self.editor.SetValue(initial_text)
		self.editor.SetSelection(-1, -1) # Select all
		self.editor.Bind(wx.EVT_CHAR_HOOK, self.onCharHook)
		mainSizer.Add(self.editor, 1, wx.EXPAND | wx.ALL, 5)
		
		btnSizer = wx.StdDialogButtonSizer()
		
		# Translators: Label for the Save button.
		saveBtn = wx.Button(self, wx.ID_OK, label=_("&Save"))
		saveBtn.Bind(wx.EVT_BUTTON, self.onSave)
		btnSizer.AddButton(saveBtn)
		
		# Translators: Label for the Cancel button.
		cancelBtn = wx.Button(self, wx.ID_CANCEL, label=_("&Cancel"))
		cancelBtn.Bind(wx.EVT_BUTTON, self.onCancel)
		btnSizer.AddButton(cancelBtn)
		
		btnSizer.Realize()
		mainSizer.Add(btnSizer, 0, wx.EXPAND | wx.ALL, 5)
		
		self.SetSizer(mainSizer)
		self.CenterOnParent()

	def onCharHook(self, evt):
		# Catch Ctrl+Enter to save
		if evt.GetKeyCode() == wx.WXK_RETURN and evt.ControlDown():
			self.onSave(evt)
		# Catch Escape to cancel
		elif evt.GetKeyCode() == wx.WXK_ESCAPE:
			self.onCancel(evt)
		else:
			evt.Skip()
		
	def onSave(self, evt):
		new_text = self.editor.GetValue()
		try:
			# Setting FormulaLocal natively supports plain text or local formulas seamlessly
			self.cell.FormulaLocal = new_text
			
			# Post-calculation error detection
			try:
				resulting_text = str(self.cell.Text).strip()
			except Exception:
				resulting_text = ""
				
			# Excel calculation errors usually start with # and end with ! or ?
			if resulting_text.startswith("#") and any(c in resulting_text for c in ['?', '!']):
				import gui
				import wx
				# Translators: Message warning that the saved formula resulted in an Excel error.
				msg = _("The formula was saved, but Excel evaluated it to an error: {error}\n\nWould you like to continue editing?").format(error=resulting_text)
				result = gui.messageBox(
					msg,
					# Translators: Title of the calculation error warning dialog.
					_("Calculation Error"),
					wx.YES_NO | wx.ICON_WARNING,
					self
				)
				if result == wx.YES:
					self.editor.SetFocus()
					return # Keep editor alive
				
			# Translators: Message confirming the cell edit was saved.
			ui.message(_("Done"))
			self.Destroy()
		except Exception as e:
			log.debug(f"BOA: Excel rejected formula: {e}")
			import wx
			# Bring up a warning, but DO NOT destroy the dialog
			wx.MessageBox(
				# Translators: Error message when Excel rejects a formula due to syntax.
				_("Excel rejected this formula. Please check for syntax errors or missing parentheses."),
				# Translators: Title of the invalid formula error dialog.
				_("Invalid Formula"),
				wx.OK | wx.ICON_ERROR,
				self
			)
			self.editor.SetFocus()

	def onCancel(self, evt):
		self.Destroy()

def _show_power_editor():
	cell = get_active_cell()
	if not cell:
		return
	gui.mainFrame.prePopup()
	dlg = PowerEditorDialog(gui.mainFrame, cell)
	dlg.ShowModal()
	gui.mainFrame.postPopup()

def open_power_editor():
	"""Spawns the multiline Power Editor safely on the main GUI thread."""
	wx.CallAfter(_show_power_editor)
