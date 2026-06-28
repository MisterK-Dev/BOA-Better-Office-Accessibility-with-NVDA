# -*- coding: UTF-8 -*-
# Copyright (C) 2026 KIRAN G T {MisterK} and Antigravity 2
# This file is covered by the GNU General Public License (GPL), version 2.
# See the file COPYING.txt for more details.

import addonHandler
addonHandler.initTranslation()

import wx  # noqa: E402
import gui  # noqa: E402

class BulkSlideOrganizer(object):
	"""
	Handles launching and applying bulk slide rearrangement operations.
	ARCHITECTURAL INTENT: Reordering multiple slides in PowerPoint natively requires dragging slides 
	with a mouse in the visual Slide Sorter. This class provides an
	accessible, programmatic way to queue multiple slide moves and execute them via COM
	all at once.
	"""
	@staticmethod
	def _get_ppt_com():
		"""
		Safely retrieves the PowerPoint COM Application object bypassing the ROT.
		Architectural Intent: GetActiveObject can fail due to UAC boundaries or lazy COM registration.
		Instead, we traverse the HWND tree from the main PPTFrameClass to the inner document window
		and forcefully extract the COM object via MSAA NativeOM injection.
		Returns (ppt_application, hwnd_main)
		"""
		import ctypes
		import comtypes.client
		import comtypes.automation
		
		hwnd_main = ctypes.windll.user32.FindWindowW("PPTFrameClass", None)
		if not hwnd_main:
			return None, None
			
		target_hwnd = [None]
		WNDENUMPROC = ctypes.WINFUNCTYPE(ctypes.c_bool, ctypes.c_int, ctypes.c_void_p)
		
		def enum_callback(hwnd, lParam):
			class_name = ctypes.create_unicode_buffer(256)
			ctypes.windll.user32.GetClassNameW(hwnd, class_name, 256)
			if class_name.value in ("mdiClass", "paneClassDC"):
				target_hwnd[0] = hwnd
				return False
			return True
			
		ctypes.windll.user32.EnumChildWindows(hwnd_main, WNDENUMPROC(enum_callback), 0)
		
		if not target_hwnd[0]:
			return None, hwnd_main
			
		oleacc = ctypes.windll.user32.oleacc if hasattr(ctypes.windll.user32, 'oleacc') else ctypes.windll.oleacc
		ptr = ctypes.POINTER(comtypes.automation.IDispatch)()
		res = oleacc.AccessibleObjectFromWindow(target_hwnd[0], -16, ctypes.byref(comtypes.automation.IDispatch._iid_), ctypes.byref(ptr))
		
		if res == 0 and ptr:
			try:
				doc_window = comtypes.client.dynamic.Dispatch(ptr)
				# Extract the main application object from the child document window
				return doc_window.Application, hwnd_main
			except Exception:
				pass
				
		return None, hwnd_main

	@staticmethod
	def _show_bulk_dialog(slides_info, hwnd):
		"""
		Creates and displays the wx.Dialog. Parses the result and triggers COM execution.
		slides_info is a list of tuples: [(SlideID, SlideName/Title, OriginalIndex), ...]
		"""
		gui.mainFrame.prePopup()
		dlg = PowerPointBulkSlideOrganizerDialog(gui.mainFrame, slides_info)
		dlg.Raise()
		res = dlg.ShowModal()
		planned_moves = dict(dlg.planned_moves) if res == wx.ID_OK else None
		dlg.Destroy()
		gui.mainFrame.postPopup()
		
		if planned_moves:
			def _apply_moves():
				"""
				Applies the scheduled moves in a background thread to prevent NVDA's
				speech engine from stalling during heavy COM operations.
				"""
				import threading
				import comtypes
				import ui
				import core
				
				def _background_move_task():
					comtypes.CoInitialize()
					try:
						import logHandler
						
						ppt, _ = BulkSlideOrganizer._get_ppt_com()
						if not ppt:
							wx.CallAfter(ui.message, _("Failed to connect to PowerPoint."))
							return
							
						pres = ppt.ActivePresentation
						
						if not pres:
							wx.CallAfter(ui.message, _("No active presentation."))
							return
							
						total_slides = pres.Slides.Count
						if total_slides <= 1:
							wx.CallAfter(ui.message, _("Not enough slides to move."))
							return

						# Build the final array order
						unmoved = [s[0] for s in slides_info if s[1] not in planned_moves]
						
						# Sort moved slides by their intended target position
						moved = [(s[0], planned_moves[s[1]]) for s in slides_info if s[1] in planned_moves]
						moved.sort(key=lambda x: x[1])
						
						final_list = unmoved[:]
						for slide_id, pos in moved:
							insert_idx = min(pos - 1, len(final_list))
							final_list.insert(insert_idx, slide_id)
							
						# 1. Focus Lock: Capture the currently active slide before moves
						active_slide_id = None
						try:
							active_slide_id = ppt.ActiveWindow.View.Slide.SlideID
						except Exception:
							pass
							
						# Apply the moves natively via COM using the immutable SlideIDs
						for i, slide_id in enumerate(final_list):
							target_index = i + 1
							slide = pres.Slides.FindBySlideID(slide_id)
							if slide.SlideIndex != target_index:
								slide.MoveTo(target_index)
								
						# 2. Focus Lock: Restore selection to the original slide
						if active_slide_id:
							try:
								restored_slide = pres.Slides.FindBySlideID(active_slide_id)
								restored_slide.Select()
							except Exception:
								pass
								
						# 3. Delayed Success Message: Give PowerPoint 500ms to settle its focus events
						# Translators: Success message when slides have been reorganized
						wx.CallAfter(core.callLater, 500, ui.message, _("Successfully reorganized slides."))
						
					except Exception as e:
						import logHandler
						logHandler.log.error(f"BOA PPT SlideMover COM error: {e}", exc_info=True)
						wx.CallAfter(ui.message, _("Failed to move slides. See log for details."))
					finally:
						comtypes.CoUninitialize()

				import winUser
				winUser.setForegroundWindow(hwnd)
				t = threading.Thread(target=_background_move_task)
				t.daemon = True
				t.start()
				
			import core
			core.callLater(100, _apply_moves)

	@staticmethod
	def launch_dialog(obj):
		"""
		Connects to PowerPoint, grabs a list of all current slides and titles, 
		and then opens the custom Bulk Slide Organizer WX dialog.
		"""
		import ui
		import wx
		import threading
		import comtypes
		
		# Immediately inform the user so they know NVDA hasn't frozen
		ui.message(_("Loading slides, please wait..."))
		
		def _background_parse_task():
			comtypes.CoInitialize()
			try:
				ppt, hwnd_ppt = BulkSlideOrganizer._get_ppt_com()
				if not ppt:
					# Translators: Error message when trying to organize slides but no presentation is active.
					wx.CallAfter(ui.message, _("Failed to hook PowerPoint."))
					return
					
				pres = ppt.ActivePresentation
				if not pres:
					# Translators: Error message when trying to organize slides but no presentation is active.
					wx.CallAfter(ui.message, _("No active presentation."))
					return
					
				total_slides = pres.Slides.Count
				if total_slides <= 1:
					wx.CallAfter(ui.message, _("Not enough slides to move."))
					return
					
				slides_info = []
				for i in range(1, total_slides + 1):
					slide = pres.Slides(i)
					slide_id = slide.SlideID
					# Try to extract the slide title
					title = ""
					try:
						if slide.Shapes.HasTitle:
							title = slide.Shapes.Title.TextFrame.TextRange.Text.strip()
					except Exception:
						pass
						
					if title:
						# Translators: Slide item format with a title (e.g., Slide 1 - Welcome)
						display_name = _("Slide {index} - {title}").format(index=i, title=title)
					else:
						# Translators: Slide item format without a title (e.g., Slide 1 - (No Title))
						display_name = _("Slide {index} - (No Title)").format(index=i)
						
					slides_info.append((slide_id, display_name, i))
				
				# Use wx.CallAfter to safely push the dialog creation onto NVDA's main GUI thread.
				wx.CallAfter(BulkSlideOrganizer._show_bulk_dialog, slides_info, hwnd_ppt)
			except Exception as e:
				# Translators: Error message when the organizer dialog fails to open.
				wx.CallAfter(ui.message, _("Error opening organizer"))
				import logHandler
				logHandler.log.error(f"PowerPoint BulkSlideOrganizer error: {e}", exc_info=True)
			finally:
				comtypes.CoUninitialize()
				
		t = threading.Thread(target=_background_parse_task)
		t.daemon = True
		t.start()

class PowerPointBulkSlideOrganizerDialog(wx.Dialog):
	"""
	A custom wxPython Dialog that provides a fully accessible interface for bulk moving slides.
	"""
	def __init__(self, parent, slides_info):
		# Translators: Title of the Bulk Slide Organizer dialog.
		super().__init__(parent, title=_("Bulk Slide Organizer"))
		self.slides_info = slides_info
		self.slide_names = [s[1] for s in self.slides_info]
		
		# Dictionary to track the user's requested moves before they press OK.
		self.planned_moves = {} 
		
		main_sizer = wx.BoxSizer(wx.VERTICAL)
		
		# --- Combo 1: Slide Selection ---
		row1 = wx.BoxSizer(wx.HORIZONTAL)
		# Translators: Label for the Slide Name field.
		row1.Add(wx.StaticText(self, label=_("Slide Name:")), 0, wx.ALL|wx.ALIGN_CENTER_VERTICAL, 5)
		self.cb_slide = wx.ComboBox(self, choices=self.slide_names, style=wx.CB_READONLY)
		if self.slide_names:
			self.cb_slide.SetSelection(0)
		self.cb_slide.Bind(wx.EVT_COMBOBOX, self.on_slide_change)
		row1.Add(self.cb_slide, 1, wx.ALL, 5)
		
		# --- Combo 2: Target Position Selection ---
		row2 = wx.BoxSizer(wx.HORIZONTAL)
		# Translators: Label for the Target Position field.
		row2.Add(wx.StaticText(self, label=_("Target Position:")), 0, wx.ALL|wx.ALIGN_CENTER_VERTICAL, 5)
		positions = [str(i) for i in range(1, len(self.slide_names) + 1)]
		self.cb_pos = wx.ComboBox(self, choices=positions, style=wx.CB_READONLY)
		if positions:
			self.cb_pos.SetSelection(0)
		self.cb_pos.Bind(wx.EVT_COMBOBOX, self.on_pos_change)
		row2.Add(self.cb_pos, 1, wx.ALL, 5)
		
		main_sizer.Add(row1, 0, wx.EXPAND)
		main_sizer.Add(row2, 0, wx.EXPAND)
		
		# --- List of Scheduled Moves ---
		# Translators: Label for the list of scheduled slide moves.
		main_sizer.Add(wx.StaticText(self, label=_("Scheduled Moves (Press Del to remove):")), 0, wx.LEFT|wx.TOP, 5)
		self.list_moves = wx.ListCtrl(self, size=(-1, 150), style=wx.LC_REPORT | wx.LC_SINGLE_SEL)
		# Translators: Column header for the slide name in the moves list.
		self.list_moves.InsertColumn(0, _("Slide"), width=250)
		# Translators: Column header for the target position in the moves list.
		self.list_moves.InsertColumn(1, _("Target Position"), width=100)
		self.list_moves.Bind(wx.EVT_LIST_KEY_DOWN, self.on_list_key_down)
		main_sizer.Add(self.list_moves, 1, wx.ALL|wx.EXPAND, 5)
		
		# --- OK / Cancel Buttons ---
		btn_sizer = wx.BoxSizer(wx.HORIZONTAL)
		# Translators: Label for the OK button.
		btn_ok = wx.Button(self, wx.ID_OK, label=_("&OK"))
		# Translators: Label for the Cancel button.
		btn_cancel = wx.Button(self, wx.ID_CANCEL, label=_("&Cancel"))
		btn_sizer.Add(btn_ok, 0, wx.ALL, 5)
		btn_sizer.Add(btn_cancel, 0, wx.ALL, 5)
		main_sizer.Add(btn_sizer, 0, wx.ALL|wx.ALIGN_RIGHT, 5)
		
		self.SetSizerAndFit(main_sizer)
		self.cb_slide.SetFocus()

	def on_slide_change(self, event):
		slide_name = self.cb_slide.GetValue()
		if slide_name in self.planned_moves:
			self.cb_pos.SetSelection(self.planned_moves[slide_name] - 1)
		else:
			if slide_name in self.slide_names:
				self.cb_pos.SetSelection(self.slide_names.index(slide_name))

	def on_pos_change(self, event):
		slide_name = self.cb_slide.GetValue()
		pos_str = self.cb_pos.GetValue()
		if slide_name and pos_str:
			pos = int(pos_str)
			self.planned_moves[slide_name] = pos
			self._refresh_list()
			import ui
			# Translators: Message announcing a move has been scheduled.
			ui.message(_("Scheduled: {slide} to position {pos}").format(slide=slide_name, pos=pos))

	def on_list_key_down(self, event):
		if event.GetKeyCode() == wx.WXK_DELETE:
			self.on_remove()
		event.Skip()

	def on_remove(self):
		sel = self.list_moves.GetFirstSelected()
		if sel != -1:
			slide_name = self.list_moves.GetItemText(sel, 0)
			if slide_name in self.planned_moves:
				del self.planned_moves[slide_name]
			self._refresh_list()
			import ui
			# Translators: Message announcing a scheduled slide move has been removed.
			ui.message(_("Move removed"))
			self.on_slide_change(None)
			
	def _refresh_list(self):
		self.list_moves.DeleteAllItems()
		for slide, pos in self.planned_moves.items():
			self.list_moves.Append([str(slide), str(pos)])
