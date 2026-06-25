# -*- coding: UTF-8 -*-
# Copyright (C) 2026 KIRAN G T {MisterK} and Antigravity 2
# This file is covered by the GNU General Public License (GPL), version 2.
# See the file COPYING.txt for more details.

import addonHandler
addonHandler.initTranslation()

import wx  # noqa: E402
import ui  # noqa: E402
from logHandler import log  # noqa: E402

class CellMonitorManager:
	"""
	Manages the 9 slots and continuous monitoring for Excel cells.
	
	Architectural Intent & Considerations:
	Windows DCOM security natively prevents standard COM event sinks for Excel cell changes 
	across out-of-process boundaries. Because we cannot "listen" for an event when a background 
	cell changes, we MUST use a non-blocking polling architecture (`core.callLater`). 
	This class manages both manual memory slots (1-9) and the active continuous background polling.
	"""
	_slots = {}  # Format: { "1": {"wb": "Book1", "sheet": "Sheet1", "cell": "$A$1", "val": "100"} }
	_monitors = {} # Format: { "Book1|Sheet1|$A$1": "100" }
	
	_isMonitoringActive = False

	@classmethod
	def _start_timer(cls, excelApp):
		"""
		Initiates the background polling loop for cell monitoring.
		
		Architectural Intent & Considerations:
		We use NVDA's native `core.callLater(150, ...)` instead of `wx.Timer` or standard `time.sleep`. 
		`time.sleep` would catastrophically block NVDA's single thread, freezing the entire screen reader. 
		`core.callLater` safely schedules the execution on NVDA's main thread loop every 150ms.
		"""
		cls._active_excel = excelApp
		if not cls._isMonitoringActive:
			cls._isMonitoringActive = True
			import core
			# 150ms feels instantaneous. We use NVDA's native core.callLater instead of wx.Timer for guaranteed execution.
			core.callLater(150, cls._check_all_monitors)
			log.info("BOA: Started Cell Monitor Loop.")

	@classmethod
	def _stop_timer(cls):
		"""
		Safely halts the continuous monitoring loop.
		
		Architectural Intent & Considerations:
		To prevent memory leaks and unnecessary CPU polling when no cells are being monitored, 
		we toggle the `_isMonitoringActive` flag. The loop function respects this flag and gracefully terminates.
		"""
		cls._isMonitoringActive = False
		cls._active_excel = None
		import gc
		gc.collect()
		log.info("BOA: Stopped Cell Monitor Loop.")

	@classmethod
	def _get_active_cell_info(cls, obj):
		"""
		Retrieves the exact workbook, sheet, address, and value of the currently focused cell.
		
		Architectural Intent & Considerations:
		Because NVDA objects don't always expose complete structural metadata natively, we must 
		query Excel's COM model directly to guarantee we have the absolute $A$1 address and parent 
		workbook name required for strict slot tracking.
		"""
		try:
			import comtypes.client
			import ctypes
			import comtypes.automation
			excel = None
			
			# Try getting active object first
			try:
				excel = comtypes.client.GetActiveObject("Excel.Application")
			except Exception:
				pass
				
			# Fallback Consideration: If GetActiveObject fails (common if Excel is in edit mode or blocked by security boundaries),
			# we must manually dig for the EXCEL7 window class handle to force a back-door connection.
			if not excel:
				hwnd7 = obj.windowHandle if getattr(obj, "windowClassName", "") == "EXCEL7" else None
				if hwnd7:
					# Dynamically load the oleacc library.
					oleacc = ctypes.windll.oleacc if hasattr(ctypes.windll, 'oleacc') else ctypes.windll.user32.oleacc
					ptr = ctypes.POINTER(comtypes.automation.IDispatch)()
					# OBJID_NATIVEOM (-16) retrieves the native COM object underneath the window.
					res = oleacc.AccessibleObjectFromWindow(hwnd7, -16, ctypes.byref(comtypes.automation.IDispatch._iid_), ctypes.byref(ptr))
					if res == 0 and ptr:
						# Safely cast the raw COM pointer back into a usable Python Excel Application object.
						excel = comtypes.client.dynamic.Dispatch(ptr).Application

			if excel:
				cell = excel.ActiveCell
				sheet = excel.ActiveSheet.Name
				wb = excel.ActiveWorkbook.Name
				address = cell.Address(False, False) # relative A1 without verbose $ symbols
				
				val = str(cell.Text) if cell.Text is not None else ""
				if "comtypes" in val.lower():
					val = ""
				val = val.strip()
				return excel, wb, sheet, address, val
		except Exception as e:
			log.debugWarning(f"BOA CellMonitor: Failed to get cell info: {e}")
		return None, None, None, None, None

	@classmethod
	def assign_slot(cls, slot_str, obj):
		"""
		Assigns the active cell to a memory slot (1-9).
		
		Architectural Intent & Considerations:
		Users need a way to quickly check specific cells without physically navigating to them.
		By assigning a cell to a dictionary slot, we cache its exact coordinate path. We also automatically 
		enable continuous monitoring for this cell so that if its value changes while the user is elsewhere, 
		they are immediately notified.
		"""
		excel, wb, sheet, address, val = cls._get_active_cell_info(obj)
		if not excel:
			# Translators: Error when the add-on cannot read a cell.
			ui.message(_("Error: Could not read Excel cell."))
			return

		is_replace = slot_str in cls._slots
		old_info = cls._slots.get(slot_str)

		cls._slots[slot_str] = {
			"wb": wb,
			"sheet": sheet,
			"cell": address,
			"val": val,
			"excel": excel
		}

		# Auto-enable continuous monitoring for slotted cells
		monitor_key = f"{wb}|{sheet}|{address}"
		cls._monitors[monitor_key] = val
		cls._start_timer(excel)

		# Remove old cell from monitors if it was only monitored via this slot
		if is_replace:
			old_monitor_key = f"{old_info['wb']}|{old_info['sheet']}|{old_info['cell']}"
			if old_monitor_key != monitor_key:
				# Check if the old cell is still assigned to ANY other slot
				is_still_slotted = False
				for s_key, s_info in cls._slots.items():
					if f"{s_info['wb']}|{s_info['sheet']}|{s_info['cell']}" == old_monitor_key:
						is_still_slotted = True
						break
				
				# If it's not in any other slot, delete it from background monitors
				if not is_still_slotted and old_monitor_key in cls._monitors:
					del cls._monitors[old_monitor_key]
			# Translators: Message when a monitored slot is overwritten.
			ui.message(_("{old_cell} has been replaced by {address} for slot {slot_str}").format(
				old_cell=old_info['cell'], address=address, slot_str=slot_str))
		else:
			# Translators: Message when a cell is successfully slotted.
			ui.message(_("{address} set to slot {slot_str}").format(address=address, slot_str=slot_str))

	@classmethod
	def read_slot(cls, slot_str, obj):
		if slot_str not in cls._slots:
			# Translators: Message when checking a slot that is currently empty.
			ui.message(_("Slot {slot_str} is empty.").format(slot_str=slot_str))
			return
			
		info = cls._slots[slot_str]
		
		excel = info.get("excel")
		active_excel, current_wb, current_sheet, _unused_addr, _unused_val = cls._get_active_cell_info(obj)
		
		if not excel:
			excel = active_excel
			
		if not excel:
			# Translators: Error when Excel cannot be accessed.
			ui.message(_("Error: Excel not accessible."))
			return

		try:
			target_wb = None
			try:
				target_wb = excel.Workbooks(info["wb"])
			except Exception:
				if active_excel:
					try:
						target_wb = active_excel.Workbooks(info["wb"])
					except Exception:
						pass
						
			if not target_wb:
				# Translators: Error when trying to read a slot from a closed workbook.
				ui.message(_("Slot {slot_str} lost. Workbook '{wb}' is closed or inaccessible.").format(
					slot_str=slot_str, wb=info['wb']))
				return
			
			sheet_names = [s.Name for s in target_wb.Sheets]
			if info["sheet"] not in sheet_names:
				# Translators: Error when trying to read a slot from a deleted sheet.
				ui.message(_("Slot {slot_str} lost. Sheet '{sheet}' was renamed or deleted.").format(
					slot_str=slot_str, sheet=info['sheet']))
				return
				
			target_sheet = target_wb.Sheets(info["sheet"])
			target_cell = target_sheet.Range(info["cell"])
			
			try:
				val = str(target_cell.Text)
				if "comtypes" in val.lower():
					val = ""
			except Exception:
				val = ""
			if not val or val.startswith("###"):
				raw_val = target_cell.Value
				val = str(raw_val) if raw_val is not None else ""
				if "comtypes" in val.lower():
					val = ""
			
			val = val.strip()
			info["val"] = val
			monitor_key = f"{info['wb']}|{info['sheet']}|{info['cell']}"
			if monitor_key in cls._monitors:
				cls._monitors[monitor_key] = val
				
			spoken_val = val
			if not spoken_val:
				spoken_val = _("Empty cell")
				
			if current_wb == info['wb'] and current_sheet == info['sheet']:
				location_str = ""
			elif current_wb == info['wb']:
				# Translators: Describes the sheet location of a monitored cell.
				location_str = _(" in {sheet}").format(sheet=info['sheet'])
			else:
				# Translators: Describes the sheet and workbook location of a monitored cell.
				location_str = _(" in {sheet} of {wb}").format(sheet=info['sheet'], wb=info['wb'])
				
			ui.message("{val} - {cell}{location_str}".format(
				val=spoken_val, cell=info['cell'], location_str=location_str))
		except Exception:
			# Translators: Error when Excel is too busy to read the slot.
			ui.message(_("Cannot read slot {slot_str}. Excel may be busy.").format(slot_str=slot_str))

	@classmethod
	def toggle_monitor(cls, obj):
		"""
		Toggles continuous background monitoring for the active cell.
		
		Architectural Intent & Considerations:
		Allows users to selectively monitor a single cell (like a total sum) without assigning it 
		to a specific 1-9 slot. If all monitors are cleared, it proactively shuts down the background 
		polling timer to conserve NVDA system resources.
		"""
		excel, wb, sheet, address, val = cls._get_active_cell_info(obj)
		if not excel:
			# Translators: Error when the add-on cannot read a cell.
			ui.message(_("Error: Could not read Excel cell."))
			return

		monitor_key = f"{wb}|{sheet}|{address}"
		
		if monitor_key in cls._monitors:
			del cls._monitors[monitor_key]
			# Translators: Message when continuous monitoring is turned off for a cell.
			ui.message(_("Continuous monitor OFF for {address}").format(address=address))
			if not cls._monitors:
				cls._stop_timer()
		else:
			cls._monitors[monitor_key] = val
			cls._start_timer(excel)
			# Translators: Message when continuous monitoring is turned on for a cell.
			ui.message(_("Continuous monitor ON for {address}").format(address=address))

	@classmethod
	def clear_all(cls, obj):
		"""
		Wipes all slots and active monitors from memory.
		
		Architectural Intent & Considerations:
		Provides a necessary reset switch for the user. Crucially, it must also call `_stop_timer()` 
		to ensure the background polling loop is fully terminated.
		"""
		cls._slots.clear()
		cls._monitors.clear()
		cls._stop_timer()
		# Translators: Message when all cell monitors are cleared.
		ui.message(_("All monitored and slotted cells cleared."))

	@classmethod
	def _check_all_monitors(cls):
		"""
		The core background polling loop that checks all registered cells for value changes.
		
		Architectural Intent & Considerations:
		This function recursively calls itself via `core.callLater`. It iterates through the `_monitors` 
		dictionary, queries the live COM cell value, and compares it against the cached value. 
		It includes strict safety gates to prevent crashing if Excel is mid-calculation or a workbook is closed.
		"""
		if not cls._isMonitoringActive:
			return

		# Reschedule first to guarantee continuous loop
		import core
		core.callLater(150, cls._check_all_monitors)

		if not cls._monitors:
			cls._stop_timer()
			return

		try:
			import comtypes.client
			try:
				# Fetch fresh excel instance dynamically to avoid stale COM proxies
				excel = comtypes.client.GetActiveObject("Excel.Application")
			except Exception:
				excel = cls._active_excel
				
			if not excel:
				return

			# Safety Gate 0: Global Excel Closure Detection
			# If the user closed Excel visually, we must immediately release cls._active_excel
			# otherwise the COM reference keeps EXCEL.EXE alive as a zombie process.
			try:
				import ctypes
				if not ctypes.windll.user32.FindWindowW("XLMAIN", None):
					cls.clear_all(None)
					return
			except Exception:
				pass

			# Safety Gate 1: Mid-Calculation Trap
			try:
				# xlDone = 0. If Excel is calculating (1) or pending (2), wait.
				if excel.CalculationState != 0:
					return
			except Exception:
				pass

			# Safety Gate 2: Ghost Workbook Cleanup & Rename Detection
			open_wbs = None
			try:
				# Fetch all open workbook names. If this fails, Excel is busy (Edit Mode).
				open_wbs = [wb.Name for wb in excel.Workbooks]
			except Exception:
				pass

			to_remove = []

			for key, last_val in cls._monitors.items():
				wb_name, sheet_name, cell_addr = key.split("|")
				
				# Active detection of renamed/lost Workbooks and Sheets
				if open_wbs is not None:
					if wb_name not in open_wbs:
						to_remove.append(key)
						continue
						
					# If workbook exists, safely verify sheet exists.
					try:
						target_wb = excel.Workbooks(wb_name)
						sheet_names = [s.Name for s in target_wb.Sheets]
						if sheet_name not in sheet_names:
							to_remove.append(key)
							continue
					except Exception:
						# If checking sheets fails due to Edit Mode, safely ignore for this tick.
						pass

				try:
					target_wb = excel.Workbooks(wb_name)
					target_sheet = target_wb.Sheets(sheet_name)
					target_cell = target_sheet.Range(cell_addr)
					
					# Safety Gate 3: Text vs Value display trap
					current_val = str(target_cell.Text) if target_cell.Text is not None else ""
					if "comtypes" in current_val.lower():
						current_val = ""
					# If column is too narrow, Excel returns ######. Fallback to raw value.
					if not current_val or current_val.startswith("###"):
						raw_val = target_cell.Value
						current_val = str(raw_val) if raw_val is not None else ""
						if "comtypes" in current_val.lower():
							current_val = ""
							
					current_val = current_val.strip()
					
					if current_val != last_val:
						cls._monitors[key] = current_val
						for s_key, s_info in cls._slots.items():
							if s_info["wb"] == wb_name and s_info["sheet"] == sheet_name and s_info["cell"] == cell_addr:
								cls._slots[s_key]["val"] = current_val
								
						import ui
						try:
							active_wb = excel.ActiveWorkbook.Name if excel.ActiveWorkbook else None
							active_sheet = excel.ActiveSheet.Name if excel.ActiveSheet else None
						except Exception:
							active_wb, active_sheet = None, None
							
						if active_wb == wb_name and active_sheet == sheet_name:
							location_str = ""
						elif active_wb == wb_name:
							# Translators: Describes the sheet location of a monitored cell that changed.
							location_str = _(" in {sheet}").format(sheet=sheet_name)
						else:
							# Translators: Describes the sheet and workbook location of a monitored cell that changed.
							location_str = _(" in {sheet} of {wb}").format(sheet=sheet_name, wb=wb_name)
							
						# Translators: Announces that a continuously monitored cell has updated its value.
						ui.message(_("{cell_addr} updated: {current_val}{location_str}").format(
							cell_addr=cell_addr, current_val=current_val, location_str=location_str))
				except Exception:
					pass

			for key in to_remove:
				del cls._monitors[key]
				wb_closed, sheet_closed, cell_addr = key.split("|")
				
				# Check if it belongs to a slot and clear it
				slot_cleared = None
				for s_key, s_info in list(cls._slots.items()):
					if s_info["wb"] == wb_closed and s_info["sheet"] == sheet_closed and s_info["cell"] == cell_addr:
						slot_cleared = s_key
						del cls._slots[s_key]
				
				import ui
				if slot_cleared:
					# Translators: Message when a monitored slot is cleared because the workbook or sheet was closed or renamed.
					ui.message(_("Monitor for Slot {slot_cleared} lost due to name change or closure.").format(
						slot_cleared=slot_cleared))
				else:
					# Translators: Message when a continuously monitored cell is cleared because the workbook or sheet was closed or renamed.
					ui.message(_("Monitor cleared: {sheet_closed} in {wb_closed} lost.").format(
						sheet_closed=sheet_closed, wb_closed=wb_closed))

		except Exception:
			pass

	_last_working_cell = None

	@classmethod
	def _jump_to_address(cls, excel, wb_name, sheet_name, cell_addr):
		try:
			# Cache current location before jumping
			try:
				if excel.ActiveWorkbook and excel.ActiveSheet and excel.ActiveCell:
					cls._last_working_cell = {
						"wb": excel.ActiveWorkbook.Name,
						"sheet": excel.ActiveSheet.Name,
						"cell": excel.ActiveCell.Address()
					}
			except Exception:
				pass
				
			wb = None
			# Robustly resolve workbook, ignoring extension mismatch issues
			for w in excel.Workbooks:
				w_name = w.Name
				if w_name == wb_name or w_name.split('.')[0] == wb_name.split('.')[0]:
					wb = w
					break
					
			if not wb:
				# Translators: Error when trying to jump to a slot in a closed workbook.
				ui.message(_("Cannot jump. The workbook '{wb}' is closed.").format(wb=wb_name))
				return
				
			try:
				sheet = wb.Sheets(sheet_name)
			except Exception:
				# Translators: Error when trying to jump to a slot in a deleted sheet.
				ui.message(_("Cannot jump. The sheet '{sheet}' was renamed or deleted.").format(sheet=sheet_name))
				return
				
			wb.Activate()
			sheet.Activate()
			sheet.Range(cell_addr).Select()
		except Exception:
			# Translators: Error when trying to jump to a slot but the address is invalid.
			ui.message(_("Cannot jump. The cell address is invalid or Excel is busy."))

	@classmethod
	def jump_to_slot(cls, slot_num, obj):
		if slot_num not in cls._slots:
			# Translators: Error when trying to jump to an empty slot.
			ui.message(_("No cell assigned to slot {slot_num}").format(slot_num=slot_num))
			return
			
		info = cls._slots[slot_num]
		excel, _unused_wb, _unused_sheet, _unused_addr, _unused_val = cls._get_active_cell_info(obj)
		if not excel:
			if cls._active_excel:
				excel = cls._active_excel
			else:
				# Translators: Error when Excel cannot be accessed.
				ui.message(_("Error: Excel not accessible."))
				return
				
		cls._jump_to_address(excel, info["wb"], info["sheet"], info["cell"])

	@classmethod
	def jump_back(cls, obj):
		if not cls._last_working_cell:
			# Translators: Error when trying to jump back but there is no history.
			ui.message(_("No previous cell to jump back to."))
			return
			
		info = cls._last_working_cell
		excel, _unused_wb, _unused_sheet, _unused_addr, _unused_val = cls._get_active_cell_info(obj)
		if not excel:
			if cls._active_excel:
				excel = cls._active_excel
			else:
				# Translators: Error when Excel cannot be accessed.
				ui.message(_("Error: Excel not accessible."))
				return
				
		cls._jump_to_address(excel, info["wb"], info["sheet"], info["cell"])
		# Clear it so we don't jump back and forth infinitely
		cls._last_working_cell = None

class ActiveMonitorsDialog(wx.Dialog):
	def __init__(self, parent, slots, monitors, excel_app):
		# Translators: Title of the active cell monitors dialog.
		super().__init__(parent, title=_("Active Cell Monitors"), style=wx.DEFAULT_DIALOG_STYLE | wx.RESIZE_BORDER)
		self.slots = slots
		self.monitors = monitors
		self.excel = excel_app
		self.mapping = [] 
		
		mainSizer = wx.BoxSizer(wx.VERTICAL)
		# Translators: Instructions for the cell monitors dialog.
		helpLabel = wx.StaticText(self, label=_("Select a cell to jump to it. Press Delete to remove it from monitors."))
		mainSizer.Add(helpLabel, 0, wx.ALL, 5)
		
		self.listBox = wx.ListBox(self, size=(500, 300))
		self.populate_list()
		self.listBox.Bind(wx.EVT_LISTBOX_DCLICK, self.onJump)
		self.listBox.Bind(wx.EVT_CHAR_HOOK, self.onCharHook)
		mainSizer.Add(self.listBox, 1, wx.EXPAND | wx.ALL, 5)
		
		btnSizer = wx.StdDialogButtonSizer()
		# Translators: Label for the Jump button.
		self.jumpBtn = wx.Button(self, wx.ID_OK, label=_("&Jump"))
		self.jumpBtn.Bind(wx.EVT_BUTTON, self.onJump)
		btnSizer.AddButton(self.jumpBtn)
		# Translators: Label for the Close button.
		closeBtn = wx.Button(self, wx.ID_CANCEL, label=_("&Close"))
		closeBtn.Bind(wx.EVT_BUTTON, self.onClose)
		btnSizer.AddButton(closeBtn)
		btnSizer.Realize()
		
		mainSizer.Add(btnSizer, 0, wx.EXPAND | wx.ALL, 5)
		self.SetSizer(mainSizer)
		self.CenterOnParent()
		
		if self.listBox.GetCount() > 0:
			self.listBox.SetSelection(0)
		self.listBox.SetFocus()
			
	def populate_list(self):
		self.listBox.Clear()
		self.mapping = []
		for slot_num in sorted(self.slots.keys()):
			info = self.slots[slot_num]
			val = info["val"]
			display_text = f"Slot {slot_num}: {info['sheet']}!{info['cell']} ({val})"
			self.listBox.Append(display_text)
			self.mapping.append((True, slot_num))
			
		for key, val in self.monitors.items():
			wb, sheet, cell = key.split("|")
			isSlotted = False
			for s_info in self.slots.values():
				if s_info["wb"] == wb and s_info["sheet"] == sheet and s_info["cell"] == cell:
					isSlotted = True
					break
			if not isSlotted:
				display_text = f"Monitor: {sheet}!{cell} ({val})"
				self.listBox.Append(display_text)
				self.mapping.append((False, key))
				
		if self.listBox.GetCount() == 0:
			# Translators: Message in dialog when no cells are monitored
			self.listBox.Append(_("No cells are currently being monitored."))
			self.mapping.append((False, None))
			if hasattr(self, 'jumpBtn'):
				self.jumpBtn.Disable()
		else:
			if hasattr(self, 'jumpBtn'):
				self.jumpBtn.Enable()

	def onCharHook(self, evt):
		if evt.GetKeyCode() == wx.WXK_DELETE:
			idx = self.listBox.GetSelection()
			if idx != wx.NOT_FOUND:
				is_slot, key = self.mapping[idx]
				if key is None:
					evt.Skip()
					return
				if is_slot:
					info = self.slots[key]
					monitor_key = f"{info['wb']}|{info['sheet']}|{info['cell']}"
					del self.slots[key]
					if monitor_key in self.monitors:
						del self.monitors[monitor_key]
					import ui
					# Translators: Message when a specific slotted cell is deleted.
					ui.message(_("Slot {slot} deleted").format(slot=key))
				else:
					del self.monitors[key]
					import ui
					# Translators: Message when a continuously monitored cell is deleted.
					ui.message(_("Monitor deleted"))
				self.populate_list()
				if self.listBox.GetCount() > 0:
					self.listBox.SetSelection(min(idx, self.listBox.GetCount() - 1))
		elif evt.GetKeyCode() == wx.WXK_ESCAPE:
			self.onClose(evt)
		elif evt.GetKeyCode() == wx.WXK_RETURN:
			self.onJump(evt)
		else:
			evt.Skip()

	def onJump(self, evt):
		idx = self.listBox.GetSelection()
		if idx != wx.NOT_FOUND:
			is_slot, key = self.mapping[idx]
			if key is None:
				return
			if is_slot:
				info = self.slots[key]
				CellMonitorManager._jump_to_address(self.excel, info["wb"], info["sheet"], info["cell"])
			else:
				wb, sheet, cell = key.split("|")
				CellMonitorManager._jump_to_address(self.excel, wb, sheet, cell)
			self.EndModal(wx.ID_OK)
			
	def onClose(self, evt):
		self.EndModal(wx.ID_CANCEL)

CellMonitorManager.open_monitor_dialog = classmethod(lambda cls, obj: _open_monitor_dialog(cls, obj))

def _open_monitor_dialog(cls, obj):
	import gui
	excel, _unused_wb, _unused_sheet, _unused_addr, _unused_val = cls._get_active_cell_info(obj)
	if not excel:
		if cls._active_excel:
			excel = cls._active_excel
		else:
			import ui
			# Translators: Error when Excel cannot be accessed.
			ui.message(_("Error: Excel not accessible."))
			return

	def _show():
		try:
			gui.mainFrame.prePopup()
			dlg = ActiveMonitorsDialog(gui.mainFrame, cls._slots, cls._monitors, excel)
			dlg.ShowModal()
		except Exception as e:
			import ui
			ui.message(f"Dialog failed: {str(e)}")
			from logHandler import log
			log.error(f"BOA Monitor Dialog Error: {e}", exc_info=True)
		finally:
			try:
				dlg.Destroy()
			except Exception:
				pass
			gui.mainFrame.postPopup()
			
	wx.CallAfter(_show)
