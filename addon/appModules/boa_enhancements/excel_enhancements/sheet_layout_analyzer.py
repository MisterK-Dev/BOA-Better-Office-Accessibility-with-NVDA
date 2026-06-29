# -*- coding: UTF-8 -*-
# Copyright (C) 2026 KIRAN G T {MisterK} and Antigravity 2
# This file is covered by the GNU General Public License (GPL), version 2.
# See the file COPYING.txt for more details.

import addonHandler
addonHandler.initTranslation()

"""
Sheet Layout Analyzer

Provides spatial awareness and data block detection for Excel sheets.
WHY THIS EXISTS (Architecture intent):
Native NVDA navigation can leave users lost in empty space if data starts at an unusual row or column.
This module uses COM SpecialCells to quickly detect fragmented data blocks and announce their locations,
giving the user an immediate understanding of the sheet's layout.
"""

import ui  # noqa: E402
from logHandler import log  # noqa: E402

import wx  # noqa: E402

class SheetLayoutAnalyzer:
	"""
	Core engine for spatial mapping and sheet boundary detection.
	
	Architectural Intent & Considerations:
	Excel provides zero native APIs for screen readers to understand where data exists in a massive grid. 
	This class uses advanced COM techniques (`SpecialCells`) to instantly mathematically deduce where 
	data islands and boundaries are, preventing the user from getting lost in empty space.
	"""
	_layout_cache = {}

	@staticmethod
	def _show_dialog(title, message, is_html=True):
		"""
		Thread-safe launcher for the NVDA browseableMessage.
		
		Wrapping this call in `wx.CallAfter` ensures that dialog creation is deferred
		until the current input gesture event loop completes, preventing focus and
		Alt+Tab Z-order corruption in the Windows OS window manager.
		"""
		wx.CallAfter(ui.browseableMessage, message, title=title, isHtml=is_html, closeButton=True, copyButton=True)

	@staticmethod
	def _col_num_to_letter(n):
		"""Converts an integer column index (1) to an Excel letter (A)."""
		s = ""
		while n > 0:
			n, remainder = divmod(n - 1, 26)
			s = chr(65 + remainder) + s
		return s

	@staticmethod
	def _get_data_areas(excel):
		"""
		Extracts the top-left and bottom-right coordinates of all data blocks in the UsedRange.
		Returns a sorted list of (row, col_letter) tuples.
		
		Architectural Intent & Considerations:
		Scanning millions of cells iteratively is impossible. By using `SpecialCells(2)` (Constants) 
		and `SpecialCells(-4123)` (Formulas), we force Excel's internal C++ engine to instantly yield 
		the specific chunk coordinates where data physically exists, bypassing the Python bottleneck.
		"""
		areas_coords = set()
		try:
			sheet = excel.ActiveSheet
			used_range = sheet.UsedRange
			
			# xlCellTypeConstants = 2, xlCellTypeFormulas = -4123
			for cell_type in (2, -4123):
				try:
					data_cells = used_range.SpecialCells(cell_type)
					if data_cells:
						for i in range(1, data_cells.Areas.Count + 1):
							area = data_cells.Areas.Item(i)
							# Collect the top-left and bottom-right cells of the area
							row = area.Row
							col = area.Column
							end_row = row + area.Rows.Count - 1
							end_col = col + area.Columns.Count - 1
							areas_coords.add((row, col, end_row, end_col))
				except Exception:
					# SpecialCells throws an exception if no cells match the type.
					pass
		except Exception as e:
			log.debug(f"BOA: Failed to get data areas: {e}")

		# Sort top-to-bottom, then left-to-right
		sorted_areas = sorted(list(areas_coords), key=lambda x: (x[0], x[1]))
		
		# Convert column index to letter for the final output
		return [(r, SheetLayoutAnalyzer._col_num_to_letter(c), 
				 er, SheetLayoutAnalyzer._col_num_to_letter(ec)) 
				for r, c, er, ec in sorted_areas]

	@staticmethod
	def announce_layout(excel):
		"""
		Manually triggered layout overview via NVDA+E, L. Detects all blocks and displays them in a dialog.
		
		Architectural Intent & Considerations:
		Generates a comprehensive snapshot of the entire sheet's topography. It intentionally caches 
		the result (`_layout_cache`) so that subsequent "Guided mode" jumps do not have to recalculate 
		the expensive COM queries.
		"""
		# Translators: Progress message spoken when starting to analyze the sheet layout.
		ui.message(_("Analyzing layout..."))
		areas = SheetLayoutAnalyzer._get_data_areas(excel)
		try:
			sheet_name = excel.ActiveSheet.Name
			SheetLayoutAnalyzer._layout_cache[sheet_name] = areas
		except Exception:
			pass
			
		html_parts = []
		# Translators: Heading for the section summarizing data blocks in the sheet.
		html_parts.append("<h2>" + _("Data Blocks Summary") + "</h2>")

		if not areas:
			# Translators: Message indicating that the entire sheet has no data.
			html_parts.append("<h3>" + _("Sheet appears to be empty.") + "</h3>")
		else:
			count = len(areas)
			if count == 1:
				r, c, er, ec = areas[0]
				if r == er and c == ec:
					# Translators: Message indicating exactly one single-cell data block was found.
					msg = _("Found 1 data block at {c}{r}.").format(c=c, r=r)
				else:
					# Translators: Message indicating exactly one multi-cell data block was found.
					msg = _("Found 1 data block: {c}{r} to {ec}{er}.").format(c=c, r=r, ec=ec, er=er)
				html_parts.append("<h3>" + msg + "</h3>")
			else:
				# Translators: Message summarizing the total number of data blocks found.
				summary = _("Found {count} data blocks in this sheet.").format(count=count)
				html_parts.append("<h3>" + summary + "</h3>")
				
				html_parts.append('<ul>')
				for i, (r, c, er, ec) in enumerate(areas):
					if r == er and c == ec:
						addr = f"{c}{r}"
					else:
						# Translators: Formats the start and end coordinates of a data block.
						addr = _("{c}{r} to {ec}{er}").format(c=c, r=r, ec=ec, er=er)
					
					# Translators: Describes a specific data block by its number and coordinates.
					block_str = _("Block {num}: {addr}").format(num=i+1, addr=addr)
					if ":" in block_str:
						label, val = block_str.split(":", 1)
						html_parts.append(f'<li><b>{label.strip()}:</b> {val.strip()}</li>')
					else:
						html_parts.append(f'<li>{block_str}</li>')
				html_parts.append("</ul>")
				
		# Append Sheet Properties
		props_html = SheetLayoutAnalyzer._get_sheet_properties(excel)
		html_parts.append(props_html)
		
		final_html = "".join(html_parts)
		# Translators: Title of the dialog displaying the sheet layout analysis.
		SheetLayoutAnalyzer._show_dialog(_("Sheet Layout Overview"), final_html)

	@staticmethod
	def auto_announce_one_time(excel):
		"""
		Instantly calculates the nearest data block and announces it when the user lands on an empty cell.
		
		Architectural Intent & Considerations:
		Fired only when the user opens a workbook or switches sheets. It calculates the Manhattan 
		distance to all blocks. It intentionally does NOT use the cache, as the user might be opening 
		a brand new file that hasn't been scanned yet.
		"""
		import ui
		try:
			cell = excel.ActiveCell
			r, c_num = cell.Row, cell.Column
			
			is_empty = False
			try:
				val = cell.Value
				text = cell.Text
				if val is None or str(text).strip() == "":
					is_empty = True
			except Exception:
				is_empty = False

			if not is_empty:
				return

			areas = SheetLayoutAnalyzer._get_data_areas(excel)
			if not areas:
				return

			closest = None
			min_dist = float('inf')
			
			for ar, ac_letter, aer, aec_letter in areas:
				ac = SheetLayoutAnalyzer._letter_to_col_num(ac_letter)
				dist = abs(r - ar) + abs(c_num - ac)
				if dist < min_dist:
					min_dist = dist
					closest = (ar, ac_letter)
					
			if closest:
				# Translators: Message spoken when querying an empty cell, pointing to the nearest data.
				msg = _("Empty cell. Nearest data at {col}{row}.").format(col=closest[1], row=closest[0])
				ui.message(msg)
		except Exception as e:
			log.debug(f"BOA: auto_announce_one_time Exception: {e}")

	@staticmethod
	def auto_announce_guided(excel):
		"""
		Announces the closest data block continuously during normal navigation.
		
		Architectural Intent & Considerations:
		Because this fires on every single keystroke in "Guided" mode, it MUST use the `_layout_cache`. 
		Executing live COM `SpecialCells` queries on every arrow key press would severely lag the NVDA interface.
		"""
		import ui
		try:
			sheet_name = excel.ActiveSheet.Name
			if sheet_name not in SheetLayoutAnalyzer._layout_cache:
				return
				
			areas = SheetLayoutAnalyzer._layout_cache[sheet_name]
			if not areas:
				return
				
			cell = excel.ActiveCell
			r, c_num = cell.Row, cell.Column
			
			is_empty = False
			try:
				val = cell.Value
				text = cell.Text
				if val is None or str(text).strip() == "":
					is_empty = True
			except Exception:
				is_empty = False

			if not is_empty:
				return

			closest = None
			min_dist = float('inf')
			
			for ar, ac_letter, aer, aec_letter in areas:
				ac = SheetLayoutAnalyzer._letter_to_col_num(ac_letter)
				dist = abs(r - ar) + abs(c_num - ac)
				if dist < min_dist:
					min_dist = dist
					closest = (ar, ac_letter)
					
			if closest:
				# Translators: Message pointing to the nearest data from the current cell.
				msg = _("Nearest data at {col}{row}").format(col=closest[1], row=closest[0])
				ui.message(msg)
		except Exception as e:
			log.debug(f"BOA: auto_announce_guided Exception: {e}")

	@staticmethod
	def _get_contiguous_hidden(sheet, start_idx, limit_idx, is_row, step):
		"""
		Calculates the exact boundary depth of hidden edge rows or columns.
		
		Architectural Intent & Considerations:
		Users often hide massive blocks of rows at the edges of sheets (e.g., hiding Rows 100-1048576). 
		This helper walks step-by-step to find exactly where the hidden block ends. It includes a strict 
		2000-iteration safety bailout to prevent NVDA from hard-freezing if a user maliciously hides millions of rows.
		"""
		last_hidden = start_idx
		curr = start_idx
		
		count = 0
		while True:
			if (step > 0 and curr > limit_idx) or (step < 0 and curr < limit_idx):
				break
			try:
				hidden = sheet.Rows(curr).Hidden if is_row else sheet.Columns(curr).Hidden
				if hidden:
					last_hidden = curr
				else:
					break
			except Exception:
				break
			curr += step
			count += 1
			if count > 2000: # Safety bailout to prevent NVDA freezing
				return -1 
		return last_hidden

	@staticmethod
	def _get_sheet_properties(excel):
		"""
		Aggregates global sheet states like Filters, Frozen Panes, and Hidden Boundaries.
		
		Architectural Intent & Considerations:
		UIA doesn't expose global sheet properties well. We manually query COM properties to build 
		a comprehensive summary string for the Layout dialog.
		"""
		props = []
		try:
			sheet = excel.ActiveSheet
			wb = excel.ActiveWorkbook
			win = excel.ActiveWindow
			
			# 1. Filter Mode
			try:
				if getattr(sheet, "AutoFilterMode", False):
					# Translators: Indicates that Excel filtering is currently active.
					props.append(_("Filter Mode: Active"))
			except Exception:
				pass
			
			# 2. Hidden Borders
			try:
				hidden_borders = []
				
				# Consideration: UsedRange ignores hidden columns at the extreme edges of the grid. 
				# To find absolute borders, we must check the literal mathematical edges of the Excel grid 
				# (Row 1, Col 1, MaxRow, MaxCol) instead of relying on the active data bounds.
				
				# Check Absolute Top Edge (Row 1)
				try:
					# If Row 1 is hidden, we scan downwards (step=1) to see how deep the hidden block goes.
					if sheet.Rows(1).Hidden:
						end_r = SheetLayoutAnalyzer._get_contiguous_hidden(sheet, 1, sheet.Rows.Count, True, 1)
						# Translators: Indicates that a very large number of top rows are hidden.
						if end_r == -1:
							# Translators: Automatically added by BOA compliance auditor
							hidden_borders.append(_("Top 2000+ Rows are hidden"))
						# Translators: Indicates that only the very first row is hidden.
						elif end_r == 1:
							# Translators: Automatically added by BOA compliance auditor
							hidden_borders.append(_("Top Row 1 is hidden"))
						# Translators: Indicates a specific range of top rows are hidden.
						else:
							# Translators: Automatically added by BOA compliance auditor
							hidden_borders.append(_("Top Rows 1 through {end_r} are hidden").format(end_r=end_r))
				except Exception:
					pass
				
				# Check Absolute Left Edge (Col 1)
				try:
					if sheet.Columns(1).Hidden:
						end_c = SheetLayoutAnalyzer._get_contiguous_hidden(sheet, 1, sheet.Columns.Count, False, 1)
						# Translators: Indicates that a very large number of left columns are hidden.
						if end_c == -1:
							# Translators: Automatically added by BOA compliance auditor
							hidden_borders.append(_("Left 2000+ Columns are hidden"))
						# Translators: Indicates that only the very first column (A) is hidden.
						elif end_c == 1:
							# Translators: Automatically added by BOA compliance auditor
							hidden_borders.append(_("Left Column A is hidden"))
						# Translators: Indicates a specific range of left columns are hidden.
						else:
							# Translators: Automatically added by BOA compliance auditor
							hidden_borders.append(_("Left Columns A through {end_col} are hidden").format(end_col=SheetLayoutAnalyzer._col_num_to_letter(end_c)))
				except Exception:
					pass
				
				# Check Absolute Bottom Edge
				try:
					# Excel's absolute max row is Rows.Count (usually 1,048,576). If the absolute bottom 
					# row is hidden, we scan upwards (step=-1) to find where the visible data stops.
					max_sheet_r = sheet.Rows.Count
					if sheet.Rows(max_sheet_r).Hidden:
						start_r = SheetLayoutAnalyzer._get_contiguous_hidden(sheet, max_sheet_r, 1, True, -1)
						# Translators: Indicates that a very large number of bottom rows are hidden.
						if start_r == -1:
							# Translators: Automatically added by BOA compliance auditor
							hidden_borders.append(_("Bottommost 2000+ Rows are hidden"))
						# Translators: Indicates that the very last row in the sheet is hidden.
						elif start_r == max_sheet_r:
							# Translators: Automatically added by BOA compliance auditor
							hidden_borders.append(_("Bottom Row {max_r} is hidden").format(max_r=max_sheet_r))
						# Translators: Indicates a specific range of bottom rows are hidden.
						else:
							# Translators: Automatically added by BOA compliance auditor
							hidden_borders.append(_("Bottom Rows {start_r} through {max_r} are hidden").format(start_r=start_r, max_r=max_sheet_r))
				except Exception:
					pass
				
				# Check Absolute Right Edge
				try:
					max_sheet_c = sheet.Columns.Count
					if sheet.Columns(max_sheet_c).Hidden:
						start_c = SheetLayoutAnalyzer._get_contiguous_hidden(sheet, max_sheet_c, 1, False, -1)
						max_c_let = SheetLayoutAnalyzer._col_num_to_letter(max_sheet_c)
						# Translators: Indicates that a very large number of right columns are hidden.
						if start_c == -1:
							# Translators: Automatically added by BOA compliance auditor
							hidden_borders.append(_("Rightmost 2000+ Columns are hidden"))
						# Translators: Indicates that the very last column in the sheet is hidden.
						elif start_c == max_sheet_c:
							# Translators: Automatically added by BOA compliance auditor
							hidden_borders.append(_("Right Column {col} is hidden").format(col=max_c_let))
						# Translators: Indicates a specific range of right columns are hidden.
						else:
							# Translators: Automatically added by BOA compliance auditor
							hidden_borders.append(_("Right Columns {start_col} through {end_col} are hidden").format(start_col=SheetLayoutAnalyzer._col_num_to_letter(start_c), end_col=max_c_let))
				except Exception:
					pass
				
				if hidden_borders:
					# Translators: Summarizes all the hidden borders (top, bottom, left, right).
					props.append(_("Hidden Borders: {borders}").format(borders=", ".join(hidden_borders)))
			except Exception:
				pass
			
			# 3. Hidden Sheets
			try:
				sheet_count = wb.Sheets.Count
				hidden_indices = []
				# Visible = -1 (xlSheetVisible), 0 (xlSheetHidden), 2 (xlSheetVeryHidden)
				for i in range(1, sheet_count + 1):
					if wb.Sheets(i).Visible != -1:
						hidden_indices.append(i)
				
				if hidden_indices:
					start_hidden = 0
					for i in range(1, sheet_count + 1):
						if i in hidden_indices:
							start_hidden += 1
						else:
							break
						
					end_hidden = 0
					for i in range(sheet_count, 0, -1):
						if i in hidden_indices:
							end_hidden += 1
						else:
							break
						
					parts = []
					if start_hidden > 0:
						# Translators: Describes how many sheets are hidden at the start of the workbook.
						parts.append(_("Top sheets 1-{count}").format(count=start_hidden) if start_hidden > 1 else _("Top sheet 1"))
					if end_hidden > 0 and (sheet_count - end_hidden >= start_hidden):
						# Translators: Describes how many sheets are hidden at the end of the workbook.
						parts.append(_("Bottom sheets {start}-{end}").format(start=sheet_count - end_hidden + 1, end=sheet_count) if end_hidden > 1 else _("Bottom sheet {num}").format(num=sheet_count))
						
					middle_count = len(hidden_indices) - start_hidden - end_hidden
					if middle_count > 0:
						# Translators: Describes how many sheets are hidden in the middle of the workbook.
						parts.append(_("{count} middle sheets").format(count=middle_count) if middle_count > 1 else _("{count} middle sheet").format(count=middle_count))
						
					if parts:
												# Translators: Summarizes the hidden sheets in the workbook.
						props.append(_("Hidden Sheets: {parts} hidden").format(parts=", ".join(parts)))
			except Exception:
				pass
			
			# 4. Sheet Protected
			try:
				if getattr(sheet, "ProtectContents", False):
					# Translators: Indicates that the current sheet is protected.
					props.append(_("Sheet Protected: True"))
			except Exception:
				pass
			
			# 5. Frozen Panes
			try:
				if getattr(win, "FreezePanes", False):
					r = win.SplitRow
					c = win.SplitColumn
					if r > 0 and c > 0:
												# Translators: Describes which rows and columns are frozen.
						props.append(_("Frozen Panes: Rows 1-{r}, Columns A-{c}").format(r=r, c=SheetLayoutAnalyzer._col_num_to_letter(c)))
					elif r > 0:
												# Translators: Describes which rows are frozen.
						props.append(_("Frozen Panes: Rows 1-{r}").format(r=r))
					elif c > 0:
												# Translators: Describes which columns are frozen.
						props.append(_("Frozen Panes: Columns A-{c}").format(c=SheetLayoutAnalyzer._col_num_to_letter(c)))
					else:
												# Translators: Indicates that frozen panes are active but specifics could not be determined.
						props.append(_("Frozen Panes: Active"))
			except Exception:
				pass
			
			# 6. Floating Objects
			try:
				c = sheet.Shapes.Count
				if c > 0:
					# Translators: Describes how many floating objects (shapes, charts) are in the sheet.
					props.append(_("Floating Objects: Contains {count} Shape(s)/Chart(s)").format(count=c))
			except Exception:
				pass
			
			# 7. Pivot Tables
			try:
				c = sheet.PivotTables().Count
				if c > 0:
					# Translators: Describes how many pivot tables are in the sheet.
					props.append(_("Pivot Tables: Contains {count} Pivot Table(s)").format(count=c))
			except Exception:
				pass
			
			# 8. View Mode
			try:
				v = win.View
				if v == 2:
					# Translators: Indicates the sheet is in Page Break Preview mode.
					props.append(_("View Mode: Page Break Preview"))
				elif v == 3:
					# Translators: Indicates the sheet is in Page Layout mode.
					props.append(_("View Mode: Page Layout"))
			except Exception:
				pass
			
		except Exception:
			pass
			
		html_props = []
		# Translators: Heading for the section summarizing overall sheet properties.
		html_props.append("<h2>" + _("Sheet Properties") + "</h2>")
		if props:
			html_props.append('<ul>')
			for prop in props:
				if ":" in prop:
					name, val = prop.split(":", 1)
					html_props.append(f'<li><b>{name.strip()}:</b> {val.strip()}</li>')
				else:
					html_props.append(f'<li>{prop}</li>')
			html_props.append("</ul>")
		else:
			# Translators: Message shown when the sheet has no special properties.
			html_props.append("<p>" + _("No special properties detected (No active filters, protection, frozen panes, or hidden borders).") + "</p>")
			
		return "".join(html_props)

	@staticmethod
	def jump_to_nearest_block(excel):
		"""
		Calculates the Manhattan distance to all known data blocks and natively moves the selection to it.
		
		Architectural Intent & Considerations:
		Instead of just telling the user where data is, this function physically transports them to the 
		closest island of data using the cached coordinates, saving them from having to press the arrow keys manually.
		"""
		import ui
		try:
			sheet_name = excel.ActiveSheet.Name
			if sheet_name not in SheetLayoutAnalyzer._layout_cache:
				# Translators: Error message instructing the user to scan the sheet layout before navigating it.
				ui.message(_("Layout cache empty. Press NVDA+E, L to scan sheet first."))
				return
				
			areas = SheetLayoutAnalyzer._layout_cache[sheet_name]
			if not areas:
				# Translators: Error message indicating the sheet has no data blocks to navigate to.
				ui.message(_("No data blocks found in cache."))
				return
				
			cell = excel.ActiveCell
			r, c_num = cell.Row, cell.Column
			
			# Robust empty cell detection
			is_empty = False
			try:
				val = cell.Value
				text = cell.Text
				# A cell is considered empty if its Value is None, OR its Text is visibly empty
				if val is None or str(text).strip() == "":
					is_empty = True
			except Exception:
				is_empty = False

			if not is_empty:
				import ui
				# Translators: Message when trying to jump to nearest data, but already on a data cell.
				ui.message(_("Already on a data cell."))
				return
				
			closest = None
			min_dist = float('inf')
			
			for ar, ac_letter, aer, aec_letter in areas:
				ac = SheetLayoutAnalyzer._letter_to_col_num(ac_letter)
				# Manhattan distance to top-left cell of the block
				dist = abs(r - ar) + abs(c_num - ac)
				if dist < min_dist:
					min_dist = dist
					closest = (ar, ac_letter)
					
			if closest:
				ar, ac_letter = closest
				excel.ActiveSheet.Cells(ar, SheetLayoutAnalyzer._letter_to_col_num(ac_letter)).Select()
				# Translators: Success message after jumping to a specific cell.
				msg = _("Jumped to {col}{row}").format(col=ac_letter, row=ar)
				ui.message(msg)
		except Exception as e:
			log.debug(f"BOA: jump_to_nearest_block Exception: {e}")

	@staticmethod
	def _letter_to_col_num(letter):
		"""Converts an Excel column letter (A) to an integer index (1)."""
		num = 0
		for c in letter:
			num = num * 26 + (ord(c.upper()) - ord('A')) + 1
		return num
