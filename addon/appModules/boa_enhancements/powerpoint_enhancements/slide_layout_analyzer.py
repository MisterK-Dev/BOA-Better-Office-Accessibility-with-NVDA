# -*- coding: UTF-8 -*-
# Copyright (C) 2026 KIRAN G T {MisterK} and Antigravity 2
# This file is covered by the GNU General Public License (GPL), version 2.
# See the file COPYING.txt for more details.

import addonHandler
addonHandler.initTranslation()

"""
PowerPoint Slide Layout Analyzer

Provides spatial awareness, accessibility metrics, and complexity analysis 
for the active PowerPoint slide.
Architectural Intent:
Because enumerating PowerPoint shapes via COM is slow and can freeze NVDA's 
main thread, we leverage `AsyncCOMTask` to gather all layout properties in the 
background. Once complete, we present a consolidated HTML report.
"""

import ui  # noqa: E402
import wx  # noqa: E402
from logHandler import log  # noqa: E402
from appModules.boa_enhancements.async_engine import AsyncCOMTask  # noqa: E402

# PowerPoint COM Constants
msoPlaceholder = 14
msoPicture = 13
msoGroup = 6
msoSmartArt = 24
msoMedia = 16
msoChart = 3
msoEmbeddedOLEObject = 7

class SlideLayoutAnalyzer:
	"""
	Core engine for spatial mapping and slide boundary detection in PowerPoint.
	"""
	
	@staticmethod
	def _show_dialog(title, message):
		"""Thread-safe launcher for the NVDA browseableMessage."""
		wx.CallAfter(ui.browseableMessage, message, title=title, isHtml=True, closeButton=True, copyButton=True)

	@staticmethod
	def _get_ppt_com():
		import ctypes
		import comtypes.client
		import comtypes.automation
		
		hwnd_main = ctypes.windll.user32.FindWindowW("PPTFrameClass", None)
		if not hwnd_main:
			return None
			
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
			return None
			
		oleacc = ctypes.windll.user32.oleacc if hasattr(ctypes.windll.user32, 'oleacc') else ctypes.windll.oleacc
		ptr = ctypes.POINTER(comtypes.automation.IDispatch)()
		res = oleacc.AccessibleObjectFromWindow(target_hwnd[0], -16, ctypes.byref(comtypes.automation.IDispatch._iid_), ctypes.byref(ptr))
		
		if res == 0 and ptr:
			try:
				doc_window = comtypes.client.dynamic.Dispatch(ptr)
				return doc_window.Application
			except Exception:
				pass
				
		return None

	@classmethod
	def analyze(cls, obj):
		"""
		Entry point. Extracts the PowerPoint application object via UIA/COM 
		and spawns the background async task.
		"""
		try:
			app = cls._get_ppt_com()
			
			if not app:
				# Translators: Error hooking PowerPoint COM
				ui.message(_("Failed to hook PowerPoint."))
				return
			
			try:
				slide = app.ActiveWindow.View.Slide
				slide_num = slide.SlideNumber
			except Exception:
				# Translators: Error when no slide is active
				ui.message(_("No active slide detected."))
				return
				
			# Translators: Announcement when starting slide analysis
			ui.message(_("Analyzing slide layout..."))
			
			task = AsyncCOMTask(
				generator=cls._analyzer_generator(slide),
				on_complete=lambda results: cls._format_and_show_results(results, slide_num)
			)
			task.start()
			
		except Exception as e:
			log.error(f"BOA SlideLayoutAnalyzer Error: {e}", exc_info=True)
			# Translators: Error during analysis
			ui.message(_("Error analyzing slide layout."))

	@staticmethod
	def _process_table(table, s_name, results):
		"""Helper to process a table COM object."""
		try:
			rows = table.Rows.Count
			cols = table.Columns.Count
			results["tables"].append({"name": s_name, "rows": rows, "cols": cols, "merged": False})
		except Exception as e:
			log.debugWarning(f"Slide Analyzer failed processing table for '{s_name}': {e}")

	@staticmethod
	def _safe_com_bool(obj, prop_name):
		"""Safely evaluates a COM property that might throw DISP_E_UNKNOWNNAME."""
		try:
			return bool(getattr(obj, prop_name))
		except Exception:
			return False

	@staticmethod
	def _process_shape(shape, results):
		"""Helper to recursively process shapes and groups."""
		try:
			s_type = shape.Type
			s_name = shape.Name
			
			try:
				top = shape.Top
				left = shape.Left
				# Off-canvas check
				if top < 0 or left < 0:
					results["off_canvas"].append(s_name)
			except Exception:
				pass
				
			# Placeholders vs Custom
			if s_type == msoPlaceholder:
				results["placeholders"] += 1
				
				# Check if placeholder is a Title (1=Title, 3=CenterTitle, 15=VerticalTitle)
				try:
					ph_type = shape.PlaceholderFormat.Type
					if ph_type in (1, 3, 15):
						if SlideLayoutAnalyzer._safe_com_bool(shape, "HasTextFrame"):
							if SlideLayoutAnalyzer._safe_com_bool(shape.TextFrame, "HasText"):
								# Only count it if it actually contains text!
								results["has_title"] = True
				except Exception:
					pass
			else:
				results["custom_shapes"] += 1
				
			# Text Density
			if SlideLayoutAnalyzer._safe_com_bool(shape, "HasTextFrame"):
				if SlideLayoutAnalyzer._safe_com_bool(shape.TextFrame, "HasText"):
					try:
						text = shape.TextFrame.TextRange.Text
						results["word_count"] += len(text.split())
					except Exception:
						pass
				
			# Alt Text check (for all non-text visual shapes)
			if s_type in (msoPicture, msoChart, msoSmartArt, msoEmbeddedOLEObject, msoMedia):
				try:
					alt = shape.AlternativeText
					if not alt or alt.isspace():
						results["missing_alt"] += 1
				except Exception:
					results["missing_alt"] += 1
					
			# Grouping (Recursive check)
			if s_type == msoGroup:
				results["groups"] += 1
				try:
					for g_idx in range(1, shape.GroupItems.Count + 1):
						SlideLayoutAnalyzer._process_shape(shape.GroupItems(g_idx), results)
				except Exception:
					pass
				
			# Media
			if s_type == msoMedia:
				results["media"] += 1
				
			# Charts
			if s_type == msoChart:
				results["charts"] += 1
				
			# Embedded OLE
			if s_type == msoEmbeddedOLEObject:
				results["ole_objects"] += 1
				
			# Tables
			if SlideLayoutAnalyzer._safe_com_bool(shape, "HasTable"):
				SlideLayoutAnalyzer._process_table(shape.Table, s_name, results)
				
		except Exception as e:
			log.debugWarning(f"Slide Analyzer skipped shape '{s_name}': {e}")


	@staticmethod
	def _analyzer_generator(slide):
		"""
		Background generator that analyzes the slide.
		Yields periodically to prevent NVDA freezing.
		"""
		results = {
			"total_shapes": 0,
			"placeholders": 0,
			"has_title": False,
			"custom_shapes": 0,
			"off_canvas": [],
			"missing_alt": 0,
			"tables": [],
			"groups": 0,
			"media": 0,
			"charts": 0,
			"ole_objects": 0,
			"z_order_warnings": [],
			"word_count": 0
		}
		
		try:
			shapes = slide.Shapes
			total = shapes.Count
			results["total_shapes"] = total
			
			shape_positions = []
			
			for i in range(1, total + 1):
				# Yield control every 3 shapes to ensure zero UI lag
				if i % 3 == 0:
					yield True
					
				try:
					shape = shapes(i)
					shape_positions.append({"name": shape.Name, "top": shape.Top, "index": i})
					SlideLayoutAnalyzer._process_shape(shape, results)
				except Exception:
					continue
					
			# Z-Order vs Visual Warning
			# Sort shape_positions by visual 'Top' coordinate
			shape_positions.sort(key=lambda x: x["top"])
			
			# If a shape is in the top 25% of the visual space but added in the last 20% of Z-order
			if total > 2 and shape_positions:
				top_shapes = shape_positions[:max(1, int(total * 0.25))]
				for ts in top_shapes:
					if ts["index"] > (total * 0.8):
						results["z_order_warnings"].append(ts["name"])
						break # Only report the most prominent warning
					
			return results
			
		except Exception as e:
			log.error(f"Slide Analyzer COM loop failed: {e}", exc_info=True)
			return results

	@staticmethod
	def _format_and_show_results(results, slide_num):
		"""
		Formats the collected dict into an accessible HTML document.
		"""
		if not results:
			return
			
		# Translators: Dialog title for Slide Layout Analyzer
		title = _("Slide {num} Layout Analysis").format(num=slide_num)
		
		# Build HTML parts
		html = f"<h1>{title}</h1>"
		
		# Translators: Section header for structural metrics
		h_structure = _("Structure & Density")
		# Translators: Reports total placeholders vs custom shapes
		s_ratio = _("Placeholders vs Custom Shapes: {p} / {c}").format(p=results['placeholders'], c=results['custom_shapes'])
		# Translators: Reports total word count on the slide
		s_words = _("Total Word Count: {w}").format(w=results['word_count'])
		
		html += f"<h2>{h_structure}</h2><ul>"
		html += f"<li>{s_ratio}</li><li>{s_words}</li>"
		html += "</ul>"
		
		if results['word_count'] > 75:
			# Translators: Warning if the slide is a wall of text
			w_dense = _("Warning: High text density (Wall of Text)")
			html += f"<h3>{w_dense}</h3>"
			
		if not results['has_title']:
			# Translators: Warning if no title placeholder is found
			w_title = _("Warning: Missing standard Title Placeholder")
			html += f"<h3>{w_title}</h3>"
		
		# Translators: Section header for Accessibility traps
		h_access = _("Accessibility Warnings")
		html += f"<h2>{h_access}</h2>"
		
		has_access = False
		
		if results["off_canvas"]:
			has_access = True
			# Translators: Header for Off-Canvas Shapes
			h_off = _("Off-Canvas Shapes")
			html += f"<h3>{h_off}</h3><ul>"
			for shape_name in results["off_canvas"]:
				html += f"<li>{shape_name}</li>"
			html += "</ul>"
		
		if results["z_order_warnings"]:
			has_access = True
			# Translators: Label for reading order mismatch warning
			z_label = _("Reading Order Mismatch")
			# Translators: Warns about reading order mismatch for a specific shape
			z_msg = _("Shape '{s}' is visually at the top but logically at the bottom.").format(s=results["z_order_warnings"][0])
			html += f"<h3>{z_label}</h3><ul><li>{z_msg}</li></ul>"
			
		if results["missing_alt"] > 0:
			has_access = True
			# Translators: Warns about missing alt text
			m_alt = _("Missing Alt-Text: {n} images").format(n=results["missing_alt"])
			html += f"<h3>{m_alt}</h3>"
			
		if not has_access:
			# Translators: Status when no accessibility warnings are found
			m_safe = _("No major spatial accessibility warnings detected.")
			html += f"<ul><li>{m_safe}</li></ul>"
			
		# Translators: Section header for Complex Objects
		h_complex = _("Complex Objects")
		html += f"<h2>{h_complex}</h2>"
		
		has_complex = False
		
		if results["groups"] > 0 or results["media"] > 0 or results["charts"] > 0 or results["ole_objects"] > 0:
			has_complex = True
			# Translators: Header for Groups, Charts, and Media
			h_gm = _("Complex Objects & Media")
			html += f"<h3>{h_gm}</h3><ul>"
			if results["groups"] > 0:
				# Translators: Number of shape groups
				n_groups = _("Grouped Objects: {n}").format(n=results["groups"])
				html += f"<li>{n_groups}</li>"
			if results["media"] > 0:
				# Translators: Number of media items
				n_media = _("Media (Audio/Video): {n}").format(n=results["media"])
				html += f"<li>{n_media}</li>"
			if results["charts"] > 0:
				# Translators: Number of charts
				n_charts = _("Charts/Graphs: {n}").format(n=results["charts"])
				html += f"<li>{n_charts}</li>"
			if results["ole_objects"] > 0:
				# Translators: Number of embedded objects
				n_ole = _("Embedded Documents (Excel/PDF): {n}").format(n=results["ole_objects"])
				html += f"<li>{n_ole}</li>"
			html += "</ul>"
			
		if results["tables"]:
			has_complex = True
			# Translators: Header for Tables
			h_tab = _("Tables")
			html += f"<h3>{h_tab}</h3><ul>"
			for t in results["tables"]:
				# Translators: Table dimensions
				t_dim = _("Table '{name}': {c} columns by {r} rows").format(name=t['name'], c=t['cols'], r=t['rows'])
				html += f"<li>{t_dim}</li>"
			html += "</ul>"
			
		if not has_complex:
			# Translators: Status when no complex objects exist
			m_clean = _("No complex tables, groups, or media.")
			html += f"<ul><li>{m_clean}</li></ul>"
		
		SlideLayoutAnalyzer._show_dialog(title, html)
