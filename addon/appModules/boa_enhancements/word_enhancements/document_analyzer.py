# -*- coding: UTF-8 -*-
# Copyright (C) 2026 KIRAN G T {MisterK} and Antigravity 2
# This file is covered by the GNU General Public License (GPL), version 2.
# See the file COPYING.txt for more details.

import addonHandler
import ui
import wx
from logHandler import log
import core

addonHandler.initTranslation()

# Word COM Constants
WD_STAT_WORDS = 0
WD_STAT_LINES = 1
WD_STAT_PAGES = 2
WD_STAT_CHARS = 3
WD_STAT_PARAGRAPHS = 4

WD_ORIENT_PORTRAIT = 0
WD_ORIENT_LANDSCAPE = 1

WD_ALIGN_VERTICAL_TOP = 0
WD_ALIGN_VERTICAL_CENTER = 1
WD_ALIGN_VERTICAL_JUSTIFY = 2
WD_ALIGN_VERTICAL_BOTTOM = 3


class WordDocumentAnalyzer:
	"""
	Core engine for analyzing a Word Document's structural, accessibility, 
	and layout health, presenting the results in an HTML dialog.
	"""

	@staticmethod
	def _show_html_dialog(title, html_content):
		"""Safely launches the UI dialog on the main thread."""
		wx.CallAfter(ui.browseableMessage, html_content, title=title, isHtml=True, closeButton=True, copyButton=True)

	@staticmethod
	def _convert_points_to_cm(points):
		"""Converts Word internal points to centimeters."""
		try:
			return points / 28.35
		except Exception:
			return 0.0

	@staticmethod
	def analyze(obj):
		"""
		Orchestrates the analysis process.
		
		:param obj: The NVDAObject representing the Word Document.
		"""
		# Translators: Analyzing document, please wait...
		ui.message(_("Analyzing document, please wait..."))
		
		# We must run the heavy COM analysis slightly asynchronously to allow 
		# the "Analyzing..." message to be spoken immediately.
		core.callLater(50, WordDocumentAnalyzer._perform_analysis, obj)

	@staticmethod
	def _perform_analysis(obj):
		try:
			doc = None
			
			# 1. Try native NVDA Legacy properties
			doc = getattr(obj, "winwordDocumentObject", None)
			if not doc:
				doc = getattr(obj, "WinwordDocumentObject", None)
				
			# 2. Try UIA Word Window Object's document property
			if not doc:
				win_obj = getattr(obj, "winwordWindowObject", None)
				if not win_obj and obj.appModule:
					win_obj = getattr(obj.appModule, "winwordWindowObject", None)
				if win_obj:
					doc = getattr(win_obj, "document", getattr(win_obj, "Document", None))
			
			# 3. Try appModule winwordApplicationObject ActiveDocument
			if not doc and obj.appModule:
				app = getattr(obj.appModule, "winwordApplicationObject", None)
				if app:
					try:
						doc = app.ActiveDocument
					except Exception:
						pass
					
			# 4. Final Fallback to ROT
			if not doc:
				try:
					import comtypes.client
					app = comtypes.client.GetActiveObject("Word.Application")
					doc = app.ActiveDocument
				except Exception:
					pass
					
			if not doc:
				# Translators: Cannot access Word application.
				ui.message(_("Cannot access Word application."))
				return
			if not doc:
				# Translators: No active document found.
				ui.message(_("No active document found."))
				return

			html_parts = []
			# Translators: Main title of the Document Analyzer report
			html_parts.append(f"<h1>{_('Word Document Analyzer Report')}</h1>")

			WordDocumentAnalyzer._append_layout_stats(doc, html_parts)
			WordDocumentAnalyzer._append_structural_stats(doc, html_parts)
			WordDocumentAnalyzer._append_accessibility_stats(doc, html_parts)
			WordDocumentAnalyzer._append_collaboration_stats(doc, html_parts)

			html_content = "".join(html_parts)
			# Translators: Word Document Analyzer
			WordDocumentAnalyzer._show_html_dialog(_("Word Document Analyzer"), html_content)

		except Exception as e:
			import traceback
			import sys
			exc_type, exc_obj, exc_tb = sys.exc_info()
			line_no = exc_tb.tb_lineno
			log.error(f"BOA Word Analyzer Error on line {line_no}: {e}", exc_info=True)
			ui.message(f"Error on line {line_no}: {exc_type.__name__}")

	@staticmethod
	def _append_layout_stats(doc, html_parts):
		# Translators: Heading for the layout section of the analyzer
		html_parts.append(f"<h2>{_('1. Document & Page Layout')}</h2>")
		html_parts.append("<ul>")
		try:
			sections = doc.Sections
			count = sections.Count
			# Translators: Total Sections: {count}
			html_parts.append(f"<li>{_('Total Sections: {count}').format(count=count)}</li>")
			
			# We will just analyze the first section to give a general overview
			if count > 0:
				first_section = sections.Item(1)
				page_setup = first_section.PageSetup
				
				# Translators: Portrait
				orient = _("Portrait") if page_setup.Orientation == WD_ORIENT_PORTRAIT else _("Landscape")
				# Translators: Orientation (Section 1): {orient}
				html_parts.append(f"<li>{_('Orientation (Section 1): {orient}').format(orient=orient)}</li>")
				
				valign_map = {
					# Translators: Top
					WD_ALIGN_VERTICAL_TOP: _("Top"),
					# Translators: Center
					WD_ALIGN_VERTICAL_CENTER: _("Center"),
					# Translators: Justified
					WD_ALIGN_VERTICAL_JUSTIFY: _("Justified"),
					# Translators: Bottom
					WD_ALIGN_VERTICAL_BOTTOM: _("Bottom")
				}
				# Translators: Unknown
				valign = valign_map.get(page_setup.VerticalAlignment, _("Unknown"))
				# Translators: Vertical Alignment: {align}
				html_parts.append(f"<li>{_('Vertical Alignment: {align}').format(align=valign)}</li>")
				
				top_cm = WordDocumentAnalyzer._convert_points_to_cm(page_setup.TopMargin)
				bot_cm = WordDocumentAnalyzer._convert_points_to_cm(page_setup.BottomMargin)
				left_cm = WordDocumentAnalyzer._convert_points_to_cm(page_setup.LeftMargin)
				right_cm = WordDocumentAnalyzer._convert_points_to_cm(page_setup.RightMargin)
				
				# Translators: Top: {t:.2f}cm, Bottom: {b:.2f}cm, Left: {l:.2f}cm, Right: {r:.2f}cm
				margins_text = _("Top: {t:.2f}cm, Bottom: {b:.2f}cm, Left: {l:.2f}cm, Right: {r:.2f}cm").format(
					t=top_cm, b=bot_cm, l=left_cm, r=right_cm)
				# Translators: Margins: {margins}
				html_parts.append(f"<li>{_('Margins: {margins}').format(margins=margins_text)}</li>")
				
				if page_setup.DifferentFirstPageHeaderFooter:
					# Translators: Warning: First page has a different header/footer.
					html_parts.append(f"<li><strong>{_('Warning: First page has a different header/footer.')}</strong></li>")
				if page_setup.OddAndEvenPagesHeaderFooter:
					# Translators: Warning: Odd and even pages have different headers/footers.
					html_parts.append(f"<li><strong>{_('Warning: Odd and even pages have different headers/footers.')}</strong></li>")
				
		except Exception as e:
			log.debug(f"BOA Layout Stats Error: {e}")
			# Translators: Could not retrieve layout statistics.
			html_parts.append(f"<li>{_('Could not retrieve layout statistics.')}</li>")
		html_parts.append("</ul>")

	@staticmethod
	def _append_structural_stats(doc, html_parts):
		# Translators: Heading for the structural section of the analyzer
		html_parts.append(f"<h2>{_('2. Structural Breakdown')}</h2>")
		html_parts.append("<ul>")
		try:
			pages = doc.ComputeStatistics(WD_STAT_PAGES, True)
			words = doc.ComputeStatistics(WD_STAT_WORDS, True)
			chars = doc.ComputeStatistics(WD_STAT_CHARS, True)
			# Architectural Note: 
			# `doc.Paragraphs.Count` returns the raw structural paragraph count. This includes 
			# completely empty lines, lines inside tables, and structural markers.
			# `doc.ComputeStatistics(WD_STAT_PARAGRAPHS)` only counts actual text paragraphs 
			# that contain words, largely ignoring tables and empty lines. We display both 
			# because screen readers navigate via the structural count, but sighted users 
			# think of paragraphs via the text count.
			paras_structural = doc.Paragraphs.Count
			paras_text = doc.ComputeStatistics(WD_STAT_PARAGRAPHS, True)
			lines = doc.ComputeStatistics(WD_STAT_LINES, True)
			
			# Translators: Pages: {pages}
			html_parts.append(f"<li>{_('Pages: {pages}').format(pages=pages)}</li>")
			# Translators: Words (including footnotes): {words}
			html_parts.append(f"<li>{_('Words (including footnotes): {words}').format(words=words)}</li>")
			# Translators: Characters (including footnotes): {chars}
			html_parts.append(f"<li>{_('Characters (including footnotes): {chars}').format(chars=chars)}</li>")
			# Translators: Text Paragraphs: {paras}
			html_parts.append(f"<li>{_('Text Paragraphs: {paras}').format(paras=paras_text)}</li>")
			# Translators: Structural Paragraphs (incl. blank lines & tables): {paras}
			html_parts.append(f"<li>{_('Structural Paragraphs (incl. blank lines & tables): {paras}').format(paras=paras_structural)}</li>")
			# Translators: Lines: {lines}
			html_parts.append(f"<li>{_('Lines: {lines}').format(lines=lines)}</li>")
		except Exception as e:
			log.debug(f"BOA Structural Stats Error: {e}")
			# Translators: Could not retrieve structural statistics.
			html_parts.append(f"<li>{_('Could not retrieve structural statistics.')}</li>")
		html_parts.append("</ul>")

	@staticmethod
	def _append_accessibility_stats(doc, html_parts):
		# Translators: Heading for the accessibility section of the analyzer
		html_parts.append(f"<h2>{_('3. Accessibility Audit')}</h2>")
		html_parts.append("<ul>")
		try:
			# Images
			inline_shapes = doc.InlineShapes
			in_count = inline_shapes.Count
			missing_alt = 0
			
			# No limit as requested by user for on-demand analysis
			if in_count > 0:
				for i in range(1, in_count + 1):
					try:
						shape = inline_shapes.Item(i)
						alt_text = getattr(shape, "AlternativeText", "")
						title_text = getattr(shape, "Title", "")
						if not alt_text and not title_text:
							missing_alt += 1
					except Exception:
						pass
			
			# Translators: Images (Inline): {count}
			html_parts.append(f"<li>{_('Images (Inline): {count}').format(count=in_count)}</li>")
			if missing_alt > 0:
				# Translators: Warning: {count} images are missing alternative text!
				html_parts.append(f"<li><strong>{_('Warning: {count} images are missing alternative text!') .format(count=missing_alt)}</strong></li>")
			
			# Floating Shapes
			shapes = doc.Shapes
			sh_count = shapes.Count
			if sh_count > 0:
				# Translators: Warning: {count} floating shapes detected. These are often inaccessible.
				html_parts.append(f"<li><strong>{_('Warning: {count} floating shapes detected. These are often inaccessible.').format(count=sh_count)}</strong></li>")
			
			# Tables
			tables = doc.Tables
			tb_count = tables.Count
			# Translators: Tables: {count}
			html_parts.append(f"<li>{_('Tables: {count}').format(count=tb_count)}</li>")
			
			# Hyperlinks
			links = doc.Hyperlinks
			link_count = links.Count
			# Translators: Hyperlinks: {count}
			html_parts.append(f"<li>{_('Hyperlinks: {count}').format(count=link_count)}</li>")
			
		except Exception as e:
			log.debug(f"BOA Accessibility Stats Error: {e}")
			# Translators: Could not retrieve accessibility statistics.
			html_parts.append(f"<li>{_('Could not retrieve accessibility statistics.')}</li>")
		html_parts.append("</ul>")

	@staticmethod
	def _append_collaboration_stats(doc, html_parts):
		# Translators: Heading for the collaboration section of the analyzer
		html_parts.append(f"<h2>{_('4. Collaboration Status')}</h2>")
		html_parts.append("<ul>")
		try:
			comments = doc.Comments
			co_count = comments.Count
			# Translators: Comments: {count}
			html_parts.append(f"<li>{_('Comments: {count}').format(count=co_count)}</li>")
			
			revisions = doc.Revisions
			rev_count = revisions.Count
			if rev_count > 0:
				# Translators: Unresolved Tracked Changes: {count}
				html_parts.append(f"<li><strong>{_('Unresolved Tracked Changes: {count}').format(count=rev_count)}</strong></li>")
			else:
				# Translators: Unresolved Tracked Changes: 0
				html_parts.append(f"<li>{_('Unresolved Tracked Changes: 0')}</li>")
				
			footnotes = doc.Footnotes
			fn_count = footnotes.Count
			# Translators: Footnotes: {count}
			html_parts.append(f"<li>{_('Footnotes: {count}').format(count=fn_count)}</li>")
			
			endnotes = doc.Endnotes
			en_count = endnotes.Count
			# Translators: Endnotes: {count}
			html_parts.append(f"<li>{_('Endnotes: {count}').format(count=en_count)}</li>")
			
		except Exception as e:
			log.debug(f"BOA Collaboration Stats Error: {e}")
			# Translators: Could not retrieve collaboration statistics.
			html_parts.append(f"<li>{_('Could not retrieve collaboration statistics.')}</li>")
		html_parts.append("</ul>")
