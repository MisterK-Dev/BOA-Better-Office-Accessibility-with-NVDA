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

class WordDocumentAnalyzer:
	"""
	Core engine for analyzing a Word Document's structural, accessibility, 
	and layout health, presenting the results in an HTML dialog.
	Now integrated with AsyncCOMTask to perform heavy analysis without freezing.
	"""

	def __init__(self, doc, selection, app):
		self.doc = doc
		self.selection = selection
		self.app = app

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
		"""
		# Translators: Analyzing document, please wait...
		ui.message(_("Analyzing document, please wait..."))
		
		try:
			doc, selection, app = WordDocumentAnalyzer._get_com_document_and_selection(obj)
			if not doc:
				# Translators: Cannot access Word application.
				ui.message(_("Cannot access Word application."))
				return
				
			auditor = WordDocumentAnalyzer(doc, selection, app)
			from appModules.boa_enhancements.async_engine import AsyncCOMTask
			task = AsyncCOMTask(auditor._analyze_generator(), on_complete=auditor._on_complete, max_tick_time=0.05)
			task.start()
		except Exception as e:
			log.error(f"BOA Word Analyzer Error: {e}", exc_info=True)
			# Translators: Message spoken when auditor fails to launch
			ui.message(_("Failed to start document analyzer."))

	@staticmethod
	def _get_com_document_and_selection(obj):
		doc = None
		selection = None
		app = None
		
		# 1. Try native NVDA Legacy properties
		doc = getattr(obj, "winwordDocumentObject", None)
		selection = getattr(obj, "winwordSelectionObject", None)
		if not doc:
			doc = getattr(obj, "WinwordDocumentObject", None)
			
		# 2. Try UIA Word Window Object's document property
		if not doc:
			win_obj = getattr(obj, "winwordWindowObject", None)
			if not win_obj and obj.appModule:
				win_obj = getattr(obj.appModule, "winwordWindowObject", None)
			if win_obj:
				doc = getattr(win_obj, "document", getattr(win_obj, "Document", None))
				selection = getattr(win_obj, "selection", getattr(win_obj, "Selection", None))
				
		# 3. Try appModule winwordApplicationObject ActiveDocument
		if obj.appModule:
			app = getattr(obj.appModule, "winwordApplicationObject", None)
			if app:
				if not doc:
					try:
						doc = app.ActiveDocument
					except Exception:
						pass
				if not selection:
					try:
						selection = app.Selection
					except Exception:
						pass
						
		# 4. Final Fallback to ROT
		if not doc or not selection:
			try:
				import comtypes.client
				app_fallback = comtypes.client.GetActiveObject("Word.Application")
				if not app:
					app = app_fallback
				if not doc:
					doc = app.ActiveDocument
				if not selection:
					selection = app.Selection
			except Exception:
				pass
				
		return doc, selection, app

	def _analyze_generator(self):
		html_parts = []
		# Translators: Main title of the Document Analyzer report
		html_parts.append(f"<h1>{_('Word Document Analyzer Report')}</h1>")
		
		# 1. Cursor Context
		# Translators: Heading for cursor context
		html_parts.append(f"<h2>{_('1. Cursor Context')}</h2><ul>")
		try:
			wdActiveEndPageNumber = 3
			wdFirstCharacterLineNumber = 10
			wdFirstCharacterColumnNumber = 9
			
			page = self.selection.Information(wdActiveEndPageNumber)
			line = self.selection.Information(wdFirstCharacterLineNumber)
			col = self.selection.Information(wdFirstCharacterColumnNumber)
			
			# Translators: Page: {page}, Line: {line}, Column: {col}
			html_parts.append(f"<li>{_('Page: {page}, Line: {line}, Column: {col}').format(page=page, line=line, col=col)}</li>")
		except Exception:
			# Translators: Error when cursor position cannot be read
			html_parts.append(f"<li>{_('Could not retrieve cursor position.')}</li>")
		html_parts.append("</ul>")
		yield
		
		# 2. Document Properties & Options
		# Translators: Heading for document properties
		html_parts.append(f"<h2>{_('2. Document Properties & Options')}</h2><ul>")
		try:
			try:
				name = self.doc.Name
				# Translators: File Name: {name}
				html_parts.append(f"<li>{_('File Name: {name}').format(name=name)}</li>")
			except Exception:
				pass
				
			props = [
				("Author", _("Author")),
				("Creation Date", _("Creation Date")),
				("Revision Number", _("Revision Number")),
				("Title", _("Title")),
				("Subject", _("Subject")),
				("Company", _("Company"))
			]
			
			for prop_name, label in props:
				try:
					val = self.doc.BuiltInDocumentProperties(prop_name).Value
					if val:
						html_parts.append(f"<li>{label}: {val}</li>")
				except Exception:
					pass
					
			if self.app:
				try:
					opt_obj = self.app.Options
					spell = _("Enabled") if opt_obj.CheckSpellingAsYouType else _("Disabled")
					gramm = _("Enabled") if opt_obj.CheckGrammarAsYouType else _("Disabled")
					# Translators: Check spelling as you type: {status}
					html_parts.append(f"<li>{_('Check spelling as you type: {status}').format(status=spell)}</li>")
					# Translators: Check grammar as you type: {status}
					html_parts.append(f"<li>{_('Check grammar as you type: {status}').format(status=gramm)}</li>")
				except Exception:
					pass
		except Exception:
			# Translators: Error when properties cannot be read
			html_parts.append(f"<li>{_('Could not retrieve document properties.')}</li>")
		html_parts.append("</ul>")
		yield
		
		# 3. Overall Statistics
		# Translators: Heading for overall statistics
		html_parts.append(f"<h2>{_('3. Overall Statistics')}</h2><ul>")
		try:
			pages = self.doc.ComputeStatistics(WD_STAT_PAGES, True)
			words = self.doc.ComputeStatistics(WD_STAT_WORDS, True)
			chars = self.doc.ComputeStatistics(WD_STAT_CHARS, True)
			
			# Architectural Note: 
			# `doc.Paragraphs.Count` returns the raw structural paragraph count. This includes 
			# completely empty lines, lines inside tables, and structural markers.
			# `doc.ComputeStatistics(WD_STAT_PARAGRAPHS)` only counts actual text paragraphs 
			# that contain words, largely ignoring tables and empty lines. We display both 
			# because screen readers navigate via the structural count, but sighted users 
			# think of paragraphs via the text count.
			paras_structural = self.doc.Paragraphs.Count
			paras_text = self.doc.ComputeStatistics(WD_STAT_PARAGRAPHS, True)
			lines = self.doc.ComputeStatistics(WD_STAT_LINES, True)
			
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
		except Exception:
			# Translators: Error when structural statistics cannot be read
			html_parts.append(f"<li>{_('Could not retrieve structural statistics.')}</li>")
		html_parts.append("</ul>")
		yield
		
		# 4. Collaboration Status
		# Translators: Heading for collaboration status
		html_parts.append(f"<h2>{_('4. Collaboration Status')}</h2><ul>")
		try:
			co_count = self.doc.Comments.Count
			# Translators: Comments: {count}
			html_parts.append(f"<li>{_('Comments: {count}').format(count=co_count)}</li>")
			
			rev_count = self.doc.Revisions.Count
			if rev_count > 0:
				# Translators: Unresolved Tracked Changes: {count}
				html_parts.append(f"<li><strong>{_('Unresolved Tracked Changes: {count}').format(count=rev_count)}</strong></li>")
			else:
				# Translators: Unresolved Tracked Changes: 0
				html_parts.append(f"<li>{_('Unresolved Tracked Changes: 0')}</li>")
				
			fn_count = self.doc.Footnotes.Count
			# Translators: Footnotes: {count}
			html_parts.append(f"<li>{_('Footnotes: {count}').format(count=fn_count)}</li>")
			
			en_count = self.doc.Endnotes.Count
			# Translators: Endnotes: {count}
			html_parts.append(f"<li>{_('Endnotes: {count}').format(count=en_count)}</li>")
		except Exception:
			# Translators: Error when collaboration stats cannot be read
			html_parts.append(f"<li>{_('Could not retrieve collaboration statistics.')}</li>")
		html_parts.append("</ul>")
		yield
		
		# 5. Accessibility Audit Summary
		# Translators: Heading for accessibility summary
		html_parts.append(f"<h2>{_('5. Accessibility Audit Summary')}</h2><ul>")
		try:
			inline_shapes = self.doc.InlineShapes
			in_count = inline_shapes.Count
			missing_alt = 0
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
					yield # Async loop yield
			# Translators: Images (Inline): {count}
			html_parts.append(f"<li>{_('Images (Inline): {count}').format(count=in_count)}</li>")
			if missing_alt > 0:
				# Translators: Warning: {count} images are missing alternative text!
				html_parts.append(f"<li><strong>{_('Warning: {count} images are missing alternative text!') .format(count=missing_alt)}</strong></li>")
			
			sh_count = self.doc.Shapes.Count
			if sh_count > 0:
				# Translators: Warning: {count} floating shapes detected. These are often inaccessible.
				html_parts.append(f"<li><strong>{_('Warning: {count} floating shapes detected. These are often inaccessible.').format(count=sh_count)}</strong></li>")
			
			link_count = self.doc.Hyperlinks.Count
			# Translators: Hyperlinks: {count}
			html_parts.append(f"<li>{_('Hyperlinks: {count}').format(count=link_count)}</li>")
		except Exception:
			# Translators: Error when accessibility stats cannot be read
			html_parts.append(f"<li>{_('Could not retrieve accessibility statistics.')}</li>")
		html_parts.append("</ul>")
		yield
		
		# 6. Layout & Structural Breakdown
		# Translators: Heading for layout breakdown
		html_parts.append(f"<h2>{_('6. Layout & Structural Breakdown')}</h2>")
		
		# Sections
		# Translators: Heading for Sections list
		html_parts.append(f"<h3>{_('Sections')}</h3>")
		try:
			sec_count = self.doc.Sections.Count
			if sec_count == 0:
				# Translators: No sections found message
				html_parts.append(f"<p>{_('No sections found.')}</p>")
			for i in range(1, sec_count + 1):
				# Translators: Section heading number
				html_parts.append(f"<h4>{_('Section {num}').format(num=i)}</h4><ul>")
				try:
					sec = self.doc.Sections.Item(i)
					ps = sec.PageSetup
					
					# Translators: Portrait Orientation
					orient = _("Portrait") if ps.Orientation == WD_ORIENT_PORTRAIT else _("Landscape")
					# Translators: Orientation info
					html_parts.append(f"<li>{_('Orientation: {orient}').format(orient=orient)}</li>")
					
					t_cm = self._convert_points_to_cm(ps.TopMargin)
					b_cm = self._convert_points_to_cm(ps.BottomMargin)
					l_cm = self._convert_points_to_cm(ps.LeftMargin)
					r_cm = self._convert_points_to_cm(ps.RightMargin)
					# Translators: Margin dimensions string
					margins = _("Top: {t:.2f}cm, Bottom: {b:.2f}cm, Left: {l:.2f}cm, Right: {r:.2f}cm").format(t=t_cm, b=b_cm, l=l_cm, r=r_cm)
					# Translators: Margins info
					html_parts.append(f"<li>{_('Margins: {margins}').format(margins=margins)}</li>")
					
					if getattr(ps, "DifferentFirstPageHeaderFooter", False):
						# Translators: Different first page header warning
						html_parts.append(f"<li><strong>{_('Warning: Different first page header/footer.')}</strong></li>")
					if getattr(ps, "OddAndEvenPagesHeaderFooter", False):
						# Translators: Odd/Even header warning
						html_parts.append(f"<li><strong>{_('Warning: Odd and even pages have different headers/footers.')}</strong></li>")
				except Exception:
					pass
				html_parts.append("</ul>")
				yield # Async yield
		except Exception:
			pass
			
		# Tables
		# Translators: Heading for Tables list
		html_parts.append(f"<h3>{_('Tables')}</h3>")
		try:
			tb_count = self.doc.Tables.Count
			if tb_count == 0:
				# Translators: No tables found message
				html_parts.append(f"<p>{_('No tables found.')}</p>")
			for i in range(1, tb_count + 1):
				# Translators: Table heading number
				html_parts.append(f"<h4>{_('Table {num}').format(num=i)}</h4><ul>")
				try:
					tb = self.doc.Tables.Item(i)
					try:
						rows = tb.Rows.Count
						cols = tb.Columns.Count
						# Translators: Table dimensions
						html_parts.append(f"<li>{_('Dimensions: {rows} Rows, {cols} Columns').format(rows=rows, cols=cols)}</li>")
					except Exception:
						# Translators: Error when table dimensions cannot be read due to merged cells
						html_parts.append(f"<li>{_('Dimensions: Complex (Contains merged/split cells preventing row count)')}</li>")
					
					if not getattr(tb, "Uniform", True):
						# Translators: Merged cells warning
						html_parts.append(f"<li><strong>{_('Warning: Contains merged or split cells.')}</strong></li>")
				except Exception:
					pass
				html_parts.append("</ul>")
				yield # Async yield
		except Exception:
			pass
			
		return html_parts
		
	def _on_complete(self, html_parts):
		if html_parts:
			html_content = "".join(html_parts)
			wx.CallAfter(ui.browseableMessage, html_content, title=_("Word Document Analyzer"), isHtml=True, closeButton=True, copyButton=True)
