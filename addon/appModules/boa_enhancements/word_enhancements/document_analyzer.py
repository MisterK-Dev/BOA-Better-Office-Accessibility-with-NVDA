# -*- coding: UTF-8 -*-
# Copyright (C) 2026 KIRAN G T {MisterK} and Antigravity 2
# Acknowledgment: COM property mappings and layout extraction logic derived from wordAccessEnhancement by paulber19.
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
		if not selection:
			selection = getattr(obj, "WinwordSelectionObject", None)
			
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
			wdWithInTable = 12
			
			get_info = getattr(self.selection, "Information", getattr(self.selection, "information", None))
			if get_info:
				page = get_info(wdActiveEndPageNumber)
				line = get_info(wdFirstCharacterLineNumber)
				col = get_info(wdFirstCharacterColumnNumber)
				in_table = get_info(wdWithInTable)
				
				# Get current section
				try:
					sec_idx = self.selection.Sections(1).Index
				except Exception:
					sec_idx = 1
				
				if in_table:
					try:
						cell = self.selection.Cells(1)
						r_idx = cell.RowIndex
						c_idx = cell.ColumnIndex
						# Translators: Context inside a table
						html_parts.append(f"<li>{_('In Section {sec}, Page {page}, Table Cell (Row {row}, Column {col})').format(sec=sec_idx, page=page, row=r_idx, col=c_idx)}</li>")
					except Exception:
						# Translators: Context inside a table without exact cell
						html_parts.append(f"<li>{_('In Section {sec}, Page {page} (Inside a Table)').format(sec=sec_idx, page=page)}</li>")
				else:
					# Translators: Context in normal text
					html_parts.append(f"<li>{_('In Section {sec}, Page {page}, Line {line}, Column {col}').format(sec=sec_idx, page=page, line=line, col=col)}</li>")
			else:
				# Translators: Error when cursor position cannot be read
				html_parts.append(f"<li>{_('Could not retrieve cursor position.')}</li>")
		except Exception:
			# Translators: General error retrieving cursor context
			html_parts.append(f"<li>{_('Error retrieving cursor context.')}</li>")
			
		html_parts.append("</ul>")
		yield
		
		# 2. Document Properties & Options
		# Translators: Heading for document properties
		html_parts.append(f"<h2>{_('2. Document Properties & Options')}</h2><ul>")
		try:
			try:
				prot = self.doc.ProtectionType
				prot_map = {
					-1: _("Unprotected"),
					0: _("Allow Only Revisions (Tracked Changes)"),
					1: _("Allow Only Comments"),
					2: _("Allow Only Form Fields")
				}
				prot_str = prot_map.get(prot, _("Unknown Protection"))
				# Translators: Document protection status
				html_parts.append(f"<li><strong>{_('Document Protection: {status}').format(status=prot_str)}</strong></li>")
			except Exception:
				pass
				
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
					prop = self.doc.BuiltInDocumentProperties(prop_name)
					try:
						val = prop.Value()
					except TypeError:
						val = prop.Value
					if val:
						html_parts.append(f"<li>{label}: {val}</li>")
				except Exception:
					pass
					
			if self.app:
				try:
					opt_obj = self.app.Options
					spell = _("Enabled") if getattr(opt_obj, "CheckSpellingAsYouType", False) else _("Disabled")
					gramm = _("Enabled") if getattr(opt_obj, "CheckGrammarAsYouType", False) else _("Disabled")
					gramm_spell = _("Enabled") if getattr(opt_obj, "CheckGrammarWithSpelling", False) else _("Disabled")
					
					# Translators: Check spelling as you type: {status}
					html_parts.append(f"<li>{_('Check spelling as you type: {status}').format(status=spell)}</li>")
					# Translators: Check grammar as you type: {status}
					html_parts.append(f"<li>{_('Check grammar as you type: {status}').format(status=gramm)}</li>")
					# Translators: Check grammar with spelling
					html_parts.append(f"<li>{_('Check grammar with spelling: {status}').format(status=gramm_spell)}</li>")
					
					track = _("Enabled") if getattr(self.doc, "TrackRevisions", False) else _("Disabled")
					# Translators: Track revisions status
					html_parts.append(f"<li>{_('Track Revisions (Recording): {status}').format(status=track)}</li>")
					
					try:
						hidden = _("Visible") if self.doc.ActiveWindow.View.ShowHiddenText else _("Hidden")
						# Translators: Hidden text visibility
						html_parts.append(f"<li>{_('Display hidden text on screen: {status}').format(status=hidden)}</li>")
					except Exception:
						pass
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
			
			try:
				list_count = self.doc.Lists.Count
			except Exception:
				list_count = 0
				
			try:
				spell_count = self.doc.SpellingErrors.Count
			except Exception:
				spell_count = 0
			
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
			# Translators: Lists: {lists}
			html_parts.append(f"<li>{_('Lists: {lists}').format(lists=list_count)}</li>")
			# Translators: Spelling errors: {errors}
			html_parts.append(f"<li>{_('Spelling errors: {errors}').format(errors=spell_count)}</li>")
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
						
					try:
						sec_start_map = {0: _("Continuous"), 1: _("New Column"), 2: _("New Page"), 3: _("Even Page"), 4: _("Odd Page")}
						sec_start = sec_start_map.get(getattr(ps, "SectionStart", 2), _("Unknown"))
						html_parts.append(f"<li>{_('Section start: {start}').format(start=sec_start)}</li>")
					except Exception:
						pass

					try:
						valign_map = {0: _("Top"), 1: _("Center"), 2: _("Justify"), 3: _("Bottom")}
						valign = valign_map.get(getattr(ps, "VerticalAlignment", 0), _("Top"))
						html_parts.append(f"<li>{_('Vertical alignment: {align}').format(align=valign)}</li>")
					except Exception:
						pass
					
					try:
						h_dist = self._convert_points_to_cm(getattr(ps, "HeaderDistance", 0))
						f_dist = self._convert_points_to_cm(getattr(ps, "FooterDistance", 0))
						html_parts.append(f"<li>{_('Distance from top to header: {d:.2f}cm').format(d=h_dist)}</li>")
						html_parts.append(f"<li>{_('Distance from bottom to footer: {d:.2f}cm').format(d=f_dist)}</li>")
					except Exception:
						pass
						
					try:
						prot_forms = _("Yes") if getattr(sec, "ProtectedForForms", False) else _("No")
						html_parts.append(f"<li>{_('Text modification only in form fields: {status}').format(status=prot_forms)}</li>")
					except Exception:
						pass

					try:
						paper_map = {7: "A4", 2: "Letter", 41: _("Custom"), 1: "11x17", 4: "Legal", 6: "A3", 9: "A5"}
						p_size = paper_map.get(getattr(ps, "PaperSize", 7), _("Standard/Other"))
						p_w = self._convert_points_to_cm(getattr(ps, "PageWidth", 0))
						p_h = self._convert_points_to_cm(getattr(ps, "PageHeight", 0))
						html_parts.append(f"<li>{_('Paper size: {size}').format(size=p_size)}</li>")
						html_parts.append(f"<li>{_('Page dimensions: {w:.2f}cm width, {h:.2f}cm height').format(w=p_w, h=p_h)}</li>")
						
						two_pages = _("Yes") if getattr(ps, "TwoPagesOnOne", False) else _("No")
						html_parts.append(f"<li>{_('Print two pages per sheet: {status}').format(status=two_pages)}</li>")
						
						dir_map = {0: _("Right-to-Left (RTL)"), 1: _("Left-to-Right (LTR)")}
						direction = dir_map.get(getattr(ps, "SectionDirection", 1), _("Left-to-Right (LTR)"))
						html_parts.append(f"<li>{_('Reading direction: {dir}').format(dir=direction)}</li>")
					except Exception:
						pass
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
						uniform = _("uniform") if getattr(tb, "Uniform", True) else _("non-uniform")
						nested = _("nested ") if getattr(tb, "NestingLevel", 1) > 1 else ""
						# Translators: Table dimensions
						html_parts.append(f"<li>{_('Dimensions: {nest}table ({uni}) of {rows} Rows, {cols} Columns').format(nest=nested, uni=uniform, rows=rows, cols=cols)}</li>")
					except Exception:
						# Translators: Error when table dimensions cannot be read due to merged cells
						html_parts.append(f"<li>{_('Dimensions: Complex (Contains merged/split cells preventing row count)')}</li>")
					
					try:
						wdActiveEndPageNumber = 3
						wdFirstCharacterLineNumber = 10
						tb_page = tb.Range.Information(wdActiveEndPageNumber)
						tb_line = tb.Range.Information(wdFirstCharacterLineNumber)
						# Translators: Localized at page {page}, line {line}
						html_parts.append(f"<li>{_('Localized at page {page}, line {line}').format(page=tb_page, line=tb_line)}</li>")
					except Exception:
						pass
						
					try:
						title = getattr(tb, "Title", "")
						if title:
							html_parts.append(f"<li>{_('Title: {t}').format(t=title)}</li>")
						desc = getattr(tb, "Descr", "")
						if desc:
							html_parts.append(f"<li>{_('Description: {d}').format(d=desc)}</li>")
					except Exception:
						pass
					
					try:
						autofit = _("Yes") if getattr(tb, "AllowAutoFit", False) else _("No")
						html_parts.append(f"<li>{_('Automatic resize to fit content: {status}').format(status=autofit)}</li>")
						
						t_pad = self._convert_points_to_cm(getattr(tb, "TopPadding", 0))
						b_pad = self._convert_points_to_cm(getattr(tb, "BottomPadding", 0))
						spacing = self._convert_points_to_cm(getattr(tb, "Spacing", 0))
						html_parts.append(f"<li>{_('Top padding: {pad:.4f}cm, Bottom padding: {bpad:.4f}cm').format(pad=t_pad, bpad=b_pad)}</li>")
						html_parts.append(f"<li>{_('Spacing between cells: {sp:.4f}cm').format(sp=spacing)}</li>")
					except Exception:
						pass
						
					try:
						borders = _("Yes") if getattr(tb.Borders, "Enable", False) else _("No")
						html_parts.append(f"<li>{_('Borders Enabled: {status}').format(status=borders)}</li>")
					except Exception:
						pass
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
