# -*- coding: UTF-8 -*-
# Copyright (C) 2026 KIRAN G T {MisterK} and Antigravity 2
# This file is covered by the GNU General Public License (GPL), version 2.

import ui
import addonHandler
from logHandler import log
from appModules.boa_enhancements.async_engine import AsyncCOMTask

addonHandler.initTranslation()

class WordFormatAuditor:
	"""
	Scans the document asynchronously for visual formatting inconsistencies
	and WCAG structural blockers, reporting them in a categorized HTML UI.
	"""
	
	@staticmethod
	def audit(obj):
		try:
			doc = WordFormatAuditor._get_com_document(obj)
			if not doc:
				# Translators: Cannot access Word application.
				ui.message(_("Cannot access Word application."))
				return
				
			# Translators: Initial message when starting the format auditor
			ui.message(_("Auditing document formatting, please wait..."))
			
			auditor = WordFormatAuditor(doc)
			task = AsyncCOMTask(auditor._audit_generator(), on_complete=auditor._on_complete, chunk_size=30)
			task.start()
		except Exception as e:
			log.error(f"BOA WordFormatAuditor Error: {e}", exc_info=True)
			# Translators: Message spoken when auditor fails to launch
			ui.message(_("Failed to start formatting audit."))
			
	@staticmethod
	def _get_com_document(obj):
		doc = getattr(obj, "winwordDocumentObject", None)
		if not doc:
			doc = getattr(obj, "WinwordDocumentObject", None)
		if not doc:
			win_obj = getattr(obj, "winwordWindowObject", None)
			if not win_obj and obj.appModule:
				win_obj = getattr(obj.appModule, "winwordWindowObject", None)
			if win_obj:
				doc = getattr(win_obj, "document", getattr(win_obj, "Document", None))
		if not doc and obj.appModule:
			app = getattr(obj.appModule, "winwordApplicationObject", None)
			if app:
				try:
					doc = app.ActiveDocument
				except Exception:
					pass
		if not doc:
			try:
				import comtypes.client
				app = comtypes.client.GetActiveObject("Word.Application")
				doc = app.ActiveDocument
			except Exception:
				pass
		return doc

	def __init__(self, doc):
		self.doc = doc
		self.layout_errors = []
		self.font_errors = []
		self.struct_errors = []
		
	def _format_loc(self, para_idx, p):
		wdActiveEndPageNumber = 3
		wdFirstCharacterLineNumber = 10
		page = -1
		line = -1
		try:
			page = p.Information(wdActiveEndPageNumber)
			line = p.Information(wdFirstCharacterLineNumber)
		except Exception:
			pass
		
		if page != -1 and line != -1:
			# Translators: Formats the location of an error using Page, Line, and Paragraph numbers.
			return _("<b>Page {page}, Line {line}</b> (Para {num})").format(page=page, line=line, num=para_idx)
		elif page != -1:
			# Translators: Formats the location using Page and Paragraph numbers.
			return _("<b>Page {page}</b> (Para {num})").format(page=page, num=para_idx)
		else:
			# Translators: Formats the location using only the Paragraph number.
			return _("<b>Paragraph {num}</b>").format(num=para_idx)
			
	def _format_table_loc(self, t_idx, t):
		wdActiveEndPageNumber = 3
		page = -1
		try:
			page = t.Range.Information(wdActiveEndPageNumber)
		except Exception:
			pass
		
		if page != -1:
			# Translators: Formats the location of a table error using Page and Table numbers.
			return _("<b>Page {page}</b> (Table {num})").format(page=page, num=t_idx)
		else:
			# Translators: Formats the location using only the Table number.
			return _("<b>Table {num}</b>").format(num=t_idx)
		
	def _audit_generator(self):
		wdUndefined = 9999999
		try:
			count = self.doc.Paragraphs.Count
		except Exception:
			count = 0
			
		consecutive_blanks = 0
		last_heading_level = 0
		
		for i in range(1, count + 1):
			try:
				p = self.doc.Paragraphs.Item(i)
				text = p.Range.Text
				loc_str = self._format_loc(i, p.Range)
				
				# 1. Visual Layout & Spacing
				
				# Empty Paragraph Check
				if text in ('\r', '\r\n', '\n', '\x0d', '\x0b'):
					consecutive_blanks += 1
					if consecutive_blanks == 3:
						# Translators: Format auditor warning about blank lines
						self.layout_errors.append(_("{loc}: 3 or more consecutive blank lines detected. Use a Page Break instead.").format(loc=loc_str))
					
					# Empty Heading Check
					style = p.Style.NameLocal
					if "heading" in style.lower() or "titre" in style.lower():
						# Translators: Format auditor warning about empty headings
						self.struct_errors.append(_("{loc}: Empty heading detected.").format(loc=loc_str))
				else:
					consecutive_blanks = 0
					
					# Manual Line Breaks (Shift+Enter) inside paragraphs
					if '\x0b' in text:
						# Translators: Format auditor warning about manual line breaks
						self.layout_errors.append(_("{loc}: Manual line break (Shift+Enter) detected inside paragraph. This breaks screen reader flow.").format(loc=loc_str))
					
				# Spacebar indent check
				if text.startswith("    "):
					# Translators: Format auditor warning about spacebar indents
					self.layout_errors.append(_("{loc}: Spacebar indent detected (4+ spaces). Use Tab or Paragraph Indent.").format(loc=loc_str))
					
				# Tab Abuse check
				if text.startswith("\t\t") or "^t^t" in text:
					# Translators: Format auditor warning about tab abuse
					self.layout_errors.append(_("{loc}: Multiple consecutive Tab characters detected. Use Center/Right alignment or Ruler Indents.").format(loc=loc_str))
					
				# 2. Font & Style
				
				# Mixed Font checks
				if p.Range.Font.Size == wdUndefined:
					# Translators: Format auditor warning about mixed font sizes
					self.font_errors.append(_("{loc}: Inconsistent font sizes detected within the same paragraph.").format(loc=loc_str))
					
				if p.Range.Font.Name == "" or p.Range.Font.Name == wdUndefined:
					# Translators: Format auditor warning about mixed font names
					self.font_errors.append(_("{loc}: Inconsistent font types detected within the same paragraph.").format(loc=loc_str))
					
				# Highlights (0 means wdNoHighlight)
				if p.Range.HighlightColorIndex != 0 and p.Range.HighlightColorIndex != wdUndefined:
					# Translators: Format auditor warning about text highlights
					self.font_errors.append(_("{loc}: Highlighted text detected.").format(loc=loc_str))
					
				# 3. Structural Accessibility (WCAG)
				
				# Fake List check
				if text.startswith("1. ") or text.startswith("2. ") or text.startswith("- "):
					if p.Range.ListFormat.ListType == 0: # 0 means wdListNoNumbering
						# Translators: Format auditor warning about fake lists
						self.struct_errors.append(_("{loc}: Fake list detected. Manual numbering used instead of built-in list.").format(loc=loc_str))
						
				# Heading Hierarchy check
				try:
					level = p.OutlineLevel
					if level >= 1 and level <= 9: # Valid headings
						if last_heading_level != 0 and (level > last_heading_level + 1):
							# Translators: Format auditor warning about skipped heading levels
							self.struct_errors.append(_("{loc}: Skipped heading level (Jumped from Level {old} directly to Level {new}).").format(loc=loc_str, old=last_heading_level, new=level))
						last_heading_level = level
				except Exception:
					pass
					
			except Exception as err:
				log.debugWarning(f"BOA Format Auditor: Failed to parse paragraph {i}: {err}")
				
			yield # Hand control back to AsyncCOMTask
			
		# Tables
		try:
			t_count = self.doc.Tables.Count
			for t_idx in range(1, t_count + 1):
				try:
					t = self.doc.Tables.Item(t_idx)
					loc_str = self._format_table_loc(t_idx, t)
					
					# Uniformity (Merged/Split cells)
					if not t.Uniform:
						# Translators: Format auditor warning about merged cells
						self.struct_errors.append(_("{loc}: Table contains merged or split cells. This breaks screen reader grid navigation.").format(loc=loc_str))
						
					# Table Headers (-1 is True in COM boolean)
					if getattr(t.Rows.Item(1), "HeadingFormat", 0) != -1:
						# Translators: Format auditor warning about missing table headers
						self.struct_errors.append(_("{loc}: Table is missing 'Repeat Header Rows'. Data will be unreadable on subsequent pages.").format(loc=loc_str))
				except Exception as err:
					log.debugWarning(f"BOA Format Auditor: Failed to parse table {t_idx}: {err}")
				
				yield # Yield per table as well to prevent freezing!
		except Exception:
			pass
			
		return (self.layout_errors, self.font_errors, self.struct_errors)
		
	def _on_complete(self, result):
		if result is None:
			return
			
		layout_errs, font_errs, struct_errs = result
		total = len(layout_errs) + len(font_errs) + len(struct_errs)
		
		if total == 0:
			# Translators: Success message when no errors found.
			ui.browseableMessage(_("<h2>Format & Accessibility Audit Complete</h2><p>No formatting inconsistencies found! The document looks clean.</p>"), _("Format Audit Results"), isHtml=True, closeButton=True, copyButton=True)
			return
			
		html_parts = []
		# Translators: Heading for the format audit report
		html_parts.append(f"<h1>{_('Format & Accessibility Audit Report')}</h1>")
		try:
			paras = self.doc.Paragraphs.Count
		except Exception:
			paras = 0
		# Translators: Summary of paragraphs and errors
		html_parts.append(f"<p><b>{_('Total Paragraphs: {paras} | Total Errors Found: {total}').format(paras=paras, total=total)}</b></p>")
		
		if struct_errs:
			# Translators: Category heading for structural errors
			html_parts.append(f"<h2>{_('1. Structural Accessibility Blockers (WCAG)')}</h2><ul>")
			for err in struct_errs:
				html_parts.append(f"<li>{err}</li>")
			html_parts.append("</ul>")
			
		if layout_errs:
			# Translators: Category heading for visual layout errors
			html_parts.append(f"<h2>{_('2. Visual Layout & Spacing Issues')}</h2><ul>")
			for err in layout_errs:
				html_parts.append(f"<li>{err}</li>")
			html_parts.append("</ul>")
			
		if font_errs:
			# Translators: Category heading for font errors
			html_parts.append(f"<h2>{_('3. Font & Style Inconsistencies')}</h2><ul>")
			for err in font_errs:
				html_parts.append(f"<li>{err}</li>")
			html_parts.append("</ul>")
		
		# Translators: Title of the browseable message window
		ui.browseableMessage("".join(html_parts), _("Format Audit Results"), isHtml=True, closeButton=True, copyButton=True)
