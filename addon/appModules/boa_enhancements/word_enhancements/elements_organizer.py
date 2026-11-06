# -*- coding: UTF-8 -*-
# Copyright (C) 2026 KIRAN G T {MisterK} and Antigravity 2
# Acknowledgment: Inspired by conceptual Word object enumeration in wordAccessEnhancement by paulber19 (GPL-2.0).
# This file is covered by the GNU General Public License (GPL), version 2.
# See the file COPYING.txt for more details.

import os
import re
import addonHandler
import ui
import wx
import core
import gui
from logHandler import log
from keyboardHandler import KeyboardInputGesture
from appModules.boa_enhancements.async_engine import AsyncCOMTask

addonHandler.initTranslation()

# Word COM Constants
WD_FIRST_CHAR_LINE = 10  # wdFirstCharacterLineNumber
WD_ACTIVE_END_PAGE = 3   # wdActiveEndPageNumber
WD_ACTIVE_END_ADJ_PAGE = 1 # wdActiveEndAdjustedPageNumber

# Revision Types mapping
REVISION_TYPES = {
	# Translators: Word revision type
	0: _("No Revision"),
	# Translators: Word revision type
	1: _("Insertion"),
	# Translators: Word revision type
	2: _("Deletion"),
	# Translators: Word revision type
	3: _("Property Changed"),
	# Translators: Word revision type
	4: _("Paragraph Number Changed"),
	# Translators: Word revision type
	5: _("Field Display Changed"),
	# Translators: Word revision type
	6: _("Reconciled Conflict"),
	# Translators: Word revision type
	7: _("Conflict"),
	# Translators: Word revision type
	8: _("Style Changed"),
	# Translators: Word revision type
	9: _("Replaced"),
	# Translators: Word revision type
	10: _("Paragraph Property Changed"),
	# Translators: Word revision type
	11: _("Table Property Changed"),
	# Translators: Word revision type
	12: _("Section Property Changed"),
	# Translators: Word revision type
	13: _("Style Definition Changed"),
	# Translators: Word revision type
	14: _("Moved From"),
	# Translators: Word revision type
	15: _("Moved To"),
	# Translators: Word revision type
	16: _("Table Cell Inserted"),
	# Translators: Word revision type
	17: _("Table Cell Deleted"),
	# Translators: Word revision type
	18: _("Table Cell Merged"),
}

# Supported Element Types
ELEMENT_TYPES = [
	# Translators: Document element type in Elements Organizer
	("comment", _("Comments")),
	# Translators: Document element type in Elements Organizer
	("revision", _("Revisions (Track Changes)")),
	# Translators: Document element type in Elements Organizer
	("footnote", _("Footnotes")),
	# Translators: Document element type in Elements Organizer
	("endnote", _("Endnotes")),
	# Translators: Document element type in Elements Organizer
	("bookmark", _("Bookmarks")),
	# Translators: Document element type in Elements Organizer
	("heading", _("Headings / Titles")),
	# Translators: Document element type in Elements Organizer
	("table", _("Tables")),
	# Translators: Document element type in Elements Organizer
	("hyperlink", _("Hyperlinks")),
	# Translators: Document element type in Elements Organizer
	("field", _("Fields")),
	# Translators: Document element type in Elements Organizer
	("formfield", _("Form Fields")),
	# Translators: Document element type in Elements Organizer
	("contentcontrol", _("Content Controls")),
	# Translators: Document element type in Elements Organizer
	("drawingshape", _("Drawing Shapes (Floating)")),
	# Translators: Document element type in Elements Organizer
	("inlineshape", _("Inline Shapes (Text Layer)")),
	# Translators: Document element type in Elements Organizer
	("frame", _("Frames")),
	# Translators: Document element type in Elements Organizer
	("section", _("Sections")),
	# Translators: Document element type in Elements Organizer
	("spellingerror", _("Spelling Errors")),
	# Translators: Document element type in Elements Organizer
	("grammaticalerror", _("Grammatical Errors")),
]

# Supported Search Scopes
SCOPE_TYPES = [
	# Translators: Search scope option in Elements Organizer
	("selection", _("Selection / Cursor Area")),
	# Translators: Search scope option in Elements Organizer
	("document", _("Whole Document")),
	# Translators: Search scope option in Elements Organizer
	("page", _("Current Page")),
	# Translators: Search scope option in Elements Organizer
	("cursortoend", _("Cursor to End of Document")),
	# Translators: Search scope option in Elements Organizer
	("starttocursor", _("Start of Document to Cursor")),
]


class WordElementItem:
	"""
	Container for a single extracted document element item.
	Contains pre-cached text and metadata, with lazy calculation of Page and Line numbers.
	"""
	def __init__(self, raw_obj, elem_type, start_pos, snippet, author="", date_str="", ref_id="", full_text="", extra=None, end_pos=None):
		self.raw_obj = raw_obj
		self.elem_type = elem_type
		self.start_pos = start_pos
		self.end_pos = end_pos if end_pos is not None else start_pos
		self.snippet = snippet
		self.author = author
		self.date_str = date_str
		self.ref_id = ref_id
		self.full_text = full_text
		self.extra = extra or {}
		self._page = None
		self._line = None

	def get_meta_display(self):
		"""
		Returns the formatted string for Column 2 ('Author / Ref') in the elements list.
		For comments, incorporates the author and total reply count for complete thread context.
		"""
		if self.elem_type == "comment":
			replies = self.extra.get("replies", [])
			reply_count = len(replies)
			# Translators: Author fallback when comment author is blank
			author = self.author or _("Unknown")
			if reply_count == 1:
				# Translators: Comment metadata in list showing author and single reply count
				return _("{author} (1 reply)").format(author=author)
			elif reply_count > 1:
				# Translators: Comment metadata in list showing author and multiple reply count
				return _("{author} ({count} replies)").format(author=author, count=reply_count)
			return author
		return self.author or self.ref_id

	def get_page_line(self, doc):
		"""
		Lazily computes the Page and Line numbers only when the item is inspected.
		This completely eliminates the COM RPC pagination freeze.
		"""
		if self._page is not None and self._line is not None:
			return self._page, self._line
		try:
			r = None
			if self.elem_type in ("spellingerror", "grammaticalerror"):
				r = self.raw_obj
			else:
				try:
					if hasattr(self.raw_obj, "Range"):
						r = self.raw_obj.Range
					elif hasattr(self.raw_obj, "Reference"):
						r = self.raw_obj.Reference
					elif hasattr(self.raw_obj, "Anchor"):
						r = self.raw_obj.Anchor
					elif hasattr(self.raw_obj, "Scope"):
						r = self.raw_obj.Scope
				except Exception:
					pass

			if not r and doc and self.start_pos >= 0:
				r = doc.Range(self.start_pos, self.start_pos)

			if r:
				self._page = r.Information(WD_ACTIVE_END_PAGE)
				self._line = r.Information(WD_FIRST_CHAR_LINE)
		except Exception:
			pass
		return self._page, self._line


def get_scope_range(doc, selection, scope_key):
	"""
	Calculates the exact Word COM Range for the chosen scope.
	"""
	try:
		if scope_key == "selection" and selection:
			if selection.Start != selection.End:
				return selection.Range
			# If cursor has no active selection span, look at the paragraph around the cursor
			try:
				return selection.Paragraphs(1).Range
			except Exception:
				return selection.Range

		elif scope_key == "page" and selection:
			# Native Word bookmark for the active page
			try:
				return doc.Bookmarks(r"\Page").Range
			except Exception:
				pass

		elif scope_key == "cursortoend" and selection:
			return doc.Range(selection.Start, doc.Content.End)

		elif scope_key == "starttocursor" and selection:
			return doc.Range(0, selection.End)

	except Exception as e:
		log.debugWarning(f"BOA ElementsOrganizer get_scope_range error: {e}")

	# Default fallback: Whole Document
	return doc.Range(0, doc.Content.End)


def is_range_in_scope(target_start, target_end, scope_start, scope_end):
	"""
	Determines if an element falls within the selected scope character boundaries.
	"""
	if scope_start is None or scope_end is None:
		return True
	return (target_start >= scope_start and target_start <= scope_end)


def extract_elements_generator(doc, selection, elem_type, scope_key, start_offset=1):
	"""
	Background generator that safely harvests elements chunk by chunk via AsyncCOMTask.
	Yields control immediately and periodically to prevent NVDA or Word from freezing.
	Returns: dict with items, has_more, next_offset, and scanned_count.
	"""
	# Immediate yield to ensure dialog renders and speaks its title
	yield True

	results = []
	has_more = False
	next_offset = None
	scanned_count = 0

	try:
		scope_range = get_scope_range(doc, selection, scope_key)
		s_start = scope_range.Start if scope_range else 0
		s_end = scope_range.End if scope_range else doc.Content.End

		# 1. COMMENTS
		if elem_type == "comment":
			comments = doc.Comments
			count = comments.Count
			batch_size = 250
			end_idx = min(start_offset + batch_size - 1, count)
			has_more = (end_idx < count)
			next_offset = (end_idx + 1) if has_more else None
			scanned_count = end_idx

			for i in range(start_offset, end_idx + 1):
				if i % 5 == 0:
					yield True
				try:
					c = comments(i)
					try:
						# Filter out replies: child comments have an Ancestor pointing to the root comment
						if getattr(c, "Ancestor", None) is not None:
							continue
					except Exception:
						pass

					c_start = c.Scope.Start
					c_end = c.Scope.End
					if not is_range_in_scope(c_start, c_end, s_start, s_end):
						continue

					body_text = c.Range.Text.strip() if c.Range and c.Range.Text else ""
					snippet = body_text.replace("\r", " ").replace("\n", " ")
					if len(snippet) > 60:
						snippet = snippet[:57] + "..."

					author = getattr(c, "Author", "")
					date_str = str(getattr(c, "Date", ""))
					is_done = bool(getattr(c, "Done", False))

					replies = []
					try:
						rep_col = getattr(c, "Replies", None)
						if rep_col:
							for r_idx in range(1, rep_col.Count + 1):
								r_item = rep_col(r_idx)
								rep_text = r_item.Range.Text.strip() if r_item.Range and r_item.Range.Text else ""
								rep_author = getattr(r_item, "Author", "")
								rep_date = str(getattr(r_item, "Date", ""))
								replies.append({
									"author": rep_author,
									"date": rep_date,
									"text": rep_text
								})
					except Exception:
						pass

					item = WordElementItem(
						raw_obj=c,
						elem_type="comment",
						start_pos=c_start,
						end_pos=c_end,
						snippet=snippet,
						author=author,
						date_str=date_str,
						full_text=body_text,
						extra={
							"is_done": is_done,
							"replies": replies,
							"scope_text": c.Scope.Text if c.Scope else ""
						}
					)
					results.append(item)
				except Exception as e:
					log.debugWarning(f"BOA ElementsOrganizer comment error at {i}: {e}")

		# 2. REVISIONS (TRACK CHANGES)
		elif elem_type == "revision":
			revs = doc.Revisions
			count = revs.Count
			batch_size = 200
			end_idx = min(start_offset + batch_size - 1, count)
			has_more = (end_idx < count)
			next_offset = (end_idx + 1) if has_more else None
			scanned_count = end_idx

			for i in range(start_offset, end_idx + 1):
				if i % 5 == 0:
					yield True
				try:
					r = revs(i)
					r_start = r.Range.Start
					r_end = r.Range.End
					if not is_range_in_scope(r_start, r_end, s_start, s_end):
						continue

					rev_type_int = int(getattr(r, "Type", 0))
					# Translators: Fallback label for an unknown revision type
					type_str = REVISION_TYPES.get(rev_type_int, _("Revision"))
					text_val = r.Range.Text.strip() if r.Range and r.Range.Text else ""
					snippet = f"[{type_str}] {text_val.replace(chr(13), ' ').replace(chr(10), ' ')}"
					if len(snippet) > 60:
						snippet = snippet[:57] + "..."

					item = WordElementItem(
						raw_obj=r,
						elem_type="revision",
						start_pos=r_start,
						end_pos=r_end,
						snippet=snippet,
						author=getattr(r, "Author", ""),
						date_str=str(getattr(r, "Date", "")),
						ref_id=type_str,
						full_text=text_val,
						extra={"rev_type": rev_type_int, "type_name": type_str}
					)
					results.append(item)
				except Exception as e:
					log.debugWarning(f"BOA ElementsOrganizer revision error at {i}: {e}")

		# 3. FOOTNOTES
		elif elem_type == "footnote":
			notes = doc.Footnotes
			count = notes.Count
			batch_size = 250
			end_idx = min(start_offset + batch_size - 1, count)
			has_more = (end_idx < count)
			next_offset = (end_idx + 1) if has_more else None
			scanned_count = end_idx

			for i in range(start_offset, end_idx + 1):
				if i % 5 == 0:
					yield True
				try:
					fn = notes(i)
					fn_start = fn.Reference.Start
					fn_end = fn.Reference.End
					if not is_range_in_scope(fn_start, fn_end, s_start, s_end):
						continue

					ref_idx = str(getattr(fn, "Index", i))
					fn_text = fn.Range.Text.strip() if fn.Range and fn.Range.Text else ""
					snippet = f"[{ref_idx}] {fn_text.replace(chr(13), ' ').replace(chr(10), ' ')}"
					if len(snippet) > 60:
						snippet = snippet[:57] + "..."

					item = WordElementItem(
						raw_obj=fn,
						elem_type="footnote",
						start_pos=fn_start,
						end_pos=fn_end,
						snippet=snippet,
						ref_id=f"#{ref_idx}",
						full_text=fn_text
					)
					results.append(item)
				except Exception as e:
					log.debugWarning(f"BOA ElementsOrganizer footnote error at {i}: {e}")

		# 4. ENDNOTES
		elif elem_type == "endnote":
			notes = doc.Endnotes
			count = notes.Count
			batch_size = 250
			end_idx = min(start_offset + batch_size - 1, count)
			has_more = (end_idx < count)
			next_offset = (end_idx + 1) if has_more else None
			scanned_count = end_idx

			for i in range(start_offset, end_idx + 1):
				if i % 5 == 0:
					yield True
				try:
					en = notes(i)
					en_start = en.Reference.Start
					en_end = en.Reference.End
					if not is_range_in_scope(en_start, en_end, s_start, s_end):
						continue

					ref_idx = str(getattr(en, "Index", i))
					en_text = en.Range.Text.strip() if en.Range and en.Range.Text else ""
					snippet = f"[{ref_idx}] {en_text.replace(chr(13), ' ').replace(chr(10), ' ')}"
					if len(snippet) > 60:
						snippet = snippet[:57] + "..."

					item = WordElementItem(
						raw_obj=en,
						elem_type="endnote",
						start_pos=en_start,
						end_pos=en_end,
						snippet=snippet,
						ref_id=f"#{ref_idx}",
						full_text=en_text
					)
					results.append(item)
				except Exception as e:
					log.debugWarning(f"BOA ElementsOrganizer endnote error at {i}: {e}")

		# 5. BOOKMARKS
		elif elem_type == "bookmark":
			bms = doc.Bookmarks
			count = bms.Count
			batch_size = 250
			end_idx = min(start_offset + batch_size - 1, count)
			has_more = (end_idx < count)
			next_offset = (end_idx + 1) if has_more else None
			scanned_count = end_idx

			for i in range(start_offset, end_idx + 1):
				if i % 5 == 0:
					yield True
				try:
					bm = bms(i)
					bm_start = bm.Range.Start
					bm_end = bm.Range.End
					if not is_range_in_scope(bm_start, bm_end, s_start, s_end):
						continue

					bm_name = bm.Name
					text_val = bm.Range.Text.strip() if bm.Range and bm.Range.Text else ""
					snippet = f"{bm_name}: {text_val[:40]}" if text_val else bm_name

					item = WordElementItem(
						raw_obj=bm,
						elem_type="bookmark",
						start_pos=bm_start,
						end_pos=bm_end,
						snippet=snippet,
						ref_id=bm_name,
						full_text=text_val
					)
					results.append(item)
				except Exception as e:
					log.debugWarning(f"BOA ElementsOrganizer bookmark error at {i}: {e}")

		# 6. HEADINGS / TITLES
		elif elem_type == "heading":
			paras = doc.Paragraphs
			count = paras.Count
			batch_size = 500
			end_idx = min(start_offset + batch_size - 1, count)
			has_more = (end_idx < count)
			next_offset = (end_idx + 1) if has_more else None
			scanned_count = end_idx

			for i in range(start_offset, end_idx + 1):
				if i % 10 == 0:
					yield True
				try:
					p = paras(i)
					level = int(getattr(p, "OutlineLevel", 10))
					# Outline levels 1 through 9 are headings
					if level < 10:
						p_start = p.Range.Start
						p_end = p.Range.End
						if not is_range_in_scope(p_start, p_end, s_start, s_end):
							continue

						h_text = p.Range.Text.strip() if p.Range and p.Range.Text else ""
						snippet = f"[H{level}] {h_text.replace(chr(13), ' ').replace(chr(10), ' ')}"
						if len(snippet) > 60:
							snippet = snippet[:57] + "..."

						item = WordElementItem(
							raw_obj=p,
							elem_type="heading",
							start_pos=p_start,
							end_pos=p_end,
							snippet=snippet,
							ref_id=f"H{level}",
							full_text=h_text,
							extra={"level": level}
						)
						results.append(item)
				except Exception as e:
					log.debugWarning(f"BOA ElementsOrganizer heading error at {i}: {e}")

		# 7. TABLES
		elif elem_type == "table":
			tables = doc.Tables
			count = tables.Count
			batch_size = 250
			end_idx = min(start_offset + batch_size - 1, count)
			has_more = (end_idx < count)
			next_offset = (end_idx + 1) if has_more else None
			scanned_count = end_idx

			for i in range(start_offset, end_idx + 1):
				if i % 5 == 0:
					yield True
				try:
					tbl = tables(i)
					tbl_start = tbl.Range.Start
					tbl_end = tbl.Range.End
					if not is_range_in_scope(tbl_start, tbl_end, s_start, s_end):
						continue

					rows_c = tbl.Rows.Count
					cols_c = tbl.Columns.Count
					title_val = getattr(tbl, "Title", "") or ""
					descr_val = getattr(tbl, "Descr", "") or ""
					has_hf = False
					try:
						has_hf = bool(tbl.Rows(1).HeadingFormat != 0)
					except Exception:
						pass

					first_cell_txt = ""
					try:
						first_cell_txt = tbl.Cell(1, 1).Range.Text.strip()
					except Exception:
						pass

					desc = title_val if title_val else first_cell_txt[:35]
					# Translators: Table dimension format: rows by columns
					snippet = _("Table {idx} ({r}x{c}): {desc}").format(idx=i, r=rows_c, c=cols_c, desc=desc)

					item = WordElementItem(
						raw_obj=tbl,
						elem_type="table",
						start_pos=tbl_start,
						end_pos=tbl_end,
						snippet=snippet,
						ref_id=f"Table {i}",
						full_text=snippet,
						extra={
							"rows": rows_c,
							"cols": cols_c,
							"heading_format": has_hf,
							"title": title_val,
							"descr": descr_val
						}
					)
					results.append(item)
				except Exception as e:
					log.debugWarning(f"BOA ElementsOrganizer table error at {i}: {e}")

		# 8. HYPERLINKS
		elif elem_type == "hyperlink":
			links = doc.Hyperlinks
			count = links.Count
			batch_size = 250
			end_idx = min(start_offset + batch_size - 1, count)
			has_more = (end_idx < count)
			next_offset = (end_idx + 1) if has_more else None
			scanned_count = end_idx

			for i in range(start_offset, end_idx + 1):
				if i % 5 == 0:
					yield True
				try:
					hl = links(i)
					hl_start = hl.Range.Start
					hl_end = hl.Range.End
					if not is_range_in_scope(hl_start, hl_end, s_start, s_end):
						continue

					disp_txt = hl.TextToDisplay or ""
					addr = getattr(hl, "Address", "") or getattr(hl, "SubAddress", "") or ""
					snippet = f"{disp_txt} -> {addr}"
					if len(snippet) > 60:
						snippet = snippet[:57] + "..."

					item = WordElementItem(
						raw_obj=hl,
						elem_type="hyperlink",
						start_pos=hl_start,
						end_pos=hl_end,
						snippet=snippet,
						ref_id=disp_txt,
						full_text=addr,
						extra={"address": addr, "display": disp_txt}
					)
					results.append(item)
				except Exception as e:
					log.debugWarning(f"BOA ElementsOrganizer hyperlink error at {i}: {e}")

		# 9. FIELDS
		elif elem_type == "field":
			fields = doc.Fields
			count = fields.Count
			batch_size = 200
			end_idx = min(start_offset + batch_size - 1, count)
			has_more = (end_idx < count)
			next_offset = (end_idx + 1) if has_more else None
			scanned_count = end_idx

			for i in range(start_offset, end_idx + 1):
				if i % 5 == 0:
					yield True
				try:
					f = fields(i)
					f_start = f.Code.Start
					f_end = f.Code.End
					if not is_range_in_scope(f_start, f_end, s_start, s_end):
						continue

					f_type = str(getattr(f, "Type", ""))
					f_code = f.Code.Text.strip() if f.Code and f.Code.Text else ""
					snippet = f"{{{f_code[:40]}}}"

					item = WordElementItem(
						raw_obj=f,
						elem_type="field",
						start_pos=f_start,
						end_pos=f_end,
						snippet=snippet,
						ref_id=f"Field {f_type}",
						full_text=f_code
					)
					results.append(item)
				except Exception as e:
					log.debugWarning(f"BOA ElementsOrganizer field error at {i}: {e}")

		# 10. FORM FIELDS
		elif elem_type == "formfield":
			ffs = doc.FormFields
			count = ffs.Count
			batch_size = 250
			end_idx = min(start_offset + batch_size - 1, count)
			has_more = (end_idx < count)
			next_offset = (end_idx + 1) if has_more else None
			scanned_count = end_idx

			for i in range(start_offset, end_idx + 1):
				if i % 5 == 0:
					yield True
				try:
					ff = ffs(i)
					ff_start = ff.Range.Start
					ff_end = ff.Range.End
					if not is_range_in_scope(ff_start, ff_end, s_start, s_end):
						continue

					ff_name = ff.Name
					ff_val = ff.Result if hasattr(ff, "Result") else ""
					snippet = f"{ff_name}: '{ff_val}'" if ff_val else ff_name

					item = WordElementItem(
						raw_obj=ff,
						elem_type="formfield",
						start_pos=ff_start,
						end_pos=ff_end,
						snippet=snippet,
						ref_id=ff_name,
						full_text=ff_val
					)
					results.append(item)
				except Exception as e:
					log.debugWarning(f"BOA ElementsOrganizer formfield error at {i}: {e}")

		# 11. CONTENT CONTROLS
		elif elem_type == "contentcontrol":
			ccs = doc.ContentControls
			count = ccs.Count
			batch_size = 250
			end_idx = min(start_offset + batch_size - 1, count)
			has_more = (end_idx < count)
			next_offset = (end_idx + 1) if has_more else None
			scanned_count = end_idx

			for i in range(start_offset, end_idx + 1):
				if i % 5 == 0:
					yield True
				try:
					cc = ccs(i)
					cc_start = cc.Range.Start
					cc_end = cc.Range.End
					if not is_range_in_scope(cc_start, cc_end, s_start, s_end):
						continue

					# Translators: Fallback title for an unnamed content control
					title = cc.Title or cc.Tag or _("Unnamed Control")
					txt = cc.Range.Text.strip() if cc.Range and cc.Range.Text else ""
					snippet = f"{title}: {txt[:40]}" if txt else title

					item = WordElementItem(
						raw_obj=cc,
						elem_type="contentcontrol",
						start_pos=cc_start,
						end_pos=cc_end,
						snippet=snippet,
						ref_id=title,
						full_text=txt
					)
					results.append(item)
				except Exception as e:
					log.debugWarning(f"BOA ElementsOrganizer contentcontrol error at {i}: {e}")

		# 12. DRAWING SHAPES (FLOATING)
		elif elem_type == "drawingshape":
			shapes = doc.Shapes
			count = shapes.Count
			batch_size = 250
			end_idx = min(start_offset + batch_size - 1, count)
			has_more = (end_idx < count)
			next_offset = (end_idx + 1) if has_more else None
			scanned_count = end_idx

			for i in range(start_offset, end_idx + 1):
				if i % 5 == 0:
					yield True
				try:
					sh = shapes(i)
					sh_start = sh.Anchor.Start
					sh_end = sh.Anchor.End
					if not is_range_in_scope(sh_start, sh_end, s_start, s_end):
						continue

					sh_name = sh.Name
					title = getattr(sh, "Title", "") or ""
					snippet = f"{sh_name} ({title})" if title else sh_name

					item = WordElementItem(
						raw_obj=sh,
						elem_type="drawingshape",
						start_pos=sh_start,
						end_pos=sh_end,
						snippet=snippet,
						ref_id=sh_name,
						full_text=getattr(sh, "AlternativeText", "") or snippet
					)
					results.append(item)
				except Exception as e:
					log.debugWarning(f"BOA ElementsOrganizer drawingshape error at {i}: {e}")

		# 13. INLINE SHAPES (TEXT LAYER)
		elif elem_type == "inlineshape":
			ishapes = doc.InlineShapes
			count = ishapes.Count
			batch_size = 250
			end_idx = min(start_offset + batch_size - 1, count)
			has_more = (end_idx < count)
			next_offset = (end_idx + 1) if has_more else None
			scanned_count = end_idx

			for i in range(start_offset, end_idx + 1):
				if i % 5 == 0:
					yield True
				try:
					ish = ishapes(i)
					ish_start = ish.Range.Start
					ish_end = ish.Range.End
					if not is_range_in_scope(ish_start, ish_end, s_start, s_end):
						continue

					title = getattr(ish, "Title", "") or ""
					snippet = f"InlineShape {i} ({title})" if title else f"InlineShape {i}"

					item = WordElementItem(
						raw_obj=ish,
						elem_type="inlineshape",
						start_pos=ish_start,
						end_pos=ish_end,
						snippet=snippet,
						ref_id=f"IShape {i}",
						full_text=getattr(ish, "AlternativeText", "") or snippet
					)
					results.append(item)
				except Exception as e:
					log.debugWarning(f"BOA ElementsOrganizer inlineshape error at {i}: {e}")

		# 14. FRAMES
		elif elem_type == "frame":
			frames = doc.Frames
			count = frames.Count
			batch_size = 250
			end_idx = min(start_offset + batch_size - 1, count)
			has_more = (end_idx < count)
			next_offset = (end_idx + 1) if has_more else None
			scanned_count = end_idx

			for i in range(start_offset, end_idx + 1):
				if i % 5 == 0:
					yield True
				try:
					fr = frames(i)
					fr_start = fr.Range.Start
					fr_end = fr.Range.End
					if not is_range_in_scope(fr_start, fr_end, s_start, s_end):
						continue

					fr_text = fr.Range.Text.strip() if fr.Range and fr.Range.Text else ""
					snippet = f"Frame {i}: {fr_text[:40]}" if fr_text else f"Frame {i}"

					item = WordElementItem(
						raw_obj=fr,
						elem_type="frame",
						start_pos=fr_start,
						end_pos=fr_end,
						snippet=snippet,
						ref_id=f"Frame {i}",
						full_text=fr_text
					)
					results.append(item)
				except Exception as e:
					log.debugWarning(f"BOA ElementsOrganizer frame error at {i}: {e}")

		# 15. SECTIONS
		elif elem_type == "section":
			secs = doc.Sections
			count = secs.Count
			batch_size = 250
			end_idx = min(start_offset + batch_size - 1, count)
			has_more = (end_idx < count)
			next_offset = (end_idx + 1) if has_more else None
			scanned_count = end_idx

			for i in range(start_offset, end_idx + 1):
				if i % 5 == 0:
					yield True
				try:
					sec = secs(i)
					sec_start = sec.Range.Start
					sec_end = sec.Range.End
					if not is_range_in_scope(sec_start, sec_end, s_start, s_end):
						continue

					sec_txt = sec.Range.Text.strip() if sec.Range and sec.Range.Text else ""
					snippet = f"Section {i} (Starts: '{sec_txt[:30]}...')" if sec_txt else f"Section {i}"

					item = WordElementItem(
						raw_obj=sec,
						elem_type="section",
						start_pos=sec_start,
						end_pos=sec_end,
						snippet=snippet,
						ref_id=f"Section {i}",
						full_text=snippet
					)
					results.append(item)
				except Exception as e:
					log.debugWarning(f"BOA ElementsOrganizer section error at {i}: {e}")

		# 16. SPELLING ERRORS
		elif elem_type == "spellingerror":
			errs = doc.SpellingErrors
			count = errs.Count
			batch_size = 100
			end_idx = min(start_offset + batch_size - 1, count)
			has_more = (end_idx < count)
			next_offset = (end_idx + 1) if has_more else None
			scanned_count = end_idx

			for i in range(start_offset, end_idx + 1):
				if i % 5 == 0:
					yield True
				try:
					err_item = errs(i)
					err_start = err_item.Start
					err_end = err_item.End
					if not is_range_in_scope(err_start, err_end, s_start, s_end):
						continue

					err_txt = err_item.Text.strip() if err_item.Text else ""
					snippet = f"'{err_txt}'"

					item = WordElementItem(
						raw_obj=err_item,
						elem_type="spellingerror",
						start_pos=err_start,
						end_pos=err_end,
						snippet=snippet,
						ref_id=err_txt,
						full_text=err_txt
					)
					results.append(item)
				except Exception as e:
					log.debugWarning(f"BOA ElementsOrganizer spellingerror error at {i}: {e}")

		# 17. GRAMMATICAL ERRORS
		elif elem_type == "grammaticalerror":
			errs = doc.GrammaticalErrors
			count = errs.Count
			batch_size = 100
			end_idx = min(start_offset + batch_size - 1, count)
			has_more = (end_idx < count)
			next_offset = (end_idx + 1) if has_more else None
			scanned_count = end_idx

			for i in range(start_offset, end_idx + 1):
				if i % 5 == 0:
					yield True
				try:
					g_item = errs(i)
					g_start = g_item.Start
					g_end = g_item.End
					if not is_range_in_scope(g_start, g_end, s_start, s_end):
						continue

					phrase_txt = g_item.Text.strip() if g_item.Text else ""
					snippet = f"'{phrase_txt[:40]}'"

					item = WordElementItem(
						raw_obj=g_item,
						elem_type="grammaticalerror",
						start_pos=g_start,
						end_pos=g_end,
						snippet=snippet,
						ref_id=phrase_txt[:20],
						full_text=phrase_txt
					)
					results.append(item)
				except Exception as e:
					log.debugWarning(f"BOA ElementsOrganizer grammaticalerror error at {i}: {e}")

	except Exception as err:
		log.error(f"BOA ElementsOrganizer generator error: {err}", exc_info=True)

	return {
		"items": results,
		"has_more": has_more,
		"next_offset": next_offset,
		"scanned_count": scanned_count
	}


class WordElementsOrganizerDialog(wx.Dialog):
	"""
	Interactive multi-element document navigation and review dialog for Word.
	Provides dynamic AsyncCOMTask background scanning across all 17 element types.
	"""
	def __init__(self, parent, doc, selection, app):
		# Translators: Title of the Word Elements Organizer dialog
		super(WordElementsOrganizerDialog, self).__init__(parent, title=_("Word Document Elements Organizer"), size=(740, 530))
		self.doc = doc
		self.selection = selection
		self.app = app
		self.items = []
		self.active_task = None
		self.batch_has_more = False
		self.batch_next_offset = None

		self._init_ui()
		self.CentreOnScreen()
		# Enforce focus on the Element Type combo box for immediate screen reader announcement
		core.callLater(50, self.combo_type.SetFocus)
		core.callLater(100, self._start_scan)

	def _init_ui(self):
		main_sizer = wx.BoxSizer(wx.VERTICAL)

		# Top Control Bar (Combos)
		top_sizer = wx.BoxSizer(wx.HORIZONTAL)

		# Combo 1: Element Type
		# Translators: Label for Element Type combo box
		lbl_type = wx.StaticText(self, label=_("&Element Type:"))
		self.combo_type = wx.Choice(self, choices=[label for _, label in ELEMENT_TYPES])
		self.combo_type.SetSelection(0)
		self.combo_type.Bind(wx.EVT_CHOICE, self.on_filter_changed)

		top_sizer.Add(lbl_type, 0, wx.ALIGN_CENTER_VERTICAL | wx.RIGHT, 5)
		top_sizer.Add(self.combo_type, 1, wx.RIGHT, 15)

		# Combo 2: Scope
		# Translators: Label for Scope combo box
		lbl_scope = wx.StaticText(self, label=_("&Scope:"))
		self.combo_scope = wx.Choice(self, choices=[label for _, label in SCOPE_TYPES])
		# Default scope is Selection / Cursor Area
		self.combo_scope.SetSelection(0)
		self.combo_scope.Bind(wx.EVT_CHOICE, self.on_filter_changed)

		top_sizer.Add(lbl_scope, 0, wx.ALIGN_CENTER_VERTICAL | wx.RIGHT, 5)
		top_sizer.Add(self.combo_scope, 1)

		main_sizer.Add(top_sizer, 0, wx.EXPAND | wx.ALL, 10)

		# Central ListCtrl
		self.list_ctrl = wx.ListCtrl(self, style=wx.LC_REPORT | wx.LC_SINGLE_SEL | wx.BORDER_SUNKEN)
		# Translators: Column header for element index
		self.list_ctrl.InsertColumn(0, _("#"), width=40)
		# Translators: Column header for element text snippet
		self.list_ctrl.InsertColumn(1, _("Text / Snippet"), width=440)
		# Translators: Column header for author or reference metadata
		self.list_ctrl.InsertColumn(2, _("Author / Ref"), width=180)

		self.list_ctrl.Bind(wx.EVT_LIST_ITEM_SELECTED, self.on_item_selected)
		self.list_ctrl.Bind(wx.EVT_LIST_ITEM_DESELECTED, self.on_item_deselected)
		self.list_ctrl.Bind(wx.EVT_LIST_ITEM_ACTIVATED, self.on_go_to)
		self.list_ctrl.Bind(wx.EVT_KEY_DOWN, self.on_list_key_down)

		main_sizer.Add(self.list_ctrl, 1, wx.EXPAND | wx.LEFT | wx.RIGHT, 10)

		# Preview Text Control (assigned Alt+P to avoid hotkey clash with Delete Alt+D)
		# Translators: Label for element preview text box
		lbl_preview = wx.StaticText(self, label=_("Element &Preview:"))
		self.preview_ctrl = wx.TextCtrl(self, style=wx.TE_MULTILINE | wx.TE_READONLY, size=(-1, 85))

		main_sizer.Add(lbl_preview, 0, wx.LEFT | wx.RIGHT | wx.TOP, 10)
		main_sizer.Add(self.preview_ctrl, 0, wx.EXPAND | wx.ALL, 10)

		# Context Action Button Bar
		self.btn_sizer = wx.BoxSizer(wx.HORIZONTAL)

		# Translators: Button to navigate Word focus to the selected element
		self.btn_goto = wx.Button(self, label=_("&Go To"))
		self.btn_goto.Bind(wx.EVT_BUTTON, self.on_go_to)
		self.btn_sizer.Add(self.btn_goto, 0, wx.RIGHT, 5)

		# Translators: Button to load more elements for capped scans
		self.btn_load_more = wx.Button(self, label=_("&Load More"))
		self.btn_load_more.Bind(wx.EVT_BUTTON, self.on_load_more)
		self.btn_load_more.Hide()
		self.btn_sizer.Add(self.btn_load_more, 0, wx.RIGHT, 5)

		# Translators: Button to open a hyperlink directly in the default web browser
		self.btn_open_link = wx.Button(self, label=_("&Open Link"))
		self.btn_open_link.Bind(wx.EVT_BUTTON, self.on_open_link)
		self.btn_open_link.Hide()
		self.btn_sizer.Add(self.btn_open_link, 0, wx.RIGHT, 5)

		# Translators: Button to rename a bookmark
		self.btn_rename = wx.Button(self, label=_("Re&name..."))
		self.btn_rename.Bind(wx.EVT_BUTTON, self.on_rename_bookmark)
		self.btn_rename.Hide()
		self.btn_sizer.Add(self.btn_rename, 0, wx.RIGHT, 5)

		# Translators: Button to apply a spelling suggestion
		self.btn_apply = wx.Button(self, label=_("&Apply..."))
		self.btn_apply.Bind(wx.EVT_BUTTON, self.on_apply_spelling)
		self.btn_apply.Hide()
		self.btn_sizer.Add(self.btn_apply, 0, wx.RIGHT, 5)

		# Translators: Button to review grammatical error suggestions directly in Word
		self.btn_review_grammar = wx.Button(self, label=_("&Review in Word..."))
		self.btn_review_grammar.Bind(wx.EVT_BUTTON, self.on_review_grammar)
		self.btn_review_grammar.Hide()
		self.btn_sizer.Add(self.btn_review_grammar, 0, wx.RIGHT, 5)

		# Translators: Button to reply to a comment
		self.btn_reply = wx.Button(self, label=_("&Reply..."))
		self.btn_reply.Bind(wx.EVT_BUTTON, self.on_reply)
		self.btn_reply.Hide()
		self.btn_sizer.Add(self.btn_reply, 0, wx.RIGHT, 5)

		# Translators: Button to accept a tracked change revision
		self.btn_accept = wx.Button(self, label=_("&Accept"))
		self.btn_accept.Bind(wx.EVT_BUTTON, self.on_accept_revision)
		self.btn_accept.Hide()
		self.btn_sizer.Add(self.btn_accept, 0, wx.RIGHT, 5)

		# Translators: Button to reject a tracked change revision
		self.btn_reject = wx.Button(self, label=_("Re&ject"))
		self.btn_reject.Bind(wx.EVT_BUTTON, self.on_reject_revision)
		self.btn_reject.Hide()
		self.btn_sizer.Add(self.btn_reject, 0, wx.RIGHT, 5)

		# Translators: Button to delete the selected element
		self.btn_delete = wx.Button(self, label=_("&Delete"))
		self.btn_delete.Bind(wx.EVT_BUTTON, self.on_delete)
		self.btn_sizer.Add(self.btn_delete, 0, wx.RIGHT, 5)

		# Translators: Button to copy all listed elements to clipboard
		self.btn_copy_all = wx.Button(self, label=_("&Copy All"))
		self.btn_copy_all.Bind(wx.EVT_BUTTON, self.on_copy_all)
		self.btn_sizer.Add(self.btn_copy_all, 0, wx.RIGHT, 5)

		self.btn_sizer.AddStretchSpacer()

		# Translators: Button to close dialog
		self.btn_close = wx.Button(self, wx.ID_CANCEL, label=_("Close"))
		self.btn_close.Bind(wx.EVT_BUTTON, lambda evt: self.EndModal(wx.ID_CANCEL))
		self.btn_sizer.Add(self.btn_close, 0)

		main_sizer.Add(self.btn_sizer, 0, wx.EXPAND | wx.ALL, 10)
		self.SetSizer(main_sizer)

		self.Bind(wx.EVT_CLOSE, self.on_dialog_close)
		self.update_action_states()

	def update_action_states(self):
		"""
		Synchronizes both visibility and enablement across all action buttons:
		- Irrelevant buttons for the active element type are hidden from UI and keyboard navigation.
		- Action buttons are disabled if the list is empty or no item is currently selected.
		- Layout is recalculated on both the button sizer and the dialog window.
		"""
		elem_idx = self.combo_type.GetSelection()
		elem_type = ELEMENT_TYPES[elem_idx][0] if 0 <= elem_idx < len(ELEMENT_TYPES) else ""
		has_items = len(self.items) > 0
		selected_idx = self.list_ctrl.GetFirstSelected() if hasattr(self, "list_ctrl") else -1
		has_selection = has_items and (selected_idx >= 0)

		# 1. Visibility (Show only buttons relevant to active element type)
		self.btn_sizer.Show(self.btn_reply, elem_type == "comment")
		self.btn_sizer.Show(self.btn_accept, elem_type == "revision")
		self.btn_sizer.Show(self.btn_reject, elem_type == "revision")
		self.btn_sizer.Show(self.btn_open_link, elem_type == "hyperlink")
		self.btn_sizer.Show(self.btn_rename, elem_type == "bookmark")
		self.btn_sizer.Show(self.btn_apply, elem_type == "spellingerror")
		self.btn_sizer.Show(self.btn_review_grammar, elem_type == "grammaticalerror")
		self.btn_sizer.Show(self.btn_load_more, getattr(self, "batch_has_more", False))

		# 2. Enablement (Enable only when an item is available and selected)
		self.btn_goto.Enable(has_selection)
		self.btn_copy_all.Enable(has_items)
		can_delete = has_selection and elem_type in (
			"comment", "revision", "footnote", "endnote", "bookmark",
			"table", "hyperlink", "drawingshape", "inlineshape", "frame"
		)
		self.btn_delete.Enable(can_delete)
		self.btn_reply.Enable(has_selection)
		self.btn_accept.Enable(has_selection)
		self.btn_reject.Enable(has_selection)
		self.btn_open_link.Enable(has_selection)
		self.btn_rename.Enable(has_selection)
		self.btn_apply.Enable(has_selection and elem_type == "spellingerror")
		self.btn_review_grammar.Enable(has_selection and elem_type == "grammaticalerror")

		self.btn_sizer.Layout()
		self.Layout()

	def on_filter_changed(self, evt):
		self._start_scan()

	def _start_scan(self, preserve_selection=False):
		elem_idx = self.combo_type.GetSelection()
		scope_idx = self.combo_scope.GetSelection()

		elem_type = ELEMENT_TYPES[elem_idx][0]
		scope_key = SCOPE_TYPES[scope_idx][0]

		if preserve_selection and hasattr(self, "list_ctrl"):
			self._restore_index = self.list_ctrl.GetFirstSelected()
		else:
			self._restore_index = 0

		self.batch_has_more = False
		self.batch_next_offset = None

		self.list_ctrl.DeleteAllItems()
		self.preview_ctrl.Clear()
		self.items = []
		self.update_action_states()

		if self.active_task:
			self.active_task.is_running = False

		gen = extract_elements_generator(self.doc, self.selection, elem_type, scope_key, start_offset=1)
		self.active_task = AsyncCOMTask(gen, on_complete=self._on_scan_complete, max_tick_time=0.05)
		self.active_task.start()

	def _on_scan_complete(self, result):
		if not self or not self.list_ctrl:
			return

		if isinstance(result, dict):
			self.items = result.get("items", [])
			self.batch_has_more = result.get("has_more", False)
			self.batch_next_offset = result.get("next_offset", None)
			scanned_count = result.get("scanned_count", 0)
		else:
			self.items = result or []
			self.batch_has_more = False
			self.batch_next_offset = None
			scanned_count = len(self.items)

		self.list_ctrl.DeleteAllItems()

		for idx, item in enumerate(self.items):
			row = self.list_ctrl.InsertItem(idx, str(idx + 1))
			self.list_ctrl.SetItem(row, 1, item.snippet)
			self.list_ctrl.SetItem(row, 2, item.get_meta_display())

		if self.items:
			target_idx = getattr(self, "_restore_index", 0)
			if target_idx < 0 or target_idx >= len(self.items):
				target_idx = 0
			# Automatically select and focus the target element so screen reader announces it upon tabbing
			self.list_ctrl.SetItemState(target_idx, wx.LIST_STATE_SELECTED | wx.LIST_STATE_FOCUSED, wx.LIST_STATE_SELECTED | wx.LIST_STATE_FOCUSED)
			self._display_item_preview(self.items[target_idx])
			self._restore_index = 0
		else:
			# Insert a clean placeholder row so screen reader announces status instead of empty silence
			row = self.list_ctrl.InsertItem(0, "-")
			# Translators: Placeholder text displayed in list when zero elements are found
			self.list_ctrl.SetItem(row, 1, _("No elements found"))
			self.list_ctrl.SetItem(row, 2, "")
			self.list_ctrl.SetItemState(0, wx.LIST_STATE_SELECTED | wx.LIST_STATE_FOCUSED, wx.LIST_STATE_SELECTED | wx.LIST_STATE_FOCUSED)
			# Translators: Announcement when zero elements are discovered
			self.preview_ctrl.SetValue(_("No items found in selected scope."))

		self.update_action_states()

		elem_idx = self.combo_type.GetSelection()
		scope_idx = self.combo_scope.GetSelection()
		elem_label = ELEMENT_TYPES[elem_idx][1]
		scope_label = SCOPE_TYPES[scope_idx][1]
		elem_type = ELEMENT_TYPES[elem_idx][0]

		if self.batch_has_more:
			if elem_type == "heading":
				# Translators: Announced when headings batch limit is reached and more paragraphs remain
				ui.message(_("Found {count} headings in first {p} paragraphs. More paragraphs remain. Press Alt+L to load more.").format(count=len(self.items), p=scanned_count))
			else:
				# Translators: Announced when element scan batch limit is reached and more items remain
				ui.message(_("Found {count} {elem_label}. Limit reached. Press Alt+L to load more.").format(count=len(self.items), elem_label=elem_label))
		else:
			if self.items:
				# Translators: Announced when elements scan completes with items found
				ui.message(_("Found {count} {elem_label} in {scope_label}.").format(count=len(self.items), elem_label=elem_label, scope_label=scope_label))
			else:
				# Translators: Announced when zero elements are found in selected scope
				ui.message(_("No {elem_label} found in {scope_label}.").format(elem_label=elem_label, scope_label=scope_label))

	def on_load_more(self, evt=None):
		if not getattr(self, "batch_has_more", False) or not getattr(self, "batch_next_offset", None):
			return

		elem_idx = self.combo_type.GetSelection()
		scope_idx = self.combo_scope.GetSelection()
		elem_type = ELEMENT_TYPES[elem_idx][0]
		scope_key = SCOPE_TYPES[scope_idx][0]

		if self.active_task:
			self.active_task.is_running = False

		# Translators: Spoken announcement when loading more elements starts
		ui.message(_("Loading more..."))
		gen = extract_elements_generator(self.doc, self.selection, elem_type, scope_key, start_offset=self.batch_next_offset)
		self.active_task = AsyncCOMTask(gen, on_complete=self._on_load_more_complete, max_tick_time=0.05)
		self.active_task.start()

	def _on_load_more_complete(self, result):
		if not self or not self.list_ctrl:
			return
		if not isinstance(result, dict):
			return

		new_items = result.get("items", [])
		self.batch_has_more = result.get("has_more", False)
		self.batch_next_offset = result.get("next_offset", None)

		start_row = len(self.items)
		self.items.extend(new_items)

		for idx, item in enumerate(new_items):
			row_idx = start_row + idx
			row = self.list_ctrl.InsertItem(row_idx, str(row_idx + 1))
			self.list_ctrl.SetItem(row, 1, item.snippet)
			self.list_ctrl.SetItem(row, 2, item.get_meta_display())

		self.update_action_states()

		elem_idx = self.combo_type.GetSelection()
		elem_label = ELEMENT_TYPES[elem_idx][1]

		if self.batch_has_more:
			# Translators: Announced when more elements were loaded and additional items still remain
			ui.message(_("Loaded {new_n} more {elem_label}. Total {total_n} loaded. Press Alt+L to load more.").format(new_n=len(new_items), elem_label=elem_label, total_n=len(self.items)))
		else:
			# Translators: Announced when more elements were loaded and end of document was reached
			ui.message(_("Loaded {new_n} more {elem_label}. Total {total_n} loaded. End reached.").format(new_n=len(new_items), elem_label=elem_label, total_n=len(self.items)))

	def _get_selected_item(self):
		idx = self.list_ctrl.GetFirstSelected()
		if 0 <= idx < len(self.items):
			return self.items[idx]
		return None

	def on_item_selected(self, evt):
		idx = evt.GetIndex()
		if 0 <= idx < len(self.items):
			self._display_item_preview(self.items[idx])
		self.update_action_states()

	def on_item_deselected(self, evt):
		if self.list_ctrl.GetFirstSelected() < 0:
			self.preview_ctrl.Clear()
		self.update_action_states()

	def _display_item_preview(self, item):
		txt_parts = []
		# Calculate page and line lazily for this focused item only
		page, line = item.get_page_line(self.doc)
		if page or line:
			# Translators: Location header in element preview
			txt_parts.append(_("[Page {p}, Line {l}]").format(p=page or "?", l=line or "?"))

		if item.elem_type == "comment":
			# Translators: Comment details header
			txt_parts.append(_("Author: {a} | Date: {d}").format(a=item.author or _("Unknown"), d=item.date_str))
			if item.extra.get("scope_text"):
				# Translators: Commented text excerpt
				txt_parts.append(_("Commented Text: \"{t}\"").format(t=item.extra['scope_text'].strip()))
			txt_parts.append(f"\n{item.full_text}")

			replies = item.extra.get("replies", [])
			if replies:
				# Translators: Replies section header
				txt_parts.append(_("\n--- Replies ({n}) ---").format(n=len(replies)))
				for r in replies:
					txt_parts.append(f"• {r['author']} ({r['date']}):\n  {r['text']}")

		elif item.elem_type == "revision":
			# Translators: Summary header for revision element preview
			txt_parts.append(_("Type: {t} | Author: {a} | Date: {d}").format(t=item.extra.get('type_name', ''), a=item.author, d=item.date_str))
			txt_parts.append(f"\n{item.full_text}")

		elif item.elem_type == "table":
			rows_c = item.extra.get("rows", "?")
			cols_c = item.extra.get("cols", "?")
			has_hf = item.extra.get("heading_format", False)
			tbl_title = item.extra.get("title", "")
			tbl_descr = item.extra.get("descr", "")
			# Translators: Table dimensions summary in preview
			txt_parts.append(_("Dimensions: {r} rows x {c} columns").format(r=rows_c, c=cols_c))
			if has_hf:
				# Translators: Indicates table row 1 repeats as a header
				txt_parts.append(_("Header Row: Present (Row 1 repeats across pages)"))
			else:
				# Translators: Indicates table has no repeated header row
				txt_parts.append(_("Header Row: None (Repeat header row not configured)"))
			if tbl_title:
				# Translators: Alternative text title of table
				txt_parts.append(_("Title: {t}").format(t=tbl_title))
			if tbl_descr:
				# Translators: Alternative text description of table
				txt_parts.append(_("Description: {d}").format(d=tbl_descr))
			txt_parts.append(f"\n{item.full_text}")

		elif item.elem_type == "hyperlink":
			# Translators: Label for hyperlink display text in preview
			txt_parts.append(_("Text: {t}").format(t=item.ref_id))
			# Translators: Label for hyperlink destination URL in preview
			txt_parts.append(_("Target URL: {u}").format(u=item.full_text))

		elif item.elem_type == "bookmark":
			# Translators: Label for bookmark name in preview
			txt_parts.append(_("Bookmark Name: {n}").format(n=item.ref_id))
			if item.full_text:
				txt_parts.append(item.full_text)

		elif item.elem_type == "spellingerror":
			# Translators: Label for spelling error text in preview
			txt_parts.append(_("Spelling Error: {word}").format(word=item.full_text))
			suggs = item.extra.get("suggestions")
			if suggs is None:
				suggs = []
				try:
					s_col = item.raw_obj.GetSpellingSuggestions()
					for s_i in range(1, min(s_col.Count, 5) + 1):
						suggs.append(s_col(s_i).Name)
				except Exception:
					pass
				item.extra["suggestions"] = suggs
			if suggs:
				# Translators: Label for spelling suggestions in preview
				txt_parts.append(_("Suggestions: {s}").format(s=", ".join(suggs)))
			else:
				# Translators: Spoken when no spelling suggestions are found
				txt_parts.append(_("Suggestions: None found"))

		elif item.elem_type == "grammaticalerror":
			# Translators: Header for grammatical error preview
			txt_parts.append(_("Type: Grammatical / Style Error"))
			# Translators: Label for sentence in grammatical error preview
			txt_parts.append(_('Sentence: "{text}"').format(text=item.full_text))
			# Translators: Explanatory note in preview for how to review grammar suggestions
			txt_parts.append(_("Suggestions: Managed by Microsoft Editor.\nPress Alt+R or click 'Review in Word' to open suggestions directly in Word."))

		else:
			if item.ref_id:
				txt_parts.append(f"Reference: {item.ref_id}")
			txt_parts.append(item.full_text)

		self.preview_ctrl.SetValue("\n".join(txt_parts))

	def on_list_key_down(self, evt):
		key = evt.GetKeyCode()
		if key == wx.WXK_RETURN:
			self.on_go_to(None)
			return
		elif key == ord('O') and wx.GetKeyState(wx.WXK_CONTROL):
			self.on_open_link(None)
			return
		elif key == wx.WXK_SPACE:
			item = self._get_selected_item()
			if item:
				if wx.GetKeyState(wx.WXK_CONTROL) and item.elem_type == "comment":
					replies = item.extra.get("replies", [])
					if replies:
						latest_rep = replies[-1]
						# Translators: Spoken announcement of latest comment reply
						ui.message(_("Latest reply from {author}: {text}").format(author=latest_rep['author'], text=latest_rep['text']))
					else:
						# Translators: Spoken announcement when Ctrl+Space is pressed on a comment with no replies
						ui.message(_("No replies to this comment."))
				else:
					ui.message(item.full_text or item.snippet)
			else:
				# Translators: Spoken announcement when spacebar is pressed on the placeholder row in an empty elements list
				ui.message(_("No elements found"))
			return
		elif key in (wx.WXK_DELETE, wx.WXK_NUMPAD_DELETE):
			self.on_delete(None)
			return
		evt.Skip()

	def on_go_to(self, evt=None):
		item = self._get_selected_item()
		if not item:
			return
		self._nav_target = {
			"start_pos": item.start_pos,
			"end_pos": item.end_pos,
			"elem_type": item.elem_type,
			"snippet": item.snippet
		}
		self._trigger_grammar_card = False
		self.EndModal(wx.ID_OK)

	def on_open_link(self, evt=None):
		item = self._get_selected_item()
		if not item or item.elem_type != "hyperlink":
			return
		addr = item.extra.get("address") or item.full_text
		if not addr:
			# Translators: Spoken when a hyperlink has no target address
			ui.message(_("No target address for this hyperlink."))
			return
		try:
			os.startfile(addr)
			# Translators: Spoken when a hyperlink is opened in browser
			ui.message(_("Opening link in browser: {url}").format(url=addr))
		except Exception as err:
			log.error(f"BOA ElementsOrganizer open link error: {err}")
			# Translators: Spoken when opening a hyperlink fails
			ui.message(_("Failed to open link: {err}").format(err=str(err)))

	def on_rename_bookmark(self, evt=None):
		item = self._get_selected_item()
		if not item or item.elem_type != "bookmark":
			return
		old_name = item.ref_id or item.snippet
		# Translators: Prompt for renaming a bookmark
		prompt_msg = _("Enter new name for bookmark '{old}':").format(old=old_name)
		# Translators: Title of rename bookmark dialog
		dlg_title = _("Rename Bookmark")
		dlg = wx.TextEntryDialog(self, prompt_msg, dlg_title, value=old_name)
		if dlg.ShowModal() == wx.ID_OK:
			new_name = dlg.GetValue().strip()
			if not new_name or new_name == old_name:
				dlg.Destroy()
				return
			if not re.match(r'^[A-Za-z][A-Za-z0-9_]*$', new_name):
				# Translators: Error message for invalid Word bookmark name
				gui.messageBox(_("Invalid bookmark name. It must start with a letter and contain only letters, numbers, and underscores (no spaces)."), _("Error"), wx.OK | wx.ICON_ERROR, self)
				dlg.Destroy()
				return
			try:
				bm = item.raw_obj
				bm_range = bm.Range
				self.doc.Bookmarks.Add(new_name, bm_range)
				bm.Delete()
				item.ref_id = new_name
				item.snippet = new_name
				item.full_text = new_name
				idx = self.list_ctrl.GetFirstSelected()
				if idx >= 0:
					self.list_ctrl.SetItem(idx, 1, new_name)
					self.list_ctrl.SetItem(idx, 2, new_name)
				self._display_item_preview(item)
				# Translators: Spoken when a bookmark is successfully renamed
				ui.message(_("Bookmark renamed to {name}").format(name=new_name))
			except Exception as err:
				log.error(f"BOA ElementsOrganizer rename bookmark error: {err}")
				# Translators: Error message when renaming a bookmark fails
				ui.message(_("Failed to rename bookmark: {err}").format(err=str(err)))
		dlg.Destroy()

	def on_review_grammar(self, evt=None):
		"""
		Navigates Word focus directly to the grammatical error sentence and signals
		the launcher to trigger Word's native Microsoft Editor suggestion card.
		"""
		item = self._get_selected_item()
		if not item or item.elem_type != "grammaticalerror":
			return
		self._nav_target = {
			"start_pos": item.start_pos,
			"end_pos": item.end_pos,
			"elem_type": item.elem_type,
			"snippet": item.snippet
		}
		self._trigger_grammar_card = True
		self.EndModal(wx.ID_OK)

	def on_apply_spelling(self, evt=None):
		item = self._get_selected_item()
		if not item or item.elem_type != "spellingerror":
			return
		suggs = item.extra.get("suggestions", [])
		if not suggs:
			# Translators: Spoken when no suggestions are available to apply
			ui.message(_("No suggestions available for this word."))
			return
		# Translators: Dialog prompt to select a spelling replacement
		prompt_msg = _("Select replacement for '{word}':").format(word=item.full_text)
		# Translators: Dialog title for applying a spelling replacement
		dlg_title = _("Apply Correction")
		dlg = wx.SingleChoiceDialog(self, prompt_msg, dlg_title, suggs)
		if dlg.ShowModal() == wx.ID_OK:
			chosen = dlg.GetStringSelection()
			if chosen:
				try:
					try:
						item.raw_obj.Text = chosen
					except Exception:
						if self.doc and item.start_pos >= 0:
							self.doc.Range(item.start_pos, item.end_pos).Text = chosen
					item.full_text = chosen
					item.snippet = f"'{chosen}' (Corrected)"
					idx = self.list_ctrl.GetFirstSelected()
					if idx >= 0:
						self.list_ctrl.SetItem(idx, 1, item.snippet)
					self._display_item_preview(item)
					# Translators: Spoken when spelling replacement is applied
					ui.message(_("Replaced with {word}").format(word=chosen))
				except Exception as err:
					log.error(f"BOA ElementsOrganizer apply spelling error: {err}")
					# Translators: Error message when replacing spelling fails
					ui.message(_("Failed to apply correction: {err}").format(err=str(err)))
		dlg.Destroy()

	def on_reply(self, evt):
		item = self._get_selected_item()
		if not item or item.elem_type != "comment":
			return

		# Translators: Dialog title for comment reply
		dlg_title = _("Reply to Comment")
		# Translators: Message prompt for comment reply text
		prompt = _("Enter reply text:")
		reply_dlg = wx.TextEntryDialog(self, prompt, dlg_title, style=wx.OK | wx.CANCEL | wx.TE_MULTILINE)
		if reply_dlg.ShowModal() == wx.ID_OK:
			rep_text = reply_dlg.GetValue().strip()
			if rep_text:
				try:
					c = item.raw_obj
					c.Replies.Add(c.Scope, rep_text)
					# Translators: Success announcement
					ui.message(_("Reply added."))
					self._start_scan(preserve_selection=True)
				except Exception as e:
					log.error(f"BOA ElementsOrganizer Reply Error: {e}")
					# Translators: Error adding reply
					ui.message(_("Failed to add reply."))
		reply_dlg.Destroy()

	def on_accept_revision(self, evt):
		item = self._get_selected_item()
		if item and item.elem_type == "revision":
			try:
				item.raw_obj.Accept()
				# Translators: Success announcement
				ui.message(_("Revision accepted."))
				self._start_scan()
			except Exception as e:
				log.error(f"BOA ElementsOrganizer Accept Error: {e}")

	def on_reject_revision(self, evt):
		item = self._get_selected_item()
		if item and item.elem_type == "revision":
			try:
				item.raw_obj.Reject()
				# Translators: Success announcement
				ui.message(_("Revision rejected."))
				self._start_scan()
			except Exception as e:
				log.error(f"BOA ElementsOrganizer Reject Error: {e}")

	def on_delete(self, evt):
		item = self._get_selected_item()
		if item:
			# Translators: Confirmation prompt before deleting an element
			msg = _("Are you sure you want to delete this {type}?").format(type=item.elem_type)
			# Translators: Confirmation title
			title = _("Confirm Deletion")
			if wx.MessageBox(msg, title, wx.YES_NO | wx.ICON_QUESTION, self) == wx.YES:
				try:
					item.raw_obj.Delete()
					# Translators: Success announcement
					ui.message(_("Item deleted."))
					self._start_scan()
				except Exception as e:
					log.error(f"BOA ElementsOrganizer Delete Error: {e}")
					# Translators: Error deleting item
					ui.message(_("Failed to delete element."))

	def on_copy_all(self, evt):
		if not self.items:
			return
		lines = []
		elem_label = self.combo_type.GetStringSelection()
		lines.append(f"=== {elem_label} ({len(self.items)} items) ===")
		for idx, itm in enumerate(self.items, start=1):
			page, line = itm.get_page_line(self.doc)
			meta = itm.author or itm.ref_id
			lines.append(f"{idx}. [Page {page or '?'}, Line {line or '?'}] {itm.snippet} | {meta}")
		summary_txt = "\n".join(lines)
		try:
			import api
			api.copyToClip(summary_txt)
			# Translators: Success announcement when items copied to clipboard
			ui.message(_("Copied {n} items to clipboard.").format(n=len(self.items)))
		except Exception as e:
			log.error(f"BOA ElementsOrganizer Copy Error: {e}")

	def clean_up(self):
		"""
		Safely halts background scans and releases all COM interface pointers.
		Idempotent: safe to invoke multiple times without error.
		"""
		if getattr(self, "_cleaned_up", False):
			return
		self._cleaned_up = True

		if self.active_task:
			self.active_task.is_running = False
			self.active_task.on_complete = None
			self.active_task = None

		if hasattr(self, "items") and self.items:
			for item in self.items:
				item.raw_obj = None
				if hasattr(item, "extra") and isinstance(item.extra, dict):
					item.extra.clear()
			self.items.clear()

		self.doc = None
		self.selection = None
		self.app = None

	def on_dialog_close(self, evt):
		self.clean_up()
		self.EndModal(wx.ID_CANCEL)


class WordElementsOrganizer:
	"""
	Launcher for the Word Document Elements Organizer.
	"""
	@staticmethod
	def _launch_dialog(doc, selection, app):
		gui.mainFrame.prePopup()
		dlg = None
		nav_target = None
		trigger_card = False
		try:
			dlg = WordElementsOrganizerDialog(gui.mainFrame, doc, selection, app)
			result = dlg.ShowModal()
			if result == wx.ID_OK:
				nav_target = getattr(dlg, "_nav_target", None)
				trigger_card = getattr(dlg, "_trigger_grammar_card", False)
		finally:
			if dlg:
				dlg.clean_up()
				dlg.Destroy()
			gui.mainFrame.postPopup()

		if nav_target and doc:
			def _perform_navigation():
				try:
					# 1. Activate Word document window
					try:
						if hasattr(doc, "ActiveWindow") and doc.ActiveWindow:
							doc.ActiveWindow.Activate()
					except Exception as win_err:
						log.debugWarning(f"BOA could not activate Word ActiveWindow: {win_err}")

					# 2. Select target range
					start_pos = nav_target.get("start_pos", 0)
					end_pos = nav_target.get("end_pos", start_pos)
					if end_pos < start_pos:
						end_pos = start_pos

					# For massive elements like entire sections (> 2000 chars),
					# collapse to start_pos to avoid massive text selection freezes.
					elem_type = nav_target.get("elem_type", "")
					if elem_type == "section" or (end_pos - start_pos > 2000):
						target_range = doc.Range(start_pos, start_pos)
					else:
						target_range = doc.Range(start_pos, end_pos)

					if target_range:
						target_range.Select()

						if trigger_card:
							# Translators: Spoken when opening Microsoft Editor suggestions
							ui.message(_("Opening Microsoft Editor suggestions..."))
							def _send_card_gesture():
								try:
									KeyboardInputGesture.fromName("alt+downArrow").send()
								except Exception as err:
									log.debugWarning(f"BOA failed to send alt+downArrow: {err}")
							core.callLater(200, _send_card_gesture)
						elif elem_type == "grammaticalerror":
							# Translators: Spoken after jumping to a grammatical error in Word
							ui.message(_("Navigated to grammatical error. Press Alt+DownArrow or Shift+F10 to review suggestions."))
					else:
						# Translators: Error navigating to element
						ui.message(_("Failed to jump to element."))
				except Exception as e:
					log.error(f"BOA ElementsOrganizer navigation error: {e}", exc_info=True)
					# Translators: Error navigating to element
					ui.message(_("Failed to jump to element."))

			# Defer slightly (60ms) to ensure Word window has fully regained OS focus
			core.callLater(60, _perform_navigation)

	@staticmethod
	def show(obj):
		try:
			from appModules.boa_enhancements.word_enhancements.document_analyzer import WordDocumentAnalyzer
			doc, selection, app = WordDocumentAnalyzer._get_com_document_and_selection(obj)
			if not doc:
				# Translators: Error message when Word COM document is unavailable
				ui.message(_("Cannot access Word document."))
				return

			# Dispatch dialog safely via wx.CallAfter to guarantee the message pump never deadlocks
			wx.CallAfter(WordElementsOrganizer._launch_dialog, doc, selection, app)

		except Exception as e:
			log.error(f"BOA WordElementsOrganizer Launcher Error: {e}", exc_info=True)
			# Translators: Error message launching dialog
			ui.message(_("Failed to open Elements Organizer."))
