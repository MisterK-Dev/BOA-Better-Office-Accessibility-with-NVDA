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
from appModules.boa_enhancements.async_engine import AsyncCOMTask

addonHandler.initTranslation()

# Word COM Constants
WD_FIRST_CHAR_LINE = 10  # wdFirstCharacterLineNumber
WD_ACTIVE_END_PAGE = 3   # wdActiveEndPageNumber
WD_ACTIVE_END_ADJ_PAGE = 1 # wdActiveEndAdjustedPageNumber

# Math Equation Types mapping
WD_OMATH_TYPES = {
	# Translators: Math equation display type
	0: _("Inline"),
	# Translators: Math equation display type
	1: _("Display (Centered)"),
}

# Word Field Types
WD_FIELD_CITATION = 96
WD_FIELD_BIBLIOGRAPHY = 56
WD_FIELD_SEQUENCE = 12

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

# Office MsoShapeType mapping
MSO_SHAPE_TYPES = {
	# Translators: Office shape type
	1: _("AutoShape"),
	# Translators: Office shape type
	2: _("Callout"),
	# Translators: Office shape type
	3: _("Chart"),
	# Translators: Office shape type
	4: _("Comment"),
	# Translators: Office shape type
	5: _("Freeform"),
	# Translators: Office shape type
	6: _("Group"),
	# Translators: Office shape type
	7: _("Embedded OLE Object"),
	# Translators: Office shape type
	8: _("Form Control"),
	# Translators: Office shape type
	9: _("Line"),
	# Translators: Office shape type
	10: _("Linked OLE Object"),
	# Translators: Office shape type
	11: _("Linked Picture"),
	# Translators: Office shape type
	12: _("OLE Control Object"),
	# Translators: Office shape type
	13: _("Picture"),
	# Translators: Office shape type
	14: _("Placeholder"),
	# Translators: Office shape type
	15: _("Text Effect"),
	# Translators: Office shape type
	16: _("Media"),
	# Translators: Office shape type
	17: _("Text Box"),
	# Translators: Office shape type
	18: _("Script Anchor"),
	# Translators: Office shape type
	19: _("Table"),
	# Translators: Office shape type
	20: _("Canvas"),
	# Translators: Office shape type
	21: _("Diagram"),
	# Translators: Office shape type
	22: _("Ink"),
	# Translators: Office shape type
	23: _("Ink Comment"),
	# Translators: Office shape type
	24: _("SmartArt"),
	# Translators: Office shape type
	25: _("Slicer"),
	# Translators: Office shape type
	26: _("Web Video"),
}

# Word WdInlineShapeType mapping
WD_INLINE_SHAPE_TYPES = {
	# Translators: Word inline shape type
	1: _("Embedded OLE Object"),
	# Translators: Word inline shape type
	2: _("Linked OLE Object"),
	# Translators: Word inline shape type
	3: _("Picture"),
	# Translators: Word inline shape type
	4: _("Linked Picture"),
	# Translators: Word inline shape type
	5: _("OLE Control Object"),
	# Translators: Word inline shape type
	6: _("Horizontal Line"),
	# Translators: Word inline shape type
	7: _("Picture with Horizontal Line"),
	# Translators: Word inline shape type
	8: _("Linked Picture with Horizontal Line"),
	# Translators: Word inline shape type
	9: _("Picture Bullet"),
	# Translators: Word inline shape type
	10: _("Script Anchor"),
	# Translators: Word inline shape type
	11: _("OWS Anchor"),
	# Translators: Word inline shape type
	12: _("Chart"),
	# Translators: Word inline shape type
	13: _("Diagram"),
	# Translators: Word inline shape type
	14: _("Locked Canvas"),
	# Translators: Word inline shape type
	15: _("SmartArt"),
	# Translators: Word inline shape type
	16: _("Web Video"),
}

# Word Content Control Types mapping
WD_CONTENT_CONTROL_TYPES = {
	# Translators: Word content control type
	0: _("Rich Text"),
	# Translators: Word content control type
	1: _("Plain Text"),
	# Translators: Word content control type
	2: _("Picture"),
	# Translators: Word content control type
	3: _("Combo Box"),
	# Translators: Word content control type
	4: _("Drop-Down List"),
	# Translators: Word content control type
	5: _("Building Block Gallery"),
	# Translators: Word content control type
	6: _("Date Picker"),
	# Translators: Word content control type
	7: _("Group"),
	# Translators: Word content control type
	8: _("Checkbox"),
	# Translators: Word content control type
	9: _("Repeating Section"),
}

# Word Story Types mapping
WD_STORY_TYPES = {
	# Translators: Word story type
	1: _("Main Text"),
	# Translators: Word story type
	2: _("Footnotes"),
	# Translators: Word story type
	3: _("Endnotes"),
	# Translators: Word story type
	4: _("Comments"),
	# Translators: Word story type
	5: _("Text Frame"),
	# Translators: Word story type
	6: _("Even Pages Header"),
	# Translators: Word story type
	7: _("Primary Header"),
	# Translators: Word story type
	8: _("Even Pages Footer"),
	# Translators: Word story type
	9: _("Primary Footer"),
	# Translators: Word story type
	10: _("First Page Header"),
	# Translators: Word story type
	11: _("First Page Footer"),
	# Translators: Word story type
	12: _("Footnote Separator"),
	# Translators: Word story type
	13: _("Footnote Continuation Separator"),
	# Translators: Word story type
	14: _("Footnote Continuation Notice"),
	# Translators: Word story type
	15: _("Endnote Separator"),
	# Translators: Word story type
	16: _("Endnote Continuation Separator"),
	# Translators: Word story type
	17: _("Endnote Continuation Notice"),
}

# Office Hyperlink Types mapping
MSO_HYPERLINK_TYPES = {
	# Translators: Hyperlink type
	0: _("Text Range"),
	# Translators: Hyperlink type
	1: _("Shape"),
	# Translators: Hyperlink type
	2: _("Inline Shape"),
}

# Word Page Orientation mapping
WD_PAGE_ORIENTATIONS = {
	# Translators: Page orientation
	0: _("Portrait"),
	# Translators: Page orientation
	1: _("Landscape"),
}

# Word Section Break Start mapping
WD_SECTION_STARTS = {
	# Translators: Section break type
	0: _("Continuous"),
	# Translators: Section break type
	1: _("New Column"),
	# Translators: Section break type
	2: _("New Page"),
	# Translators: Section break type
	3: _("Even Page"),
	# Translators: Section break type
	4: _("Odd Page"),
}

# Word Paper Size mapping
WD_PAPER_SIZES = {
	0: "10x14",
	1: "11x17",
	2: "Letter",
	3: "Letter Small",
	4: "Legal",
	5: "Executive",
	6: "A3",
	7: "A4",
	8: "A4 Small",
	9: "A5",
	10: "B4",
	11: "B5",
	12: "C Sheet",
	13: "D Sheet",
	14: "E Sheet",
	18: "Folio",
	19: "Ledger",
	20: "Note",
	21: "Quarto",
	22: "Statement",
	23: "Tabloid",
	# Translators: Custom paper size
	41: _("Custom"),
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
	# Translators: Document element type in Elements Organizer
	("equation", _("Math Equations")),
	# Translators: Document element type in Elements Organizer
	("citation", _("Citations & References")),
	# Translators: Document element type in Elements Organizer
	("caption", _("Captions (Figures & Tables)")),
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
		elif self.elem_type in ("drawingshape", "inlineshape"):
			return self.extra.get("type_name", self.ref_id)
		elif self.elem_type == "contentcontrol":
			return self.extra.get("type_name", self.ref_id)
		elif self.elem_type == "heading":
			style_name = self.extra.get("style_name")
			if style_name:
				return f"{self.ref_id} ({style_name})"
			return self.ref_id
		elif self.elem_type == "bookmark":
			story_name = self.extra.get("story_name")
			# Translators: Check if bookmark is outside main text story
			if story_name and story_name != _("Main Text"):
				return f"{self.ref_id} ({story_name})"
			return self.ref_id
		elif self.elem_type == "section":
			start_type = self.extra.get("start_type")
			if start_type:
				return start_type
			return self.ref_id
		elif self.elem_type == "hyperlink":
			type_name = self.extra.get("type_name")
			if type_name:
				return type_name
			return self.author or self.ref_id
		elif self.elem_type == "equation":
			return self.extra.get("type_name", self.ref_id)
		elif self.elem_type == "citation":
			return self.extra.get("tag", self.ref_id)
		elif self.elem_type == "caption":
			return self.extra.get("label", self.ref_id)
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

					is_empty = False
					try:
						is_empty = bool(getattr(bm, "Empty", False))
					except Exception:
						pass

					story_int = 1
					story_name = ""
					try:
						story_int = int(getattr(bm, "StoryType", 1))
						# Translators: Fallback story type
						story_name = WD_STORY_TYPES.get(story_int, _("Main Text"))
					except Exception:
						pass

					item = WordElementItem(
						raw_obj=bm,
						elem_type="bookmark",
						start_pos=bm_start,
						end_pos=bm_end,
						snippet=snippet,
						ref_id=bm_name,
						full_text=text_val,
						extra={
							"is_empty": is_empty,
							"story_type_int": story_int,
							"story_name": story_name,
						}
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

						style_name = ""
						try:
							if hasattr(p, "Style") and p.Style:
								style_name = getattr(p.Style, "NameLocal", "") or ""
						except Exception:
							pass

						item = WordElementItem(
							raw_obj=p,
							elem_type="heading",
							start_pos=p_start,
							end_pos=p_end,
							snippet=snippet,
							ref_id=f"H{level}",
							full_text=h_text,
							extra={"level": level, "style_name": style_name}
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

					is_uniform = True
					try:
						is_uniform = bool(tbl.Uniform)
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
							"descr": descr_val,
							"is_uniform": is_uniform,
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

					hl_type_int = int(getattr(hl, "Type", 0))
					# Translators: Fallback label for an unknown hyperlink type
					type_name = MSO_HYPERLINK_TYPES.get(hl_type_int, _("Text Range"))
					disp_txt = getattr(hl, "TextToDisplay", "") or ""
					addr = getattr(hl, "Address", "") or ""
					sub_addr = getattr(hl, "SubAddress", "") or ""
					full_target = addr
					if addr and sub_addr:
						full_target = f"{addr}#{sub_addr}"
					elif sub_addr:
						full_target = f"#{sub_addr}"

					target_snippet = full_target or addr or sub_addr
					snippet = f"{disp_txt} -> {target_snippet}" if (disp_txt and target_snippet) else (disp_txt or target_snippet or f"Hyperlink {i}")
					if len(snippet) > 60:
						snippet = snippet[:57] + "..."

					item = WordElementItem(
						raw_obj=hl,
						elem_type="hyperlink",
						start_pos=hl_start,
						end_pos=hl_end,
						snippet=snippet,
						ref_id=disp_txt or f"Hyperlink {i}",
						full_text=full_target,
						extra={
							"address": addr,
							"sub_address": sub_addr,
							"display": disp_txt,
							"type_int": hl_type_int,
							"type_name": type_name,
						}
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
					result_txt = ""
					try:
						if f.Result and f.Result.Text:
							result_txt = f.Result.Text.strip()
					except Exception:
						pass

					if result_txt:
						snippet = f"{result_txt[:30]} {{{f_code[:25]}}}"
					else:
						snippet = f"{{{f_code[:40]}}}"
					if len(snippet) > 60:
						snippet = snippet[:57] + "..."

					item = WordElementItem(
						raw_obj=f,
						elem_type="field",
						start_pos=f_start,
						end_pos=f_end,
						snippet=snippet,
						ref_id=f"Field {f_type}",
						full_text=f_code,
						extra={
							"field_type": f_type,
							"code_text": f_code,
							"result_text": result_txt,
						}
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
					status_text = ""
					help_text = ""
					try:
						status_text = getattr(ff, "StatusText", "") or ""
					except Exception:
						pass
					try:
						help_text = getattr(ff, "HelpText", "") or ""
					except Exception:
						pass

					snippet = f"{ff_name}: '{ff_val}'" if ff_val else ff_name
					if len(snippet) > 60:
						snippet = snippet[:57] + "..."

					item = WordElementItem(
						raw_obj=ff,
						elem_type="formfield",
						start_pos=ff_start,
						end_pos=ff_end,
						snippet=snippet,
						ref_id=ff_name,
						full_text=ff_val,
						extra={
							"name": ff_name,
							"status_text": status_text,
							"help_text": help_text,
							"type": getattr(ff, "Type", 0),
						}
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

					cc_type_int = int(getattr(cc, "Type", 0))
					# Translators: Fallback label for an unknown content control type
					type_name = WD_CONTENT_CONTROL_TYPES.get(cc_type_int, _("Content Control"))
					is_checkbox = (cc_type_int == 8)  # wdContentControlCheckbox
					is_checked = False
					if is_checkbox:
						try:
							is_checked = bool(getattr(cc, "Checked", False))
						except Exception:
							pass

					# Translators: Fallback title for an unnamed content control
					title = cc.Title or cc.Tag or _("Unnamed Control")
					txt = cc.Range.Text.strip() if cc.Range and cc.Range.Text else ""

					if is_checkbox:
						if is_checked:
							# Translators: Checkbox checked prefix in elements list
							state_str = _("[Checked]")
						else:
							# Translators: Checkbox unchecked prefix in elements list
							state_str = _("[Not Checked]")
						snippet = f"{state_str} {title}"
					else:
						snippet = f"[{type_name}] {title}: {txt[:40]}" if txt else f"[{type_name}] {title}"

					if len(snippet) > 60:
						snippet = snippet[:57] + "..."

					item = WordElementItem(
						raw_obj=cc,
						elem_type="contentcontrol",
						start_pos=cc_start,
						end_pos=cc_end,
						snippet=snippet,
						ref_id=title,
						full_text=txt,
						extra={
							"type_int": cc_type_int,
							"type_name": type_name,
							"is_checkbox": is_checkbox,
							"is_checked": is_checked,
							"title": getattr(cc, "Title", "") or "",
							"tag": getattr(cc, "Tag", "") or "",
						}
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

					sh_type_int = int(getattr(sh, "Type", 0))
					# Translators: Fallback label for an unknown shape type
					type_name = MSO_SHAPE_TYPES.get(sh_type_int, _("Shape"))
					sh_name = sh.Name or f"Shape {i}"
					title = getattr(sh, "Title", "") or ""
					alt_text = getattr(sh, "AlternativeText", "") or ""
					has_alt = bool(alt_text.strip() or title.strip())

					# Check for text box or text frame
					is_textbox = False
					box_text = ""
					try:
						if sh_type_int == 17 or getattr(sh, "HasTextFrame", False):
							tf = getattr(sh, "TextFrame", None)
							if tf and getattr(tf, "HasText", False):
								box_text = tf.TextRange.Text.strip()
								if box_text:
									is_textbox = True
					except Exception:
						pass

					# Dimensions in centimeters
					width_pt = float(getattr(sh, "Width", 0.0) or 0.0)
					height_pt = float(getattr(sh, "Height", 0.0) or 0.0)
					width_cm = round(width_pt * 0.03528, 1)
					height_cm = round(height_pt * 0.03528, 1)

					if is_textbox and box_text:
						clean_box = box_text.replace("\r", " ").replace("\n", " ")
						# Translators: Prefix for text box element snippet
						snippet = _("[Text Box] {text}").format(text=clean_box[:40])
					elif has_alt:
						desc = title or alt_text
						snippet = f"{sh_name}: {desc[:40]}"
					else:
						# Translators: Prefix for graphic missing alternative text
						snippet = _("[No Alt Text] {name}").format(name=sh_name)

					item = WordElementItem(
						raw_obj=sh,
						elem_type="drawingshape",
						start_pos=sh_start,
						end_pos=sh_end,
						snippet=snippet,
						ref_id=sh_name,
						full_text=alt_text or box_text or snippet,
						extra={
							"shape_index": i,
							"type_int": sh_type_int,
							"type_name": type_name,
							"has_alt": has_alt,
							"alt_text": alt_text,
							"title": title,
							"is_textbox": is_textbox,
							"box_text": box_text,
							"width_cm": width_cm,
							"height_cm": height_cm,
						}
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

					ish_type_int = int(getattr(ish, "Type", 0))
					# Translators: Fallback label for an unknown inline shape type
					type_name = WD_INLINE_SHAPE_TYPES.get(ish_type_int, _("Inline Graphic"))
					title = getattr(ish, "Title", "") or ""
					alt_text = getattr(ish, "AlternativeText", "") or ""
					has_alt = bool(alt_text.strip() or title.strip())
					ish_name = title or f"Graphic {i}"

					# Dimensions in centimeters
					width_pt = float(getattr(ish, "Width", 0.0) or 0.0)
					height_pt = float(getattr(ish, "Height", 0.0) or 0.0)
					width_cm = round(width_pt * 0.03528, 1)
					height_cm = round(height_pt * 0.03528, 1)

					if has_alt:
						desc = alt_text or title
						snippet = f"{ish_name}: {desc[:40]}"
					else:
						# Translators: Prefix for inline graphic missing alternative text
						snippet = _("[No Alt Text] {name}").format(name=ish_name)

					item = WordElementItem(
						raw_obj=ish,
						elem_type="inlineshape",
						start_pos=ish_start,
						end_pos=ish_end,
						snippet=snippet,
						ref_id=f"IShape {i}",
						full_text=alt_text or snippet,
						extra={
							"type_int": ish_type_int,
							"type_name": type_name,
							"has_alt": has_alt,
							"alt_text": alt_text,
							"title": title,
							"is_textbox": False,
							"box_text": "",
							"width_cm": width_cm,
							"height_cm": height_cm,
						}
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

					width_pt = 0.0
					height_pt = 0.0
					try:
						width_pt = float(getattr(fr, "Width", 0.0) or 0.0)
					except Exception:
						pass
					try:
						height_pt = float(getattr(fr, "Height", 0.0) or 0.0)
					except Exception:
						pass
					width_cm = round(width_pt * 0.03528, 1)
					height_cm = round(height_pt * 0.03528, 1)

					fr_text = fr.Range.Text.strip() if fr.Range and fr.Range.Text else ""
					snippet = f"Frame {i}: {fr_text[:40]}" if fr_text else f"Frame {i}"
					if len(snippet) > 60:
						snippet = snippet[:57] + "..."

					item = WordElementItem(
						raw_obj=fr,
						elem_type="frame",
						start_pos=fr_start,
						end_pos=fr_end,
						snippet=snippet,
						ref_id=f"Frame {i}",
						full_text=fr_text,
						extra={
							"width_cm": width_cm,
							"height_cm": height_cm,
						}
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

					ps = None
					try:
						ps = sec.PageSetup
					except Exception:
						pass

					orientation_name = ""
					paper_size_name = ""
					start_type_name = ""
					left_cm = right_cm = top_cm = bottom_cm = 0.0
					has_headers = False
					has_footers = False

					if ps:
						try:
							orient_val = getattr(ps, "Orientation", 0)
							# Translators: Fallback orientation
							orientation_name = WD_PAGE_ORIENTATIONS.get(orient_val, _("Portrait"))
						except Exception:
							pass
						try:
							paper_val = getattr(ps, "PaperSize", 7)
							# Translators: Fallback paper size
							paper_size_name = WD_PAPER_SIZES.get(paper_val, _("Standard"))
						except Exception:
							pass
						try:
							start_val = getattr(ps, "SectionStart", 2)
							# Translators: Fallback section start
							start_type_name = WD_SECTION_STARTS.get(start_val, _("New Page"))
						except Exception:
							pass
						try:
							left_cm = round(float(getattr(ps, "LeftMargin", 0.0) or 0.0) * 0.03528, 1)
							right_cm = round(float(getattr(ps, "RightMargin", 0.0) or 0.0) * 0.03528, 1)
							top_cm = round(float(getattr(ps, "TopMargin", 0.0) or 0.0) * 0.03528, 1)
							bottom_cm = round(float(getattr(ps, "BottomMargin", 0.0) or 0.0) * 0.03528, 1)
						except Exception:
							pass

					try:
						if sec.Headers and sec.Headers.Count > 0:
							has_headers = True
					except Exception:
						pass

					try:
						if sec.Footers and sec.Footers.Count > 0:
							has_footers = True
					except Exception:
						pass

					sec_txt = sec.Range.Text.strip() if sec.Range and sec.Range.Text else ""
					snippet = f"Section {i} (Starts: '{sec_txt[:30]}...')" if sec_txt else f"Section {i}"
					if len(snippet) > 60:
						snippet = snippet[:57] + "..."

					item = WordElementItem(
						raw_obj=sec,
						elem_type="section",
						start_pos=sec_start,
						end_pos=sec_end,
						snippet=snippet,
						ref_id=f"Section {i}",
						full_text=snippet,
						extra={
							"section_index": i,
							"orientation": orientation_name,
							"paper_size": paper_size_name,
							"start_type": start_type_name,
							"left_cm": left_cm,
							"right_cm": right_cm,
							"top_cm": top_cm,
							"bottom_cm": bottom_cm,
							"has_headers": has_headers,
							"has_footers": has_footers,
						}
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

		# 18. MATH EQUATIONS
		elif elem_type == "equation":
			omaths = getattr(doc, "OMaths", None)
			count = omaths.Count if omaths else 0
			batch_size = 200
			end_idx = min(start_offset + batch_size - 1, count)
			has_more = (end_idx < count)
			next_offset = (end_idx + 1) if has_more else None
			scanned_count = end_idx

			for i in range(start_offset, end_idx + 1):
				if i % 5 == 0:
					yield True
				try:
					om = omaths(i)
					om_r = getattr(om, "Range", None)
					if not om_r:
						continue
					om_start = om_r.Start
					om_end = om_r.End
					if not is_range_in_scope(om_start, om_end, s_start, s_end):
						continue

					om_type_val = int(getattr(om, "Type", 0))
					# Translators: Fallback equation type name
					type_name = WD_OMATH_TYPES.get(om_type_val, _("Equation"))

					raw_math = om_r.Text.strip() if om_r.Text else ""
					clean_math = raw_math.replace("\r", " ").replace("\n", " ").strip()

					para_text = ""
					try:
						p_obj = om_r.Paragraphs(1)
						if p_obj and p_obj.Range and p_obj.Range.Text:
							para_text = p_obj.Range.Text.strip().replace("\r", " ").replace("\n", " ")
					except Exception:
						pass

					snippet = f"[{type_name}] {clean_math}" if clean_math else f"[{type_name}] Equation {i}"
					if len(snippet) > 60:
						snippet = snippet[:57] + "..."

					item = WordElementItem(
						raw_obj=om,
						elem_type="equation",
						start_pos=om_start,
						end_pos=om_end,
						snippet=snippet,
						ref_id=f"Equation {i}",
						full_text=clean_math,
						extra={
							"type_val": om_type_val,
							"type_name": type_name,
							"math_text": clean_math,
							"para_text": para_text,
							"equation_index": i,
						}
					)
					results.append(item)
				except Exception as e:
					log.debugWarning(f"BOA ElementsOrganizer equation error at {i}: {e}")

		# 19. CITATIONS & REFERENCES
		elif elem_type == "citation":
			# Pre-cache bibliography sources for fast O(1) lookup
			bib_sources = {}
			try:
				b_col = getattr(doc, "Bibliography", None)
				if b_col:
					s_col = getattr(b_col, "Sources", None)
					if s_col:
						for s_idx in range(1, s_col.Count + 1):
							try:
								s_obj = s_col(s_idx)
								tag_key = getattr(s_obj, "Tag", "").strip()
								if tag_key:
									s_title = ""
									s_year = ""
									s_author = ""
									try:
										s_title = s_obj.Field("Title") or ""
									except Exception:
										pass
									try:
										s_year = s_obj.Field("Year") or ""
									except Exception:
										pass
									try:
										s_xml = getattr(s_obj, "XML", "")
										m_last = re.search(r'<b:Last>(.*?)</b:Last>', s_xml)
										m_first = re.search(r'<b:First>(.*?)</b:First>', s_xml)
										if m_last and m_first:
											s_author = f"{m_first.group(1)} {m_last.group(1)}"
										elif m_last:
											s_author = m_last.group(1)
										else:
											s_author = s_obj.Field("Author") or ""
									except Exception:
										pass
									bib_sources[tag_key.lower()] = {
										"tag": tag_key,
										"title": s_title,
										"year": s_year,
										"author": s_author,
									}
							except Exception:
								pass
			except Exception as b_err:
				log.debugWarning(f"BOA ElementsOrganizer bibliography scan error: {b_err}")

			fields = doc.Fields
			count = fields.Count
			batch_size = 250
			end_idx = min(start_offset + batch_size - 1, count)
			has_more = (end_idx < count)
			next_offset = (end_idx + 1) if has_more else None
			scanned_count = end_idx

			for i in range(start_offset, end_idx + 1):
				if i % 5 == 0:
					yield True
				try:
					f = fields(i)
					f_type = getattr(f, "Type", 0)
					code_txt = f.Code.Text.strip() if f.Code and f.Code.Text else ""

					is_cit = (f_type == WD_FIELD_CITATION) or code_txt.upper().startswith("CITATION")
					if not is_cit:
						continue

					f_start = f.Code.Start
					f_end = f.Result.End if f.Result else f.Code.End
					if not is_range_in_scope(f_start, f_end, s_start, s_end):
						continue

					# Parse citation key / tag
					m = re.search(r'CITATION\s+([^\s\\]+)', code_txt, re.IGNORECASE)
					cit_tag = m.group(1).strip('"') if m else "Citation"

					# Visual in-text result (e.g. "(Knuth, 1984)" or "Error! Invalid source specified.")
					res_txt = f.Result.Text.strip() if f.Result and f.Result.Text else ""
					clean_res = res_txt.replace("\r", " ").replace("\n", " ").strip()

					is_invalid = "invalid source" in clean_res.lower()
					source_info = bib_sources.get(cit_tag.lower())

					# Enclosing sentence / paragraph
					sentence_txt = ""
					try:
						sent_col = f.Code.Sentences
						if sent_col and sent_col.Count > 0:
							sentence_txt = sent_col(1).Text.strip().replace("\r", " ").replace("\n", " ")
					except Exception:
						try:
							para_col = f.Code.Paragraphs
							if para_col and para_col.Count > 0:
								sentence_txt = para_col(1).Range.Text.strip().replace("\r", " ").replace("\n", " ")
						except Exception:
							pass

					if is_invalid or not source_info:
						# Translators: Prefix in list when a citation source is not found in bibliography
						status_str = _("[Unresolved Source]")
						snippet = f"{status_str} {cit_tag}"
					else:
						snippet = f"{clean_res}: {sentence_txt[:35]}" if clean_res else cit_tag

					if len(snippet) > 60:
						snippet = snippet[:57] + "..."

					item = WordElementItem(
						raw_obj=f,
						elem_type="citation",
						start_pos=f_start,
						end_pos=f_end,
						snippet=snippet,
						ref_id=cit_tag,
						full_text=clean_res or cit_tag,
						extra={
							"tag": cit_tag,
							"code_text": code_txt,
							"result_text": clean_res,
							"sentence_text": sentence_txt,
							"is_resolved": (source_info is not None) and not is_invalid,
							"author": source_info.get("author", "") if source_info else "",
							"title": source_info.get("title", "") if source_info else "",
							"year": source_info.get("year", "") if source_info else "",
						}
					)
					results.append(item)
				except Exception as e:
					log.debugWarning(f"BOA ElementsOrganizer citation error at {i}: {e}")

		# 20. CAPTIONS (FIGURES, TABLES, EQUATIONS)
		elif elem_type == "caption":
			fields = doc.Fields
			count = fields.Count
			batch_size = 250
			end_idx = min(start_offset + batch_size - 1, count)
			has_more = (end_idx < count)
			next_offset = (end_idx + 1) if has_more else None
			scanned_count = end_idx

			for i in range(start_offset, end_idx + 1):
				if i % 5 == 0:
					yield True
				try:
					f = fields(i)
					f_type = getattr(f, "Type", 0)
					code_txt = f.Code.Text.strip() if f.Code and f.Code.Text else ""

					is_seq = (f_type == WD_FIELD_SEQUENCE) or code_txt.upper().startswith("SEQ")
					if not is_seq:
						continue

					f_start = f.Code.Start
					f_end = f.Result.End if f.Result else f.Code.End
					if not is_range_in_scope(f_start, f_end, s_start, s_end):
						continue

					# Parse SEQ label (e.g. Figure, Table, Equation)
					m = re.search(r'SEQ\s+([^\s\\]+)', code_txt, re.IGNORECASE)
					label = m.group(1).strip('"') if m else "Caption"

					# Sequence number rendered result (e.g. "1", "2")
					seq_num = f.Result.Text.strip() if f.Result and f.Result.Text else ""

					# Full caption paragraph
					full_caption = ""
					p = None
					try:
						p = f.Code.Paragraphs(1)
						if p and p.Range and p.Range.Text:
							full_caption = p.Range.Text.strip().replace("\r", " ").replace("\n", " ")
					except Exception:
						pass

					if not full_caption:
						full_caption = f"{label} {seq_num}".strip()

					# Extract description by stripping label & sequence number
					desc = re.sub(r'^' + re.escape(label) + r'\s+' + re.escape(seq_num) + r'[:\.\-\s]*', '', full_caption, flags=re.IGNORECASE).strip()
					if not desc:
						desc = full_caption

					# Proximity check for adjacent graphic or table
					adjacent_type = ""
					if p:
						try:
							prev_p = p.Previous()
							if prev_p:
								if getattr(prev_p.Range, "InlineShapes", None) and prev_p.Range.InlineShapes.Count > 0:
									# Translators: Proximity audit indicating graphic above caption
									adjacent_type = _("Image / Shape Present (Above caption)")
								elif getattr(prev_p.Range, "Tables", None) and prev_p.Range.Tables.Count > 0:
									# Translators: Proximity audit indicating table above caption
									adjacent_type = _("Table Present (Above caption)")
						except Exception:
							pass

						if not adjacent_type:
							try:
								next_p = p.Next()
								if next_p:
									if getattr(next_p.Range, "Tables", None) and next_p.Range.Tables.Count > 0:
										# Translators: Proximity audit indicating table below caption
										adjacent_type = _("Table Present (Below caption)")
									elif getattr(next_p.Range, "InlineShapes", None) and next_p.Range.InlineShapes.Count > 0:
										# Translators: Proximity audit indicating graphic below caption
										adjacent_type = _("Image / Shape Present (Below caption)")
							except Exception:
								pass

					if not adjacent_type:
						# Translators: Proximity audit when no table or image is directly adjacent
						adjacent_type = _("None detected nearby (Floating shape or text-only caption)")

					snippet = full_caption
					if len(snippet) > 60:
						snippet = snippet[:57] + "..."

					item = WordElementItem(
						raw_obj=f,
						elem_type="caption",
						start_pos=f_start,
						end_pos=f_end,
						snippet=snippet,
						ref_id=label,
						full_text=full_caption,
						extra={
							"label": label,
							"seq_num": seq_num,
							"code_text": code_txt,
							"full_caption": full_caption,
							"description": desc,
							"adjacent_type": adjacent_type,
						}
					)
					results.append(item)
				except Exception as e:
					log.debugWarning(f"BOA ElementsOrganizer caption error at {i}: {e}")

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

		# Translators: Button to edit or add alternative text to a graphic or shape
		self.btn_edit_alt = wx.Button(self, label=_("Edit Alt &Text..."))
		self.btn_edit_alt.Bind(wx.EVT_BUTTON, self.on_edit_alt_text)
		self.btn_edit_alt.Hide()
		self.btn_sizer.Add(self.btn_edit_alt, 0, wx.RIGHT, 5)

		# Translators: Button to reply to a comment
		self.btn_reply = wx.Button(self, label=_("&Reply..."))
		self.btn_reply.Bind(wx.EVT_BUTTON, self.on_reply)
		self.btn_reply.Hide()
		self.btn_sizer.Add(self.btn_reply, 0, wx.RIGHT, 5)

		# Translators: Button to modify the text content of a comment, footnote, or endnote
		self.btn_modify_text = wx.Button(self, label=_("&Modify Text..."))
		self.btn_modify_text.Bind(wx.EVT_BUTTON, self.on_modify_text)
		self.btn_modify_text.Hide()
		self.btn_sizer.Add(self.btn_modify_text, 0, wx.RIGHT, 5)

		# Translators: Button to modify the author name of a comment
		self.btn_modify_author = wx.Button(self, label=_("Modify A&uthor..."))
		self.btn_modify_author.Bind(wx.EVT_BUTTON, self.on_modify_author)
		self.btn_modify_author.Hide()
		self.btn_sizer.Add(self.btn_modify_author, 0, wx.RIGHT, 5)

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
		self.btn_sizer.Show(self.btn_modify_text, elem_type in ("comment", "footnote", "endnote"))
		self.btn_sizer.Show(self.btn_modify_author, elem_type == "comment")
		self.btn_sizer.Show(self.btn_accept, elem_type == "revision")
		self.btn_sizer.Show(self.btn_reject, elem_type == "revision")
		self.btn_sizer.Show(self.btn_open_link, elem_type == "hyperlink")
		self.btn_sizer.Show(self.btn_rename, elem_type == "bookmark")
		self.btn_sizer.Show(self.btn_apply, elem_type == "spellingerror")
		self.btn_sizer.Show(self.btn_edit_alt, elem_type in ("drawingshape", "inlineshape"))
		self.btn_sizer.Show(self.btn_load_more, getattr(self, "batch_has_more", False))

		# 2. Enablement (Enable only when an item is available and selected)
		self.btn_goto.Enable(has_selection)
		self.btn_copy_all.Enable(has_items)
		can_delete = has_selection and elem_type in (
			"comment", "revision", "footnote", "endnote", "bookmark",
			"table", "hyperlink", "drawingshape", "inlineshape", "frame",
			"equation", "citation", "caption"
		)
		self.btn_delete.Enable(can_delete)
		self.btn_reply.Enable(has_selection)
		self.btn_modify_text.Enable(has_selection and elem_type in ("comment", "footnote", "endnote"))
		self.btn_modify_author.Enable(has_selection and elem_type == "comment")
		self.btn_accept.Enable(has_selection)
		self.btn_reject.Enable(has_selection)
		self.btn_open_link.Enable(has_selection)
		self.btn_rename.Enable(has_selection)
		self.btn_apply.Enable(has_selection and elem_type == "spellingerror")
		self.btn_edit_alt.Enable(has_selection and elem_type in ("drawingshape", "inlineshape"))

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
			if target_idx < 0:
				target_idx = 0
			elif target_idx >= len(self.items):
				target_idx = max(0, len(self.items) - 1)
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
			is_uniform = item.extra.get("is_uniform", True)
			# Translators: Table dimensions summary in preview
			txt_parts.append(_("Dimensions: {r} rows x {c} columns").format(r=rows_c, c=cols_c))
			if is_uniform:
				# Translators: Indicates table has a uniform rectangular grid structure
				txt_parts.append(_("Grid Structure: Uniform (Clean rectangular grid)"))
			else:
				# Translators: Indicates table has merged or split cells
				txt_parts.append(_("Grid Structure: Irregular (Contains merged or split cells)"))
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
			type_name = item.extra.get("type_name", "")
			if type_name:
				# Translators: Hyperlink type in preview
				txt_parts.append(_("Hyperlink Type: {t}").format(t=type_name))
			# Translators: Label for hyperlink display text in preview
			txt_parts.append(_("Text: {t}").format(t=item.ref_id))
			addr = item.extra.get("address", "")
			sub_addr = item.extra.get("sub_address", "")
			if addr:
				# Translators: Label for hyperlink destination URL in preview
				txt_parts.append(_("Target URL: {u}").format(u=addr))
			if sub_addr:
				# Translators: Label for internal bookmark or destination in preview
				txt_parts.append(_("Internal Destination: {dest}").format(dest=sub_addr))

		elif item.elem_type == "bookmark":
			# Translators: Label for bookmark name in preview
			txt_parts.append(_("Bookmark Name: {n}").format(n=item.ref_id))
			st_name = item.extra.get("story_name", "")
			if st_name:
				# Translators: Bookmark story location in preview
				txt_parts.append(_("Location Story: {story}").format(story=st_name))
			is_empty = item.extra.get("is_empty", False)
			if is_empty:
				# Translators: Bookmark empty status
				txt_parts.append(_("Type: Empty Position Marker (No enclosed text)"))
			else:
				# Translators: Bookmark enclosing text status
				txt_parts.append(_("Type: Encloses Text"))
			if item.full_text:
				# Translators: Bookmarked text content
				txt_parts.append(_("Bookmarked Text:\n{t}").format(t=item.full_text))

		elif item.elem_type == "field":
			f_type = item.extra.get("field_type", "")
			res_txt = item.extra.get("result_text", "")
			code_txt = item.extra.get("code_text", item.full_text)
			if f_type:
				# Translators: Field type header in preview
				txt_parts.append(_("Field Type: {t}").format(t=f_type))
			if res_txt:
				# Translators: Rendered visual result of a field
				txt_parts.append(_("Visual Result:\n{res}").format(res=res_txt))
			# Translators: Formula or code underlying a field
			txt_parts.append(_("Field Code:\n{{{code}}}").format(code=code_txt))

		elif item.elem_type == "formfield":
			# Translators: Form field name in preview
			txt_parts.append(_("Field Name: {name}").format(name=item.ref_id))
			st = item.extra.get("status_text", "")
			ht = item.extra.get("help_text", "")
			if st:
				# Translators: Status bar prompt configured for a form field
				txt_parts.append(_("Status Prompt: {text}").format(text=st))
			if ht:
				# Translators: Help text prompt configured for a form field
				txt_parts.append(_("Help Text: {text}").format(text=ht))
			if item.full_text:
				# Translators: Form field current value
				txt_parts.append(_("Current Value: {val}").format(val=item.full_text))

		elif item.elem_type == "contentcontrol":
			t_name = item.extra.get("type_name", "")
			if t_name:
				# Translators: Content control type in preview
				txt_parts.append(_("Control Type: {t}").format(t=t_name))
			if item.extra.get("title"):
				# Translators: Content control title in preview
				txt_parts.append(_("Title: {title}").format(title=item.extra["title"]))
			if item.extra.get("tag"):
				# Translators: Content control tag in preview
				txt_parts.append(_("Tag: {tag}").format(tag=item.extra["tag"]))
			if item.extra.get("is_checkbox"):
				chk = item.extra.get("is_checked", False)
				if chk:
					# Translators: Checkbox content control state in preview
					txt_parts.append(_("State: Checked (Yes)"))
				else:
					# Translators: Checkbox content control state in preview
					txt_parts.append(_("State: Not Checked (No)"))
			if item.full_text:
				# Translators: Content control text contents in preview
				txt_parts.append(_("Content:\n{text}").format(text=item.full_text))

		elif item.elem_type == "heading":
			lvl = item.extra.get("level", "?")
			s_name = item.extra.get("style_name", "")
			if s_name:
				# Translators: Heading level and Word style name in preview
				txt_parts.append(_("Level: Heading {lvl} | Style: {style}").format(lvl=lvl, style=s_name))
			else:
				# Translators: Heading level in preview
				txt_parts.append(_("Level: Heading {lvl}").format(lvl=lvl))
			if item.full_text:
				txt_parts.append(f"\n{item.full_text}")

		elif item.elem_type == "frame":
			w_cm = item.extra.get("width_cm", 0.0)
			h_cm = item.extra.get("height_cm", 0.0)
			if w_cm > 0 and h_cm > 0:
				# Translators: Frame dimensions in preview pane (centimeters)
				txt_parts.append(_("Dimensions: {w} cm wide x {h} cm high").format(w=w_cm, h=h_cm))
			if item.full_text:
				# Translators: Frame text content in preview pane
				txt_parts.append(_("Frame Content:\n{text}").format(text=item.full_text))

		elif item.elem_type == "section":
			orient = item.extra.get("orientation", "")
			paper = item.extra.get("paper_size", "")
			start_t = item.extra.get("start_type", "")
			top_m = item.extra.get("top_cm", 0.0)
			bot_m = item.extra.get("bottom_cm", 0.0)
			left_m = item.extra.get("left_cm", 0.0)
			right_m = item.extra.get("right_cm", 0.0)
			has_h = item.extra.get("has_headers", False)
			has_f = item.extra.get("has_footers", False)

			# Translators: Section layout summary in preview
			txt_parts.append(_("Orientation: {orient} | Paper Size: {paper} | Break: {start_t}").format(orient=orient, paper=paper, start_t=start_t))
			# Translators: Margins summary in centimeters
			txt_parts.append(_("Margins: Top {top} cm, Bottom {bot} cm, Left {left} cm, Right {right} cm").format(top=top_m, bot=bot_m, left=left_m, right=right_m))
			# Translators: Header presence in section
			h_str = _("Present") if has_h else _("None")
			# Translators: Footer presence in section
			f_str = _("Present") if has_f else _("None")
			# Translators: Presence of headers and footers in section
			txt_parts.append(_("Headers: {h} | Footers: {f}").format(h=h_str, f=f_str))
			if item.full_text:
				txt_parts.append(f"\n{item.full_text}")

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
			txt_parts.append(_("Suggestions: Managed by Microsoft Editor.\nPress Enter or click 'Go To' to navigate to this error in Word, then press Alt+DownArrow or Shift+F10 to review suggestions."))

		elif item.elem_type in ("drawingshape", "inlineshape"):
			# Translators: Fallback label for a graphic element
			type_name = item.extra.get("type_name", _("Graphic"))
			width_cm = item.extra.get("width_cm", 0.0)
			height_cm = item.extra.get("height_cm", 0.0)
			has_alt = item.extra.get("has_alt", False)
			alt_text = item.extra.get("alt_text", "")
			title = item.extra.get("title", "")
			is_textbox = item.extra.get("is_textbox", False)
			box_text = item.extra.get("box_text", "")

			# Translators: Shape type in preview pane
			txt_parts.append(_("Type: {t}").format(t=type_name))
			if width_cm > 0 and height_cm > 0:
				# Translators: Graphic dimensions in preview pane (centimeters)
				txt_parts.append(_("Dimensions: {w} cm wide x {h} cm high").format(w=width_cm, h=height_cm))

			if is_textbox and box_text:
				# Translators: Label for text box contents in preview pane
				txt_parts.append(_("Text Content:\n{text}").format(text=box_text))
			elif has_alt:
				if title:
					# Translators: Title of a graphic in preview pane
					txt_parts.append(_("Title: {t}").format(t=title))
				if alt_text:
					# Translators: Alternative text description of a graphic in preview pane
					txt_parts.append(_("Alternative Text:\n{desc}").format(desc=alt_text))
			else:
				# Translators: Status indicator when a graphic has no alternative text
				txt_parts.append(_("Alt Text Status: MISSING (Inaccessible to screen readers)"))
				# Translators: Guidance for editing missing alternative text
				txt_parts.append(_("Press Alt+T or click 'Edit Alt Text' to add a description."))

		elif item.elem_type == "equation":
			t_name = item.extra.get("type_name", "")
			math_txt = item.extra.get("math_text", item.full_text)
			para_txt = item.extra.get("para_text", "")
			if t_name:
				# Translators: Math equation type header in preview
				txt_parts.append(_("Equation Type: {t}").format(t=t_name))
			if math_txt:
				# Translators: Header for visual rendered formula in preview
				txt_parts.append(_("Visual Formula Rendering:\n{math}").format(math=math_txt))
			if para_txt:
				# Translators: Enclosing context header for equation preview
				txt_parts.append(_("Enclosing Context:\n\"{context}\"").format(context=para_txt))

		elif item.elem_type == "citation":
			cit_tag = item.extra.get("tag", item.ref_id)
			res_txt = item.extra.get("result_text", "")
			is_resolved = item.extra.get("is_resolved", False)
			sent_txt = item.extra.get("sentence_text", "")
			code_txt = item.extra.get("code_text", "")

			if res_txt:
				# Translators: Header for visual rendered in-text citation in preview
				txt_parts.append(_("Visual In-Text Citation:\n{res}").format(res=res_txt))
			# Translators: Citation tag / key header in preview
			txt_parts.append(_("Citation Key: {key}").format(key=cit_tag))

			if is_resolved:
				# Translators: Indicates citation key was verified in master bibliography
				txt_parts.append(_("Source Status: Verified in Bibliography"))
				auth = item.extra.get("author", "")
				title = item.extra.get("title", "")
				yr = item.extra.get("year", "")
				ref_parts = []
				if auth:
					ref_parts.append(auth)
				if yr:
					ref_parts.append(f"({yr})")
				if title:
					ref_parts.append(f'"{title}"')
				if ref_parts:
					# Translators: Master bibliography source entry in preview
					txt_parts.append(_("Bibliographic Reference:\n{ref}").format(ref=" ".join(ref_parts)))
			else:
				# Translators: Indicates citation key could not be found in bibliography
				txt_parts.append(_("Source Status: UNRESOLVED (Source missing from Bibliography)"))

			if sent_txt:
				# Translators: Context sentence for citation in preview
				txt_parts.append(_("Cited In Sentence:\n\"{sent}\"").format(sent=sent_txt))
			if code_txt:
				# Translators: Raw citation field code in preview
				txt_parts.append(_("Field Code: {{{code}}}").format(code=code_txt))

		elif item.elem_type == "caption":
			lbl = item.extra.get("label", "")
			s_num = item.extra.get("seq_num", "")
			full_cap = item.extra.get("full_caption", item.full_text)
			descr = item.extra.get("description", "")
			adj_obj = item.extra.get("adjacent_type", "")

			if full_cap:
				# Translators: Header for rendered visual caption in preview
				txt_parts.append(_("Visual Caption (As Rendered):\n{cap}").format(cap=full_cap))
			# Translators: Category and sequence number for caption in preview
			txt_parts.append(_("Category: {cat} | Number: {num}").format(cat=lbl, num=s_num))
			if descr and descr != full_cap:
				# Translators: Caption description text without label
				txt_parts.append(_("Description: {desc}").format(desc=descr))
			if adj_obj:
				# Translators: Proximity check audit in preview
				txt_parts.append(_("Associated Object: {obj}").format(obj=adj_obj))

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
		elif key == ord('T') and wx.GetKeyState(wx.WXK_CONTROL):
			item = self._get_selected_item()
			if item and item.elem_type in ("drawingshape", "inlineshape"):
				self.on_edit_alt_text(None)
				return
		elif key == ord('M') and wx.GetKeyState(wx.WXK_CONTROL):
			item = self._get_selected_item()
			if item and item.elem_type in ("comment", "footnote", "endnote"):
				self.on_modify_text(None)
				return
		elif key == ord('U') and wx.GetKeyState(wx.WXK_CONTROL):
			item = self._get_selected_item()
			if item and item.elem_type == "comment":
				self.on_modify_author(None)
				return
		elif key == wx.WXK_F2:
			item = self._get_selected_item()
			if item and item.elem_type == "bookmark":
				self.on_rename_bookmark(None)
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
			"snippet": item.snippet,
			"ref_id": item.ref_id,
			"shape_index": item.extra.get("shape_index") if item.extra else None,
			"is_textbox": item.extra.get("is_textbox", False) if item.extra else False,
		}
		self.EndModal(wx.ID_OK)

	def on_open_link(self, evt=None):
		item = self._get_selected_item()
		if not item or item.elem_type != "hyperlink":
			return
		addr = item.extra.get("address", "") or item.full_text
		sub_addr = item.extra.get("sub_address", "")
		if not addr and sub_addr:
			# Translators: Spoken when a hyperlink points to an internal bookmark or heading
			ui.message(_("Internal document link to '{dest}'. Press Enter or click 'Go To' to jump to it in Word.").format(dest=sub_addr))
			return
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

	def on_edit_alt_text(self, evt=None):
		item = self._get_selected_item()
		if not item or item.elem_type not in ("drawingshape", "inlineshape"):
			return
		curr_alt = item.extra.get("alt_text", "")
		# Translators: Fallback name for a graphic without a title
		item_name = item.ref_id or _("Graphic")
		# Translators: Prompt message to enter alternative text for a graphic
		prompt_msg = _("Enter alternative text description for '{name}':").format(name=item_name)
		# Translators: Title of the alternative text editor dialog
		dlg_title = _("Edit Alternative Text")
		dlg = wx.TextEntryDialog(self, prompt_msg, dlg_title, value=curr_alt, style=wx.OK | wx.CANCEL | wx.TE_MULTILINE)
		if dlg.ShowModal() == wx.ID_OK:
			new_alt = dlg.GetValue().strip()
			try:
				raw = item.raw_obj
				if raw:
					raw.AlternativeText = new_alt
				item.extra["alt_text"] = new_alt
				item.extra["has_alt"] = bool(new_alt)
				item.full_text = new_alt or item.snippet
				if new_alt:
					item.snippet = f"{item_name}: {new_alt[:40]}"
				else:
					# Translators: Prefix for graphic missing alternative text
					item.snippet = _("[No Alt Text] {name}").format(name=item_name)
				idx = self.list_ctrl.GetFirstSelected()
				if idx >= 0:
					self.list_ctrl.SetItem(idx, 1, item.snippet)
				self._display_item_preview(item)
				# Translators: Spoken announcement when alternative text is successfully updated
				ui.message(_("Alternative text updated."))
			except Exception as err:
				log.error(f"BOA ElementsOrganizer edit alt text error: {err}")
				# Translators: Error message when updating alternative text fails
				ui.message(_("Failed to update alternative text: {err}").format(err=str(err)))
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

	def on_modify_text(self, evt=None):
		item = self._get_selected_item()
		if not item or item.elem_type not in ("comment", "footnote", "endnote"):
			return

		# Safe Rich Content Guard for Footnotes and Endnotes
		if item.elem_type in ("footnote", "endnote"):
			note = item.raw_obj
			note_range = getattr(note, "Range", None)
			has_rich = False
			if note_range:
				try:
					if getattr(note_range, "InlineShapes", None) and note_range.InlineShapes.Count > 0:
						has_rich = True
				except Exception:
					pass
				try:
					if getattr(note_range, "Fields", None) and note_range.Fields.Count > 0:
						has_rich = True
				except Exception:
					pass
				try:
					if hasattr(note_range, "OMaths") and note_range.OMaths.Count > 0:
						has_rich = True
				except Exception:
					pass

			# Translators: Element type label for footnote
			fn_label = _("footnote")
			# Translators: Element type label for endnote
			en_label = _("endnote")
			note_label = fn_label if item.elem_type == "footnote" else en_label

			if has_rich:
				# Translators: Warning dialog message when attempting to edit a note containing rich elements
				warning_msg = _(
					"This {note_type} contains non-text elements (such as embedded fields, math equations, or images). "
					"Editing it in plain text would remove those objects.\n\n"
					"To protect your document formatting, please use 'Go To' to edit this note directly in Microsoft Word."
				).format(note_type=note_label)
				# Translators: Title of warning dialog for rich note content
				warning_title = _("Rich Content Warning")
				gui.messageBox(warning_msg, warning_title, wx.OK | wx.ICON_WARNING, self)
				return

			# Translators: Prompt message for modifying footnote or endnote text
			prompt_msg = _("Enter new text for {note_type}:").format(note_type=note_label)
			# Translators: Title of footnote text modification dialog
			fn_title = _("Modify Footnote")
			# Translators: Title of endnote text modification dialog
			en_title = _("Modify Endnote")
			dlg_title = fn_title if item.elem_type == "footnote" else en_title

			dlg = wx.TextEntryDialog(self, prompt_msg, dlg_title, value=item.full_text, style=wx.OK | wx.CANCEL | wx.TE_MULTILINE)
			if dlg.ShowModal() == wx.ID_OK:
				new_text = dlg.GetValue().strip()
				if not new_text:
					dlg.Destroy()
					return
				try:
					note.Range.Text = new_text
					item.full_text = new_text
					ref_idx = item.ref_id.replace("#", "")
					clean_snippet = new_text.replace("\r", " ").replace("\n", " ")
					snippet = f"[{ref_idx}] {clean_snippet}"
					if len(snippet) > 60:
						snippet = snippet[:57] + "..."
					item.snippet = snippet
					idx = self.list_ctrl.GetFirstSelected()
					if idx >= 0:
						self.list_ctrl.SetItem(idx, 1, snippet)
					self._display_item_preview(item)
					# Translators: Success announcement when note text is modified
					ui.message(_("{note_type} text updated.").format(note_type=note_label.capitalize()))
				except Exception as e:
					log.error(f"BOA ElementsOrganizer modify note error: {e}")
					# Translators: Error message when modifying note text fails
					ui.message(_("Failed to update {note_type} text.").format(note_type=note_label))
			dlg.Destroy()
			return

		# Comment Text Modification
		elif item.elem_type == "comment":
			# Translators: Prompt message for modifying comment text
			prompt_msg = _("Enter new text for comment:")
			# Translators: Title of comment text modification dialog
			dlg_title = _("Modify Comment Text")
			dlg = wx.TextEntryDialog(self, prompt_msg, dlg_title, value=item.full_text, style=wx.OK | wx.CANCEL | wx.TE_MULTILINE)
			if dlg.ShowModal() == wx.ID_OK:
				new_text = dlg.GetValue().strip()
				if not new_text:
					dlg.Destroy()
					return
				try:
					c = item.raw_obj
					c.Range.Text = new_text
					item.full_text = new_text
					clean_txt = new_text.replace("\r", " ").replace("\n", " ")
					snippet = clean_txt[:57] + "..." if len(clean_txt) > 60 else clean_txt
					item.snippet = snippet
					idx = self.list_ctrl.GetFirstSelected()
					if idx >= 0:
						self.list_ctrl.SetItem(idx, 1, snippet)
					self._display_item_preview(item)
					# Translators: Success announcement when comment text is modified
					ui.message(_("Comment text updated."))
				except Exception as e:
					log.error(f"BOA ElementsOrganizer modify comment text error: {e}")
					# Translators: Error message when modifying comment text fails
					ui.message(_("Failed to update comment text."))
			dlg.Destroy()

	def on_modify_author(self, evt=None):
		item = self._get_selected_item()
		if not item or item.elem_type != "comment":
			return

		old_author = item.author
		# Translators: Prompt message for modifying comment author name
		prompt_msg = _("Enter new author name for comment:")
		# Translators: Title of comment author modification dialog
		dlg_title = _("Modify Comment Author")
		dlg = wx.TextEntryDialog(self, prompt_msg, dlg_title, value=old_author)
		if dlg.ShowModal() == wx.ID_OK:
			new_author = dlg.GetValue().strip()
			if not new_author or new_author == old_author:
				dlg.Destroy()
				return
			try:
				c = item.raw_obj
				c.Author = new_author
				item.author = new_author
				idx = self.list_ctrl.GetFirstSelected()
				if idx >= 0:
					self.list_ctrl.SetItem(idx, 2, item.get_meta_display())
				self._display_item_preview(item)
				# Translators: Success announcement when comment author is modified
				ui.message(_("Comment author updated to {name}").format(name=new_author))
			except Exception as e:
				log.error(f"BOA ElementsOrganizer modify comment author error: {e}")
				# Translators: Error message when modifying comment author fails
				ui.message(_("Failed to update comment author."))
		dlg.Destroy()

	def on_accept_revision(self, evt=None):
		item = self._get_selected_item()
		if item and item.elem_type == "revision":
			try:
				item.raw_obj.Accept()
				# Translators: Success announcement when a revision is accepted
				ui.message(_("Revision accepted."))
				self._start_scan(preserve_selection=True)
			except Exception as e:
				log.error(f"BOA ElementsOrganizer Accept Error: {e}")
				# Translators: Error message when accepting revision fails
				ui.message(_("Failed to accept revision."))

	def on_reject_revision(self, evt=None):
		item = self._get_selected_item()
		if item and item.elem_type == "revision":
			try:
				item.raw_obj.Reject()
				# Translators: Success announcement when a revision is rejected
				ui.message(_("Revision rejected."))
				self._start_scan(preserve_selection=True)
			except Exception as e:
				log.error(f"BOA ElementsOrganizer Reject Error: {e}")
				# Translators: Error message when rejecting revision fails
				ui.message(_("Failed to reject revision."))

	def on_delete(self, evt=None):
		item = self._get_selected_item()
		if not item:
			return

		elem_dict = dict(ELEMENT_TYPES)
		type_label = elem_dict.get(item.elem_type, item.elem_type.capitalize())
		item_name = item.ref_id or item.snippet
		if len(item_name) > 40:
			item_name = item_name[:37] + "..."

		# Special handling for tracked revisions (Word COM does not have Revision.Delete, only Revision.Reject)
		if item.elem_type == "revision":
			# Translators: Confirmation prompt before rejecting a tracked change revision
			msg = _("Are you sure you want to reject '{name}' ({type})?").format(name=item_name, type=type_label)
			# Translators: Title of confirmation dialog before rejecting a revision
			title = _("Confirm Rejection")
			if wx.MessageBox(msg, title, wx.YES_NO | wx.ICON_QUESTION, self) == wx.YES:
				try:
					item.raw_obj.Reject()
					# Translators: Success announcement when a revision is rejected
					ui.message(_("Revision rejected."))
					self._start_scan(preserve_selection=True)
				except Exception as e:
					log.error(f"BOA ElementsOrganizer Reject Error: {e}")
					# Translators: Error message when rejecting revision fails
					ui.message(_("Failed to reject revision."))
			return

		# Standard deletion for comments, bookmarks, footnotes, endnotes, tables, shapes, frames, hyperlinks, equations, citations, captions
		# Translators: Confirmation prompt before deleting an element with its name and type
		msg = _("Are you sure you want to delete '{name}' ({type})?").format(name=item_name, type=type_label)
		# Translators: Title of confirmation dialog before deleting an element
		title = _("Confirm Deletion")
		if wx.MessageBox(msg, title, wx.YES_NO | wx.ICON_QUESTION, self) == wx.YES:
			try:
				if item.elem_type == "equation":
					if hasattr(item.raw_obj, "Remove"):
						item.raw_obj.Remove()
					else:
						item.raw_obj.Range.Delete()
				else:
					item.raw_obj.Delete()
				# Translators: Success announcement
				ui.message(_("Item deleted."))
				self._start_scan(preserve_selection=True)
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
		try:
			dlg = WordElementsOrganizerDialog(gui.mainFrame, doc, selection, app)
			result = dlg.ShowModal()
			if result == wx.ID_OK:
				nav_target = getattr(dlg, "_nav_target", None)
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

					# 2. Select target range or text frame
					is_textbox = nav_target.get("is_textbox", False)
					shape_index = nav_target.get("shape_index", None)
					if is_textbox and shape_index:
						try:
							sh = doc.Shapes(shape_index)
							if sh and getattr(sh, "HasTextFrame", False):
								tf = getattr(sh, "TextFrame", None)
								if tf and getattr(tf, "HasText", False):
									tf_range = tf.TextRange
									tf_range.Collapse(1)  # wdCollapseStart
									tf_range.Select()
									return
						except Exception as tb_err:
							log.debugWarning(f"BOA could not navigate into text box: {tb_err}")

					start_pos = nav_target.get("start_pos", 0)
					end_pos = nav_target.get("end_pos", start_pos)
					if end_pos < start_pos:
						end_pos = start_pos

					elem_type = nav_target.get("elem_type", "")
					if elem_type == "formfield":
						ref_id = nav_target.get("ref_id", "")
						if ref_id and hasattr(doc, "FormFields"):
							try:
								doc.FormFields(ref_id).Select()
								return
							except Exception as ff_sel_err:
								log.debugWarning(f"BOA FormField native Select failed, falling back to Range: {ff_sel_err}")

					# Error elements require full block highlight for Word's correction engines (Editor / Spellcheck);
					# all navigation & structural elements use collapsed caret placement for safety and clean speech.
					if elem_type in ("spellingerror", "grammaticalerror"):
						target_range = doc.Range(start_pos, end_pos)
					else:
						target_range = doc.Range(start_pos, start_pos)

					if target_range:
						target_range.Select()

						if elem_type == "grammaticalerror":
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
