# -*- coding: UTF-8 -*-
# Copyright (C) 2026 KIRAN G T {MisterK} and Antigravity 2
# This file is covered by the GNU General Public License (GPL), version 2.
# See the file COPYING.txt for more details.

import addonHandler
addonHandler.initTranslation()

import ui
import wx
from logHandler import log
from appModules.boa_enhancements.async_engine import AsyncCOMTask
from .slide_layout_analyzer import SlideLayoutAnalyzer

class PowerPointDocumentAnalyzer:
	"""
	Analyzes the entire PowerPoint presentation in the background.
	Generates a deeply navigable HTML report acting as a Table of Contents
	and accessibility health overview.
	"""

	@staticmethod
	def _show_dialog(title, message):
		wx.CallAfter(ui.browseableMessage, message, title=title, isHtml=True, closeButton=True, copyButton=True)

	@classmethod
	def analyze(cls, obj):
		try:
			# Re-use the existing robust COM hooking from SlideLayoutAnalyzer
			app = SlideLayoutAnalyzer._get_ppt_com()
			
			if not app:
				# Translators: Error hooking PowerPoint COM
				ui.message(_("Failed to hook PowerPoint."))
				return
			
			try:
				pres = app.ActivePresentation
				total_slides = pres.Slides.Count
				pres_name = pres.Name
			except Exception:
				# Translators: Error when no presentation is active
				ui.message(_("No active presentation detected."))
				return
				
			# Translators: Announcement when starting document analysis
			ui.message(_("Analyzing document layout in background..."))
			
			task = AsyncCOMTask(
				generator=cls._analyzer_generator(pres),
				on_complete=lambda results: cls._format_and_show_results(results, pres_name)
			)
			task.start()
			
		except Exception as e:
			log.error(f"BOA PowerPointDocumentAnalyzer Error: {e}", exc_info=True)
			# Translators: Error during analysis
			ui.message(_("Error starting document analysis."))

	@staticmethod
	def _analyzer_generator(pres):
		results = {
			"total_slides": 0,
			"hidden_slides": 0,
			"toc": [],
			"reading_order_issues": [],
			"missing_alt": [],
			"off_canvas": [],
			"tables": 0,
			"groups": 0,
			"media": 0,
			"charts": 0,
			"ole_objects": 0,
			"dense_slides": []
		}
		
		try:
			slides = pres.Slides
			total = slides.Count
			results["total_slides"] = total
			
			for slide_idx in range(1, total + 1):
				yield True
				
				try:
					slide = slides(slide_idx)
					slide_num = slide.SlideNumber
					
					# Hidden check
					if SlideLayoutAnalyzer._safe_com_bool(slide.SlideShowTransition, "Hidden"):
						results["hidden_slides"] += 1
						
					# Extract Title for TOC
					slide_title = ""
					if SlideLayoutAnalyzer._safe_com_bool(slide.Shapes, "HasTitle"):
						try:
							slide_title = slide.Shapes.Title.TextFrame.TextRange.Text.strip()
						except Exception:
							pass
					
					if not slide_title:
						# Translators: Fallback title for slides missing a Title placeholder
						slide_title = _("Untitled Slide")
					
					results["toc"].append({"num": slide_num, "title": slide_title})
					
					# Deep Shape Analysis for this slide
					shapes = slide.Shapes
					s_count = shapes.Count
					
					slide_missing_alt = 0
					slide_off_canvas = False
					slide_word_count = 0
					shape_positions = []
					
					for i in range(1, s_count + 1):
						# Yield periodically inside heavy slides
						if i % 10 == 0:
							yield True
							
						try:
							shape = shapes(i)
							s_type = shape.Type
							
							# Positions for Z-order check
							try:
								top = shape.Top
								shape_positions.append({"name": shape.Name, "top": top, "index": i})
								if top < 0 or shape.Left < 0:
									slide_off_canvas = True
							except Exception:
								pass
								
							# Alt text
							# 13=Picture, 3=Chart, 24=SmartArt, 7=OLE, 16=Media
							if s_type in (13, 3, 24, 7, 16):
								try:
									alt = shape.AlternativeText
									if not alt or alt.isspace():
										slide_missing_alt += 1
								except Exception:
									slide_missing_alt += 1
									
							# Complexity Mapping
							if s_type == 6: # msoGroup
								results["groups"] += 1
							elif s_type == 16: # msoMedia
								results["media"] += 1
							elif s_type == 3: # msoChart
								results["charts"] += 1
							elif s_type == 7: # msoEmbeddedOLEObject
								results["ole_objects"] += 1
							elif SlideLayoutAnalyzer._safe_com_bool(shape, "HasTable"):
								results["tables"] += 1
								
							# Word Count
							if SlideLayoutAnalyzer._safe_com_bool(shape, "HasTextFrame"):
								if SlideLayoutAnalyzer._safe_com_bool(shape.TextFrame, "HasText"):
									try:
										slide_word_count += len(shape.TextFrame.TextRange.Text.split())
									except Exception:
										pass
										
						except Exception:
							continue
							
					# Post-process slide stats
					if slide_missing_alt > 0:
						results["missing_alt"].append({"num": slide_num, "count": slide_missing_alt})
						
					if slide_off_canvas:
						results["off_canvas"].append(slide_num)
						
					if slide_word_count > 75:
						results["dense_slides"].append({"num": slide_num, "words": slide_word_count})
						
					# Reading order check
					shape_positions.sort(key=lambda x: x["top"])
					if s_count > 2 and shape_positions:
						top_shapes = shape_positions[:max(1, int(s_count * 0.25))]
						for ts in top_shapes:
							if ts["index"] > (s_count * 0.8):
								results["reading_order_issues"].append(slide_num)
								break
								
				except Exception:
					continue
					
			return results
			
		except Exception as e:
			log.error(f"Document Analyzer COM loop failed: {e}", exc_info=True)
			return results

	@staticmethod
	def _format_and_show_results(results, pres_name):
		if not results:
			return
			
		# Translators: Dialog title for Document Analyzer
		title = _("Document Analysis: {name}").format(name=pres_name)
		
		html = f"<h1>{title}</h1>"
		
		# Overview
		# Translators: Header for Document Overview
		html += f"<h2>{_('Overview')}</h2><ul>"
		# Translators: Total slides in the document
		html += f"<li>{_('Total Slides: {n}').format(n=results['total_slides'])}</li>"
		if results["hidden_slides"] > 0:
			# Translators: Hidden slides in the document
			html += f"<li>{_('Hidden Slides: {n}').format(n=results['hidden_slides'])}</li>"
		html += "</ul>"
		
		# Virtual Table of Contents
		# Translators: Header for Table of Contents
		html += f"<h2>{_('Table of Contents')}</h2><ul>"
		for item in results["toc"]:
			# Translators: TOC entry
			html += f"<li>{_('Slide {num}: {title}').format(num=item['num'], title=item['title'])}</li>"
		html += "</ul>"
		
		# Accessibility Health
		# Translators: Header for Accessibility Health
		html += f"<h2>{_('Accessibility Health')}</h2>"
		has_access_issues = False
		
		if results["reading_order_issues"]:
			has_access_issues = True
			# Translators: Header for Reading Order Mismatch
			html += f"<h3>{_('Reading Order Mismatch (Z-Order vs Visual)')}</h3><ul>"
			for num in results["reading_order_issues"]:
				# Translators: Indicates a slide has reading order issues
				html += f"<li>{_('Slide {num}').format(num=num)}</li>"
			html += "</ul>"
			
		if results["missing_alt"]:
			has_access_issues = True
			# Translators: Header for Missing Alt Text
			html += f"<h3>{_('Missing Alt-Text')}</h3><ul>"
			for item in results["missing_alt"]:
				# Translators: Missing alt text on a specific slide
				html += f"<li>{_('Slide {num}: {count} images missing alt-text').format(num=item['num'], count=item['count'])}</li>"
			html += "</ul>"
			
		if results["off_canvas"]:
			has_access_issues = True
			# Translators: Header for Off-Canvas Elements
			html += f"<h3>{_('Off-Canvas Elements')}</h3><ul>"
			for num in results["off_canvas"]:
				html += f"<li>{_('Slide {num}').format(num=num)}</li>"
			html += "</ul>"
			
		if not has_access_issues:
			html += f"<ul><li>{_('No major accessibility issues detected.')}</li></ul>"
			
		# Density & Complexity
		# Translators: Header for Density & Complexity
		html += f"<h2>{_('Density & Complexity')}</h2>"
		
		if results["dense_slides"]:
			# Translators: Header for dense slides
			html += f"<h3>{_('High Text Density (Wall of Text)')}</h3><ul>"
			for item in results["dense_slides"]:
				# Translators: Dense slide word count
				html += f"<li>{_('Slide {num}: {words} words').format(num=item['num'], words=item['words'])}</li>"
			html += "</ul>"
			
		html += f"<h3>{_('Complex Objects')}</h3><ul>"
		html += f"<li>{_('Tables: {n}').format(n=results['tables'])}</li>"
		html += f"<li>{_('Grouped Objects: {n}').format(n=results['groups'])}</li>"
		html += f"<li>{_('Media: {n}').format(n=results['media'])}</li>"
		html += f"<li>{_('Charts: {n}').format(n=results['charts'])}</li>"
		html += f"<li>{_('Embedded Documents: {n}').format(n=results['ole_objects'])}</li>"
		html += "</ul>"
		
		PowerPointDocumentAnalyzer._show_dialog(title, html)
