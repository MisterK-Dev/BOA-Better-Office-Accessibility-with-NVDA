# -*- coding: UTF-8 -*-
# Copyright (C) 2026 KIRAN G T {MisterK} and Antigravity 2
# This file is covered by the GNU General Public License (GPL), version 2.
# See the file COPYING.txt for more details.

import addonHandler
addonHandler.initTranslation()


class ConditionalFormattingTracker:
	"""
	Analyzes and reads Excel Conditional Formatting rules and active visual results.
	
	Architectural Intent & Considerations:
	NVDA natively reads the raw underlying value of a cell. If a cell turns red because a value is negative, 
	a blind user has no idea the cell is visually flagged. This class intercepts the COM `FormatConditions` 
	collection to decode the rules, and uses `DisplayFormat` to read the final rendered color/font on the screen.
	"""
	@staticmethod
	def _get_color_name(color_int):
		"""
		Converts an Excel COM color integer (BBGGRR) to NVDA's native color name.
		
		Architectural Intent & Considerations:
		Excel outputs raw 24-bit BGR integers. To make this accessible, we extract the RGB bytes using bitwise 
		shifts and map them against NVDA's internal global color dictionary, ensuring the user hears "Red" 
		instead of "255".
		"""
		import colors
		try:
			# Handle strange COM variants
			if color_int is None:
				# Translators: Fallback string when a color cannot be identified.
				return _("Unknown Color")
			
			try:
				color_int = int(color_int)
			except Exception:
				# Translators: Fallback string when a color cannot be identified.
				return _("Unknown Color")
				
			if color_int < 0:
				# Translators: String indicating no value or none.
				return _("None")
			
			# Excel stores colors as BBGGRR
			r = color_int & 255
			g = (color_int >> 8) & 255
			b = (color_int >> 16) & 255
			
			# Hook into NVDA's massive native color dictionary
			return colors.RGB(r, g, b).name
		except Exception:
			# Translators: Fallback string when a color cannot be identified.
			return _("Unknown Color")

	@staticmethod
	def _get_condition_value_str(cv):
		"""
		Extracts human-readable thresholds for data bars and color scales.
		
		Architectural Intent & Considerations:
		Gradient rules (like Color Scales) have dynamic anchors (e.g., "Highest Value" or "50th Percentile") 
		rather than static numbers. This helper queries the exact `ConditionValue` type enum to accurately 
		explain the gradient's anchors.
		"""
		try:
			t = cv.Type
			# Translators: Describes a conditional formatting rule based on the lowest value.
			if t == 1:
				return _("Lowest Value")
			# Translators: Describes a conditional formatting rule based on the highest value.
			if t == 2:
				return _("Highest Value")
			
			val = cv.Value
			# Translators: Describes a conditional formatting rule based on a specific number.
			if t == 0:
				return _("Number {val}").format(val=val)
			# Translators: Describes a conditional formatting rule based on a percentage.
			if t == 3:
				return _("{val}%").format(val=val)
			# Translators: Describes a conditional formatting rule based on a formula.
			if t == 4:
				return _("Formula {val}").format(val=val)
			# Translators: Describes a conditional formatting rule based on a percentile.
			if t == 5:
				return _("{val}th Percentile").format(val=val)
			return str(val)
		except Exception:
			# Translators: Fallback string for unknown conditional formatting criteria.
			return _("Unknown Criteria")

	@staticmethod
	def check_quick_format(excel):
		"""
		Lightweight checker called on every selection change. Returns a brief string if formatting is present.
		
		Architectural Intent & Considerations:
		Fired on EVERY arrow key press. It MUST be incredibly fast. It does not parse the complex rules; it 
		only checks `FormatConditions.Count` and does a quick `DisplayFormat` comparison to see if the cell 
		has changed color. Heavy parsing is deferred to `announce_deep_dive`.
		"""
		from logHandler import log
		try:
			cell = excel.ActiveCell
			if not cell:
				return None
				
			# VT_EMPTY safety check
			try:
				count = cell.FormatConditions.Count
				if count == 0:
					return None
			except Exception:
				return None
				
			# Translators: Initial message indicating the cell has conditional formatting.
			msg = _("Has conditional formatting")
			
			# Check if it actually triggered a visual change
			try:
				df = cell.DisplayFormat
				
				# Check background color
				bg_color = df.Interior.Color
				base_bg = cell.Interior.Color
				if bg_color != base_bg and bg_color != 16777215: # 16777215 is white
					color_name = ConditionalFormattingTracker._get_color_name(bg_color)
					# Translators: Appended to indicate the background color of the formatted cell.
					msg += _(": {color_name} Background").format(color_name=color_name)
				else:
					# Check font color
					font_color = df.Font.Color
					base_font = cell.Font.Color
					if font_color != base_font and font_color != 0: # 0 is black
						color_name = ConditionalFormattingTracker._get_color_name(font_color)
						# Translators: Appended to indicate the text color of the formatted cell.
						msg += _(": {color_name} Text").format(color_name=color_name)
			except Exception:
				# DisplayFormat can sometimes fail on complex merged cells or specific views
				pass
				
			return msg
		except Exception as e:
			log.debug(f"BOA: check_quick_format failed: {e}")
			return None

	@staticmethod
	def announce_deep_dive(excel):
		"""
		Heavy analyzer called only when the user presses NVDA+E, F. Reads all rules and results.
		
		Architectural Intent & Considerations:
		Because Excel has 12 entirely different formatting rule structures (Data Bars, Top 10, Formulas, etc.), 
		they all possess completely different COM properties. This method manually unpacks the specific COM 
		interfaces for each rule type to generate a readable English translation of the logic.
		"""
		import ui
		from logHandler import log
		try:
			cell = excel.ActiveCell
			if not cell:
				# Translators: Error message when analyzing a cell but no cell is active.
				ui.message(_("No active cell found."))
				return
				
			try:
				count = cell.FormatConditions.Count
				if count == 0:
					# Translators: Message when the cell has no conditional formatting rules.
					ui.message(_("No conditional formatting rules applied to this cell."))
					return
			except Exception:
				# Translators: Error message when formatting rules cannot be read.
				ui.message(_("Failed to read format conditions."))
				return
				
			rules_msgs = []
			# Translators: Summary of the number of conditional formatting rules found.
			summary_msg = _("Found {count} conditional formatting rules.").format(count=count)
			
			# 1. READ THE EXACT RULES
			for i in range(1, count + 1):
				try:
					fc = cell.FormatConditions.Item(i)
					rule_type = fc.Type
					
					# Stop if true and Applies To
					# Translators: Indicates that rule evaluation stops if this rule is true.
					stop_str = _(" (Stops if true)") if getattr(fc, "StopIfTrue", False) else ""
					try:
						# Translators: Indicates the range of cells the rule applies to.
						applies_str = _(" [Applies to {address}]").format(address=fc.AppliesTo.Address())
					except Exception:
						applies_str = ""
						
					# Consideration: Excel uses the `xlFormatConditionType` enum to classify rules. 
					# Each rule type requires accessing completely different COM properties (e.g., `xlColorScale` 
					# uses `ColorScaleCriteria`, while `xlCellValue` uses `Operator` and `Formula1`).
					# Mapping of common xlFormatConditionType values:
					# Translators: Generic descriptor for a conditional formatting rule.
					type_str = _("Rule")
					if rule_type == 1: # xlCellValue (Basic cell value thresholds)
						op = fc.Operator
						# Translators: Operators used in conditional formatting rules.
						op_map = {1: _("Between"), 2: _("Not Between"), 3: _("Equal to"), 4: _("Not Equal to"), 5: _("Greater than"), 6: _("Less than"), 7: _("Greater than or equal to"), 8: _("Less than or equal to")}
						# Translators: Fallback operator description.
						op_str = op_map.get(op, _("Operator {op}").format(op=op))
						
						try:
							f1 = fc.Formula1
						except Exception:
							# Translators: Fallback for an unknown formula value.
							f1 = _("Unknown")
							
						if op in [1, 2]:
							try:
								f2 = fc.Formula2
								# Translators: Describes a rule based on two cell values (e.g. between).
								type_str = _("Cell Value is {op} {f1} and {f2}").format(op=op_str, f1=f1, f2=f2)
							except Exception:
								# Translators: Describes a rule based on a single cell value.
								type_str = _("Cell Value is {op} {f1}").format(op=op_str, f1=f1)
						else:
							# Translators: Describes a rule based on a single cell value.
							type_str = _("Cell Value is {op} {f1}").format(op=op_str, f1=f1)
					elif rule_type == 2: # xlExpression (Custom Formula)
						try:
							# Translators: Describes a rule based on a formula.
							type_str = _("Formula: {form}").format(form=fc.Formula1)
						except Exception:
							# Translators: Generic descriptor for a formula-based rule.
							type_str = _("Formula rule")
					elif rule_type == 3: # xlColorScale (2 or 3 color gradients)
						try:
							cc = fc.ColorScaleCriteria
							count_crit = cc.Count
							c_strs = []
							for j in range(1, count_crit + 1):
								crit = cc.Item(j)
								c_name = ConditionalFormattingTracker._get_color_name(crit.FormatColor.Color)
								val_str = ConditionalFormattingTracker._get_condition_value_str(crit)
								# Translators: Describes a point in a color scale gradient.
								c_strs.append(_("Point {num} ({val}) is {color}").format(num=j, val=val_str, color=c_name))
							# Translators: Describes a color scale gradient rule.
							type_str = _("{count}-Color Scale gradient. {points}").format(count=count_crit, points=", ".join(c_strs))
						except Exception:
							# Translators: Generic descriptor for a color scale gradient.
							type_str = _("Color Scale gradient")
					elif rule_type == 4: # xlDatabar (In-cell horizontal bar graphs)
						try:
							c = ConditionalFormattingTracker._get_color_name(fc.BarColor.Color)
							min_val = ConditionalFormattingTracker._get_condition_value_str(fc.MinPoint)
							max_val = ConditionalFormattingTracker._get_condition_value_str(fc.MaxPoint)
							# Translators: Indicates that the cell value is hidden by the rule.
							hide_str = _(" (Note: Cell value is visually hidden)") if getattr(fc, "ShowValue", True) is False else ""
							# Translators: Describes a data bar rule.
							type_str = _("Data Bar gradient ({color}). From {min_val} to {max_val}{hide_str}").format(color=c, min_val=min_val, max_val=max_val, hide_str=hide_str)
						except Exception:
							# Translators: Generic descriptor for a data bar rule.
							type_str = _("Data Bar")
					elif rule_type == 6: # xlIconSet (Traffic lights, flags, arrows)
						try:
							# Translators: Indicates that the cell value is hidden by the icon set rule.
							hide_str = _(" (Note: Cell value is visually hidden)") if getattr(fc, "ShowIconOnly", False) is True else ""
							# Translators: Describes an icon set rule.
							type_str = _("Icon Set ({count} icons){hide_str}").format(count=fc.IconSet.Count, hide_str=hide_str)
						except Exception:
							# Translators: Generic descriptor for an icon set rule.
							type_str = _("Icon Set")
					elif rule_type == 8: # xlUniqueValues
						# Translators: Describes a rule formatting unique or duplicate values.
						type_str = _("Unique or Duplicate Values")
					elif rule_type == 9: # xlTextString
						try:
							# Translators: Describes a rule formatting cells that contain specific text.
							type_str = _("Text contains: '{text}'").format(text=fc.TextString)
						except Exception:
							# Translators: Generic descriptor for a text string rule.
							type_str = _("Text String rule")
					elif rule_type == 10: # xlBlanksCondition
						# Translators: Describes a rule formatting blank cells.
						type_str = _("Format Blank cells")
					elif rule_type == 12: # xlErrorsCondition
						# Translators: Describes a rule formatting error cells.
						type_str = _("Format Error cells")
					elif rule_type == 5: # xlTop10
						try:
							if fc.TopBottom == 1: # xlTop10Top
								# Translators: Describes a rule formatting the top ranked values.
								type_str = _("Top {rank} values").format(rank=fc.Rank)
							else:
								# Translators: Describes a rule formatting the bottom ranked values.
								type_str = _("Bottom {rank} values").format(rank=fc.Rank)
						except Exception:
							# Translators: Generic descriptor for a top/bottom ranking rule.
							type_str = _("Top/Bottom ranking rule")
					else:
						# Translators: Fallback descriptor for an unknown rule type.
						type_str = _("Formatting Rule Type {type}").format(type=rule_type)
					
					# Translators: Assembles the complete description of a rule.
					rules_msgs.append(_("Rule {num}{stop}: {type}.{applies}").format(num=i, stop=stop_str, type=type_str, applies=applies_str))
				except Exception:
					# Translators: Fallback when a rule cannot be parsed.
					rules_msgs.append(_("Rule {num}: Unknown type.").format(num=i))
					
			# 2. READ THE ANALYZED RESULTS
			results_msgs = []
			try:
				df = cell.DisplayFormat
				
				# Background
				if df.Interior.ColorIndex != -4142: # Not xlColorIndexNone
					color_name = ConditionalFormattingTracker._get_color_name(df.Interior.Color)
					# Translators: Describes the background color result of formatting.
					results_msgs.append(_("Background is {color}.").format(color=color_name))
					
				# Font
				# Translators: Describes bold font style.
				font_weight = _("Bold") if df.Font.Bold else ""
				# Translators: Describes italic font style.
				font_italic = _("Italic") if df.Font.Italic else ""
				# Translators: Describes strikethrough font style.
				font_strike = _("Strikethrough") if df.Font.Strikethrough else ""
				
				font_styles = [s for s in [font_weight, font_italic, font_strike] if s]
				if font_styles:
					# Translators: Describes the font styles result of formatting.
					results_msgs.append(_("Font is {styles}.").format(styles=" and ".join(font_styles)))
					
				color_name = ConditionalFormattingTracker._get_color_name(df.Font.Color)
				# Translators: Condition checking for specific text colors.
				if color_name != _("Black") and color_name != _("Unknown Color") and color_name != "Unknown Color":
					# Translators: Describes the text color result of formatting.
					results_msgs.append(_("Text is {color}.").format(color=color_name))
					
				# Number Format
				if df.NumberFormat != cell.NumberFormat:
					# Translators: Describes a change in number format due to conditional formatting.
					results_msgs.append(_("Number format changed to {format}.").format(format=df.NumberFormat))
					
			except Exception:
				# Translators: Error message when final display format cannot be determined.
				results_msgs.append(_("Could not analyze final display format."))
				
			if not results_msgs:
				# Translators: Message when conditional formatting is active but doesn't change visuals.
				results_msgs.append(_("No active visual changes detected."))
				
			# Build HTML content
			html_parts = []
			# Translators: Heading for the rules summary section in the analysis dialog.
			html_parts.append("<h2>" + _("Rules Summary") + "</h2>")
			html_parts.append("<h3>" + summary_msg + "</h3>")
			
			if rules_msgs:
				html_parts.append("<ul>")
				for r_msg in rules_msgs:
					if ":" in r_msg:
						label, val = r_msg.split(":", 1)
						html_parts.append(f"<li><b>{label.strip()}:</b> {val.strip()}</li>")
					else:
						html_parts.append(f"<li>{r_msg}</li>")
				html_parts.append("</ul>")
				
			# Translators: Heading for the analyzed results section in the analysis dialog.
			html_parts.append("<h2>" + _("Analyzed Results") + "</h2>")
			is_fallback = False
			fallback_text = ""
			if len(results_msgs) == 1:
				raw_text = results_msgs[0]
				# Translators: Condition checking for specific result states.
				if raw_text == _("No active visual changes detected.") or raw_text == _("Could not analyze final display format."):
					is_fallback = True
					fallback_text = raw_text
					
			if is_fallback:
				html_parts.append(f"<p>{fallback_text}</p>")
			else:
				html_parts.append("<ul>")
				for res in results_msgs:
					html_parts.append(f"<li>{res}</li>")
				html_parts.append("</ul>")
				
			final_html = "".join(html_parts)
			
			# Using our fully accessible SheetLayoutAnalyzer dialog
			from .sheet_layout_analyzer import SheetLayoutAnalyzer
			# Translators: Title of the conditional formatting analysis dialog.
			SheetLayoutAnalyzer._show_dialog(_("Conditional Formatting Analysis"), final_html)
			
		except Exception as e:
			log.debug(f"BOA: announce_deep_dive failed: {e}")
			# Translators: Error message when conditional formatting analysis fails.
			ui.message(_("Failed to analyze conditional formatting."))
