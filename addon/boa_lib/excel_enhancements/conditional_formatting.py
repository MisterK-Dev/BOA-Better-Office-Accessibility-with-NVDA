import math

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
                return "Unknown Color"
            
            try:
                color_int = int(color_int)
            except Exception:
                return "Unknown Color"
                
            if color_int < 0:
                return "None"
            
            # Excel stores colors as BBGGRR
            r = color_int & 255
            g = (color_int >> 8) & 255
            b = (color_int >> 16) & 255
            
            # Hook into NVDA's massive native color dictionary
            return colors.RGB(r, g, b).name
        except Exception:
            return "Unknown Color"

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
            if t == 1: return "Lowest Value"
            if t == 2: return "Highest Value"
            
            val = cv.Value
            if t == 0: return f"Number {val}"
            if t == 3: return f"{val}%"
            if t == 4: return f"Formula {val}"
            if t == 5: return f"{val}th Percentile"
            return str(val)
        except Exception:
            return "Unknown Criteria"

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
                
            msg = "Has conditional formatting"
            
            # Check if it actually triggered a visual change
            try:
                df = cell.DisplayFormat
                
                # Check background color
                bg_color = df.Interior.Color
                base_bg = cell.Interior.Color
                if bg_color != base_bg and bg_color != 16777215: # 16777215 is white
                    color_name = ConditionalFormattingTracker._get_color_name(bg_color)
                    msg += f": {color_name} Background"
                else:
                    # Check font color
                    font_color = df.Font.Color
                    base_font = cell.Font.Color
                    if font_color != base_font and font_color != 0: # 0 is black
                        color_name = ConditionalFormattingTracker._get_color_name(font_color)
                        msg += f": {color_name} Text"
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
                ui.message("No active cell found.")
                return
                
            try:
                count = cell.FormatConditions.Count
                if count == 0:
                    ui.message("No conditional formatting rules applied to this cell.")
                    return
            except Exception:
                ui.message("Failed to read format conditions.")
                return
                
            rules_msgs = []
            rules_msgs.append(f"Found {count} conditional formatting rules.")
            
            # 1. READ THE EXACT RULES
            for i in range(1, count + 1):
                try:
                    fc = cell.FormatConditions.Item(i)
                    rule_type = fc.Type
                    
                    # Stop if true and Applies To
                    stop_str = " (Stops if true)" if getattr(fc, "StopIfTrue", False) else ""
                    try:
                        applies_str = f" [Applies to {fc.AppliesTo.Address()}]"
                    except Exception:
                        applies_str = ""
                        
                    # Consideration: Excel uses the `xlFormatConditionType` enum to classify rules. 
                    # Each rule type requires accessing completely different COM properties (e.g., `xlColorScale` 
                    # uses `ColorScaleCriteria`, while `xlCellValue` uses `Operator` and `Formula1`).
                    # Mapping of common xlFormatConditionType values:
                    type_str = "Rule"
                    if rule_type == 1: # xlCellValue (Basic cell value thresholds)
                        op = fc.Operator
                        op_map = {1: "Between", 2: "Not Between", 3: "Equal to", 4: "Not Equal to", 5: "Greater than", 6: "Less than", 7: "Greater than or equal to", 8: "Less than or equal to"}
                        op_str = op_map.get(op, "Operator " + str(op))
                        
                        try:
                            f1 = fc.Formula1
                        except Exception:
                            f1 = "Unknown"
                            
                        if op in [1, 2]:
                            try:
                                f2 = fc.Formula2
                                type_str = f"Cell Value is {op_str} {f1} and {f2}"
                            except Exception:
                                type_str = f"Cell Value is {op_str} {f1}"
                        else:
                            type_str = f"Cell Value is {op_str} {f1}"
                    elif rule_type == 2: # xlExpression (Custom Formula)
                        try:
                            type_str = f"Formula: {fc.Formula1}"
                        except Exception:
                            type_str = "Formula rule"
                    elif rule_type == 3: # xlColorScale (2 or 3 color gradients)
                        try:
                            cc = fc.ColorScaleCriteria
                            count_crit = cc.Count
                            c_strs = []
                            for j in range(1, count_crit + 1):
                                crit = cc.Item(j)
                                c_name = ConditionalFormattingTracker._get_color_name(crit.FormatColor.Color)
                                val_str = ConditionalFormattingTracker._get_condition_value_str(crit)
                                c_strs.append(f"Point {j} ({val_str}) is {c_name}")
                            type_str = f"{count_crit}-Color Scale gradient. " + ", ".join(c_strs)
                        except Exception:
                            type_str = "Color Scale gradient"
                    elif rule_type == 4: # xlDatabar (In-cell horizontal bar graphs)
                        try:
                            c = ConditionalFormattingTracker._get_color_name(fc.BarColor.Color)
                            min_val = ConditionalFormattingTracker._get_condition_value_str(fc.MinPoint)
                            max_val = ConditionalFormattingTracker._get_condition_value_str(fc.MaxPoint)
                            hide_str = " (Note: Cell value is visually hidden)" if getattr(fc, "ShowValue", True) is False else ""
                            type_str = f"Data Bar gradient ({c}). From {min_val} to {max_val}{hide_str}"
                        except Exception:
                            type_str = "Data Bar"
                    elif rule_type == 6: # xlIconSet (Traffic lights, flags, arrows)
                        try:
                            hide_str = " (Note: Cell value is visually hidden)" if getattr(fc, "ShowIconOnly", False) is True else ""
                            type_str = f"Icon Set ({fc.IconSet.Count} icons){hide_str}"
                        except Exception:
                            type_str = "Icon Set"
                    elif rule_type == 8: # xlUniqueValues
                        type_str = "Unique or Duplicate Values"
                    elif rule_type == 9: # xlTextString
                        try:
                            type_str = f"Text contains: '{fc.TextString}'"
                        except Exception:
                            type_str = "Text String rule"
                    elif rule_type == 10: # xlBlanksCondition
                        type_str = "Format Blank cells"
                    elif rule_type == 12: # xlErrorsCondition
                        type_str = "Format Error cells"
                    elif rule_type == 5: # xlTop10
                        try:
                            if fc.TopBottom == 1: # xlTop10Top
                                type_str = f"Top {fc.Rank} values"
                            else:
                                type_str = f"Bottom {fc.Rank} values"
                        except Exception:
                            type_str = "Top/Bottom ranking rule"
                    else:
                        type_str = f"Formatting Rule Type {rule_type}"
                    
                    rules_msgs.append(f"Rule {i}{stop_str}: {type_str}.{applies_str}")
                except Exception as e:
                    rules_msgs.append(f"Rule {i}: Unknown type.")
                    
            # 2. READ THE ANALYZED RESULTS
            results_msgs = []
            try:
                df = cell.DisplayFormat
                
                # Background
                if df.Interior.ColorIndex != -4142: # Not xlColorIndexNone
                    color_name = ConditionalFormattingTracker._get_color_name(df.Interior.Color)
                    results_msgs.append(f"Background is {color_name}.")
                    
                # Font
                font_weight = "Bold" if df.Font.Bold else ""
                font_italic = "Italic" if df.Font.Italic else ""
                font_strike = "Strikethrough" if df.Font.Strikethrough else ""
                
                font_styles = [s for s in [font_weight, font_italic, font_strike] if s]
                if font_styles:
                    results_msgs.append("Font is " + " and ".join(font_styles) + ".")
                    
                color_name = ConditionalFormattingTracker._get_color_name(df.Font.Color)
                if color_name != "Black" and color_name != "Unknown Color":
                    results_msgs.append(f"Text is {color_name}.")
                    
                # Number Format
                if df.NumberFormat != cell.NumberFormat:
                    results_msgs.append(f"Number format changed to {df.NumberFormat}.")
                    
            except Exception:
                results_msgs.append("Could not analyze final display format.")
                
            if not results_msgs:
                results_msgs.append("No active visual changes detected.")
                
            final_msg = "\n".join(rules_msgs) + "\n\nAnalyzed Results:\n" + "\n".join(results_msgs)
            
            # Using our fully accessible SheetLayoutAnalyzer dialog
            from .sheet_layout_analyzer import SheetLayoutAnalyzer
            SheetLayoutAnalyzer._show_dialog("Conditional Formatting Analysis", final_msg)
            
        except Exception as e:
            log.debug(f"BOA: announce_deep_dive failed: {e}")
            ui.message("Failed to analyze conditional formatting.")
