import os

fixes = {
    # cell_navigation_tracker.py
    'ui.message(_("{address} selected").format(address=spoken_address))': '                    # Translators: Announced when a cell or range is selected.\n',
    'ui.message(_("Row {row_num} hidden").format(row_num=first_hidden))': '                            # Translators: Announced when navigating past a single hidden row.\n',
    'ui.message(_("Rows {start} through {end} hidden").format(start=min_r + 1, end=max_r - 1))': '                                # Translators: Announced when navigating past multiple hidden rows.\n',
    'ui.message(_("Rows {start} through {end} hidden").format(start=first_hidden, end=last_hidden))': '                            # Translators: Announced when navigating past multiple hidden rows.\n',
    'ui.message(_("Crossed heavily fragmented hidden rows"))': '                            # Translators: Announced when navigating past many non-contiguous hidden rows.\n',
    'ui.message(_("Column {col_letter} hidden").format(col_letter=col_num_to_letter(first_hidden)))': '                            # Translators: Announced when navigating past a single hidden column.\n',
    'ui.message(_("Columns {start} through {end} hidden").format(start=col_num_to_letter(min_c + 1), end=col_num_to_letter(max_c - 1)))': '                                # Translators: Announced when navigating past multiple hidden columns.\n',
    'ui.message(_("Columns {start} through {end} hidden").format(start=col_num_to_letter(first_hidden), end=col_num_to_letter(last_hidden)))': '                            # Translators: Announced when navigating past multiple hidden columns.\n',
    'ui.message(_("Crossed heavily fragmented hidden columns"))': '                            # Translators: Announced when navigating past many non-contiguous hidden columns.\n',
    'ui.message(_("No more data below"))': '                                    # Translators: Announced when trying to jump to edge of data but there is none below.\n',
    'ui.message(_("No more data above"))': '                                    # Translators: Announced when trying to jump to edge of data but there is none above.\n',
    'ui.message(_("No more data to the right"))': '                                    # Translators: Announced when trying to jump to edge of data but there is none to the right.\n',
    'ui.message(_("No more data to the left"))': '                                    # Translators: Announced when trying to jump to edge of data but there is none to the left.\n',
    'ui.message(_("Panes frozen"))': '                    # Translators: Announced when frozen panes are enabled.\n',
    'ui.message(_("Panes unfrozen"))': '                    # Translators: Announced when frozen panes are disabled.\n',
    'ui.message(_("{sheet_name} deleted").format(sheet_name=_last_structural_sheet))': '                    # Translators: Announced when a sheet is deleted.\n',
    'ui.message(_("{sheet} hidden").format(sheet=sheet.Name))': '                            # Translators: Announced when a sheet is hidden.\n',
    'ui.message(_("Sheet hidden"))': '                            # Translators: Announced when a sheet is hidden.\n',
    'ui.message(_("Sheet unhidden"))': '                            # Translators: Announced when a sheet is unhidden.\n',
    'category=_("BOA (Better Office Accessibility)")': '        # Translators: The category name for NVDA gestures belonging to this add-on.\n',
    'address_str = _("Row {num}").format(num=start_row)': '        # Translators: Describes a single row.\n',
    'address_str = _("Rows {start} through {end}").format(start=start_row, end=end_row)': '        # Translators: Describes a range of rows.\n',
    'address_str = _("Column {col}").format(col=start_letter)': '        # Translators: Describes a single column.\n',
    'address_str = _("Columns {start} through {end}").format(start=start_letter, end=end_letter)': '        # Translators: Describes a range of columns.\n',
    'state_str = _("hidden") if is_hidden else _("unhidden")': '        # Translators: Describes the visibility state of rows or columns.\n',
    'ui.message(_("{address} {state}").format(address=address_str, state=state_str))': '        # Translators: Announced when rows or columns visibility state changes.\n',

    # conditional_formatting.py
    'return _("Unknown Color")': '    # Translators: Fallback string when a color cannot be identified.\n',
    'return _("None")': '    # Translators: String indicating no value or none.\n',
    'if t == 1: return _("Lowest Value")': '    # Translators: Describes a conditional formatting rule based on the lowest value.\n',
    'if t == 2: return _("Highest Value")': '    # Translators: Describes a conditional formatting rule based on the highest value.\n',
    'if t == 0: return _("Number {val}").format(val=val)': '    # Translators: Describes a conditional formatting rule based on a specific number.\n',
    'if t == 3: return _("{val}%").format(val=val)': '    # Translators: Describes a conditional formatting rule based on a percentage.\n',
    'if t == 4: return _("Formula {val}").format(val=val)': '    # Translators: Describes a conditional formatting rule based on a formula.\n',
    'if t == 5: return _("{val}th Percentile").format(val=val)': '    # Translators: Describes a conditional formatting rule based on a percentile.\n',
    'return _("Unknown Criteria")': '    # Translators: Fallback string for unknown conditional formatting criteria.\n',
    'msg = _("Has conditional formatting")': '            # Translators: Initial message indicating the cell has conditional formatting.\n',
    'msg += _(": {color_name} Background").format(color_name=color_name)': '                # Translators: Appended to indicate the background color of the formatted cell.\n',
    'msg += _(": {color_name} Text").format(color_name=color_name)': '                # Translators: Appended to indicate the text color of the formatted cell.\n',
    'ui.message(_("No active cell found."))': '            # Translators: Error message when analyzing a cell but no cell is active.\n',
    'ui.message(_("No conditional formatting rules applied to this cell."))': '            # Translators: Message when the cell has no conditional formatting rules.\n',
    'ui.message(_("Failed to read format conditions."))': '            # Translators: Error message when formatting rules cannot be read.\n',
    'summary_msg = _("Found {count} conditional formatting rules.").format(count=count)': '        # Translators: Summary of the number of conditional formatting rules found.\n',
    'stop_str = _(" (Stops if true)") if getattr(fc, "StopIfTrue", False) else ""': '            # Translators: Indicates that rule evaluation stops if this rule is true.\n',
    'applies_str = _(" [Applies to {address}]").format(address=fc.AppliesTo.Address())': '            # Translators: Indicates the range of cells the rule applies to.\n',
    'type_str = _("Rule")': '            # Translators: Generic descriptor for a conditional formatting rule.\n',
    'op_map = {1: _("Between"), 2: _("Not Between"), 3: _("Equal to"), 4: _("Not Equal to"), 5: _("Greater than"), 6: _("Less than"), 7: _("Greater than or equal to"), 8: _("Less than or equal to")}': '            # Translators: Operators used in conditional formatting rules.\n',
    'op_str = op_map.get(op, _("Operator {op}").format(op=op))': '            # Translators: Fallback operator description.\n',
    'f1 = _("Unknown")': '                # Translators: Fallback for an unknown formula value.\n',
    'type_str = _("Cell Value is {op} {f1} and {f2}").format(op=op_str, f1=f1, f2=f2)': '                # Translators: Describes a rule based on two cell values (e.g. between).\n',
    'type_str = _("Cell Value is {op} {f1}").format(op=op_str, f1=f1)': '                # Translators: Describes a rule based on a single cell value.\n',
    'type_str = _("Formula: {form}").format(form=fc.Formula1)': '                # Translators: Describes a rule based on a formula.\n',
    'type_str = _("Formula rule")': '                # Translators: Generic descriptor for a formula-based rule.\n',
    'c_strs.append(_("Point {num} ({val}) is {color}").format(num=j, val=val_str, color=c_name))': '                    # Translators: Describes a point in a color scale gradient.\n',
    'type_str = _("{count}-Color Scale gradient. {points}").format(count=count_crit, points=", ".join(c_strs))': '                # Translators: Describes a color scale gradient rule.\n',
    'type_str = _("Color Scale gradient")': '                # Translators: Generic descriptor for a color scale gradient.\n',
    'hide_str = _(" (Note: Cell value is visually hidden)") if getattr(fc, "ShowValue", True) is False else ""': '            # Translators: Indicates that the cell value is hidden by the rule.\n',
    'type_str = _("Data Bar gradient ({color}). From {min_val} to {max_val}{hide_str}").format(color=c, min_val=min_val, max_val=max_val, hide_str=hide_str)': '                # Translators: Describes a data bar rule.\n',
    'type_str = _("Data Bar")': '                # Translators: Generic descriptor for a data bar rule.\n',
    'hide_str = _(" (Note: Cell value is visually hidden)") if getattr(fc, "ShowIconOnly", False) is True else ""': '            # Translators: Indicates that the cell value is hidden by the icon set rule.\n',
    'type_str = _("Icon Set ({count} icons){hide_str}").format(count=fc.IconSet.Count, hide_str=hide_str)': '                # Translators: Describes an icon set rule.\n',
    'type_str = _("Icon Set")': '                # Translators: Generic descriptor for an icon set rule.\n',
    'type_str = _("Unique or Duplicate Values")': '                # Translators: Describes a rule formatting unique or duplicate values.\n',
    'type_str = _("Text contains: \'{text}\'").format(text=fc.TextString)': '                # Translators: Describes a rule formatting cells that contain specific text.\n',
    'type_str = _("Text String rule")': '                # Translators: Generic descriptor for a text string rule.\n',
    'type_str = _("Format Blank cells")': '                # Translators: Describes a rule formatting blank cells.\n',
    'type_str = _("Format Error cells")': '                # Translators: Describes a rule formatting error cells.\n',
    'type_str = _("Top {rank} values").format(rank=fc.Rank)': '                # Translators: Describes a rule formatting the top ranked values.\n',
    'type_str = _("Bottom {rank} values").format(rank=fc.Rank)': '                # Translators: Describes a rule formatting the bottom ranked values.\n',
    'type_str = _("Top/Bottom ranking rule")': '                # Translators: Generic descriptor for a top/bottom ranking rule.\n',
    'type_str = _("Formatting Rule Type {type}").format(type=rule_type)': '                # Translators: Fallback descriptor for an unknown rule type.\n',
    'rules_msgs.append(_("Rule {num}{stop}: {type}.{applies}").format(num=i, stop=stop_str, type=type_str, applies=applies_str))': '            # Translators: Assembles the complete description of a rule.\n',
    'rules_msgs.append(_("Rule {num}: Unknown type.").format(num=i))': '            # Translators: Fallback when a rule cannot be parsed.\n',
    'results_msgs.append(_("Background is {color}.").format(color=color_name))': '            # Translators: Describes the background color result of formatting.\n',
    'font_weight = _("Bold") if df.Font.Bold else ""': '            # Translators: Describes bold font style.\n',
    'font_italic = _("Italic") if df.Font.Italic else ""': '            # Translators: Describes italic font style.\n',
    'font_strike = _("Strikethrough") if df.Font.Strikethrough else ""': '            # Translators: Describes strikethrough font style.\n',
    'results_msgs.append(_("Font is {styles}.").format(styles=" and ".join(font_styles)))': '                # Translators: Describes the font styles result of formatting.\n',
    'if color_name != _("Black") and color_name != _("Unknown Color") and color_name != "Unknown Color":': '            # Translators: Condition checking for specific text colors.\n',
    'results_msgs.append(_("Text is {color}.").format(color=color_name))': '                # Translators: Describes the text color result of formatting.\n',
    'results_msgs.append(_("Number format changed to {format}.").format(format=df.NumberFormat))': '            # Translators: Describes a change in number format due to conditional formatting.\n',
    'results_msgs.append(_("Could not analyze final display format."))': '            # Translators: Error message when final display format cannot be determined.\n',
    'results_msgs.append(_("No active visual changes detected."))': '            # Translators: Message when conditional formatting is active but doesn\'t change visuals.\n',
    'html_parts.append("<h2>" + _("Rules Summary") + "</h2>")': '        # Translators: Heading for the rules summary section in the analysis dialog.\n',
    'html_parts.append("<h2>" + _("Analyzed Results") + "</h2>")': '        # Translators: Heading for the analyzed results section in the analysis dialog.\n',
    'if raw_text == _("No active visual changes detected.") or raw_text == _("Could not analyze final display format."):': '            # Translators: Condition checking for specific result states.\n',
    'SheetLayoutAnalyzer._show_dialog(_("Conditional Formatting Analysis"), final_html)': '        # Translators: Title of the conditional formatting analysis dialog.\n',
    'ui.message(_("Failed to analyze conditional formatting."))': '        # Translators: Error message when conditional formatting analysis fails.\n',
    
    # Remaining in sheet_layout_analyzer.py with \n
    '# Translators: Summarizes the hidden sheets in the workbook.\\n                        props.append(_("Hidden Sheets: {parts} hidden").format(parts=", ".join(parts)))': '                        # Translators: Summarizes the hidden sheets in the workbook.\n                        props.append(_("Hidden Sheets: {parts} hidden").format(parts=", ".join(parts)))',
    '# Translators: Describes which rows and columns are frozen.\\n                        props.append(_("Frozen Panes: Rows 1-{r}, Columns A-{c}").format(r=r, c=SheetLayoutAnalyzer._col_num_to_letter(c)))': '                        # Translators: Describes which rows and columns are frozen.\n                        props.append(_("Frozen Panes: Rows 1-{r}, Columns A-{c}").format(r=r, c=SheetLayoutAnalyzer._col_num_to_letter(c)))',
    '# Translators: Describes which rows are frozen.\\n                        props.append(_("Frozen Panes: Rows 1-{r}").format(r=r))': '                        # Translators: Describes which rows are frozen.\n                        props.append(_("Frozen Panes: Rows 1-{r}").format(r=r))',
    '# Translators: Describes which columns are frozen.\\n                        props.append(_("Frozen Panes: Columns A-{c}").format(c=SheetLayoutAnalyzer._col_num_to_letter(c)))': '                        # Translators: Describes which columns are frozen.\n                        props.append(_("Frozen Panes: Columns A-{c}").format(c=SheetLayoutAnalyzer._col_num_to_letter(c)))',
    '# Translators: Indicates that frozen panes are active but specifics could not be determined.\\n                        props.append(_("Frozen Panes: Active"))': '                        # Translators: Indicates that frozen panes are active but specifics could not be determined.\n                        props.append(_("Frozen Panes: Active"))',
}

files = [
    "addon/appModules/boa_enhancements/excel_enhancements/cell_navigation_tracker.py",
    "addon/appModules/boa_enhancements/excel_enhancements/conditional_formatting.py",
    "addon/appModules/boa_enhancements/excel_enhancements/sheet_layout_analyzer.py"
]

for file_path in files:
    if not os.path.exists(file_path):
        continue
    with open(file_path, "r", encoding="utf-8") as f:
        lines = f.readlines()
        
    new_lines = []
    for i, line in enumerate(lines):
        inserted = False
        for k, v in fixes.items():
            if k in line:
                if "sheet_layout_analyzer" in file_path and "\\n" in line:
                    # special case to fix the literal \n strings
                    line = line.replace(k, v)
                    new_lines.append(line)
                    inserted = True
                    break
                else:
                    if i == 0 or "# Translators:" not in lines[i-1]:
                        indent = len(line) - len(line.lstrip())
                        new_lines.append(" " * indent + v.lstrip())
                    new_lines.append(line)
                    inserted = True
                    break
        if not inserted:
            new_lines.append(line)
            
    with open(file_path, "w", encoding="utf-8") as f:
        f.writelines(new_lines)
print("Finished fixing all files")
