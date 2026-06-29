import re

path = "addon/appModules/boa_enhancements/excel_enhancements/sheet_layout_analyzer.py"

with open(path, "r", encoding="utf-8") as f:
    lines = f.readlines()

new_lines = []

translations = {
    'ui.message(_("Analyzing layout..."))': '        # Translators: Progress message spoken when starting to analyze the sheet layout.\n',
    'html_parts.append("<h2>" + _("Data Blocks Summary") + "</h2>")': '        # Translators: Heading for the section summarizing data blocks in the sheet.\n',
    'html_parts.append("<h3>" + _("Sheet appears to be empty.") + "</h3>")': '            # Translators: Message indicating that the entire sheet has no data.\n',
    'msg = _("Found 1 data block at {c}{r}.").format(c=c, r=r)': '                    # Translators: Message indicating exactly one single-cell data block was found.\n',
    'msg = _("Found 1 data block: {c}{r} to {ec}{er}.").format(c=c, r=r, ec=ec, er=er)': '                    # Translators: Message indicating exactly one multi-cell data block was found.\n',
    'summary = _("Found {count} data blocks in this sheet.").format(count=count)': '                # Translators: Message summarizing the total number of data blocks found.\n',
    'addr = _("{c}{r} to {ec}{er}").format(c=c, r=r, ec=ec, er=er)': '                        # Translators: Formats the start and end coordinates of a data block.\n',
    'block_str = _("Block {num}: {addr}").format(num=i+1, addr=addr)': '                    # Translators: Describes a specific data block by its number and coordinates.\n',
    'SheetLayoutAnalyzer._show_dialog(_("Sheet Layout Overview"), final_html)': '        # Translators: Title of the dialog displaying the sheet layout analysis.\n',
    'msg = _("Empty cell. Nearest data at {col}{row}.").format(col=closest[1], row=closest[0])': '                # Translators: Message spoken when querying an empty cell, pointing to the nearest data.\n',
    'msg = _("Nearest data at {col}{row}").format(col=closest[1], row=closest[0])': '                # Translators: Message pointing to the nearest data from the current cell.\n',
    'props.append(_("Filter Mode: Active"))': '                    # Translators: Indicates that Excel filtering is currently active.\n',
    'if end_r == -1: hidden_borders.append(_("Top 2000+ Rows are hidden"))': '                        # Translators: Indicates that a very large number of top rows are hidden.\n',
    'elif end_r == 1: hidden_borders.append(_("Top Row 1 is hidden"))': '                        # Translators: Indicates that only the very first row is hidden.\n',
    'else: hidden_borders.append(_("Top Rows 1 through {end_r} are hidden").format(end_r=end_r))': '                        # Translators: Indicates a specific range of top rows are hidden.\n',
    'if end_c == -1: hidden_borders.append(_("Left 2000+ Columns are hidden"))': '                        # Translators: Indicates that a very large number of left columns are hidden.\n',
    'elif end_c == 1: hidden_borders.append(_("Left Column A is hidden"))': '                        # Translators: Indicates that only the very first column (A) is hidden.\n',
    'else: hidden_borders.append(_("Left Columns A through {end_col} are hidden").format(end_col=SheetLayoutAnalyzer._col_num_to_letter(end_c)))': '                        # Translators: Indicates a specific range of left columns are hidden.\n',
    'if start_r == -1: hidden_borders.append(_("Bottommost 2000+ Rows are hidden"))': '                        # Translators: Indicates that a very large number of bottom rows are hidden.\n',
    'elif start_r == max_sheet_r: hidden_borders.append(_("Bottom Row {max_r} is hidden").format(max_r=max_sheet_r))': '                        # Translators: Indicates that the very last row in the sheet is hidden.\n',
    'else: hidden_borders.append(_("Bottom Rows {start_r} through {max_r} are hidden").format(start_r=start_r, max_r=max_sheet_r))': '                        # Translators: Indicates a specific range of bottom rows are hidden.\n',
    'if start_c == -1: hidden_borders.append(_("Rightmost 2000+ Columns are hidden"))': '                        # Translators: Indicates that a very large number of right columns are hidden.\n',
    'elif start_c == max_sheet_c: hidden_borders.append(_("Right Column {col} is hidden").format(col=max_c_let))': '                        # Translators: Indicates that the very last column in the sheet is hidden.\n',
    'else: hidden_borders.append(_("Right Columns {start_col} through {end_col} are hidden").format(start_col=SheetLayoutAnalyzer._col_num_to_letter(start_c), end_col=max_c_let))': '                        # Translators: Indicates a specific range of right columns are hidden.\n',
    'props.append(_("Hidden Borders: {borders}").format(borders=", ".join(hidden_borders)))': '                    # Translators: Summarizes all the hidden borders (top, bottom, left, right).\n',
    'parts.append(_("Top sheets 1-{count}").format(count=start_hidden) if start_hidden > 1 else _("Top sheet 1"))': '                        # Translators: Describes how many sheets are hidden at the start of the workbook.\n',
    'parts.append(_("Bottom sheets {start}-{end}").format(start=sheet_count - end_hidden + 1, end=sheet_count) if end_hidden > 1 else _("Bottom sheet {num}").format(num=sheet_count))': '                        # Translators: Describes how many sheets are hidden at the end of the workbook.\n',
    'parts.append(_("{count} middle sheets").format(count=middle_count) if middle_count > 1 else _("{count} middle sheet").format(count=middle_count))': '                        # Translators: Describes how many sheets are hidden in the middle of the workbook.\n',
    'props.append(_("Hidden Sheets: {parts} hidden").format(parts=", ".join(parts)))': '                        # Translators: Summarizes the hidden sheets in the workbook.\n',
    'props.append(_("Sheet Protected: True"))': '                    # Translators: Indicates that the current sheet is protected.\n',
    'props.append(_("Frozen Panes: Rows 1-{r}, Columns A-{c}").format(r=r, c=SheetLayoutAnalyzer._col_num_to_letter(c)))': '                        # Translators: Describes which rows and columns are frozen.\n',
    'props.append(_("Frozen Panes: Rows 1-{r}").format(r=r))': '                        # Translators: Describes which rows are frozen.\n',
    'props.append(_("Frozen Panes: Columns A-{c}").format(c=SheetLayoutAnalyzer._col_num_to_letter(c)))': '                        # Translators: Describes which columns are frozen.\n',
    'props.append(_("Frozen Panes: Active"))': '                        # Translators: Indicates that frozen panes are active but specifics could not be determined.\n',
    'props.append(_("Floating Objects: Contains {count} Shape(s)/Chart(s)").format(count=c))': '                    # Translators: Describes how many floating objects (shapes, charts) are in the sheet.\n',
    'props.append(_("Pivot Tables: Contains {count} Pivot Table(s)").format(count=c))': '                    # Translators: Describes how many pivot tables are in the sheet.\n',
    'props.append(_("View Mode: Page Break Preview"))': '                    # Translators: Indicates the sheet is in Page Break Preview mode.\n',
    'props.append(_("View Mode: Page Layout"))': '                    # Translators: Indicates the sheet is in Page Layout mode.\n',
    'html_props.append("<h2>" + _("Sheet Properties") + "</h2>")': '        # Translators: Heading for the section summarizing overall sheet properties.\n',
    'html_props.append("<p>" + _("No special properties detected (No active filters, protection, frozen panes, or hidden borders).") + "</p>")': '            # Translators: Message shown when the sheet has no special properties.\n',
    'ui.message(_("Layout cache empty. Press NVDA+E, L to scan sheet first."))': '                # Translators: Error message instructing the user to scan the sheet layout before navigating it.\n',
    'ui.message(_("No data blocks found in cache."))': '                # Translators: Error message indicating the sheet has no data blocks to navigate to.\n',
    'ui.message(_("Already on a data cell."))': '                # Translators: Message when trying to jump to nearest data, but already on a data cell.\n',
    'msg = _("Jumped to {col}{row}").format(col=ac_letter, row=ar)': '                # Translators: Success message after jumping to a specific cell.\n',
}

for i, line in enumerate(lines):
    # Check if this line matches any translation exactly
    inserted = False
    for k, v in translations.items():
        if k in line:
            # Check if previous line is already a translator comment
            if i == 0 or "# Translators:" not in lines[i-1]:
                # match indentation
                indent = len(line) - len(line.lstrip())
                new_lines.append(" " * indent + v.strip() + "\\n")
            inserted = True
            break
    new_lines.append(line)

with open(path, "w", encoding="utf-8") as f:
    f.writelines(new_lines)

print("Done")
