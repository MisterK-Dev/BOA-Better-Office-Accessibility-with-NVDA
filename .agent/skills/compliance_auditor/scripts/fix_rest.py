import os

fixes = {
    # powerpnt.py and winword.py
    'scriptCategory = _("BOA (Better Office Accessibility)")': '    # Translators: The category name for NVDA gestures belonging to this add-on.\n',
    'script_triggerCommandPrefix.__doc__ = _("Triggers the BOA command prefix for PowerPoint. Press this, followed by a specific command key.")': '    # Translators: Describes the script that triggers the BOA command prefix for PowerPoint.\n',
    'script_triggerCommandPrefix.__doc__ = _("Triggers the BOA command prefix for Word. Press this, followed by a specific command key.")': '    # Translators: Describes the script that triggers the BOA command prefix for Word.\n',

    # boa_gui.py
    'title = _("BOA Office Enhancements")': '        # Translators: Title of the BOA general settings dialog.\n',
    'choices = [_("Off"), _("One-time Announcement"), _("Guided Announcement")]': '        # Translators: Options for the sheet layout announcement mode.\n',
    'cb.SetToolTip(_("Select layout announcement mode. \'Guided Announcement\' requires you to press NVDA+E, L at least once per sheet to scan the layout."))': '        # Translators: Tooltip explaining the layout announcement mode setting.\n',
    'choices = [_("Off"), _("Strict Memory Check (CountA)"), _("Visible Data Only (Math Engine)")]': '        # Translators: Options for the engine used to detect the end of data.\n',
    'cb.SetToolTip(_("Select the engine used to detect the end of data in a row or column."))': '        # Translators: Tooltip explaining the end of data radar engine setting.\n',
    '"grid_mover": _("&Excel: Enable Bulk Sheet Organizer and Quick Sheet Mover"),': '            # Translators: Label for the Bulk Sheet Organizer setting checkbox.\n',
    '"sheet_rename": _("Use accessible Sheet &Rename dialog instead of native edit field"),': '            # Translators: Label for the accessible Sheet Rename setting checkbox.\n',
    '"safe_rich_edit": _("Prevent NVDA &crashes in Excel text fields (SafeRichEdit)"),': '            # Translators: Label for the SafeRichEdit in Excel setting checkbox.\n',
    '"unselect_tracking": _("Announce when a m&ulti-cell selection is unexpectedly lost"),': '            # Translators: Label for the multi-cell selection tracking setting checkbox.\n',
    '"hidden_row_skip": _("Proactively announce when navigating past &hidden rows or columns"),': '            # Translators: Label for the hidden row skip announcement setting checkbox.\n',
    '"sheet_layout_analyzer": _("Enable Sheet &Layout Analyzer via NVDA+E, L"),': '            # Translators: Label for the Sheet Layout Analyzer setting checkbox.\n',
    '"auto_announce_first_block": _("Sheet Layout Auto-Announce &Mode:"),': '            # Translators: Label for the auto-announce mode dropdown setting.\n',
    '"conditional_formatting": _("Conditional &Formatting and color"),': '            # Translators: Label for the conditional formatting tracking setting checkbox.\n',
    '"cell_monitor": _("Enable Cell Moni&tor (slots 1-9 and continuous background monitoring)"),': '            # Translators: Label for the Cell Monitor setting checkbox.\n',
    '"end_of_data_radar": _("Announce when there is no more data in the direction you are &navigating"),': '            # Translators: Label for the End of Data Radar setting checkbox.\n',
    '"formula_auditing_announcements": _("Enable Formula Auditor &features"),': '            # Translators: Label for the Formula Auditor features setting checkbox.\n',
    '"merged_cell_tracking": _("Explicitly announce m&erged cells")': '            # Translators: Label for the merged cell tracking setting checkbox.\n',
    'create_group("excel", _("Excel Enhancements"), excel_features)': '        # Translators: Title of the group box containing Excel specific settings.\n',
    '"standard_color_grid": _("&PowerPoint: Read hidden Hex codes when navigating the Standard Color hexagon grid"),': '            # Translators: Label for the PowerPoint color grid setting checkbox.\n',
    '"hex_edit": _("Ensure the He&x Color edit field is properly labeled"),': '            # Translators: Label for the PowerPoint hex color edit field setting checkbox.\n',
    '"rgb_edit": _("Ensure the R&GB Color edit fields are properly labeled"),': '            # Translators: Label for the PowerPoint RGB color edit field setting checkbox.\n',
    '"safe_rich_edit": _("Prevent NVDA cr&ashes in PowerPoint text fields (SafeRichEdit)")': '            # Translators: Label for the SafeRichEdit in PowerPoint setting checkbox.\n',
    'create_group("powerpoint", _("PowerPoint Enhancements"), ppt_features)': '        # Translators: Title of the group box containing PowerPoint specific settings.\n',
    '"safe_rich_edit": _("&Word: Prevent NVDA crashes in Word text fields (SafeRichEdit)")': '            # Translators: Label for the SafeRichEdit in Word setting checkbox.\n',
    'create_group("word", _("Word Enhancements"), word_features)': '        # Translators: Title of the group box containing Word specific settings.\n',

    # accessible_rename.py
    'dlg = wx.TextEntryDialog(gui.mainFrame, _("Enter new sheet name:"), _("Rename Sheet"), initial_name)': '        # Translators: Dialog text prompting the user to enter a new sheet name.\n',
    'ui.message(_("Renaming to {name}").format(name=clean_name))': '        # Translators: Message announcing the sheet is being renamed.\n',
    'return _("Rename sheet")': '        # Translators: The script description for renaming a sheet.\n',

    # bulk_sheet_organizer.py
    'ui.message(_("Bulk arrangement complete"))': '        # Translators: Success message when all sheet moves are finished.\n',
    'ui.message(_("Error during bulk move: {error}").format(error=e))': '        # Translators: Error message if the bulk sheet move fails.\n',
    'ui.message(_("Could not find Excel grid."))': '                # Translators: Error message when Excel grid cannot be found.\n',
    'ui.message(_("Failed to hook Excel."))': '                # Translators: Error message when the add-on fails to connect to Excel.\n',
    'ui.message(_("No active workbook."))': '                # Translators: Error message when trying to organize sheets but no workbook is active.\n',
    'ui.message(_("Error opening organizer"))': '            # Translators: Error message when the organizer dialog fails to open.\n',
    'super().__init__(parent, title=_("Bulk Sheet Organizer"))': '        # Translators: Title of the Bulk Sheet Organizer dialog.\n',
    'row1.Add(wx.StaticText(self, label=_("Sheet Name:")), 0, wx.ALL|wx.ALIGN_CENTER_VERTICAL, 5)': '        # Translators: Label for the Sheet Name field.\n',
    'row2.Add(wx.StaticText(self, label=_("Target Position:")), 0, wx.ALL|wx.ALIGN_CENTER_VERTICAL, 5)': '        # Translators: Label for the Target Position field.\n',
    'main_sizer.Add(wx.StaticText(self, label=_("Scheduled Moves (Press Del to remove):")), 0, wx.LEFT|wx.TOP, 5)': '        # Translators: Label for the list of scheduled sheet moves.\n',
    'self.list_moves.InsertColumn(0, _("Sheet"), width=150)': '        # Translators: Column header for the sheet name in the moves list.\n',
    'self.list_moves.InsertColumn(1, _("Target Position"), width=100)': '        # Translators: Column header for the target position in the moves list.\n',
    'btn_ok = wx.Button(self, wx.ID_OK, label=_("&OK"))': '        # Translators: Label for the OK button.\n',
    'btn_cancel = wx.Button(self, wx.ID_CANCEL, label=_("&Cancel"))': '        # Translators: Label for the Cancel button.\n',
    'ui.message(_("Scheduled: {sheet} to position {pos}").format(sheet=sheet, pos=pos))': '        # Translators: Message announcing a move has been scheduled.\n',
    'ui.message(_("Move removed"))': '        # Translators: Message announcing a scheduled move has been removed.\n',

    # cell_editor.py
    'ui.message(_("No formula"))': '                # Translators: Message when trying to edit a formula but the cell has none.\n',
    'parent, title=_("BOA Power Editor"), style=wx.DEFAULT_DIALOG_STYLE | wx.RESIZE_BORDER': '        # Translators: Title of the BOA Power Editor dialog.\n',
    'helpLabel = wx.StaticText(self, label=_("&Edit cell content (press Ctrl+Enter to save, Esc to cancel):"))': '        # Translators: Instructions for the Power Editor dialog.\n',
    'saveBtn = wx.Button(self, wx.ID_OK, label=_("&Save"))': '        # Translators: Label for the Save button.\n',
    'cancelBtn = wx.Button(self, wx.ID_CANCEL, label=_("&Cancel"))': '        # Translators: Label for the Cancel button.\n',
    'msg = _("The formula was saved, but Excel evaluated it to an error: {error}\\n\\nWould you like to continue editing?").format(error=resulting_text)': '                # Translators: Message warning that the saved formula resulted in an Excel error.\n',
    '_("Calculation Error"),': '                # Translators: Title of the calculation error warning dialog.\n',
    'ui.message(_("Done"))': '            # Translators: Message confirming the cell edit was saved.\n',
    '_("Excel rejected this formula. Please check for syntax errors or missing parentheses."),': '            # Translators: Error message when Excel rejects a formula due to syntax.\n',
    '_("Invalid Formula"),': '            # Translators: Title of the invalid formula error dialog.\n',

    # cell_monitor.py
    'ui.message(_("Error: Could not read Excel cell."))': '                    # Translators: Error when the add-on cannot read a cell.\n',
    'ui.message(_("{old_cell} has been replaced by {address} for slot {slot_str}").format(': '                # Translators: Message when a monitored slot is overwritten.\n',
    'ui.message(_("{address} set to slot {slot_str}").format(address=address, slot_str=slot_str))': '                # Translators: Message when a cell is successfully slotted.\n',
    'ui.message(_("Slot {slot_str} is empty.").format(slot_str=slot_str))': '            # Translators: Message when checking a slot that is currently empty.\n',
    'ui.message(_("Error: Excel not accessible."))': '            # Translators: Error when Excel cannot be accessed.\n',
    'ui.message(_("Slot {slot_str} lost. Workbook \'{wb}\' is closed or inaccessible.").format(': '                    # Translators: Error when trying to read a slot from a closed workbook.\n',
    'ui.message(_("Slot {slot_str} lost. Sheet \'{sheet}\' was renamed or deleted.").format(': '                    # Translators: Error when trying to read a slot from a deleted sheet.\n',
    "location_str = _(\" in {sheet}\").format(sheet=info['sheet'])": '                # Translators: Describes the sheet location of a monitored cell.\n',
    "location_str = _(\" in {sheet} of {wb}\").format(sheet=info['sheet'], wb=info['wb'])": '                # Translators: Describes the sheet and workbook location of a monitored cell.\n',
    'ui.message(_("Cannot read slot {slot_str}. Excel may be busy.").format(slot_str=slot_str))': '            # Translators: Error when Excel is too busy to read the slot.\n',
    'ui.message(_("Continuous monitor OFF for {address}").format(address=address))': '                # Translators: Message when continuous monitoring is turned off for a cell.\n',
    'ui.message(_("Continuous monitor ON for {address}").format(address=address))': '                # Translators: Message when continuous monitoring is turned on for a cell.\n',
    'ui.message(_("All monitored and slotted cells cleared."))': '        # Translators: Message when all cell monitors are cleared.\n',
    'location_str = _(" in {sheet}").format(sheet=sheet_name)': '            # Translators: Describes the sheet location of a monitored cell that changed.\n',
    'location_str = _(" in {sheet} of {wb}").format(sheet=sheet_name, wb=wb_name)': '            # Translators: Describes the sheet and workbook location of a monitored cell that changed.\n',
    'ui.message(_("{cell_addr} updated: {current_val}{location_str}").format(': '            # Translators: Announces that a continuously monitored cell has updated its value.\n',
    'ui.message(_("Cannot jump. The workbook \'{wb}\' is closed.").format(wb=wb_name))': '            # Translators: Error when trying to jump to a slot in a closed workbook.\n',
    'ui.message(_("Cannot jump. The sheet \'{sheet}\' was renamed or deleted.").format(sheet=sheet_name))': '            # Translators: Error when trying to jump to a slot in a deleted sheet.\n',
    'ui.message(_("Cannot jump. The cell address is invalid or Excel is busy."))': '            # Translators: Error when trying to jump to a slot but the address is invalid.\n',
    'ui.message(_("No cell assigned to slot {slot_num}").format(slot_num=slot_num))': '            # Translators: Error when trying to jump to an empty slot.\n',
    'ui.message(_("No previous cell to jump back to."))': '            # Translators: Error when trying to jump back but there is no history.\n',
    'super().__init__(parent, title=_("Active Cell Monitors"), style=wx.DEFAULT_DIALOG_STYLE | wx.RESIZE_BORDER)': '        # Translators: Title of the active cell monitors dialog.\n',
    'helpLabel = wx.StaticText(self, label=_("Select a cell to jump to it. Press Delete to remove it from monitors."))': '        # Translators: Instructions for the cell monitors dialog.\n',
    'jumpBtn = wx.Button(self, wx.ID_OK, label=_("&Jump"))': '        # Translators: Label for the Jump button.\n',
    'closeBtn = wx.Button(self, wx.ID_CANCEL, label=_("&Close"))': '        # Translators: Label for the Close button.\n',
    'ui.message(_("Slot {slot} deleted").format(slot=key))': '                # Translators: Message when a specific slotted cell is deleted.\n',
    'ui.message(_("Monitor deleted"))': '                # Translators: Message when a continuously monitored cell is deleted.\n',

    # cell_navigation_tracker.py (remaining)
    'ui.message(_("unselected"))': '                        # Translators: Announced when a selection is cleared and nothing is selected.\n',
    'ui.message(_("Merged cell, {address}").format(address=spoken_address))': '                            # Translators: Announced when navigating into a merged cell.\n',
}

files = [
    "addon/appModules/powerpnt.py",
    "addon/appModules/winword.py",
    "addon/appModules/boa_enhancements/boa_gui.py",
    "addon/appModules/boa_enhancements/excel_enhancements/accessible_rename.py",
    "addon/appModules/boa_enhancements/excel_enhancements/bulk_sheet_organizer.py",
    "addon/appModules/boa_enhancements/excel_enhancements/cell_editor.py",
    "addon/appModules/boa_enhancements/excel_enhancements/cell_monitor.py",
    "addon/appModules/boa_enhancements/excel_enhancements/cell_navigation_tracker.py"
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

print("Finished fixing remaining files")
