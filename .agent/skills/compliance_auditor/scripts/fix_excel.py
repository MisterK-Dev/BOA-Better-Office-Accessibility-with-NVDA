import os

fixes = {
    'scriptCategory = _("BOA (Better Office Accessibility)")': '    # Translators: The category name for NVDA gestures belonging to this add-on.\n',
    'script_triggerCommandPrefix.__doc__ = _("Triggers the BOA command prefix. Press this, followed by a specific command key.")': '    # Translators: Describes the script that triggers the BOA command prefix for Excel.\n',
    'script_bulkSheetOrganizer.__doc__ = _("Launches the Bulk Sheet Organizer dialog to reorder sheets.")': '    # Translators: Describes the script that launches the Bulk Sheet Organizer.\n',
    'script_announceSheetLayout.__doc__ = _("Analyzes and announces the layout structure of the current sheet.")': '    # Translators: Describes the script that announces the sheet layout.\n',
    'script_jumpToNearestBlock.__doc__ = _("Jumps the focus to the nearest cell block.")': '    # Translators: Describes the script that jumps focus to the nearest cell block.\n',
    'script_announceConditionalFormatting.__doc__ = _("Announces conditional formatting details for the current cell.")': '    # Translators: Describes the script that announces conditional formatting.\n',
    'script_toggleCellMonitor.__doc__ = _("Toggles monitoring on the active cell.")': '    # Translators: Describes the script that toggles monitoring on a cell.\n',
    'script_clearAllCellMonitors.__doc__ = _("Clears all monitored cells.")': '    # Translators: Describes the script that clears all cell monitors.\n',
    'script_tracePrecedents.__doc__ = _("Traces precedents and opens a dialog to jump to cells that provide data to the active cell.")': '    # Translators: Describes the script that traces formula precedents.\n',
    'script_traceDependents.__doc__ = _("Traces dependents and opens a dialog to jump to cells that depend on the active cell.")': '    # Translators: Describes the script that traces formula dependents.\n',
    'script_speakFormula.__doc__ = _("Announces the raw formula of the active cell.")': '    # Translators: Describes the script that announces the raw cell formula.\n',
    'script_openPowerEditor.__doc__ = _("Opens the accessible Power Editor dialog to edit the active cell\'s contents.")': '    # Translators: Describes the script that opens the accessible Power Editor.\n',
    'script_openMonitorDialog.__doc__ = _("Opens the Active Cell Monitors dialog to manage and jump to monitored cells.")': '    # Translators: Describes the script that opens the Active Cell Monitors dialog.\n',
    'script_jumpBack.__doc__ = _("Jumps focus back to the cell you were working on before navigating away.")': '    # Translators: Describes the script that jumps focus back to the previous cell.\n',
}

file_path = "addon/appModules/excel.py"

if os.path.exists(file_path):
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

print("Finished fixing excel.py")
