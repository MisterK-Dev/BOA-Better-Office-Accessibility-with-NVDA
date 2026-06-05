"""
Sheet Layout Analyzer

Provides spatial awareness and data block detection for Excel sheets.
WHY THIS EXISTS (Architecture intent):
Native NVDA navigation can leave users lost in empty space if data starts at an unusual row or column.
This module uses COM SpecialCells to quickly detect fragmented data blocks and announce their locations,
giving the user an immediate understanding of the sheet's layout.
"""

import ui
from logHandler import log
import comtypes.client

import wx
import gui
import api

class LayoutDialog(wx.Dialog):
    def __init__(self, parent, title, message):
        super(LayoutDialog, self).__init__(parent, title=title, size=(400, 300))
        
        mainSizer = wx.BoxSizer(wx.VERTICAL)
        
        self.textCtrl = wx.TextCtrl(self, style=wx.TE_MULTILINE | wx.TE_READONLY | wx.HSCROLL)
        self.textCtrl.SetValue(message)
        mainSizer.Add(self.textCtrl, proportion=1, flag=wx.EXPAND | wx.ALL, border=10)
        
        btnSizer = wx.BoxSizer(wx.HORIZONTAL)
        
        self.copyBtn = wx.Button(self, label="Copy (Ctrl+Shift+C)")
        self.copyBtn.Bind(wx.EVT_BUTTON, self.onCopy)
        btnSizer.Add(self.copyBtn, flag=wx.RIGHT, border=10)
        
        # Using &Close automatically binds Alt+C natively in wxPython
        self.closeBtn = wx.Button(self, id=wx.ID_CANCEL, label="&Close")
        self.closeBtn.Bind(wx.EVT_BUTTON, lambda evt: self.EndModal(wx.ID_CANCEL))
        btnSizer.Add(self.closeBtn)
        
        mainSizer.Add(btnSizer, flag=wx.ALIGN_RIGHT | wx.ALL, border=10)
        
        self.SetSizer(mainSizer)
        self.CenterOnParent()
        
        self.Bind(wx.EVT_CHAR_HOOK, self.onCharHook)
        
    def onCharHook(self, evt):
        # Check for Ctrl+Shift+C (Copy)
        if evt.ControlDown() and evt.ShiftDown() and evt.GetKeyCode() == ord('C'):
            self.onCopy(None)
        elif evt.GetKeyCode() == wx.WXK_ESCAPE:
            self.EndModal(wx.ID_CANCEL)
        else:
            evt.Skip()
            
    def onCopy(self, evt):
        api.copyToClip(self.textCtrl.GetValue())
        import ui
        ui.message("Copied to clipboard")
        self.textCtrl.SetFocus()

class SheetLayoutAnalyzer:
    _layout_cache = {}

    @staticmethod
    def _show_dialog(title, message):
        def run():
            gui.mainFrame.prePopup()
            d = LayoutDialog(gui.mainFrame, title, message)
            d.ShowModal()
            d.Destroy()
            gui.mainFrame.postPopup()
        wx.CallAfter(run)

    @staticmethod
    def _col_num_to_letter(n):
        s = ""
        while n > 0:
            n, remainder = divmod(n - 1, 26)
            s = chr(65 + remainder) + s
        return s

    @staticmethod
    def _get_data_areas(excel):
        """
        Extracts the top-left coordinates of all data blocks in the UsedRange.
        Returns a sorted list of (row, col_letter) tuples.
        """
        areas_coords = set()
        from logHandler import log
        try:
            sheet = excel.ActiveSheet
            used_range = sheet.UsedRange
            
            # xlCellTypeConstants = 2, xlCellTypeFormulas = -4123
            for cell_type in (2, -4123):
                try:
                    data_cells = used_range.SpecialCells(cell_type)
                    if data_cells:
                        for i in range(1, data_cells.Areas.Count + 1):
                            area = data_cells.Areas.Item(i)
                            # Collect the top-left and bottom-right cells of the area
                            row = area.Row
                            col = area.Column
                            end_row = row + area.Rows.Count - 1
                            end_col = col + area.Columns.Count - 1
                            areas_coords.add((row, col, end_row, end_col))
                except Exception:
                    # SpecialCells throws an exception if no cells match the type.
                    pass
        except Exception as e:
            log.debug(f"BOA: Failed to get data areas: {e}")

        # Sort top-to-bottom, then left-to-right
        sorted_areas = sorted(list(areas_coords), key=lambda x: (x[0], x[1]))
        
        # Convert column index to letter for the final output
        return [(r, SheetLayoutAnalyzer._col_num_to_letter(c), 
                 er, SheetLayoutAnalyzer._col_num_to_letter(ec)) 
                for r, c, er, ec in sorted_areas]

    @staticmethod
    def announce_layout(excel):
        """
        Manually triggered layout overview. Detects all blocks and displays them in a custom dialog.
        """
        areas = SheetLayoutAnalyzer._get_data_areas(excel)
        try:
            sheet_name = excel.ActiveSheet.Name
            SheetLayoutAnalyzer._layout_cache[sheet_name] = areas
        except Exception:
            pass
            
        if not areas:
            msg = "Sheet appears to be empty."
        else:
            count = len(areas)
            if count == 1:
                r, c, er, ec = areas[0]
                if r == er and c == ec:
                    msg = f"Found 1 data block at {c}{r}."
                else:
                    msg = f"Found 1 data block: {c}{r} to {ec}{er}."
            else:
                block_strings = []
                for r, c, er, ec in areas:
                    if r == er and c == ec:
                        block_strings.append(f"{c}{r}")
                    else:
                        block_strings.append(f"{c}{r} to {ec}{er}")
                        
                blocks_msg = "\n".join([f"Block {i+1}: {addr}" for i, addr in enumerate(block_strings)])
                msg = f"Found {count} data blocks in this sheet.\n{blocks_msg}"
                
        # Append Sheet Properties
        props = SheetLayoutAnalyzer._get_sheet_properties(excel)
        if props:
            msg += props
            
        SheetLayoutAnalyzer._show_dialog("Sheet Layout Overview", msg)

    @staticmethod
    def auto_announce_one_time(excel):
        """
        Instantly calculates the nearest data block and announces it when the user lands on an empty cell
        after opening a workbook or switching sheets. Does not rely on cache.
        """
        from logHandler import log
        import ui
        import core
        try:
            cell = excel.ActiveCell
            r, c_num = cell.Row, cell.Column
            
            is_empty = False
            try:
                val = cell.Value
                text = cell.Text
                if val is None or str(text).strip() == "":
                    is_empty = True
            except Exception:
                is_empty = False

            if not is_empty:
                return

            areas = SheetLayoutAnalyzer._get_data_areas(excel)
            if not areas:
                return

            closest = None
            min_dist = float('inf')
            
            for ar, ac_letter, aer, aec_letter in areas:
                ac = SheetLayoutAnalyzer._letter_to_col_num(ac_letter)
                dist = abs(r - ar) + abs(c_num - ac)
                if dist < min_dist:
                    min_dist = dist
                    closest = (ar, ac_letter)
                    
            if closest:
                msg = f"Empty cell. Nearest data at {closest[1]}{closest[0]}."
                ui.message(msg)
        except Exception as e:
            log.debug(f"BOA: auto_announce_one_time Exception: {e}")

    @staticmethod
    def auto_announce_guided(excel):
        """
        Calculates the Manhattan distance to all known data blocks from the cache and announces 
        the closest block when landing on an empty cell during normal navigation.
        """
        from logHandler import log
        import ui
        import core
        try:
            sheet_name = excel.ActiveSheet.Name
            if sheet_name not in SheetLayoutAnalyzer._layout_cache:
                return
                
            areas = SheetLayoutAnalyzer._layout_cache[sheet_name]
            if not areas:
                return
                
            cell = excel.ActiveCell
            r, c_num = cell.Row, cell.Column
            
            is_empty = False
            try:
                val = cell.Value
                text = cell.Text
                if val is None or str(text).strip() == "":
                    is_empty = True
            except Exception:
                is_empty = False

            if not is_empty:
                return

            closest = None
            min_dist = float('inf')
            
            for ar, ac_letter, aer, aec_letter in areas:
                ac = SheetLayoutAnalyzer._letter_to_col_num(ac_letter)
                dist = abs(r - ar) + abs(c_num - ac)
                if dist < min_dist:
                    min_dist = dist
                    closest = (ar, ac_letter)
                    
            if closest:
                msg = f"Nearest data at {closest[1]}{closest[0]}"
                ui.message(msg)
        except Exception as e:
            log.debug(f"BOA: auto_announce_guided Exception: {e}")

    @staticmethod
    def _get_contiguous_hidden(sheet, start_idx, limit_idx, is_row, step):
        last_hidden = start_idx
        curr = start_idx
        
        # Fast path: check if the rest of the entire sheet is hidden
        try:
            rng = sheet.Range(sheet.Rows(start_idx), sheet.Rows(limit_idx)) if is_row else sheet.Range(sheet.Columns(start_idx), sheet.Columns(limit_idx))
            if getattr(rng, "Hidden", False) is True:
                return limit_idx
        except Exception: pass
            
        count = 0
        while True:
            if (step > 0 and curr > limit_idx) or (step < 0 and curr < limit_idx): break
            try:
                hidden = sheet.Rows(curr).Hidden if is_row else sheet.Columns(curr).Hidden
                if hidden:
                    last_hidden = curr
                else:
                    break
            except Exception:
                break
            curr += step
            count += 1
            if count > 2000: # Safety bailout to prevent NVDA freezing
                return -1 
        return last_hidden

    @staticmethod
    def _get_sheet_properties(excel):
        props = []
        try:
            sheet = excel.ActiveSheet
            wb = excel.ActiveWorkbook
            win = excel.ActiveWindow
            
            # 1. Filter Mode
            try:
                if getattr(sheet, "AutoFilterMode", False):
                    props.append("Filter Mode: Active")
            except Exception: pass
            
            # 2. Hidden Borders
            try:
                ur = sheet.UsedRange
                min_r = ur.Row
                min_c = ur.Column
                max_r = min_r + ur.Rows.Count - 1
                max_c = min_c + ur.Columns.Count - 1
                
                hidden_borders = []
                
                # Check absolute top edge (Row 1)
                try:
                    if sheet.Rows(1).Hidden:
                        end_r = SheetLayoutAnalyzer._get_contiguous_hidden(sheet, 1, sheet.Rows.Count, True, 1)
                        if end_r == -1: hidden_borders.append("Top Rows 1 through 2000+ are hidden")
                        elif end_r == 1: hidden_borders.append("Top Row 1 is hidden")
                        else: hidden_borders.append(f"Top Rows 1 through {end_r} are hidden")
                except Exception: pass
                
                # Check absolute left edge (Col 1)
                try:
                    if sheet.Columns(1).Hidden:
                        end_c = SheetLayoutAnalyzer._get_contiguous_hidden(sheet, 1, sheet.Columns.Count, False, 1)
                        if end_c == -1: hidden_borders.append("Left Columns A through 2000+ are hidden")
                        elif end_c == 1: hidden_borders.append("Left Column A is hidden")
                        else: hidden_borders.append(f"Left Columns A through {SheetLayoutAnalyzer._col_num_to_letter(end_c)} are hidden")
                except Exception: pass
                
                # Check bottom edge of used range
                try:
                    if sheet.Rows(max_r + 1).Hidden:
                        end_r = SheetLayoutAnalyzer._get_contiguous_hidden(sheet, max_r + 1, sheet.Rows.Count, True, 1)
                        if end_r == -1: hidden_borders.append(f"Bottom Rows {max_r + 1} through 2000+ are hidden")
                        elif end_r == max_r + 1: hidden_borders.append(f"Bottom Row {max_r + 1} is hidden")
                        else: hidden_borders.append(f"Bottom Rows {max_r + 1} through {end_r} are hidden")
                except Exception: pass
                
                # Check right edge of used range
                try:
                    if sheet.Columns(max_c + 1).Hidden:
                        end_c = SheetLayoutAnalyzer._get_contiguous_hidden(sheet, max_c + 1, sheet.Columns.Count, False, 1)
                        if end_c == -1: hidden_borders.append(f"Right Columns {SheetLayoutAnalyzer._col_num_to_letter(max_c + 1)} through 2000+ are hidden")
                        elif end_c == max_c + 1: hidden_borders.append(f"Right Column {SheetLayoutAnalyzer._col_num_to_letter(max_c + 1)} is hidden")
                        else: hidden_borders.append(f"Right Columns {SheetLayoutAnalyzer._col_num_to_letter(max_c + 1)} through {SheetLayoutAnalyzer._col_num_to_letter(end_c)} are hidden")
                except Exception: pass
                
                if hidden_borders:
                    props.append("Hidden Borders: " + ", ".join(hidden_borders))
            except Exception: pass
            
            # 3. Hidden Sheets
            try:
                sheet_count = wb.Sheets.Count
                hidden_indices = []
                # Visible = -1 (xlSheetVisible), 0 (xlSheetHidden), 2 (xlSheetVeryHidden)
                for i in range(1, sheet_count + 1):
                    if wb.Sheets(i).Visible != -1:
                        hidden_indices.append(i)
                
                if hidden_indices:
                    start_hidden = 0
                    for i in range(1, sheet_count + 1):
                        if i in hidden_indices: start_hidden += 1
                        else: break
                        
                    end_hidden = 0
                    for i in range(sheet_count, 0, -1):
                        if i in hidden_indices: end_hidden += 1
                        else: break
                        
                    parts = []
                    if start_hidden > 0:
                        parts.append(f"Top sheets 1-{start_hidden}" if start_hidden > 1 else "Top sheet 1")
                    if end_hidden > 0 and (sheet_count - end_hidden >= start_hidden):
                        parts.append(f"Bottom sheets {sheet_count - end_hidden + 1}-{sheet_count}" if end_hidden > 1 else f"Bottom sheet {sheet_count}")
                        
                    middle_count = len(hidden_indices) - start_hidden - end_hidden
                    if middle_count > 0:
                        parts.append(f"{middle_count} middle sheet{'s' if middle_count > 1 else ''}")
                        
                    if parts:
                        props.append("Hidden Sheets: " + ", ".join(parts) + " hidden")
            except Exception: pass
            
            # 4. Sheet Protected
            try:
                if getattr(sheet, "ProtectContents", False):
                    props.append("Sheet Protected: True")
            except Exception: pass
            
            # 5. Frozen Panes
            try:
                if getattr(win, "FreezePanes", False):
                    r = win.SplitRow
                    c = win.SplitColumn
                    if r > 0 and c > 0:
                        props.append(f"Frozen Panes: Rows 1-{r}, Columns A-{SheetLayoutAnalyzer._col_num_to_letter(c)}")
                    elif r > 0:
                        props.append(f"Frozen Panes: Rows 1-{r}")
                    elif c > 0:
                        props.append(f"Frozen Panes: Columns A-{SheetLayoutAnalyzer._col_num_to_letter(c)}")
                    else:
                        props.append("Frozen Panes: Active")
            except Exception: pass
            
            # 6. Floating Objects
            try:
                c = sheet.Shapes.Count
                if c > 0:
                    props.append(f"Floating Objects: Contains {c} Shape(s)/Chart(s)")
            except Exception: pass
            
            # 7. Pivot Tables
            try:
                c = sheet.PivotTables().Count
                if c > 0:
                    props.append(f"Pivot Tables: Contains {c} Pivot Table(s)")
            except Exception: pass
            
            # 8. View Mode
            try:
                v = win.View
                if v == 2:
                    props.append("View Mode: Page Break Preview")
                elif v == 3:
                    props.append("View Mode: Page Layout")
            except Exception: pass
            
        except Exception:
            pass
            
        if props:
            return "\n\n--- Sheet Properties ---\n" + "\n".join(props)
        else:
            return "\n\n--- Sheet Properties ---\nNo special properties detected (No active filters, protection, frozen panes, or hidden borders)."

    @staticmethod
    def jump_to_nearest_block(excel):
        """
        Calculates the Manhattan distance to all known data blocks and natively moves the Excel 
        selection to the closest block.
        """
        from logHandler import log
        import ui
        try:
            sheet_name = excel.ActiveSheet.Name
            if sheet_name not in SheetLayoutAnalyzer._layout_cache:
                ui.message("Layout cache empty. Press NVDA+E, L to scan sheet first.")
                return
                
            areas = SheetLayoutAnalyzer._layout_cache[sheet_name]
            if not areas:
                ui.message("No data blocks found in cache.")
                return
                
            cell = excel.ActiveCell
            r, c_num = cell.Row, cell.Column
            
            # Robust empty cell detection
            is_empty = False
            try:
                val = cell.Value
                text = cell.Text
                # A cell is considered empty if its Value is None, OR its Text is visibly empty
                if val is None or str(text).strip() == "":
                    is_empty = True
            except Exception:
                is_empty = False

            if not is_empty:
                import ui
                ui.message("Already on a data cell.")
                return
                
            closest = None
            min_dist = float('inf')
            
            for ar, ac_letter, aer, aec_letter in areas:
                ac = SheetLayoutAnalyzer._letter_to_col_num(ac_letter)
                # Manhattan distance to top-left cell of the block
                dist = abs(r - ar) + abs(c_num - ac)
                if dist < min_dist:
                    min_dist = dist
                    closest = (ar, ac_letter)
                    
            if closest:
                ar, ac_letter = closest
                excel.ActiveSheet.Cells(ar, SheetLayoutAnalyzer._letter_to_col_num(ac_letter)).Select()
                msg = f"Jumped to {ac_letter}{ar}"
                ui.message(msg)
        except Exception as e:
            log.debug(f"BOA: jump_to_nearest_block Exception: {e}")

    @staticmethod
    def _letter_to_col_num(letter):
        num = 0
        for c in letter:
            num = num * 26 + (ord(c.upper()) - ord('A')) + 1
        return num
