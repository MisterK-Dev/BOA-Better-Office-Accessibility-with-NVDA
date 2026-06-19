# -*- coding: UTF-8 -*-
# Copyright (C) 2026 KIRAN G T {MisterK} and Antigravity 2
# This file is covered by the GNU General Public License (GPL), version 2.
# See the file COPYING.txt for more details.

"""
Excel Formula Auditor Enhancements
Handles safe precedents and dependents navigation, native hotkey hooks,
"""

import wx
import api
import ui
import core
import re
from logHandler import log
import addonHandler
from appModules.boa_enhancements import boa_config
from .cell_navigation_tracker import _get_excel_app
import comtypes

addonHandler.initTranslation()

class AuditingResultDialog(wx.Dialog):
    """
    Dialog to display a list of precedents or dependents.
    """
    def __init__(self, parent, title, items):
        super(AuditingResultDialog, self).__init__(parent, title=title, size=(400, 300))
        self.items = items
        
        sizer = wx.BoxSizer(wx.VERTICAL)
        
        instruction = wx.StaticText(self, label=_("&Select a reference and press Enter to jump to it:"))
        sizer.Add(instruction, 0, wx.ALL, 5)
        
        self.list_box = wx.ListBox(self, choices=[item["display"] for item in self.items])
        sizer.Add(self.list_box, 1, wx.EXPAND | wx.ALL, 5)
        
        btn_sizer = wx.StdDialogButtonSizer()
        
        self.jump_btn = wx.Button(self, wx.ID_OK, label=_("&Jump"))
        self.jump_btn.SetDefault()
        btn_sizer.AddButton(self.jump_btn)
        
        self.cancel_btn = wx.Button(self, wx.ID_CANCEL, label=_("&Cancel"))
        btn_sizer.AddButton(self.cancel_btn)
        btn_sizer.Realize()
        
        sizer.Add(btn_sizer, 0, wx.EXPAND | wx.ALL, 5)
        self.SetSizer(sizer)
        
        self.list_box.Bind(wx.EVT_LISTBOX_DCLICK, self.on_jump)
        self.Bind(wx.EVT_BUTTON, self.on_jump, id=wx.ID_OK)
        self.Bind(wx.EVT_BUTTON, self.on_cancel, id=wx.ID_CANCEL)
        self.Bind(wx.EVT_CLOSE, self.on_close)
        
        if self.items:
            self.list_box.SetSelection(0)
            self.list_box.SetFocus()
            
    def on_jump(self, event):
        sel = self.list_box.GetSelection()
        if sel != wx.NOT_FOUND:
            item = self.items[sel]
            if item.get("is_broken"):
                import ui
                ui.message(_("Cannot jump to a closed external workbook."))
                return
            try:
                sheet_name = item["sheet"]
                addr = item["address"]
                wb_name = item.get("workbook")
                
                # Grab a fresh COM object on the Main Thread!
                from .cell_navigation_tracker import _get_excel_app
                excel_app = _get_excel_app()
                if not excel_app: return
                
                # Jump safely across sheets/workbooks
                if wb_name:
                    excel_app.Workbooks(wb_name).Activate()
                excel_app.Worksheets(sheet_name).Activate()
                excel_app.ActiveSheet.Range(addr).Select()
            except Exception as e:
                log.error("BOA: Failed to jump to reference: {}".format(e))
        self.EndModal(wx.ID_OK)

    def on_cancel(self, event):
        self.EndModal(wx.ID_CANCEL)

    def on_close(self, event):
        self.EndModal(wx.ID_CANCEL)

def _extract_closed_links(formula):
    """
    Scans the formula string for external workbook references.
    Example: ='C:\Path\[Workbook.xlsx]SheetName'!A1
    """
    links = []
    if not formula or not formula.startswith("="):
        return links
    
    # Regex matches external references with or without paths
    pattern = r"(?:'[^']+'|\[[^\]]+\][^!]+)![A-Za-z0-9\$]+"
    matches = set(re.findall(pattern, formula))
    for m in matches:
        links.append({
            "display": _("(External) {}").format(m),
            "sheet": None,
            "address": m,
            "workbook": None,
            "is_broken": True
        })
    return links

def _prop(obj):
    """
    Safely evaluates a comtypes NamedProperty or returns the object if it's not dynamic.
    """
    if not obj: return obj
    try:
        if "NamedProperty" in str(type(obj)) or "MethodCaller" in str(type(obj)):
            return obj()
    except Exception:
        pass
    return obj

def _gather_references(is_precedent, excel, original_cell, original_sheet, orig_sheet_name, skip_draw=False):
    """
    Gathers references by safely looping NavigateArrow.
    Catches COM errors and guarantees ScreenUpdating and Selection restoration.
    """
    import comtypes
    items = []
    try:
        original_addr = _prop(original_cell.Address)
        try:
            active_wb = _prop(excel.ActiveWorkbook)
            original_wb = _prop(active_wb.Name)
        except Exception:
            original_wb = ""
            
        is_prec_val = 1 if is_precedent else 0
        
        arrow_num = 1
        while True:
            link_num = 1
            arrow_found = False
            
            while True:
                try:
                    # Reactivate the origin sheet before navigating, or Excel loses tracer context across workbooks!
                    original_sheet.Parent.Activate()
                    original_sheet.Activate()
                    
                    target = original_cell.NavigateArrow(is_prec_val, arrow_num, link_num)
                    
                    target_addr = _prop(target.Address)
                    target_parent = _prop(target.Parent)
                    target_sheet_name = _prop(target_parent.Name)
                    
                    # If target is original, we've exhausted links for this arrow
                    if target_addr == original_addr and target_sheet_name == orig_sheet_name:
                        break
                        
                    arrow_found = True
                    
                    val_str = ""
                    try:
                        val = _prop(target.Value)
                        if val is not None:
                            if isinstance(val, tuple):
                                val_str = " - <Multiple Values>"
                            else:
                                val_str = " - {}".format(val)
                    except comtypes.COMError:
                        pass
                        
                    wb_name = ""
                    try:
                        target_parent_parent = _prop(target_parent.Parent)
                        wb_name = _prop(target_parent_parent.Name)
                    except comtypes.COMError:
                        pass
                        
                    wb_display = "[{}]".format(wb_name) if wb_name and wb_name != original_wb else ""
                    
                    display = "{wb}{sheet}!{addr}{val}".format(
                        wb=wb_display, 
                        sheet=target_sheet_name, 
                        addr=target_addr.replace("$", ""), 
                        val=val_str
                    )
                    
                    # Deduplicate items
                    is_dup = False
                    for existing in items:
                        if existing["address"] == target_addr and existing["sheet"] == target_sheet_name and existing["workbook"] == wb_name:
                            is_dup = True
                            break
                            
                    if not is_dup:
                        items.append({
                            "display": display,
                            "sheet": target_sheet_name,
                            "address": target_addr,
                            "workbook": wb_name,
                            "type": "native"
                        })
                        
                    link_num += 1
                    
                except comtypes.COMError as e:
                    log.debug("BOA: Loop finished or error in NavigateArrow: {}".format(e))
                    break # Break inner loop
            
            if not arrow_found:
                break # All arrows exhausted
            
            arrow_num += 1
                
        # Append broken external links if we are checking precedents
        if is_precedent:
            try:
                formula = _prop(original_cell.Formula)
                if formula:
                    closed_links = _extract_closed_links(formula)
                    for c in closed_links:
                        # Dedup against native NavigateArrow items
                        is_dup = False
                        for existing in items:
                            # Normalize addresses to prevent A1 vs $A$1 mismatches
                            if (existing["workbook"].lower() == c["workbook"].lower() and 
                                existing["sheet"].lower() == c["sheet"].lower() and 
                                existing["address"].replace('$', '') == c["address"].replace('$', '')):
                                is_dup = True
                                break
                        if not is_dup:
                            items.append(c)
            except comtypes.COMError:
                pass
                
    except Exception as e:
        log.error("BOA: Fatal error gathering references: {}".format(e))
    finally:
        # STRICT CLEANUP
        if not skip_draw:
            try:
                if is_precedent:
                    original_cell.ShowPrecedents(Remove=True)
                else:
                    original_cell.ShowDependents(Remove=True)
            except Exception:
                pass
            
        try:
            orig_parent = _prop(original_sheet.Parent)
            orig_parent.Activate()
            original_sheet.Activate()
            original_sheet.Range(original_addr).Select()
        except comtypes.COMError:
            pass
            
    return items

def _show_auditing_dialog(is_precedent, skip_native_draw=False):
    import threading
    import comtypes
    from .cell_navigation_tracker import _get_excel_app
    
    # Grab the exact XLMAIN HWND from the main thread's cached COM pointer!
    main_excel = _get_excel_app()
    if not main_excel:
        return
        
    try:
        target_hwnd = main_excel.Hwnd
    except Exception as e:
        log.error("BOA: Failed to get target hwnd: {}".format(e))
        return
        
    def worker(xlmain_hwnd):
        try:
            comtypes.CoInitialize()
        except Exception:
            pass
            
        def _do_work():
            # 1. Thread-Local COM Connection: Bypass cached _get_excel_app to prevent RPC_E_WRONG_THREAD
            import comtypes.client
            import comtypes.automation
            import ctypes
            
            excel = None
            
            # Drill down from the EXACT xlmain_hwnd to find its specific EXCEL7 window
            try:
                hwnd7 = None
                if xlmain_hwnd:
                    xldesk = ctypes.windll.user32.FindWindowExW(xlmain_hwnd, 0, "XLDESK", None)
                    if xldesk:
                        hwnd7 = ctypes.windll.user32.FindWindowExW(xldesk, 0, "EXCEL7", None)
                
                if hwnd7:
                    oleacc = ctypes.windll.oleacc if hasattr(ctypes.windll, 'oleacc') else ctypes.windll.user32.oleacc
                    ptr = ctypes.POINTER(comtypes.automation.IDispatch)()
                    res = oleacc.AccessibleObjectFromWindow(hwnd7, -16, ctypes.byref(comtypes.automation.IDispatch._iid_), ctypes.byref(ptr))
                    if res == 0 and ptr:
                        import comtypes.client.dynamic
                        excel = comtypes.client.dynamic.Dispatch(ptr).Application
            except Exception as e:
                log.error("BOA: Thread fallback COM retrieval failed: {}".format(e))
                
            if not excel:
                log.error("BOA: Failed to get thread-local Excel COM object.")
                return
                
            try:
                cell = _prop(excel.ActiveCell)
                sheet = _prop(excel.ActiveSheet)
                
                if not cell or not sheet:
                    log.error("BOA: Failed to retrieve ActiveCell or ActiveSheet (Excel might be busy parsing the ribbon command).")
                    return
                    
                orig_sheet_name = _prop(sheet.Name)
                
                # Turn off screen updating so NavigateArrow doesn't flash the screen!
                try:
                    excel.ScreenUpdating = False
                except Exception:
                    pass
                    
                if not skip_native_draw:
                    try:
                        sheet.ClearArrows()
                        if is_precedent:
                            cell.ShowPrecedents()
                        else:
                            cell.ShowDependents()
                    except comtypes.COMError as e:
                        log.debug("BOA: ShowPrecedents/Dependents failed natively (likely closed links): {}".format(e))
                        pass
                    
                try:
                    items = _gather_references(is_precedent, excel, cell, sheet, orig_sheet_name, skip_draw=skip_native_draw)
                except comtypes.COMError as e:
                    log.error("BOA: Gather references failed: {}".format(e))
                    items = []
                
                # Fallback to DirectPrecedents/Dependents for same-sheet accuracy if NavigateArrow fails
                try:
                    if is_precedent:
                        direct = _prop(cell.DirectPrecedents)
                    else:
                        direct = _prop(cell.DirectDependents)
                        
                    areas = _prop(direct.Areas)
                    count = _prop(areas.Count)
                    for i in range(1, count + 1):
                        area = _prop(areas.Item(i))
                        val_str = ""
                        try:
                            v = _prop(area.Value)
                            if v is not None:
                                if isinstance(v, tuple):
                                    val_str = " - <Multiple Values>"
                                else:
                                    val_str = " - {}".format(v)
                        except Exception:
                            pass
                            
                        area_addr = _prop(area.Address)
                        display = "{sheet}!{addr}{val}".format(
                            sheet=orig_sheet_name,
                            addr=area_addr.replace('$', ''),
                            val=val_str
                        )
                        
                        is_dup = False
                        for existing in items:
                            if existing["address"] == area_addr and existing["sheet"] == orig_sheet_name:
                                is_dup = True
                                break
                                
                        if not is_dup:
                            items.append({
                                "display": display,
                                "sheet": orig_sheet_name,
                                "address": area_addr,
                                "workbook": "",
                                "type": "direct_fallback"
                            })
                except Exception as e:
                    log.debug("BOA: Direct fallback failed or no references found: {}".format(e))
                    
                def _show_gui():
                    import gui
                    import ui
                    try:
                        if not items:
                            ui.message(_("No precedents found.") if is_precedent else _("No dependents found."))
                            return
                            
                        gui.mainFrame.prePopup()
                        title = _("Precedents") if is_precedent else _("Dependents")
                        
                        # Parent to gui.mainFrame, use ShowModal, and clean up safely
                        dlg = AuditingResultDialog(gui.mainFrame, title, items)
                        dlg.Raise()
                        dlg.ShowModal()
                        dlg.Destroy()
                        gui.mainFrame.postPopup()
                    except Exception as e:
                        log.error("BOA: Failed to show auditing dialog: {}".format(e))
                
                import gui.guiHelper
                gui.guiHelper.wxCallOnMain(_show_gui)
                
            except Exception as e:
                log.error("BOA: Thread error in auditing logic: {}".format(e))
            finally:
                try:
                    if excel:
                        excel.ScreenUpdating = True
                except Exception:
                    pass
                
                # Explicitly delete COM objects so they are garbage collected BEFORE CoUninitialize
                try:
                    del excel
                    del cell
                    del sheet
                except Exception:
                    pass

        try:
            _do_work()
        except Exception as e:
            log.error("BOA: Thread wrapper error in auditing: {}".format(e))
        finally:
            try:
                comtypes.CoUninitialize()
            except Exception:
                pass
                
    t = threading.Thread(target=worker, args=(target_hwnd,))
    t.daemon = True
    t.start()

def show_precedents_dialog(obj):
    _show_auditing_dialog(True)

def show_dependents_dialog(obj):
    _show_auditing_dialog(False)

