import win32com.client

def test():
    try:
        excel = win32com.client.Dispatch("Excel.Application")
        wb = excel.Workbooks.Add()
        sheet = wb.ActiveSheet
        
        # Hide column E
        sheet.Columns("E:E").EntireColumn.Hidden = True
        
        # Gap range from B to Y
        gap = sheet.Range("B1:Y1")
        print(f"Gap count: {gap.Cells.Count}")
        
        visible = gap.SpecialCells(12)
        print(f"Visible Areas: {visible.Areas.Count}, Address: {visible.Address}")
        
        # What if it's 2 hidden columns, E and G?
        sheet.Columns("G:G").EntireColumn.Hidden = True
        visible2 = gap.SpecialCells(12)
        print(f"Visible Areas 2: {visible2.Areas.Count}, Address: {visible2.Address}")
        
        wb.Close(False)
        excel.Quit()
    except Exception as e:
        print(f"Exception: {e}")

if __name__ == "__main__":
    test()
