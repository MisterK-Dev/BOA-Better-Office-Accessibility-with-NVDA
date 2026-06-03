import win32com.client

def test():
    try:
        excel = win32com.client.Dispatch("Excel.Application")
        wb = excel.Workbooks.Add()
        sheet = wb.ActiveSheet
        
        # Test 1D column range (1 row high, 3 columns wide)
        gap = sheet.Range("C1:E1")
        print(f"Gap count: {gap.Cells.Count}")
        
        visible = gap.SpecialCells(12)
        print(f"Visible Areas: {visible.Areas.Count}, Address: {visible.Address}")
        
        # Now what if it's 2 cells?
        gap2 = sheet.Range("C1:D1")
        visible2 = gap2.SpecialCells(12)
        print(f"Gap2 visible address: {visible2.Address}")
        
        wb.Close(False)
        excel.Quit()
    except Exception as e:
        print(f"Exception: {e}")

if __name__ == "__main__":
    test()
