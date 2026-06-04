import win32com.client

def test():
    try:
        excel = win32com.client.Dispatch("Excel.Application")
        wb = excel.Workbooks.Add()
        sheet = wb.ActiveSheet
        
        # Test SpecialCells on massive gap
        gap = sheet.Range("A1:A1048576")
        
        try:
            visible = gap.SpecialCells(12)
            print(f"Visible Areas: {visible.Areas.Count}")
        except Exception as e:
            print(f"SpecialCells threw: {e}")
            
        wb.Close(False)
        excel.Quit()
    except Exception as e:
        print(f"Exception: {e}")

if __name__ == "__main__":
    test()
