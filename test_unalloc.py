import win32com.client

def test():
    try:
        excel = win32com.client.Dispatch("Excel.Application")
        wb = excel.Workbooks.Add()
        sheet = wb.ActiveSheet
        
        # Test SpecialCells on unallocated 100-row gap
        gap = sheet.Range("A5:A105")
        
        try:
            visible = gap.SpecialCells(12)
            print(f"Visible Address: {visible.Address}")
        except Exception as e:
            print(f"SpecialCells threw: {e}")
            
        wb.Close(False)
        excel.Quit()
    except Exception as e:
        print(f"Exception: {e}")

if __name__ == "__main__":
    test()
