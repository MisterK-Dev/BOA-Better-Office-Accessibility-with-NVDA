import win32com.client

def test():
    try:
        excel = win32com.client.Dispatch("Excel.Application")
        wb = excel.Workbooks.Add()
        sheet = wb.ActiveSheet
        
        gap = sheet.Range(sheet.Cells(1, 5), sheet.Cells(1, 5))
        
        if gap.Cells.Count == 1:
            print("It IS 1")
        else:
            print(f"It is NOT 1! It is {gap.Cells.Count} (type: {type(gap.Cells.Count)})")
            
        wb.Close(False)
        excel.Quit()
    except Exception as e:
        print(f"Exception: {e}")

if __name__ == "__main__":
    test()
