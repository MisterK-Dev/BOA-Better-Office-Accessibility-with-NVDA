import win32com.client

def test():
    try:
        excel = win32com.client.Dispatch("Excel.Application")
        wb = excel.Workbooks.Add()
        sheet = wb.ActiveSheet
        
        gap = sheet.Range(sheet.Cells(1, 5), sheet.Cells(1, 5))
        print(f"Count property: {gap.Count}")
        print(f"Cells.Count property: {gap.Cells.Count}")
        
        wb.Close(False)
        excel.Quit()
    except Exception as e:
        print(f"Exception: {e}")

if __name__ == "__main__":
    test()
