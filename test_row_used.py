import win32com.client

def test():
    try:
        excel = win32com.client.Dispatch("Excel.Application")
        excel.Visible = True
        wb = excel.Workbooks.Add()
        sheet = wb.ActiveSheet
        
        sheet.Range("A1").Value = "Data"
        print(f"UsedRange before hiding Row 3: {sheet.UsedRange.Address}")
        
        # Hide Row 3
        sheet.Rows("3:3").EntireRow.Hidden = True
        print(f"UsedRange after hiding Row 3: {sheet.UsedRange.Address}")
        
        wb.Close(False)
        excel.Quit()
    except Exception as e:
        print(f"Exception: {e}")

if __name__ == "__main__":
    test()
