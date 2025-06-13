import win32com.client

excel = win32com.client.Dispatch("Excel.Application")
excel.Visible = False

wb = excel.Workbooks.Open(r"C:\Path\To\Workbook.xlsm")
excel.Application.Run("Module1.MyMacro")
wb.Save()
wb.Close(False)
excel.Quit()
