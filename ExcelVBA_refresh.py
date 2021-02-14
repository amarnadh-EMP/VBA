import os
from win32com.client import Dispatch


##Refresh excel
try:
    if os.path.exists("D:\\WK_New\\WorkDay_Daily.xlsm"):
        xl = Dispatch('Excel.Application')
        xl.visible=False
        wb=xl.Workbooks.Open(Filename = "D:\\WK_New\\WorkDay_Daily.xlsm", ReadOnly=1)
        xl.Application.Run("Ref_all")
        xl.DisplayAlerts = False
        wb.Save()
        wb.Close(True)
        xl.quit()
        del xl
        print("Macro refresh completed!")
except Exception:
        #xl = Dispatch('Excel.Application')
        #wb=xl.Workbooks.Open(Filename = "D:\\WK_New\\WorkDay_Daily.xlsm", ReadOnly=1)
        #wb.Close(True)
        #del xl

        print("This is not a valid path!")
     

