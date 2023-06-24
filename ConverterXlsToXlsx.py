import openpyxl
import win32com.client as win32
#Converts .xls file to .xlxs file 
# uploadpath - path of the .xls file. downloadpath - path and name where to save .xlsx file
def ConvertXlsToXlsx(uploadpath,downloadpath):
    excel = win32.gencache.EnsureDispatch('Excel.Application')
    wb = excel.Workbooks.Open(uploadpath)
    wb.SaveAs(downloadpath, FileFormat = 51)    #FileFormat = 51 is for .xlsx
    wb.Close()                               
    excel.Application.Quit()

ws = openpyxl.load_workbook("Carriers zone range.xlsx")["UPS zip ranges"]
for ZipCodeNumber in range(2,904):
    zipzone=ws.cell(row = ZipCodeNumber, column = 2).value[:3]#zone for zipcodes for ZipCodeNumbers row 
    ConvertXlsToXlsx(uploadpath=f'xls_files/{zipzone}.xls',downloadpath=f'xlsx_files/{zipzone}.xlsx')
