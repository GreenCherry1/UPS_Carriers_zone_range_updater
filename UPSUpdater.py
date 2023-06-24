import openpyxl
import requests
import xlwt
import win32com.client as win32

#creates empty xls file
def CreateEmptyXlsFile(path):
    wb=xlwt.Workbook()
    wb.add_sheet('1')
    wb.save(path)
    return wb
    
#writes .xls file from the server content in existing file
#zipzone = 3 first numbers of the zipcode
def DownloadXlsFile(zipzone,path,workbook):
    headers = { 'User-Agent': 'Chrome/92.0.4515.107'}
    resp=requests.get(url+zipzone+'.xls',headers=headers)
    open(path,'wb').write(resp.content)
    workbook.save(path)

#Converts .xls file to .xlxs file 
# uploadpath - path of the .xls file. downloadpath - path and name where to save .xlsx file
def ConvertXlsToXlsx(uploadpath,downloadpath):
    excel = win32.gencache.EnsureDispatch('Excel.Application')
    wb = excel.Workbooks.Open(uploadpath)
    wb.SaveAs(downloadpath, FileFormat = 51)    #FileFormat = 51 is for .xlsx
    wb.Close()                               
    excel.Application.Quit()

def UpdateZip(ZipCodeNumber,worksheet):
    ws.cell(row = ZipCodeNumber, column = 2).value=s[39:42]+s[43:45]#copies zip range of downloaded zone file to cell in Carriers zone range
    ws.cell(row = ZipcodeNumber, column = 3).value=s[49:52]+s[53:55]#copies zip range of downloaded zone file to cell in Carriers zone range

url = 'https://www.ups.com/media/us/currentrates/zone-csv/'
wb=openpyxl.load_workbook("Carriers zone range.xlsx")
ws = wb["UPS zip ranges"]
for ZipCodeNumber in range(2,904):
    zipzone=ws.cell(row = ZipCodeNumber, column = 2).value[:3]#value from Carriers zone range
    path=f'D:/UPS_Carriers_zone_range_updater/xls_files/{zipzone}.xls'#Path where file should be saved
    wb=CreateEmptyXlsFile(path)
    DownloadXlsFile(zipzone=zipzone,path=path,workbook=wb)
    ConvertXlsToXlsx(uploadpath=f'xls_files/{zipzone}.xls',downloadpath=f'xlsx_files/{zipzone}.xlsx')
    UpdateZip(ZipcodeNumber=ZipcodeNumber,worksheet=ws)
wb.save("Carriers zone range.xlsx")
