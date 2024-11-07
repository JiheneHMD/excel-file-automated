from openpyxl import Workbook 
from win32com.client import Dispatch


workbook= Workbook()
sheet=workbook.active

sheet["A1"]="hello"
sheet["B1"]="world"

workbook.save(filename="hello.xlsx")
xl=Dispatch("Excel.Application")
xl.Visible =True

wb = xl.workbooks.Open(r'C:\Users\jihen\Desktop\xml_file_generated_withpython\hello.xlsx')
