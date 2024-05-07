import xlwings as xw

# Create a new workbook
app = xw.App(visible=True, spec='/Applications/Microsoft Excel.app')

jls_extract_var = xw
wb = xw.Book("/Users/samihabdeljabbar/Desktop/excel_auto/excel_auto/example.xlsx")

# Rest of your code...
sheet = wb.sheets["No"]

sheet["A1"].value = "1000"
