import openpyxl

File_Sheet_Row = "$transit_list_location$"
Filename = File_Sheet_Row

xls = openpyxl.load_workbook(Filename, read_only=True, keep_links=False, data_only=True)

print(xls.sheetnames)
