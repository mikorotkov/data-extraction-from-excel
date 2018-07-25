import openpyxl

#countries=("Austria", "Europe", "France", "Germany", "Italy", "Spain", "Switzerland", "UK", "USA")
#countries=("Austria","Europe")
#for country in countries:
country="Austria"
wb = openpyxl.load_workbook('C:\\Users\\Michael\\Downloads\\source_file_' + country + '.xlsx',read_only=True, data_only=True)
sheets =  wb.get_sheet_names()
print(sheets)



wb.close
print('Done')