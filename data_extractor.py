import openpyxl

#countries=("Austria", "Europe", "France", "Germany", "Italy", "Spain", "Switzerland", "UK", "USA")
#countries=("Austria","Europe")
#for country in countries:
country="Austria"
wb = openpyxl.load_workbook('C:\\Users\\Michael\\Downloads\\source_file_' + country + '.xlsx',read_only=True, data_only=True)
sheets =  wb.sheetnames
for i in range(0,3):
    sheets.pop(i)
sheets.pop()
print(sheets)
dates=[]
costs=[]
aff_group=[]
country_name=[]
data=[]
sheet_name='Trovit_AT'
sheet = wb[sheet_name]
for row in range(3, 408):
    datum = sheet['A' + str(row)].value
    cost = sheet['G' + str(row)].value
    aff=sheet['A502'].value
    if datum != None:
        dates.append(datum)
        if cost !=None:
            costs.append(cost)
            aff_group.append(aff)
            country_name.append(country)
data=list(zip(dates,costs,aff_group,country_name))
print(data)



wb.close
print('Done')