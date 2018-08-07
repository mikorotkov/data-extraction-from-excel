import openpyxl
import csv
import datetime

now=datetime.datetime.now()
#dictionary for iterating between files 
#countries={9:"Austria", 3:"Europe", 6:"France", 1:"Germany", 11:"Italy", 13:"Spain", 7:"Switzerland", 2:"UK", 12:"USA", 14:"Netherlands", 10:"Belgium"}
countries={9:"Austria"}
print(f'Started at: {now.strftime("%H:%M")}')
for c_id, country in countries.items():

    #path='C:\\Users\\mKorotkov\\Documents\\'
    #input_file='Channel Controlling 2018 '
    #open the workbook
    path='C:\\Users\\Michael\\Downloads\\'
    input_file='source_file_'
    wb = openpyxl.load_workbook(path + input_file + country + '.xlsx',read_only=True, data_only=True)

    sheets =  wb.sheetnames #list of sheet names

    #removing summary, beispiel and data pivot sheets
    i=0
    while i <=1:
        popped = sheets.pop(0)
        i+=1
    sheets.pop()
    #print(sheets)

    #preparing lists for data that will be extracted from each sheet
    dates=[]
    costs=[]
    aff_group=[]
    country_id=[]
    data=[]
    #iterating between the sheets and extracting the data
    for sheet in sheets:
        sheet=wb[sheet]
        sheet_name=sheet[502][0].value
        # extracting only cost and date
        for row in sheet.iter_rows(min_row=3,max_row=408, min_col=1, max_col=7):
            for cell in row:
                if cell.column==1:
                    datum = cell.value
                elif cell.column==7:
                    cost = cell.value
                    #removing blank cells
                    if datum != None:
                        dates.append(datum.strftime("%Y-%m-%d"))
                        if cost !=None:
                            #adding data to list
                            costs.append(cost)
                            aff_group.append(sheet_name)
                            country_id.append(c_id)
        #putting lists together into a list of tuples 
        data=list(zip(dates,aff_group,country_id,costs))
        #print(data)
        cur_date=now.strftime("%Y%m%d")
        
#wiriting into the csv file
        with open(path + cur_date +'_output'+ '.csv', 'w', newline='') as csvfile:
            filewriter = csv.writer(csvfile, delimiter=',',
                                    quotechar='|', quoting=csv.QUOTE_MINIMAL)
            filewriter.writerow(['Date', 'AffiliateGroup', 'CountryId','Cost'])
            for value in data:
                filewriter.writerow(value)


    wb.close
print(f'Ended at: {now.strftime("%H:%M")}')