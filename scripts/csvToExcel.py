import csv
import xlrd
import xlwt
from xlutils.copy import copy

# get excel template
workbook = xlrd.open_workbook('F://dropshipping//Standtempl.xls')
# xlwt can not update existing excel file so copy new file to update
new_exl = copy(workbook)
new_exl_sheet = new_exl.get_sheet(0)

orderCsv = 'F://dropshipping//orders_export-2.csv'
with open(orderCsv) as f:
    reader = csv.DictReader(f)
    recordNumber = 0
    for row in reader:
        
        name = row['Shipping Name']
        street = row['Shipping Street']
        city = row['Shipping City']
        postcode = row['Shipping Zip']
        province = row['Shipping Province']
        country_english = row[ 'Shipping Country']
        phone = row['Shipping Phone']
        # excel文件模版多一行
        recordNumber = recordNumber+1
        if 'AU' in country_english.upper():
            new_exl_sheet.write(recordNumber,8,'澳大利亚')
            new_exl_sheet.write(recordNumber,9,'Australia')
        if 'US' in country_english.upper():
            new_exl_sheet.write(recordNumber,8,'美国')
            new_exl_sheet.write(recordNumber,9,'United States')
        if 'CA' in country_english.upper():
            new_exl_sheet.write(recordNumber,8,'加拿大')
            new_exl_sheet.write(recordNumber,9,'Canada')

        new_exl_sheet.write(recordNumber,3,name)
        new_exl_sheet.write(recordNumber,4,street)
        new_exl_sheet.write(recordNumber,5,city)
        new_exl_sheet.write(recordNumber,6,province)
        new_exl_sheet.write(recordNumber,7,postcode)
            
        new_exl_sheet.write(recordNumber,10,'垫子')
        new_exl_sheet.write(recordNumber,11,'垫子')
        new_exl_sheet.write(recordNumber,12,'Ryan N Riley Mats')
        new_exl_sheet.write(recordNumber,13,'1')
        new_exl_sheet.write(recordNumber,14,'430')
        new_exl_sheet.write(recordNumber,15,'3')
        new_exl_sheet.write(recordNumber,16,'CN')
        new_exl_sheet.write(recordNumber,17,)       
new_exl.save('F://dropshipping//temp.xls')