import xlrd
import xlwt
from xlutils.copy import copy

workbook = xlrd.open_workbook('F://dropshipping//temp1.xls')
# xlwt can not update existing excel file so copy new file to update
new_exl = copy(workbook)
new_exl_sheet = new_exl.get_sheet(0)
worksheet = workbook.sheet_by_index(0)
print(worksheet.name)

totalRows = worksheet.nrows
totalCols = worksheet.ncols

print('total: ' + str(totalRows) + ' records ')
for r in range(0,totalRows):
    
    name = worksheet.cell_value(r,1)
    street = worksheet.cell_value(r,2)
    city = worksheet.cell_value(r,3)
    postcode = worksheet.cell_value(r,4)
    province = worksheet.cell_value(r,5)
    country_english = worksheet.cell_value(r,6)
    phone = worksheet.cell_value(r,7)
    
    #excel文件模版多一行
    i=r+1
    if 'AU' in country_english.upper():
        new_exl_sheet.write(i,8,'澳大利亚')
        new_exl_sheet.write(i,9,'Australia')
    if 'US' in country_english.upper():
        new_exl_sheet.write(i,8,'美国')
        new_exl_sheet.write(i,9,'United States')
    if 'CA' in country_english.upper():
        new_exl_sheet.write(i,8,'加拿大')
        new_exl_sheet.write(i,9,'Canada')

    new_exl_sheet.write(i,3,name)
    new_exl_sheet.write(i,4,street)
    new_exl_sheet.write(i,5,city)
    new_exl_sheet.write(i,6,province)
    new_exl_sheet.write(i,7,postcode)
        
    new_exl_sheet.write(i,10,'垫子')
    new_exl_sheet.write(i,11,'垫子')
    new_exl_sheet.write(i,12,'Ryan N Riley Mats')
    new_exl_sheet.write(i,13,'1')
    new_exl_sheet.write(i,14,'430')
    new_exl_sheet.write(i,15,'3')
    new_exl_sheet.write(i,16,'CN')
    new_exl_sheet.write(i,17,)
       
new_exl.save('F://dropshipping//temp.xls')