# Tutorial 1: Create a simple XLSX file
# Tutorial 2: Adding formatting to the XLSX File
# Tutorial3 : Adding different types of data types
#different types of data  in date validation

import xlsxwriter

import datetime
from datetime import datetime

report_card = (

        ['manish', 77,'2011-01-22',2122],
        ['suresh', 78,'2011-01-12',4325],
        ['kamapisachi',83,'2011-01-8',5645],

    )

workbook = xlsxwriter.Workbook('student.xlsx')
worksheet = workbook.add_worksheet()
currency_format = workbook.add_format({'num_format': '$#,##0.00'})
bold = workbook.add_format({'bold': True})


date_format =  workbook.add_format({'num_format': 'mmmm d yyyy'})
#money_format = workbook.add_format({'num_format': '$#,##0'})
cell_format1 = workbook.add_format({'bold':True,'italics':True})
cell_format =  workbook.add_format({'bold':True})
revenue_format = workbook.add_format({''})

# Date Validation in  xlsx writer
worksheet.data_validation('A2',{'validate':'revenue',)

row = 0
col = 0

worksheet.write(row,col,'Name',bold)
worksheet.write(row,col+1,'Marks',bold)
worksheet.write(row,col+2,'Date',bold)
worksheet.write(row,col+3,'Fees',bold)
worksheet.write(row,col+4,'remarks')


#write formats
worksheet.write_string(12,14,'PRATIK')
for name, marks,date_str,fees  in report_card:

    #convert date to datetime object
    date = datetime.strptime(date_str, "%Y-%m-%d")

    worksheet.write(row+1, col, name)
    worksheet.write(row+1, col + 1, marks)
    worksheet.write(row+1,col+2,date)
    worksheet.write(row + 1, col + 3, fees)
    row = row + 1


for name,marks,date_str,fees in report_card:

    #convert date to datetime object
    date = datetime.strptime(date_str,'%Y-%m-%d')

    worksheet.write(row+1,col,name,cell_format)
    worksheet.write_string(row+2,col+3,marks,cell_format)
    worksheet.write(row+1,col+2, date,date_format)

for name,num,revenue, date_str:
    # conversion to date object
    worksheet.write(12,7,date_str,cell_format),
    workbook.write('H1', cell_format)
    workbook.write(23,2,cell_format)
    row = row +2


import datetime

for name,num,revenue,date_str:
    #conversion  in date object
    worksheet.write(12,7,date_str,cell_format)
    worksheet.xlsxwriter(10,7,revenue,cell_format)
    row = row + 1


reported_data = (

    ['hdwbch',77, '2019-02-12',1232],
    ['huedu',88,'2019-02-12',3241],
    ['bhdhebd',89,'2019-01-11',1121]

)

for name,data,id,date,revenue in reported_data:

#datevalivation

    worksheet.write(row,col+1,data,cell_format),
    worksheet.write(row+9,col+2,id,),
    worksheet.write(row+9,col+3,date,date_format ,hiddenn = True),
    worksheet.write(row+9 ,col+4,revenue,revenue_format)
    row = row  + 1


for data,id,revenue in reported_data:

    #conversion to datetime object
    worksheet.write(row +3 , col+10,cell_format),
    worksheet.write(row+2,col+11,cell_format)
    row = row+1
    print(row)


import datetime

for  data,id,date,revenue in reported_data:
    #conversion to date object
    date  = datetime.timedelta(days=7)
    # convert to date object

    worksheet.write(row+4,col+1,data,cell_format),
    worksheet.write(row+4,col+2,id),
    worksheet.write(row+4,col+3,date),
    worksheet.write(row+4,col+4,revenue),



for data,id,date,revenue in reported_data:
    #converting to date object
    worksheet.write(row+17,col+1,data,cell_format)
    worksheet.write(row+17,col+2,id,cell_format)
    worksheet.write(row+17,col+3,revenue,revenue_format)
    row = row+1

