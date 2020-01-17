# Tutorial 1: Create a simple XLSX file
# Tutorial 2: Adding formatting to the XLSX File
# Tutorial3 : Adding different types of data types
#different types of data  in date validation



import xlsxwriter
from datetime import datetime



report_card = (

        ['manish', 77,'2011-01-22',2122],
        ['suresh', 78,'2011-01-12',4325],
        ['kamapisachi',83,'2011-01-8',5645],

    )


workbook = xlsxwriter.Workbook('student.xlsx')
worksheet = workbook.add_worksheet()

bold = workbook.add_format({'bold': True})


date_format =  workbook.add_format({'num_format': 'mmmm d yyyy'})
money_format = workbook.add_format({'num_format': '$#,##0'})


row = 0
col = 0


worksheet.write(row,col,'Name',bold)
worksheet.write(row,col+1,'Marks',bold)
worksheet.write(row,col+2,'Date',bold)
worksheet.write(row,col+3,'Fees',bold)




#write formats
worksheet.write_string(12,14,'PRATIK')


for name, marks,date_str,fees  in report_card:

    #convert date to datetime object
    date = datetime.strptime(date_str, "%Y-%m-%d")

    worksheet.write(row+1, col, name)
    worksheet.write(row+1, col + 1, marks)
    worksheet.write(row+1,col+2,date,date_format)
    worksheet.write(row + 1, col + 3, fees ,money_format)
    row = row + 1


workbook.close()









