# Worksheet class

import xlsxwriter

workbook = xlsxwriter.Workbook('sample.xlsx',{'strings_to_numbers':  False,
                               'strings_to_formulas': True,
                               'strings_to_urls':     True})


data = ('Foo','Bar','Zin')
worksheet =  workbook.add_worksheet()
cell_format = workbook.add_format({'bold': True, 'italic': True})


worksheet.write(0, 0, 'Hello')                                        # write_string()
worksheet.write(1, 0, 'World')                                        # write_string()
worksheet.write(2, 0, 2)                                              # write_number()
worksheet.write(3, 0, 3.00001)                                        # write_number()
worksheet.write(4, 0, '=SIN(PI()/4)')                                 # write_formula()
worksheet.write_string(4,2,cell_format)

worksheet.write(5, 0, '')                                             # write_blank()
worksheet.write(6, 0, None)                                           # write_blank()
worksheet.write('B1','shyan',cell_format)                             # positional statement B1
worksheet.write('B2', True)                                           #write boolean
worksheet.write('B3','https://www.python.org/')                       # write url
#worksheet.write_rich_string('A1', 'This is ', cell_format, 'bold')    #write rich string
#worksheet.write_rich_string('D1')
worksheet.set_row(0,20,cell_format)                                # sets  row for 1 to 20 to   cell_format tyre

worksheet.set_row(2,11,cell_format)                                  #set row attribute
worksheet.set_col(3,2,cell_format)                                   # set col attribute
worksheet.set_col_header(0,0,1,cell_format)                          # sets col header

#write row and write column
worksheet.write_row('C1',data)                                        #writerow #writes all value in data in rows
worksheet.write_column('D1',data)                                     # write column # writes all value indata in column

# set row and set column
#worksheet.insert_image('E1','python.jpg')                             #inserts image at purticular position


worksheet.set_autofillers(0,0,10)
worksheet.set_row(0,0,10)
worksheet.set_x_axis(2,1,'Revenue')



worksheet.set_row_header('A1', ' This is row header 1')
worksheet.set_col_header(0,0,'this is col_header1')
chart = worksheet .add_chart({type,'column'})
worksheet.insert_chart('G1', chart)
workbook.insert_chart('G2', 'Revenue')
worksheet.insert_chart({'type':'bar'})


workbook.close('')