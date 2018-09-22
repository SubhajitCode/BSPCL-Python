import xlrd
import xlwt

from xlrd.sheet import ctype_text
def parse_xlsx():
    m=0
    book = xlwt.Workbook(encoding="utf-8")
    sheet1 = book.add_sheet("Sheet 1")
    file_location = "output.xlsx"
    workbook = xlrd.open_workbook(file_location)
    for i in range (3,141):
        active_sheet = workbook.sheet_by_index(i)
        num_rows = active_sheet.nrows
        num_cols = active_sheet.ncols
        for row_idx in range(0, num_rows):
            if active_sheet.cell(row_idx,14).value=='EE' and active_sheet.cell(row_idx,6).value=='EBC':
                for col_idx in range(0,num_cols):
                    cell_value=active_sheet.cell(row_idx,col_idx).value
                    print(m)
                    sheet1.write(m,col_idx,cell_value)
                m=m+1
    

    book.save("cleaned_ebc.xls")
                
                
            

parse_xlsx()