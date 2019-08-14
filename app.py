import xlrd

file_name= 'Details.xlsx'
workbook = xlrd.open_workbook(file_name)
worksheet = workbook.sheet_by_name("Sheet1")

num_rows = worksheet.nrows
num_cols = worksheet.ncols

file_data =[]
for row in range(0, num_rows):
    row_data = []
    for col in range(0, num_cols):
        data = worksheet.cell_value(row, col)
        row_data.append(data)
    file_data.append(row_data)

print(file_data)
print(file_data[1][1])