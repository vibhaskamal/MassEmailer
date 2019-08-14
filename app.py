import xlrd
import smtplib
import ssl

file_name= 'Details.xlsx'
sheet_name = "Sheet1"

def read_data(file_name, sheet):
    workbook = xlrd.open_workbook(file_name)
    worksheet = workbook.sheet_by_name(sheet)

    num_rows = worksheet.nrows
    num_cols = worksheet.ncols

    file_data =[]
    for row in range(0, num_rows):
        row_data = []
        for col in range(0, num_cols):
            data = worksheet.cell_value(row, col)
            row_data.append(data)
        file_data.append(row_data)
    
    return file_data


file_data = read_data(file_name, sheet_name)
# print(file_data)
# print(file_data[1][1])

port = 587
smtp_server = "smtp.gmail.com"
sender_email = ""
receiver_email = ""
password = ""
message = """\
Subject: Testing

Testing the proper functioning of this app."""

context = ssl.create_default_context()
with smtplib.SMTP(smtp_server, port) as server:
    server.ehlo()
    server.starttls(context=context)
    server.ehlo()
    server.login(sender_email, password)
    server.sendmail(sender_email, receiver_email, message)