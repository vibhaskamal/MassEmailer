import xlrd
import smtplib
import ssl
import email
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

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


def sendMail():
    # Setting up the SMTP server details
    s = smtplib.SMTP(host='smtp.gmail.com', port=587)

    # Start the TLS session
    s.starttls()

    s.login("", "")      

    msg = MIMEMultipart()

    # setup the parameters of the message
    msg['From']=""
    msg['To']=""
    msg['Subject']="Python app Part 2"
    
    message = "Hello, how are you?"

    # add in the message body
    msg.attach(MIMEText(message, 'plain'))
    
    # send the message via the server set up earlier.
    s.send_message(msg)

    del msg
    
    # Terminate the SMTP session and close the connection
    s.quit()

    print("Done")


sendMail()

print("Successful")