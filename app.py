import xlrd
import smtplib
import ssl
import email
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from string import Template

"""
This function reads the data from a sheet in an excel file

@parameter filename: Name of the excel file (with the extension)
@parameter sheet: Name of the sheet within the excel file, from where the data has to be read
@return: Data within the specified sheet in the form of arrays
"""
def read_data(file_name, sheet):
    # Opening the excel file and the specified sheet
    workbook = xlrd.open_workbook(file_name)
    worksheet = workbook.sheet_by_name(sheet)

    # Finding number of rows and columns with data in it
    num_rows = worksheet.nrows
    num_cols = worksheet.ncols

    # Looping through each row and storing values for all columns within a row in an array
    # All the arrays will be part of a parent array: file_data
    file_data =[]
    for row in range(0, num_rows):
        row_data = []
        for col in range(0, num_cols):
            data = worksheet.cell_value(row, col)
            row_data.append(data)
        file_data.append(row_data)
    
    return file_data


file_name= 'Details.xlsx'
sheet_name = "Sheet1"

file_data = read_data(file_name, sheet_name)


def sendMail(sender_email, receiver_email, password, subject, message, host="smtp.gmail.com", port=587):
    # Setting up the SMTP server details
    s = smtplib.SMTP(host, port)

    # Start the TLS session
    s.starttls()

    s.login(sender_email, password)      

    msg = MIMEMultipart()

    # setup the parameters of the message
    msg['From'] = sender_email
    msg['To'] = receiver_email
    msg['Subject'] = subject
    
    message_body = message

    # add in the message body
    msg.attach(MIMEText(message_body, 'plain'))
    
    # send the message via the server set up earlier.
    s.send_message(msg)

    del msg
    
    # Terminate the SMTP session and close the connection
    s.quit()

    print("Done")


def readFile(filename):
    file = open(filename, "r")
    data = file.read()
    return data


def createMessage(template_body, person_name, money_value):
    body = template_body.format(NAME=person_name, AMOUNT=money_value)
    return body


def main():
    file_name= 'Details.xlsx'
    sheet_name = "Sheet1"

    file_data = read_data(file_name, sheet_name)

    for i in range(1, len(file_data)):
        name = file_data[i][1]
        email = file_data[i][3]
        amount = file_data[i][4]

        text_file = readFile("Body.txt")

        msg_body = createMessage(text_file, name, amount)

        sendMail("", email, "", "Amount due", msg_body)

main()

print("Successful")

