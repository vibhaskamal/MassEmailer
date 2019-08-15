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


"""
This functions sets up the SMTP server, starts a TLS session and enables the user to log into their account

@parameter sender_email: Sender's email
@parameter password: Password for sender's email
@parameter host
@parameter port
@return : SMTP server connection instance
"""
def setupServerConnection(sender_email, password, host="smtp.gmail.com", port=587):
    # Setting up the SMTP server details
    server_connection = smtplib.SMTP(host, port)

    # Start the TLS session
    server_connection.starttls()

    server_connection.login(sender_email, password)

    return server_connection


"""
This function terminates the SMTP server

@parameter server_connection: An SMTP server connection instance
"""
def terminateServerSession(server_connection):
    server_connection.quit()


"""
This function sends the emails from the sender's account to the receiver accounts

@parameter server_connection: An SMTP server connection instance
@parameter sender_email: Sender's email
@parameter password
@parameter password: Password for sender's email
"""
def sendMail(server_connection, sender_email, password, receiver_email, subject, message):
    msg = MIMEMultipart()

    # Setting up the email's contents
    msg['From'] = sender_email
    msg['To'] = receiver_email
    msg['Subject'] = subject
    
    message_body = message

    # Adding the message body
    msg.attach(MIMEText(message_body, 'plain'))
    
    # Send the email using the SMTP server connection instance
    server_connection.send_message(msg)

    del msg
    
    # Testing purposes
    # print("Done")


"""
This function reads the data in the filename given as a parameter

@parameter filname: Name of the file from which the data is to be read
@return : Data in the file
"""
def readFile(filename):
    file = open(filename, "r")
    data = file.read()
    return data


"""
This function creates the message to be sent

@parameter text: The text which has to be sent as the body of the email
@parameter person_name: Name of the person (this parameter is hardcoded and will have to be modified based on the structure of                                      Body.txt)
@parameter money_value: Amount of money (this parameter is hardcoded and will have to be modified based on the structure of                                      Body.txt)
@return : Body of the email
"""
def createMessage(text, person_name, money_value):
    body = text.format(NAME=person_name, AMOUNT=money_value)
    return body



def main():
    # Excel file and sheets from where the user data is to be extracted
    file_name= 'Details.xlsx'
    sheet_name = "Sheet1"

    file_data = read_data(file_name, sheet_name)

    sender_email = ""
    sender_password = ""

    connection = setupServerConnection(sender_email, sender_password)

    # Looping through each row in the excel sheet and sending emails created using the data in the excel sheet and Body.txt
    for i in range(1, len(file_data)):
        name = file_data[i][1]
        receiver_email = file_data[i][3]
        amount = file_data[i][4]

        text_file = readFile("Body.txt")

        msg_body = createMessage(text_file, name, amount)

        sendMail(connection, sender_email, sender_password, receiver_email, "Amount due", msg_body)

    terminateServerSession(connection)

main()

print("END OF PROGRAM")

