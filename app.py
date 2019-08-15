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

# msd = email.MIMEMultipart()


# def sendMail(sender_email, receiver_email, password, smtp_server="localhost"):
def sendMail():
    # port = 1025  # For SSL
    # message = """\
    # Subject: Test message\
    
    # This is a test message."""

    # msg = email.MIMEMultipart()
    # msg['From'] = sender_email
    # msg['To'] = receiver_email
    # msg['Subject'] = "Python email"

    # body = "Python test mail"
    # msg.attach(email.MIMEText(body, 'plain'))

    # text = msg.as_string()

    # text = """From: Adam
    # To: Sandler
    # Subject: Python email
    
    # Python email app.
    # """    
    # # context = ssl.create_default_context()
    # # with smtplib.SMTP_SSL(smtp_server, port, context=context) as server:
    # #     server.login(sender_email, password)
    # #     server.sendmail(sender_email, receiver_email, text)

    # with smtplib.SMTP(smtp_server, port) as server:
    #     # server.ehlo()  # Can be omitted
    #     # server.starttls(context=context)
    #     # server.ehlo()  # Can be omitted
    #     server.login(sender_email, password)
    #     server.sendmail(sender_email, receiver_email, text)

    # port = 587
    # smtp_server = "smtp.gmail.com"
    # # sender_email = ""
    # # receiver_email = ""
    # # password = ""
    # message = """\
    # Subject: Testing\n

    # Testing the proper functioning of this app."""
    
    # context = ssl.create_default_context()
    # with smtplib.SMTP(smtp_server, port) as server:
    #     server.ehlo()
    #     server.starttls(context=context)
    #     server.ehlo()
    #     server.login(sender_email, password)
    #     server.sendmail(sender_email, receiver_email, message)
    # 
    # set up the SMTP server
    s = smtplib.SMTP(host='smtp.gmail.com', port=587)
    s.starttls()
    s.login("email", "password")      

    msg = MIMEMultipart()       # create a message

    # add in the actual person name to the message template
    # message = message_template.substitute(PERSON_NAME=name.title())

    # Prints out the message body for our sake
    # print(message)

    # setup the parameters of the message
    msg['From']=""
    msg['To']=""
    msg['Subject']="Python app"
    
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