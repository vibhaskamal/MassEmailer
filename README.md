# MassEmailer


This program is an email sender application built using Python 3.

It uses an Excel file to get the details of the email recipients and a text file (.txt) to get the body of the email to be sent.



# Usage

## Requirements

- Python 3


## Instructions to Run the Program

1) Open the **Details.xlsx** file.

2) Enter the details under the relavant columns.

3) Modify the **Body.txt** file based on your email body.

4) If you modify the {} enclosed text, make the corresponding change in the **app.py** file under the **createMessage()** function.

5) Go to **app.py** file and enter your email ID and password in the **sender_email** and **sender_password** variables.

6) Ensure that **app.py**, **Details.xlsx** and **Body.txt** are in the same folder.

7) Run **app.py**