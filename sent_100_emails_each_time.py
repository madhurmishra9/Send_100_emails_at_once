# code to send many emails automatically by reading an .xlsx file.

import smtplib
from email import encoders
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.image import MIMEImage
import openpyxl
import getpass
import os.path



# this code is used to read a new generation excel i.e .xlsx files and send email to more 95 email id automatically with predefined text.
# to use it with xls and xlsx the old generation excel format we can use xlrd or xlwt
def send_email(send_to):
    email = 'sender email id'
    password = ''
    subject = 'test'
    message = 'Test message'
    #file_location = r'Attachment file location in case any '

    msg = MIMEMultipart()
    msg['From'] = email
    msg['To'] = send_to
    msg['Subject'] = subject

    msg.attach(MIMEText(message, 'plain'))

    # Adding the attachment
    filename = os.path.basename(file_location)
    attachment = open(file_location, "rb").read()
    image = MIMEImage(attachment, name=filename)
    msg.attach(image)

    server = smtplib.SMTP('smtp.gmail.com', 587)#using gmail server
    server.starttls()
    server.login(email, password)# passing credentials these are defined above or can be used as command line parameters
    text = msg.as_string()
    server.sendmail(email, send_to, text)
    server.quit()

def read_excel():
    wb = openpyxl.load_workbook(r'Excel file path containg list of email addresses',read_only = True)
    sheet = wb.active
    list_emails = list()# defineing empty list
    count = 1
    for i in range(count,95):#have restricted to send at most 95 emails as gmail have email restrictions of 100 emails.
        cell = 'A' + str(count)# column in excel containing the email addresses
        if sheet[cell].value != None:
            #print("Success")
            #print(sheet[cell].value)
            list_emails.append(str(sheet[cell].value)) #fetching email list from file
        count += 1

    #wb.save(r'file that was read') #saving excel in case if te file is opened as read_only = False
    #print(list_emails)
    #print (count)
    for i in range(0,95):#
        print(list_emails[i])# printing email address for which email is sent
        send_email(list_emails[i]) # calling the send email function


read_excel()
