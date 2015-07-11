import sys
import time
import os
from os.path import basename
from openpyxl import load_workbook
from ConfigParser import ConfigParser
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
from email.mime.base import MIMEBase
import email


SLEEP_TIME = 0

#Excel Sheet variables
EXCEL_SHEET_FILENAME = "sample.xlsx"
PAN_COL_ALPHABET = "A"
NAME_COL_ALPHABET = "B"
EMAIL_COL_ALPHABET = "C"

#Email server
SMTP_HOST = "md-77.webhostbox.net"
SMTP_PORT = 465
SMTP_USERNAME = "test@mittalentrp.com"
SMTP_PASSWORD = "simplepass"
SMTP_FROM = SMTP_USERNAME

#Email
MAIL_SUBJECT = "TDS Certificate"
MAIL_FROM = "Test Name"
BODY = """\
<html>
    <head></head>
    <body>
        Dear %s,
        <p>Please find as attachment a copy of your TDS certificate (PDF)</p>
        <p>
            From,<br/>
            %s
        </p>
    </body>
</html>
""" 


###########################  Settings functions ##############################

def cfg_load(section, key):
    config = ConfigParser()
    retval = False
    try:
        config.read('settings.ini')
        if config.has_section(section):
            retval = config.get(section, key)
    except Exception:
        pass
    return retval


def cfg_save(section, key, value):
    config = ConfigParser()
    try:
        config.read('settings.ini')
        if not config.has_section(section):
            config.add_section(settings) 
        config.set(section, key, value)
        with open("settings.ini", 'w') as cfg:
            config.write(cfg)
        return True
    except Exception:
        pass
    return False

###########################  Email functions ##############################


s = smtplib.SMTP_SSL()

def connect_smtp():
    s.connect(SMTP_HOST, SMTP_PORT)
    #s.starttls()
    s.login(SMTP_USERNAME, SMTP_PASSWORD)

def disconnect_smtp():
    s.quit()

def send_mail(to_address, body_str, file):
    msg = MIMEMultipart('')
    msg["Subject"] = MAIL_SUBJECT
    msg["From"] = SMTP_FROM
    msg["To"] = to_address
    body = body_str
    htmlMIME = MIMEText(body, 'html')
    msg.attach(htmlMIME)

    attachFile = MIMEBase('application', 'octet-stream')
    file_basename = basename(file)
    attachFile.add_header('Content-Disposition', 'attachment;', filename='%s' % file_basename)

    with open(file, "rb") as fil:
        attachFile.set_payload(fil.read())

    email.Encoders.encode_base64(attachFile)
    msg.attach(attachFile)

    s.set_debuglevel(1)
    s.sendmail(SMTP_FROM, to_address, msg.as_string())


###########################  Script starts here ##############################

#Start reading the excel file
print "[INFO] Beginning to read excel sheet....."
wb = load_workbook(filename = EXCEL_SHEET_FILENAME)  
sheet = wb.worksheets[0]

#Collect some statistics needed for later
row_count = int(sheet.get_highest_row())
column_count = sheet.get_highest_column()
print "[INFO Total of %s rows (including header row)" % (row_count)

#Load the last processed row, so we can resume from the next row
last_processed_row = cfg_load("settings", "last_processed_row")
# Assuming the first row contains the HEADER columns
# we skip to row 2
row = 2
#Or we resume from the row after the last processed row
#Delete the 'last_processed_row' line to start
#from the beginning
if last_processed_row:
    row = int(last_processed_row)
    row += 1

if row > row_count:
    print "[INFO] All rows have been processed....."
else:
    print "[INFO] Starting from row: %s....." % row
    connect_smtp()

    while(row <= row_count):
        panCell = PAN_COL_ALPHABET + str(row)
        nameCell = NAME_COL_ALPHABET + str(row)
        emailCell = EMAIL_COL_ALPHABET + str(row)

        currPANCell = sheet[panCell]
        currNameCell = sheet[nameCell]
        currEmailCell = sheet[emailCell]

        custName = currNameCell.value
        custEmail = currEmailCell.value
        custPan = currPANCell.value

        try:
            print("Sending to: [ROW NO:%s] %s (%s)" % (row, custName.encode('utf8'), custEmail))
        except Exception:
            pass

        try:
            if os.path.isdir('pdfs'):
                pdf_path = "pdfs\%s.%s" % (custPan, "pdf")
                if os.path.exists(pdf_path):
                    body = BODY % (custName, MAIL_FROM)
                    send_mail(custEmail, body, pdf_path)

                    #Save the current row number
                    cfg_save('settings', 'last_processed_row', row)

                    print("Email sent !")
                else:
                    print "[ERROR][ROW NO: %s] PDF file for this name is missing" % row
            else:
                print "All pdf files must be in a folder called 'pdfs' in the same folder as the script"
        except Exception, e:
            print "[ERROR]  %s" % e

        time.sleep(SLEEP_TIME)
        row += 1

    disconnect_smtp()




            
