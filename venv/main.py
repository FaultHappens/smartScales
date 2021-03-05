import datetime
import time
import pandas as pd
import smtplib, ssl
from threading import Timer
from datetime import timedelta
from openpyxl import load_workbook
from email import encoders
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

port = 465
smtp_server = "smtp.gmail.com"
sender_email = "smartscalesdigitec@gmail.com"
receiver_email = "tskitishvili04@gmail.com"
password = "digiTec61"
emailSendingTime = 23
filename = "output.xlsx"
waitTime = 3

def sendEmail():
    print("Sending Email!")
    x = datetime.datetime.today()
    y = x.replace(day=x.day, hour=emailSendingTime, minute=0, second=0, microsecond=0) + timedelta(days=1)
    delta_t = y - x
    secs = delta_t.total_seconds()
    t = Timer(secs, sendEmail)
    t.start()
    print(f"Next Email will be sent on {y}")
    message = MIMEMultipart()
    message["From"] = sender_email
    message["To"] = receiver_email
    t=datetime.date.today()
    message["Subject"] = t.strftime('%m/%d/%Y')
    message["Bcc"] = receiver_email
    message.attach(MIMEText(f"გაზომილი {datetime.date.today()}", "plain"))
    with open(filename, "rb") as attachment:
        part = MIMEBase("application", "octet-stream")
        part.set_payload(attachment.read())
    encoders.encode_base64(part)
    part.add_header(
        "Content-Disposition",
        f"attachment; filename= {filename}",
    )
    message.attach(part)
    text = message.as_string()
    context = ssl.create_default_context()
    with smtplib.SMTP_SSL("smtp.gmail.com", 465, context=context) as server:
        server.login(sender_email, password)
        server.sendmail(sender_email, receiver_email, text)
    resetExcel()

x = datetime.datetime.today()
y = x.replace(day=x.day, hour=emailSendingTime, minute=0, second=0, microsecond=0)
delta_t = y - x
secs = delta_t.total_seconds()
t = Timer(secs, sendEmail)
t.start()

def resetExcel():
    df = pd.DataFrame({'გაზომილი(კგ)': [],
                       'თარიღი და დრო': []})
    writer = pd.ExcelWriter(filename, engine='openpyxl')
    df.to_excel(writer, sheet_name='Sheet1', index=False)
    writer.save()

def writeToExcel(weight):
    book = load_workbook(filename)
    ws = book.worksheets[0]
    newRow = (weight, datetime.datetime.now())
    ws.append(newRow)
    book.save(filename)

def readWeight():
    #TODO: reading weight from hx711 and returning (in kg)
    weight = 1
    return weight

while True:
    if readWeight() >= 1:
        for i in range(waitTime):
            if(readWeight()) < 1:
                break
            time.sleep(1)
        writeToExcel(readWeight())