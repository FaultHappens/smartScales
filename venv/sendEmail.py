from email import encoders
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

def sendEmail():
    port = 465
    smtp_server = "smtp.gmail.com"
    sender_email = "smartscalesdigitec@gmail.com"
    receiver_email = "tskitishvili04@gmail.com"
    password = "digiTec61"
    emailSendingTime = 23
    filename = "output.xlsx"
    print("Sending Email!")
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
