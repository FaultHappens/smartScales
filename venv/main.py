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
from hx711 import HX711
from pyGPIO.gpio import connector, gpio
import smbus
from sendEmail import sendEmail


scaleRatio = 26.33
ledPin = connector.GPIOp12
buttonPin = connector.GPIOp11
sendButtonPin = connector.GPIOp13
minWeight = 1000
hx = HX711(dout_pin=connector.GPIOp8, pd_sck_pin=connector.GPIOp7)
hx.set_scale_ratio(scaleRatio)
hx.zero()
gpio.setcfg(buttonPin, gpio.INPUT)
gpio.setcfg(sendButtonPin, gpio.INPUT)
gpio.setcfg(ledPin, gpio.OUTPUT)
gpio.pullup(buttonPin, gpio.PULLDOWN)
gpio.pullup(sendButtonPin, gpio.PULLDOWN)
port = 465
smtp_server = "smtp.gmail.com"
sender_email = "smartscalesdigitec@gmail.com"
receiver_email = "tskitishvili04@gmail.com"
password = "digiTec61"
emailSendingTime = 22
filename = "/etc/smartScales/venv/output.xlsx"
waitTime = 3

gpio.output(ledPin, 0)


# Define some device parameters
I2C_ADDR  = 0x27 # I2C device address
LCD_WIDTH = 16   # Maximum characters per line

# Define some device constants
LCD_CHR = 1 # Mode - Sending data
LCD_CMD = 0 # Mode - Sending command

LCD_LINE_1 = 0x80 # LCD RAM address for the 1st line
LCD_LINE_2 = 0xC0 # LCD RAM address for the 2nd line
LCD_LINE_3 = 0x94 # LCD RAM address for the 3rd line
LCD_LINE_4 = 0xD4 # LCD RAM address for the 4th line

LCD_BACKLIGHT  = 0x08  # On
#LCD_BACKLIGHT = 0x00  # Off

ENABLE = 0b00000100 # Enable bit

# Timing constants
E_PULSE = 0.0005
E_DELAY = 0.0005

#Open I2C interface
bus = smbus.SMBus(0)

def lcd_init():
  # Initialise display
  lcd_byte(0x33,LCD_CMD) # 110011 Initialise
  lcd_byte(0x32,LCD_CMD) # 110010 Initialise
  lcd_byte(0x06,LCD_CMD) # 000110 Cursor move direction
  lcd_byte(0x0C,LCD_CMD) # 001100 Display On,Cursor Off, Blink Off 
  lcd_byte(0x28,LCD_CMD) # 101000 Data length, number of lines, font size
  lcd_byte(0x01,LCD_CMD) # 000001 Clear display
  time.sleep(E_DELAY)

def lcd_byte(bits, mode):
  # Send byte to data pins
  # bits = the data
  # mode = 1 for data
  #        0 for command

  bits_high = mode | (bits & 0xF0) | LCD_BACKLIGHT
  bits_low = mode | ((bits<<4) & 0xF0) | LCD_BACKLIGHT

  # High bits
  bus.write_byte(I2C_ADDR, bits_high)
  lcd_toggle_enable(bits_high)

  # Low bits
  bus.write_byte(I2C_ADDR, bits_low)
  lcd_toggle_enable(bits_low)

def lcd_toggle_enable(bits):
  # Toggle enable
  time.sleep(E_DELAY)
  bus.write_byte(I2C_ADDR, (bits | ENABLE))
  time.sleep(E_PULSE)
  bus.write_byte(I2C_ADDR,(bits & ~ENABLE))
  time.sleep(E_DELAY)

def lcd_string(message,line):
  # Send string to display

  message = message.ljust(LCD_WIDTH," ")

  lcd_byte(line, LCD_CMD)

  for i in range(LCD_WIDTH):
    lcd_byte(ord(message[i]),LCD_CHR)


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
    weight = hx.get_weight_mean(20)
    return weight
    
def writeToLcd(weightt):
    message = f"     {int(weightt)} g"
    lcd_string(message,LCD_LINE_1)

lcd_init()
while True:
    if gpio.input(buttonPin):
        while gpio.input(buttonPin):
            pass
        hx.zero()
        print("Set zero!")
    weight = readWeight()
    if weight <= 0:
        writeToLcd(0)
    else:
        writeToLcd(weight)
    print(weight)
    if weight >= minWeight:
        gpio.output(ledPin, 1)
        for i in range(waitTime):
            weight = readWeight()
            writeToLcd(weight)
            if weight < minWeight:
                break
            time.sleep(1)
        writeToExcel(weight)
        writeToLcd(weight)
        gpio.output(ledPin, 0)
        while weight > minWeight:
            weight = readWeight()
            writeToLcd(weight)
    gpio.output(ledPin, 0)
