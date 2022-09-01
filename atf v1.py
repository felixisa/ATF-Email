#atf email v0.1
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
import sqlite3
import xlrd

def find_classes(name):
    conn = sqlite3.connect('ATFplacement.db')
    cur = conn.cursor()
    cur.execute("SELECT ClassName FROM Classes WHERE Members LIKE '%{}%'".format(name))
    classes = cur.fetchall()
    res = ''
    for row in classes:
        res = str(row[0])+ "\n" + res 
    return res
    conn.close()

def find_email(name):
    conn = sqlite3.connect('ATFplacement.db')
    cur = conn.cursor()
    cur.execute("SELECT Email FROM Emails WHERE Name LIKE '%{}%'".format(name))
    email = cur.fetchone()
    return str(email[0])

def send(name):
    recipient = find_email(name)
    classes = find_classes(name)
    fromaddr = "admin@acrossthefloor.com"
    toaddr = recipient
   
    # instance of MIMEMultipart
    msg = MIMEMultipart()
  
    # storing the senders email address  
    msg['From'] = fromaddr
  
    # storing the receivers email address 
    msg['To'] = toaddr
  
    # storing the subject 
    msg['Subject'] = "ATF Fall 2022 Schedule ({})".format(name)
  
    # string to store the body of the mail
    body = "We hope you are all enjoying your summer. While there's still some time left to enjoy, it's also time to start thinking and planning your child's schedule for the upcoming school year. Below is your child's placement and recommended class schedule.\n\n" + classes + "\nIf you have a schedule conflict or If your child would like to take a different class, please let us know and we can try to find the best option. \n\nDid you know we have other exciting classes to offer?\n\nMaybe your child likes to sing!  Our Musical Theater classes are perfect for those who are interested in learning to sing as well as those who want to train vocally and take their singing to a higher level. Group lessons and private options, including audition prep are available.\n\nWe also have a breakdancing program, which is a street-style form of dance, which will have its Olympic Debut in 2024! Breaking is physical, builds strength and confidence, and encourages self-expression.\n\nClass space is limited and some fill up quicker than others. We highly recommend that you register as early as possible to reserve your child's spot in the class you want.\n\nAnd if you register by August 31, you will receive a 20% off coupon to use at Main Street Dancewear where you will find everything you need for classes. \n\nTo register is easy! Just click the link below to access the Parent Portal\nhttps://app.jackrabbitclass.com/jr3.0/ParentPortal/Login?orgId=504571 \n\nAttached is the financial contract with the tuition rates for 22-23.\n\nCan't wait to see you back in class! Classes begin on Sept 12th!"
  
    # attach the body with the msg instance
    msg.attach(MIMEText(body, 'plain'))
  
    # open the file to be sent 
    filename = "{}.pdf".format(name)
    attachment = open(filename, "rb")
  
    # instance of MIMEBase and named as p
    p = MIMEBase('application', 'pdf')
  
    # To change the payload into encoded form
    p.set_payload((attachment).read())
  
    # encode into base64
    encoders.encode_base64(p)
   
    p.add_header('Content-Disposition', "attachment; filename= %s" % filename)
  
    # attach the instance 'p' to instance 'msg'
    msg.attach(p)
  
    # creates SMTP session
    s = smtplib.SMTP('smtp.gmail.com', 587)
  
    # start TLS for security
    s.starttls()
  
    # Authentication
    s.login(fromaddr, "ATFdance1996")
  
    # Converts the Multipart msg into a string
    text = msg.as_string()
  
    # sending the mail
    s.sendmail(fromaddr, toaddr, text)
  
    # terminating the session
    s.quit()

def main():
    
    conn = sqlite3.connect('ATFplacement.db')
    cur = conn.cursor()
    cur.execute("SELECT Name FROM Emails")
    #students = cur.fetchall()[262:]
    #for student in students:
        #send(str(student[0]))
    print(find_classes("Dempsey A"))

main()
