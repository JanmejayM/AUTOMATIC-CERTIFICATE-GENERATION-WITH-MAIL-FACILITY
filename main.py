

#                                   Automatic Certificate Generation With Mail Facility
#                                           Author-->  Janmejay Mohapatra


import  cv2
import matplotlib.pyplot as pt
import tkinter as tk
from tkinter import filedialog as fd
import datetime as dt
import pandas as pd
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders

d=dt.datetime.now();

path=fd.askopenfilename()

u=pd.read_excel(path,dtype=str)
#to be displayed


Mail=u['Mail']
Id=u['Certificate Id']
Id=Id.to_list()
Name=u['Name']
Name=Name.to_list()
Instructor=u['Course Ins']
Project=u['Project']
Project=Project.to_list()
Instructor=Instructor.to_list()
dtw=d.strftime("%d-%b-%y")

image = cv2.imread('')#put location of the certificate template


#use only for getting Co-ordinates....
   #pt.imshow(image)
   #pt.show()




for i in range(0,len(Name)):
    image = cv2.imread('')#put location of the certificate template

    im = cv2.putText(image, Name[i], (722, 802), cv2.FONT_HERSHEY_SIMPLEX, 3, (1, 0, 0), 5, cv2.LINE_AA)
    im = cv2.putText(im, dtw, (1536, 1312), cv2.FONT_HERSHEY_SIMPLEX, 1.5, (1, 0, 0), 3, cv2.LINE_AA)
    im = cv2.putText(im, Id[i], (1858, 135), cv2.FONT_HERSHEY_SIMPLEX, 1.5, (1, 0, 0), 2, cv2.LINE_AA)
    im = cv2.putText(im, Project[i], (594, 1073), cv2.FONT_HERSHEY_SIMPLEX, 3, (1, 0, 0), 5, cv2.LINE_AA)
    im = cv2.putText(im, Instructor[i], (492, 1312), cv2.FONT_HERSHEY_SIMPLEX, 1.5, (1, 0, 0), 2, cv2.LINE_AA)


    cv2.imwrite('Certificate.png',im)
    fromaddr=""#put your mail id
    toaddr = Mail[i:i+1]
     # instance of MIMEMultipart
    msg = MIMEMultipart()
    # storing the senders email address
    msg['From'] = fromaddr
    # storing the receivers email address
    #msg['To'] = ""
    # storing the subject
    msg['Subject'] = "Certificate for Completion of Course"
    # string body

    body = "The Certificate is attached with this Mail"+"   "+Name[i]
    # attach the body
    msg.attach(MIMEText(body,'plain'))
    # open the file to be sent
    filename = "Certificate.png"
    attachment = open("/Users/janmejaymohapatra/PycharmProjects/p1/Certificate.png", "rb")
    # instance of MIMEBase as p
    p = MIMEBase('application', 'octet-stream')
    # To change the payload into encoded form
    p.set_payload((attachment).read())
    # encode into base64
    encoders.encode_base64(p)
    p.add_header('Content-Disposition', "attachment; filename= %s" % filename)
    # attach the instance 'p' to instance 'msg'
    msg.attach(p)
    # creates SMTP session
    s = smtplib.SMTP('smtp.gmail.com',587)
    # start TLS for security
    s.starttls()
    # Authentication
    s.login(fromaddr,"")#put your password
    # Converts the Multipart msg into a string
    text = msg.as_string()
    # sending the mail
    s.sendmail(fromaddr, toaddr, text)
    # terminating the session
    im=image

s.quit()
print('Task Done')






