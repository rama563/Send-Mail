# -*- coding: utf-8 -*-
"""
Created on Mon Jun 18 13:23:00 2018

@author: RamaKrishna Ventrapragada
"""

import smtplib
import os
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email.utils import formatdate
from email import encoders
import datetime
os.chdir("M:/Monashee Projects/Daily/Dealog Database/MDD/MDD Output")
now = datetime.datetime.now()
month=str(now.month)
if(len(month)==1):
        month="0"+month
date=str(now.day)
if(len(date)==1):
        date="0"+date
year=str(now.year)
file_name="MDD_"+month+date+year+"_MMLS.xlsx"
file_path="M:/Monashee Projects/Daily/Dealog Database/MDD/MDD Output/"+file_name

files_avl=os.listdir()


if(file_name in files_avl):
    server = smtplib.SMTP('smtp.gmail.com', 587)    
    server.ehlo()
    server.starttls()
    server.login("ghc@goldenhillsindia.com", "GHChyd2017")
    part = MIMEBase('application', "octet-stream")
    part.set_payload(open(file_path, "rb").read())
    encoders.encode_base64(part)
    part.add_header('Content-Disposition', 'attachment; filename="MDD.xlsx"')
    msg = MIMEMultipart()
    text="Hi Griffin/Jeff, \n\n PFA the Latest MDD \n\n Thanks & Regards,\n Abyn"
    msg.attach(MIMEText(text))
    msg.attach(part)
    #to="abyn.scaria@goldenhillsindia.com,sravya.v@goldenhillsindia.com"
    to="griffin@monasheecap.com,jeff@monasheecap.com,operations@monasheecap.com"
    #cc="shradha.singhal@goldenhillsindia.com"
    cc="shradha.singhal@goldenhillsindia.com,anirudh.vajjah@goldenhillsindia.com,abyn.scaria@goldenhillsindia.com"    
    msg['From']="ghc@goldenhillsindia.com"
    msg['To']=to
    msg['Cc']=cc
    rcpt =cc.split(",") + to.split(",")
    msg['Subject']="MDD for "+month+"-"+date+"-"+year
    server.sendmail("ghc@goldenhillsindia.com",rcpt, msg.as_string())
else:
    server = smtplib.SMTP('smtp.gmail.com', 587)    
    server.ehlo()
    server.starttls()
    server.login("ghc@goldenhillsindia.com", "GHChyd2017")
    msg = MIMEMultipart()
    text="Hi Team, \n\n MMLS MDD File - ("+file_name+") is not available in the MDD Output folder, Please Check. \n\n Regards,\n Bot"
    msg.attach(MIMEText(text))
    to="abyn.scaria@goldenhillsindia.com,sravya.v@goldenhillsindia.com"
    cc="shradha.singhal@goldenhillsindia.com,ramakrishna.v@goldenhillsindia.com"
    msg['From']="ghc@goldenhillsindia.com"
    msg['To']=to
    msg['Cc']=cc
    rcpt =cc.split(",") + to.split(",")
    msg['Subject']='Failure Mail - '+"MDD for"+month+"-"+date+"-"+year
    server.sendmail("ghc@goldenhillsindia.com",rcpt, msg.as_string())
