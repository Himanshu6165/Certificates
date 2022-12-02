import cv2
import pandas as pd
import os 
import smtplib
from email.message import EmailMessage
from getpass import getpass
df0=pd.read_excel("For Email.xlsx")
list_names=df0["NAME"].values

for index, name in enumerate(list_names):
    template= cv2.imread('Generic-Certificate-of-Completion.jpg')
    cv2.putText(template,name,(187,175),cv2.FONT_HERSHEY_COMPLEX,0.5,(0,0,255),1,cv2.LINE_AA)
    cv2.imwrite(f'E:\python\generated certificates\{name}.jpg',template)
    print('Processing certificate')
sender_email="himanshubansal2107@gmail.com"
sender_pass="xjmormuyechnayev"

df=pd.read_excel("For Email.xlsx")
receivers_email=df["EMAIL_ID"].values
sub=("Test Mail ")
attach_files=df["Files to be attached"]
name=df["NAME"].values


zipped=zip(receivers_email,attach_files,name)

for(a,b,c) in zipped:
    
    msg=EmailMessage()
    files=[(r"E:\python\generated certificates\{}.jpg".format(b))]
    
    for file in files:
        
        with open(file,'rb') as f:
            
            file_data=f.read()
            file_name=f.name
            
        msg['From']=sender_email
        msg['To']=a
        msg['Subject']=sub
        msg.set_content(f"hello {c}! I have something for you.")
        msg.add_attachment(file_data,maintype='application',subtype='octet-stream',filename="{}.jpg".format(b))
        
        with smtplib.SMTP_SSL('smtp.gmail.com',465) as smtp:
            
            smtp.login(sender_email,sender_pass)
            
            smtp.send_message(msg)
            
print("All mail sent!")
