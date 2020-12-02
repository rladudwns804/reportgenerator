import pyodbc
import pandas as pd
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders

#Connect to SQL Server:
conenct = pyodbc.connect(
'Driver=MySQL ODBC 8.0 ANSI Driver;'
'SERVER=localhost;'
'DATABASE=datbasename'
'UID=username'
'PWD=password'
'charset=utf8mb4'
)

#SQL querey statment:
sql= 'SELECT * FROM test' #Put necessary querey stament here

#Read the SQL using pandas read_sql:
data = pd.read_sql(sql,connect)

#Convert the qurery data to dataframe
df = pd.DataFrame(data)
df.to_excel("ouput.xlsx", index=False)

#Setup for email(gmail)

gmail_user= "useremail@gmail.com"
gmail_p[assword= "userpassword"
sent_from = gmail_user
to= "receiver@gmail.com" #For multiple recipent, use list 
subject= "Subject of email"
body= "This is the body of email"

#Setup MIME
message= MIMEMultipart()
message["From"] = gmail_user
message["To"] = to #If using list, must convert to str use ",".join(to) 
message["Subject"] = subject

#Setup body and attachemtns for the mail:
message.attach(MIMEText(body))

attach_file_name = "output.xlsx"
attach_file = open(attach_file_name, "rb") #Open file as binary mode

payload = MIMEBase("application", "octate-stream")
payload.set_payload((attach_file).read())
encoders.encode_base64(payload) #Encode the attachment

#Add payload header with filename:
payload.add_header("Content-Disposition", "attachment", filename= atach_file_name)
message.attach(payload)

#Create SMTP session for sending the email:

server=smtplib.SMTP('smtp.gmail.com')
server.starttls() #Start secure connection
server.login(gmail_user,gmail_password)
text = message.as_string()
server.sendmail(gmail_user, to, text)
server.quit()


