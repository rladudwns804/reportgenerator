import pyodbc
import pandas as pd
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from premailer import transform

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


#Format DataFrame(If want to have table on the body and not as attachment) : 

pre_html = df.style.set_table_styles([{'selector':'th',
                            'props':[('background' ,'#4F81BD')]},
                            {'selector': 'tr:nth-of-type(odd)',
                             'props': [('background', '#DCE6F1')]}]).hide_index().render()

html  transform(pre_html, pretty_print =True) #inline CSS



#Setup for email(gmail)

gmail_user= "useremail@gmail.com"
style.set_table_style
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




#Setupt MIME for tables on the body(not attachment) :
message.attach(MIMEText(body))
part = MIMEText(html,'html')
mesasge.attach(part)


#Setup attachemtns for the mail:

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


