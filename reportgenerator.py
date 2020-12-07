import pyodbc
import pandas as pd
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from premailer import transform
from bs4 import BeautifulSoup as sp

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

style = df.style.set_table_styles([
                                    {'selector': 'table',
                                     'props':[('border','1px solid')]},
                                    {'selector': 'td',
                                     'props':[('border','1px solid')]},
                                    {'selector': 'tr',
                                     'props':[('border','1px solid')]},
                                    {'selector':'th',
                            'props':[('background' ,'#4F81BD')]}]).hide_index().render()


#Inline CSS via Transform and use Beautiful soup to find Table ID:
html = transform(style, pretty_print=True).replace("\n","")
s= sp(html, "html.parser")
tableId = s.find('table')['id']



#Further Format DataFrame(As outlook and gmail does not support some CSS functions):
collapse = html.replace("table id=\"{}\"".format(tableId), #collapse table
                      "table id=\"{}\" style=\"border-collapse: collapse;\"".format(tableId)) 
noSpace = ' '.join(collapse.split()) #remove whitespaces

#Hightlight odd rows:
for x in range(len(df)):
    
    if x == 0:
        temp = noSpace.replace("<tr style=\"border:1px solid\"> <td id=\"{}row{}_col0\" class=\"data row{} col0\" style=\"border:1px solid\">".format(tableId,x,x),
                   
                  
                  "<tr style=\"border:1px solid;background-color:#DCE6F1\"> <td id=\"{}row{}_col0\" class=\"data row{} col0\" style=\"border:1px solid\">".format(tableId,x,x))
        
    elif x%2 ==0:
        temp2 = temp.replace("<tr style=\"border:1px solid\"> <td id=\"{}row{}_col0\" class=\"data row{} col0\" style=\"border:1px solid\">".format(tableId,x,x),
                   
                  
                  "<tr style=\"border:1px solid;background-color:#DCE6F1\"> <td id=\"{}row{}_col0\" class=\"data row{} col0\" style=\"border:1px solid\">".format(tableId,x,x))
    
        temp = temp2
        formated = temp2



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


