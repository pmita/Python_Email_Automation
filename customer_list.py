import pandas as pd
import smtplib
from email.mime.email_body import MIMEText
from email.mime.multipart import MIMEMultipart

#Enter your credentials
your_email = 'winautomation.example@gmail.com'
your_password = 'Mmita1995@'
smtp_protocol = 'smtp.gmail.com' # This is unique as per email service of your choice

#Initiate SMTP connection to your email service. You can google these if unsure
server = smtplib.SMTP_SSL(smtp_protocol, 465)
server.ehlo()
server.login(your_email, your_password)

#Access excel detaisl with pandas and introduce new classes to store all data.
email_list = pd.read_excel('C:/Users/pmita/Desktop/Python/Email/drafts/list_of_customers.xlsx')
names = email_list['NAME'] #This will return a new class containing names
emails = email_list['EMAIL'] #This will return a new class containing emails
invoices = email_list['INVOICE']
dates = email_list['DATE']


#Loop through each individual item and send email one at a time
for email_index in range(len(emails)): 
    print("Sending Email ... Please Wait \n")

    #Extract data from each line as per email_index
    name = names[email_index]
    email = emails[email_index]
    invoice = invoices[email_index]
    date = dates[email_index]

    #Construct the multipart email with all necessary details
    msg = MIMEMultipart()
    msg['From'] = your_email
    msg['To'] = email
    msg['Subject'] = 'Invoice ' + invoice
    message = 'Dear ' + name + ',\n\n Thank you so much for choosing our service.\n\n Your order was comolpeted on ' + date + 'with a invoice number of ' + invoice + '. Please keep those in hand. \n\n Regards,\n The X team'

    msg.attach(MIMEText(message, 'plain'))
    email_body = msg.as_string()

    #Send the combined email
    server.sendmail(your_email, [email], email_body)
    print("Email was sent to %s \n" % name)

#Shut down the connection with the server once finished
server.close()
