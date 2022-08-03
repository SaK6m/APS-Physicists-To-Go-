'''
    Python code to send email to a list of emails from a spreadsheet
'''

import maskpass 
import pandas as pd
import smtplib 
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

#---- link to survey ---

excel = input('Enter the name of the file (with the extension): ')
# reading the spreadsheet
email_list = pd.read_excel(excel, sheet_name = 'matched group') # Email testing excel.xlsx

print('-----------------------')
print("let's login into your email...")
user = input('Enter your email address: ')
password = maskpass.askpass(prompt="Password: ", mask="*")

# your email details
SERVER = 'smtp.gmail.com'  # your smtp server
PORT  = 587    # your port number
FROM  = user    # your from email id
PASS  = password # your email id password  ffhzaategahnoapo

# Authentication part
server = smtplib.SMTP(SERVER,PORT)

server.set_debuglevel(1) # set 1 to help in debugging
server.ehlo()
server.starttls() # start TLS connection which is secure connection

server.login(FROM,PASS)

print('Getting the names and the emails................')
# getting the names and the emails
teacher_names = email_list['Teacher_name']
teacher_emails = email_list['Teacher_email']
physicists_names = email_list['Physicist_name']
physicists_emails = email_list['Physicist_email']
common_topics = email_list['Common Topics']
priority_numbers = email_list['Priority Number']
months = email_list['Month Matched']
meeting_preferences = email_list['Type of Visit(T)']
school_lvls = email_list['School Level(T)']

# email composing
print('Composing Email................')

# iterate through the records
for i in range(len(teacher_emails)):
  
    # for every record get the name and the email addresses
    teacher_name = teacher_names[i]
    physicist_name = physicists_names[i]
    to = teacher_emails[i]
    cc = physicists_emails[i]

    common_topic = common_topics[i]
    priority_number = priority_numbers[i]
    month = months[i]
    meeting = meeting_preferences[i]
    school = school_lvls[i]

    #----- Email body ------
    msg = MIMEMultipart('alternative')
    msg['Subject'] = "You've been matched with Physicist To-Go"
    msg['To'] = to
    msg['Cc'] = cc

    html_message = """\
        <head> </head>
        <body>
            <p>
                Hi {teacher_name}, <br>
                {physicist_name} will be the physicists visiting your class! Your prioritized # {priority_number} topic: {common_topic}
                has been selected to match you with {physicist_name} over a {meeting} for your selected month {month}. The next step is for you to reach out to 
                {physicist_name} to pick a day/time for your visit. They are CC'd to this email. <br>
                
                More information: <br>
                    <ul>
                    <li>Example 1</li>
                    <li>Example 2</li>
                    <li>Example 3</li>
                    </ul>                

                Once you've scheduled your visit, let us know here <a href="www.google.com"> link here </a> so we can follow up with you and see how it went. <br>

                If you have any questions or need additional information, please contact us at physicscentral@aps.org <br>
                <br>
                Thanks and we look forward to hearing about your visit! <br>
                Physicist To-Go Team
            </p>
        </body>
        """.format(**locals())

    rcpt = [to] + [cc]

    part = MIMEText(html_message, 'html')

    msg.attach(part)

    # sending the email
    server.sendmail(FROM, rcpt, msg.as_string())

# close the smtp server 
server.close()

print("Email sent!")
print("Thank you for using this.....")