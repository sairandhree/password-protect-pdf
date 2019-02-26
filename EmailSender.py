import smtplib 
from email.mime.multipart import MIMEMultipart 
from email.mime.text import MIMEText 
from email.mime.base import MIMEBase 
from email import encoders 
from dateutil.relativedelta import relativedelta
from datetime import datetime


def sendEmail(to, attachmentFile, s, config):
    today = datetime.now()
    one_month_ago = today - relativedelta(months = 1)
    # print "one month ago date time: %s" % one_month_ago
    last_month_text = one_month_ago.strftime('%B')
    last_year_full = one_month_ago.strftime('%Y')


    #print("Salary SLIP for month of {} {} ".format(last_month_text,last_year_full) )
  

    toaddr = str(config["cc_emailId"])
    
    # instance of MIMEMultipart 
    msg = MIMEMultipart() 
    
    # storing the senders email address   
    msg['From'] = str(config["Email_from_Name"])
    # storing the receivers email address  
    msg['To'] = to 
    
    # storing the subject  
    msg['Subject'] = "Salary SLIP for month of {} {} ".format(last_month_text,last_year_full)
    
    # string to store the body of the mail 
    body = ''' <div>Hi,</div>

    <div> Attached is your password protected salary slip for the month of {} {} and the password to open the file is your <strong> PAN in uppercase.</strong> </div>
    <br/>
    <div> <font color="red"> Note :  Please make sure that the net amount payable in your salary slip is the same as the sum of the amount credited into your bank account and food card.
    </font></div>
    <br/>
    <div>Thanks and Regards,</div>

    <div>Niyuj Finance</div>'''.format(last_month_text,last_year_full)
    
    # attach the body with the msg instance 
    msg.attach(MIMEText(body, 'html', 'utf-8')) 
    
    # open the file to be sent  
    filename = attachmentFile.split("\\")[-1:][0]
    attachment = open(attachmentFile, "rb") 
    
    # instance of MIMEBase and named as p 
    p = MIMEBase('application', 'octet-stream') 
    
    # To change the payload into encoded form 
    p.set_payload((attachment).read()) 
    
    # encode into base64 
    encoders.encode_base64(p) 
    
    p.add_header('Content-Disposition', "attachment; filename= %s" % filename) 
    
    # attach the instance 'p' to instance 'msg' 
    msg.attach(p) 
  
 
    # Converts the Multipart msg into a string 
    text = msg.as_string() 
    
    # sending the mail 
    s.sendmail("Niyuj Finance", toaddr, text) 
    
    