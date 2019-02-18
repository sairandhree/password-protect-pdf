import EmailSender
import time
import smtplib
import os
import PyPDF2
import glob
import shutil
import sys
import win32com.client 
import configparser


bookName = ""
masterSheetName = ""
emailColumn = ""
config = None


def getEmailId(empName, doc):
    global bookName, masterSheetName, emailColumn
    try:
        sheet = doc.Worksheets(masterSheetName)
        
        matchingCell = sheet.UsedRange.Find(empName)
        
        emailCell = "{}{}".format(emailColumn,matchingCell.Address[-1:])
        
        
        email = sheet.Range(emailCell)
       
        return str(email)
    except Exception as e:
        print(str(e))
        return ""



def main():
    global bookName, masterSheetName, emailColumn, config
    config = configparser.ConfigParser()
    config.read('config.ini')
    defConfig = config["DEFAULT"]

    bookName = str(defConfig["Workbook"])
    masterSheetName = str(defConfig["Master_Sheet"])
    emailColumn = str(defConfig["EmailId_column"])

    path = os.getcwd().replace('\'', '\\') + '\\'
    excel = win32com.client.gencache.EnsureDispatch('Excel.Application')
    doc = excel.Workbooks.Open(path+bookName, ReadOnly=True)
    excel.Visible = False
    
    fromaddr = str(defConfig["Sender_EmailId"])
    print(fromaddr)
    # creates SMTP session 
    s = smtplib.SMTP('smtp.gmail.com', 587) 


    # start TLS for security 
    s.starttls() 
 
    # Authentication 
    s.login(fromaddr, str(defConfig["Sender_Email_Password"])) 
    
    pdfFiles = glob.glob("protected"+'\*.pdf')

    print(pdfFiles,path)
    for file in pdfFiles:
        empName = file[:-4]
        empName = empName.split("\\")[-1:][0]
        
        email = getEmailId(empName,doc)
        print("sending {} to {}".format(file,email))
        EmailSender.sendEmail(email, file, s, defConfig)

    
    doc.Close(SaveChanges=False)
    excel.Quit()


    # terminating the session
    s.quit()


if __name__ == "__main__":
    main()
