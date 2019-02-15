import test
import time
import smtplib
import os
import PyPDF2
import glob
import shutil
import sys
import win32com.client 


bookName = "Salaries.xlsx"
masterSheetName = "Master"
passwordColumn = "B"
emailColumn = "C"


def getEmailId(empName, doc):
    global bookName, masterSheetName, passwordColumn
    try:
        sheet = doc.Worksheets(masterSheetName)
        
        matchingCell = sheet.UsedRange.Find(empName)
        
        passwordCell = "{}{}".format(passwordColumn,matchingCell.Address[-1:])
        
        
        password = sheet.Range(passwordCell)
       
        return str(password)
    except Exception as e:
        return ""



def main():
    global bookName, masterSheetName, passwordColumn
    if len(sys.argv) != 4:
        print("Usage:  py PasswordProtectPDFs.py Salaries.xlsx Master D ")
        sys.exit(1)

    bookName = sys.argv[1]
    masterSheetName = sys.argv[2]
    passwordColumn = sys.argv[3]

    path = os.getcwd().replace('\'', '\\') + '\\'
    excel = win32com.client.gencache.EnsureDispatch('Excel.Application')
    doc = excel.Workbooks.Open(path+bookName, ReadOnly=True)
    excel.Visible = False
    
    


    fromaddr = "sairandhree.sule@niyuj.com"
    # creates SMTP session 
    s = smtplib.SMTP('smtp.gmail.com', 587) 

    # start TLS for security 
    s.starttls() 

    # Authentication 
    s.login(fromaddr, "Ananya12") 
    
    pdfFiles = glob.glob("protected"+'\*.pdf')

    print(pdfFiles,path)
    for file in pdfFiles:
        empName = file[:-4]
        empName = empName.split("\\")[-1:][0]
        print(empName)
        
        email = getEmailId(empName,doc)
        print("sending {} to {}".format(file,email))
        test.sendEmail(email, file, s)

    
    doc.Close(SaveChanges=False)
    excel.Quit()


    # terminating the session
    s.quit()


if __name__ == "__main__":
    main()
