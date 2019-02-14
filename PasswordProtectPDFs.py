import os
import PyPDF2
import glob

import sys
import win32com.client 


bookName = "Salaries.xlsx"
masterSheetName = "Master"
passwordColumn = "B"
emailColumn = "C"


def set_password(input_file, user_pass, owner_pass):
    
    if user_pass == "": 
        return

    try:
        path, filename = os.path.split(input_file)
        path = os.getcwd().replace('\'', '\\') + '\\'
        print("path ",path , filename)

        if not os.path.exists("protected"):
            os.makedirs("protected")

        output_file = os.path.join(path+"protected",   filename)

        output = PyPDF2.PdfFileWriter()
        
        input_stream = PyPDF2.PdfFileReader(open(input_file, "rb"))
        
        for i in range(0, input_stream.getNumPages()):
            output.addPage(input_stream.getPage(i))
        
        outputStream = open(output_file, "wb")
        # Set user and owner password to pdf file
        output.encrypt(user_pass, owner_pass, use_128bit=True)
        output.write(outputStream)
        
        outputStream.close()
        
    except Exception as e:
        print('Exception setting password for ',str(e), input_file)
        pass


def exportToPdf(doc, masterSheetName):
    path = os.getcwd().replace('\'', '\\') + '\\'
    for x in range(0, len(doc.Worksheets)):
        try:
            sheet = doc.Worksheets[x]
            #sheet.PageSetup.PrintGridLines = 1
            #print(sheet.Name)
            # 57 is PDF format even though it isn't listed as such in Microsofts documentation.
            if sheet.Name != masterSheetName:
                sheet.SaveAs(path+sheet.Name+".pdf",  FileFormat=57)
        except:
           pass

    pdfFiles = glob.glob('*.pdf')

    for file in pdfFiles:
        empName = file[:-4]
        #print(empName)
        password = getPassword(empName,doc)
        print("password $$$$$$$$$$", password)
        set_password(file, password, "MasterPassword")


def getPassword(empName, doc):
    global bookName, masterSheetName, passwordColumn
    try:
        sheet = doc.Worksheets(masterSheetName)
        #print("sheet ###############", empName)
        matchingCell = sheet.UsedRange.Find(empName)
        
        passwordCell = "{}{}".format(passwordColumn,matchingCell.Address[-1:])
        #print("password cell", passwordCell)
        
        password = sheet.Range(passwordCell)
        #print("password is ",empName, password)
        return str(password)
    except Exception as e:
        print ("exception getting password",str(e))
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
    excel =  win32com.client.gencache.EnsureDispatch('Excel.Application')
    doc = excel.Workbooks.Open(path+bookName, ReadOnly=True)

    excel.Visible = False
    exportToPdf(doc,  masterSheetName)

    doc.Close(SaveChanges=False)
    excel.Quit()


if __name__ == "__main__":
    main()
