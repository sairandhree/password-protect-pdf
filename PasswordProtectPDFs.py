import os
import PyPDF2
import glob
import shutil
import sys
import win32com.client 
import configparser


bookName = ""
masterSheetName = "aster"
passwordColumn = ""
emailColumn = ""
config = None


def set_password(input_file, user_pass, owner_pass):
    
    if user_pass == "": 
        return

    try:
        path, filename = os.path.split(input_file)
        path = os.getcwd().replace('\'', '\\') + '\\'
       
        if not os.path.exists("protected"):
            os.makedirs("protected")

        output_file = os.path.join(path+"protected",   filename)

        output = PyPDF2.PdfFileWriter()
        inputFile = open(input_file,"rb")

        input_stream = PyPDF2.PdfFileReader(inputFile)
        
        for i in range(0, input_stream.getNumPages()):
            output.addPage(input_stream.getPage(i))
        
        outputStream = open(output_file, "wb")
        # Set user and owner password to pdf file
        output.encrypt(user_pass, owner_pass, use_128bit=True)
        output.write(outputStream)
        
        outputStream.close()
        inputFile.close()
        
    except Exception as e:
        print('Exception setting password for ',str(e), input_file)
        pass


def getPassword(empName, doc):
    global bookName, masterSheetName, passwordColumn
    try:
        sheet = doc.Worksheets(masterSheetName)
        
        matchingCell = sheet.UsedRange.Find(empName)
        
        passwordCell = "{}{}".format(passwordColumn,matchingCell.Address[-1:])
        
        
        password = sheet.Range(passwordCell)
       
        return str(password)
    except Exception as e:
        print("exception", str(e))
        return ""


def exportToPdf(doc, masterSheetName):
    path = os.getcwd().replace('\'', '\\') + '\\'
    for x in range(1, len(doc.Worksheets)+1):
        try:
            sheet = doc.Worksheets[x]
            #sheet.PageSetup.PrintGridLines = 1
            #print(sheet.Name)
            # 57 is PDF format even though it isn't listed as such in Microsofts documentation.
            if sheet.Name != masterSheetName:
                sheet.SaveAs(path+sheet.Name+".pdf",  FileFormat=57)
        except Exception as e:
            print("exception ", str(e))
            pass

    pdfFiles = glob.glob('*.pdf')

    for file in pdfFiles:
        empName = file[:-4]
      
        password = getPassword(empName,doc)

        set_password(file, password, "MasterPassword")

    
def moveUnprotectedFiles():
    path = os.getcwd().replace('\'', '\\') + '\\'

    destination = "unprotected"
    if not os.path.exists(destination):
        os.makedirs(destination)

    pdfFiles = glob.glob('*.pdf')

    for f in pdfFiles:
        shutil.move(f, path+destination)

def main():
    global bookName, masterSheetName, passwordColumn, config

    config = configparser.ConfigParser()
    config.read('config.ini')
    defConfig = config["DEFAULT"]


    bookName = str(defConfig["Workbook"])
    masterSheetName = str(defConfig["Master_Sheet"])
    passwordColumn = str(defConfig["Password_Column"])

    print(bookName, masterSheetName, passwordColumn)

    path = os.getcwd().replace('\'', '\\') + '\\'
    excel =  win32com.client.gencache.EnsureDispatch('Excel.Application')
    doc = excel.Workbooks.Open(path+bookName, ReadOnly=True)

    excel.Visible = False
 



    exportToPdf(doc,  masterSheetName)

    moveUnprotectedFiles()


    
    doc.Close(SaveChanges=False)
    excel.Quit()


if __name__ == "__main__":
    main()
