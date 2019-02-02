# Reading an excel file using Python 
import xlwings as xw
import os
import PyPDF2
import glob
import win32com.client 

def getPassword(empName,bookName,sheetName):
    

    try:
        wb = xw.Book(bookName)
        sht = wb.sheets[sheetName]
        myCell = wb.sheets[sheetName].api.UsedRange.Find(empName)
        password = sht.range('B'+str(myCell.row)).value
        print ("retriving password for ",empName)
        return password
    except Exception:   
        print("exceptions dfsd")
        return ""



def set_password(input_file, user_pass, owner_pass):
   
  
    print(" protecting ", input_file)

    try :
        path, filename = os.path.split(input_file)

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
    except Exception:
        print('Exception 1')
        pass
   

def exportToPdf( bookName = 'Salaries.xlsx'):
        
    excel = win32com.client.Dispatch('Excel.Application')
    excel.Visible = False
    
    path = os.getcwd().replace('\'', '\\') + '\\'
    doc = excel.WorkBooks.Open(path+bookName, ReadOnly=True)

    for x in range(0,len(doc.Worksheets)):
        try:
            sheet = doc.Worksheets[x]
            sheet.PageSetup.PrintGridLines = 1
            print(sheet.name)
            # 57 is PDF format even though it isn't listed as such in Microsofts documentation.
            if sheet.name != "Master":
                sheet.SaveAs(path+sheet.name+".pdf",  FileFormat=57)
        except:
            pass

    doc.Close(SaveChanges=False)
    excel.Quit()





def main():
    bookName = 'Salaries.xlsx'
    sheetName = 'Master'
    exportToPdf(bookName)
    pdfFiles = glob.glob('*.pdf')

    for file in pdfFiles :
        empName  = file[:-4]
        print(empName)
       
        set_password(file, getPassword(empName,bookName,sheetName) , "MasterPassword")



if __name__ == "__main__":
    main()