# Reading an excel file using Python 
import xlwings as xw
import os
import argparse  
import PyPDF2
import glob

def getPassword(name):
    
    bookName = 'Salaries.xlsm'
    sheetName = 'Master'
    
    
    try:
        wb = xw.Book(bookName)
        sht = wb.sheets[sheetName]
        myCell = wb.sheets[sheetName].api.UsedRange.Find(name)
        print('---------------')
        password = sht.range('B'+str(myCell.row)).value
        print ("retriving password for ",name)
        return password
    except Exception:   
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
   




def main():
    pdfFiles = glob.glob('*.pdf')
    for file in pdfFiles :
        empName  = file[:-4]
        print(empName)
       
        set_password(file, getPassword(empName) , "MasterPassword")



if __name__ == "__main__":
    main()