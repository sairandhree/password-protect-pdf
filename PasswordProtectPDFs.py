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
        print (password)
        return password
    except Exception   
        return ""



def set_password(input_file, user_pass, owner_pass):
    """
    Function creates new temporary pdf file with same content,
    assigns given password to pdf and rename it with original file.
    """
   
    try :
        path, filename = os.path.split(input_file)
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
        pass
    # Rename temporary output file with original filename, this
    # will automatically delete temporary file
    #os.rename(output_file, input_file)




def main():
    pdfFiles = glob.glob('*.pdf')
    for file in pdfFiles :
        empName  = file[:-4]
        print(empName)
       
        set_password(file, getPassword(empName) , "MasterPassword")



if __name__ == "__main__":
    main()