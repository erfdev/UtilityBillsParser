import PyPDF2
import os
import tabula

import xlsxwriter
#For pdf2image to work you need to download poppler include the bin/ dir in $PATH
from pdf2image import convert_from_path 


import pdfminer3
from pdfminer3.pdfinterp import PDFResourceManager, PDFPageInterpreter
from pdfminer3.converter import TextConverter
from pdfminer3.layout import LAParams
from pdfminer3.pdfpage import PDFPage
from io import StringIO
from PIL import Image
import os 
from tkinter import *
from tkinter import filedialog

cwd = os.getcwd()
ls  = os.listdir()

print(cwd)
print(ls)

def convert_pdf_to_txt(path, pages=None):
    if not pages:
        pagenums = set()
    else:
        pagenums = set(pages)
    output = StringIO()
    manager = PDFResourceManager()
    converter = TextConverter(manager, output, laparams=LAParams())
    interpreter = PDFPageInterpreter(manager, converter)

    infile = open(path, 'rb')
    for page in PDFPage.get_pages(infile, pagenums):
        interpreter.process_page(page)
    infile.close()
    converter.close()
    text = output.getvalue()
    output.close()
    return text

#######################################################################################################
Month = 'September' #Enter Month Here
Bill = Month+'\document-0.pdf'

text = convert_pdf_to_txt(Bill, pages=[1])

#print(text)
lines = text.splitlines()
#for i in range(len(lines)):
#    print('[',i,']',lines[i])

gas_charge = float(lines[85].replace("$",""))
print("TOTAL GAS CHARGE:", gas_charge)

print("60% BILL", round(gas_charge*.6,2))
print("40% BILL", round(gas_charge*.4,2))

pages = convert_from_path(Bill, 100)

billP = pages[1].crop((40,45,425,465))
#billP.show()
billP.save(Month+'\Bill-0.jpg', 'JPEG')


#######################################################################################################
Bill = Month+'\document-1.pdf'
text = convert_pdf_to_txt(Bill, pages=[0])

#print(text)
lines = text.splitlines()
#for i in range(len(lines)):
#    print('[',i,']',lines[i])


elexicon_charge = float(lines[17].replace("$",""))
print("TOTAL ELECTRICITY CHARGE:", elexicon_charge)

print("60% BILL", round(elexicon_charge*.6,2))
print("40% BILL", round(elexicon_charge*.4,2))

pages = convert_from_path(Bill, 100)
billP = pages[0].crop((550,70,830,350))
#billP.show()
billP.save(Month+'\Bill-1.jpg', 'JPEG')

write_to_wb = False

if write_to_wb :
    workbook = xlsxwriter.Workbook(Month+'\\'+Month+'_Utilities.xlsx')
    cell_format = workbook.add_format({'bold': True, 'italic': True})
    cell_format2 = workbook.add_format({'bold': True, 'fg_color':'#FFFF00'})
    worksheet = workbook.add_worksheet('44Main')
    worksheet.set_column(0,0,25)
    worksheet.write('A1', 'Enbridge Gas')
    worksheet.write('B1', round(gas_charge*.6,2))
    worksheet.write('A2', 'Elexicon Electricity')
    worksheet.write('B2', round(elexicon_charge*.6,2))
    worksheet.write('A3', 'TOTAL DUE',cell_format)
    worksheet.write('B3', '=B1+B2',cell_format2)
    worksheet.insert_image('B5', Month+'\Bill-0.jpg')
    worksheet.insert_image('H5', Month+'\Bill-1.jpg')
    worksheet = workbook.add_worksheet('44Basement')
    worksheet.set_column(0,0,25)
    worksheet.write('A1', 'Enbridge Gas')
    worksheet.write('B1', round(gas_charge*.4,2))
    worksheet.write('A2', 'Elexicon Electricity')
    worksheet.write('B2', round(elexicon_charge*.4,2))
    worksheet.write('A3', 'TOTAL DUE',cell_format)
    worksheet.write('B3', '=B1+B2',cell_format2)
    worksheet.insert_image('B5', Month+'\Bill-0.jpg')
    worksheet.insert_image('H5', Month+'\Bill-1.jpg')

    workbook.close()

class App:
    def __init__(self, master):

        self.numbills = 0
        frame = Frame(master)
        frame.pack()

        
        self.v = StringVar()
        self.v.set("Enter number of utilities")
        self.instructions = Label(frame, textvariable=self.v)
        self.instructions.pack(side=TOP)
        self.button = Button(
            frame, text="QUIT", fg="red", command=frame.quit
            )
        self.button.pack(side=BOTTOM)

        self.buttontext = StringVar()
        self.buttontext.set("Enter")
        self.bill = Button(frame, textvariable=self.buttontext, command=self.set_numbills)
        self.bill.pack(side=RIGHT)

        self.num_bills_e = Entry(frame)
        self.num_bills_e.pack(side=LEFT)
        self.num_bills_e.delete(0,END)
        self.num_bills_e.insert(0,"0")
        


    def set_numbills(self):
        self.numbills = self.num_bills_e.get()
        print ("hi there, everyone! You have entered:",self.numbills)
        self.v.set(('You have entered :'+self.numbills+'\nPlease upload first bill'))
        self.buttontext.set("Upload")
        self.num_bills_e.pack_forget()
        self.bill.pack()
       
        


root = Tk()
root.title("UtilityParser [Alpha0.1 by erfdev]")

app = App(root)


w = Label(root, text="Hello, world!")
w.pack()
#root.filename =  filedialog.askopenfilename(initialdir = "/",title = "Select file",filetypes = (("jpeg files","*.jpg"),("all files","*.*")))
#print (root.filename)

root.mainloop()