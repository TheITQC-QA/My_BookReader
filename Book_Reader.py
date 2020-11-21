import pyttsx3
import PyPDF2

path = "E://PythonStuff//pywinauto.pdf"   #You can add your file path/location here
reader_init = pyttsx3.init()
file=open(path,"rb")
reader=PyPDF2.PdfFileReader(file)
pages = reader.numPages
page1=reader.getPage(8)
pdfData=page1.extractText()
reader_init.say(pdfData)
reader_init.runAndWait()
