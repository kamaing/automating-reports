import sys
import os # for getting the current path
from win32com.client import constants
import win32com.client

word = win32com.client.DispatchEx("Word.Application")
cwd = os.getcwd() # get current path
doc = word.Documents.Open('C:\\Users\\KamaIng\\real.docx')
doc.TablesOfContents(1).Update()
doc.Close(SaveChanges=True)
word.Quit()
