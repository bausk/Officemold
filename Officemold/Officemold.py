import win32com.client
import glob

Application = win32com.client.Dispatch("Excel.Application")
print('Hello World')
Workbook = Application.Workbooks.Add()

Base = Workbook.ActiveSheet

wordapp = win32com.client.Dispatch("Word.Application") # Create new Word Object
wordapp.Visible = 0 # Word Application should`t be visible
worddoc = wordapp.Documents.Add() # Create new Document Object
worddoc.PageSetup.Orientation = 1 # Make some Setup to the Document:
worddoc.PageSetup.LeftMargin = 20
worddoc.PageSetup.TopMargin = 20
worddoc.PageSetup.BottomMargin = 20
worddoc.PageSetup.RightMargin = 20
worddoc.Content.Font.Size = 11
worddoc.Tables.Add()
worddoc.Close()
wordapp.Quit()