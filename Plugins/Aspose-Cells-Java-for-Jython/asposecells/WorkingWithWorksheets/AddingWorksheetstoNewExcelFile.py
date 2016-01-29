from asposecells import Settings
from com.aspose.cells import Workbook
from com.aspose.cells import SaveFormat


class AddingWorksheetstoNewExcelFile:

    def __init__(self):
        dataDir = Settings.dataDir + 'WorkingWithWorksheets/AddingWorksheetstoNewExcelFile/'
     
        workbook = Workbook(dataDir + "Book1.xls")

        #Adding a new worksheet to the Workbook object
        worksheets = workbook.getWorksheets()

        sheetIndex = worksheets.add()
        worksheet = worksheets.get(sheetIndex)

        #Setting the name of the newly added worksheet
        worksheet.setName("My Worksheet")

        #Saving the Excel file
        workbook.save(dataDir + "book.out.xls")

        #Print Message
        print "Sheet added successfully."

if __name__ == '__main__':        
    AddingWorksheetstoNewExcelFile()