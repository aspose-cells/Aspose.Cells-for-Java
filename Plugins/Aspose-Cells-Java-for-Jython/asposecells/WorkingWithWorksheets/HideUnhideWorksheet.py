from asposecells import Settings
from com.aspose.cells import Workbook


class HideUnhideWorksheet:

    def __init__(self):
        dataDir = Settings.dataDir + 'WorkingWithWorksheets/HideUnhideWorksheet/'
     
        workbook = Workbook(dataDir + "Book1.xls")
        
        #Accessing the first worksheet in the Excel file
        worksheets = workbook.getWorksheets()
        worksheet = worksheets.get(0)

        #Hiding the first worksheet of the Excel file
        worksheet.setVisible(False)

        #Saving the modified Excel file in default (that is Excel 2003) format
        workbook.save(dataDir + "output.xls")

        # Print message
        print "Worksheet 1 is now hidden, please check the output document."

if __name__ == '__main__':        
    HideUnhideWorksheet()