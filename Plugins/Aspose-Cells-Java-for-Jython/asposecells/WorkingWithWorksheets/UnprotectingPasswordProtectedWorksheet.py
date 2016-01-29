from asposecells import Settings
from com.aspose.cells import Workbook
from com.aspose.cells import SaveFormat
from com.aspose.cells import FileFormatType;


class UnprotectingPasswordProtectedWorksheet:

    def __init__(self):
        dataDir = Settings.dataDir + 'WorkingWithWorksheets/UnprotectingPasswordProtectedWorksheet/'
        
        filesFormatType = FileFormatType

        #Instantiating a Workbook object
        workbook = Workbook(dataDir + "book1.xls")

        worksheets = workbook.getWorksheets()
        worksheet = worksheets.get(0)

        protection = worksheet.getProtection()

        #Unprotecting the worksheet with a password
        worksheet.unprotect("aspose")

        # Save the excel file.
        workbook.save(dataDir + "output.xls", filesFormatType.EXCEL_97_TO_2003)

        #Print Message
        print "Worksheet unprotected successfully."

if __name__ == '__main__':        
    UnprotectingPasswordProtectedWorksheet()