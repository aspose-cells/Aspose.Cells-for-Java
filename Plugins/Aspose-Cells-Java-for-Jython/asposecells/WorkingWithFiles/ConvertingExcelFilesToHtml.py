from asposecells import Settings
from com.aspose.cells import Workbook
from com.aspose.cells import SaveFormat


class ConvertingExcelFilesToHtml:

    def __init__(self):
        dataDir = Settings.dataDir + 'WorkingWithFiles/ConvertingExcelFilesToHtml/'
        
        saveFormat = SaveFormat

        workbook = Workbook(dataDir + "Book1.xls")

        #Save the document in PDF format
        workbook.save(dataDir + "OutBook1.html", saveFormat.HTML)

        # Print message
        print "\n Excel to HTML conversion performed successfully."
 
if __name__ == '__main__':        
    ConvertingExcelFilesToHtml()