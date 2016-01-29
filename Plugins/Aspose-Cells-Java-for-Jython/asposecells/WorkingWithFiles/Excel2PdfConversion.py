from asposecells import Settings
from com.aspose.cells import Workbook
from com.aspose.cells import SaveFormat



class Excel2PdfConversion:

    def __init__(self):
        dataDir = Settings.dataDir + 'WorkingWithFiles/Excel2PdfConversion/'
        
        saveFormat = SaveFormat

        workbook = Workbook(dataDir + "Book1.xls")

        #Save the document in PDF format
        workbook.save(dataDir + "OutBook1.pdf", saveFormat.PDF)

        # Print message
        print "\n Excel to PDF conversion performed successfully."
 
if __name__ == '__main__':        
    Excel2PdfConversion()