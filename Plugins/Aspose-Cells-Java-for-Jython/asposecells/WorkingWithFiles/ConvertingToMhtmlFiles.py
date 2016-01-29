from asposecells import Settings
from com.aspose.cells import Workbook
from com.aspose.cells import HtmlSaveOptions
from com.aspose.cells import SaveFormat


class ConvertingToMhtmlFiles:

    def __init__(self):
        dataDir = Settings.dataDir + 'WorkingWithFiles/ConvertingToMhtmlFiles/'
        
        saveFormat = SaveFormat

        #Specify the file path
        filePath = dataDir + "Book1.xlsx"

        #Specify the HTML saving options
        sv = HtmlSaveOptions(saveFormat.M_HTML)

        #Instantiate a workbook and open the template XLSX file
        wb = Workbook(filePath)

        #Save the MHT file
        wb.save(filePath + ".out.mht", sv)

        # Print message
        print "Excel to MHTML conversion performed successfully."
 
if __name__ == '__main__':        
    ConvertingToMhtmlFiles()