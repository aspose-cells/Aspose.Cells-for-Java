from asposecells import Settings
from com.aspose.cells import Workbook
from com.aspose.cells import SaveFormat


class SplitPanes:

    def __init__(self):
        dataDir = Settings.dataDir + 'WorkingWithWorksheets/SplitPanes/'
        
        saveFormat = SaveFormat;
     
        workbook = Workbook(dataDir + "Book1.xls")

        #Set the active cell
        workbook.getWorksheets().get(0).setActiveCell("A20");

        #Split the worksheet window
        workbook.getWorksheets().get(0).split();

        #Save the excel file
        workbook.save(dataDir + "book.out.xls", saveFormat.EXCEL_97_TO_2003);

        #Print Message
        print "Panes split successfully."

if __name__ == '__main__':        
    SplitPanes()