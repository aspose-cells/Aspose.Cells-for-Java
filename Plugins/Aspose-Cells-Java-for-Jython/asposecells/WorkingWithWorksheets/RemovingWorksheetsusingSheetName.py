from asposecells import Settings
from com.aspose.cells import Workbook
from java.io import FileInputStream;


class RemovingWorksheetsusingSheetName:

    def __init__(self):
        dataDir = Settings.dataDir + 'WorkingWithWorksheets/RemovingWorksheetsusingSheetName/'
          
        #Creating a file stream containing the Excel file to be opened
        fstream = FileInputStream(dataDir + "Book1.xls");

        #Instantiating a Workbook object with the stream
        workbook = Workbook(fstream);

        #Removing a worksheet using its sheet name
        workbook.getWorksheets().removeAt("Sheet1");

        #Saving the Excel file
        workbook.save(dataDir + "book.out.xls");

        #Closing the file stream to free all resources
        fstream.close();

        #Print Message
        print "Sheet removed successfully.";

if __name__ == '__main__':        
    RemovingWorksheetsusingSheetName()