from asposecells import Settings
from com.aspose.cells import Workbook
from com.aspose.cells import FileFormatType

class HelloWorld:

    def __init__(self):
        dataDir = Settings.dataDir + 'quickstart/'
        
        workbook = Workbook()
        
        sheet = workbook.getWorksheets().get(0)
        
        cell = sheet.getCells().get("A1")
        
        cell.setValue("Hello World!")
        
        file_format_type = FileFormatType
        
        workbook.save(dataDir + "HelloWorld.xls" , file_format_type.EXCEL_97_TO_2003 )
        
        print "Document has been saved, please check the output file.";

if __name__ == '__main__':        
    HelloWorld()