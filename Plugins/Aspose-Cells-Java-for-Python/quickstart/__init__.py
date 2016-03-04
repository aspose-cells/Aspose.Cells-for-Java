__author__ = 'fahadadeel'
import jpype

class HelloWorld:

    def __init__(self,dataDir):

        self.dataDir = dataDir
        self.Workbook = jpype.JClass("com.aspose.cells.Workbook")
        self.FileFormatType = jpype.JClass("com.aspose.cells.FileFormatType")
        
    def main(self):
        
        workbook = self.Workbook()
        
        sheet = workbook.getWorksheets().get(0)
        
        cell = sheet.getCells().get("A1")
        
        cell.setValue("Hello World!")
        
        file_format_type = self.FileFormatType
        
        workbook.save(self.dataDir + "HelloWorld.xls" , file_format_type.EXCEL_97_TO_2003 )
        
        print "Document has been saved, please check the output file.";