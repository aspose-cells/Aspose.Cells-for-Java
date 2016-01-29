from asposecells import Settings
from com.aspose.cells import Workbook


class FreezePanes:

    def __init__(self):
        dataDir = Settings.dataDir + 'WorkingWithWorksheets/FreezePanes/'
     
        workbook = Workbook(dataDir + "Book1.xls")
        
        #Accessing the first worksheet in the Excel file
        worksheets = workbook.getWorksheets()
        worksheet = worksheets.get(0)

        #Applying freeze panes settings
        worksheet.freezePanes(3,2,3,2)

        #Saving the modified Excel file in default format
        workbook.save(dataDir + "book.out.xls")

        #Print Message
        print "Panes freeze successfull."

if __name__ == '__main__':        
    FreezePanes()