from asposecells import Settings
from com.aspose.cells import Workbook


class DisplayHideGridlines:

    def __init__(self):
        dataDir = Settings.dataDir + 'WorkingWithWorksheets/DisplayHideGridlines/'
     
        workbook = Workbook(dataDir + "Book1.xls")
        
        #Accessing the first worksheet in the Excel file
        worksheets = workbook.getWorksheets()

        worksheet = worksheets.get(0)

        #Hiding the grid lines of the first worksheet of the Excel file
        worksheet.setGridlinesVisible(False)

        #Saving the modified Excel file in default (that is Excel 2000) format
        workbook.save(dataDir + "output.xls")

        # Print message
        print "Grid lines are now hidden on sheet 1, please check the output document."
 
if __name__ == '__main__':        
    DisplayHideGridlines()