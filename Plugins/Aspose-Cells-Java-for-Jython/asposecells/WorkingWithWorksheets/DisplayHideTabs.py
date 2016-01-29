from asposecells import Settings
from com.aspose.cells import Workbook


class DisplayHideTabs:

    def __init__(self):
        dataDir = Settings.dataDir + 'WorkingWithWorksheets/DisplayHideTabs/'
     
        workbook = Workbook(dataDir + "Book1.xls")
        
        #Hiding the tabs of the Excel file
        workbook.getSettings().setShowTabs(False)

        #Saving the modified Excel file in default (that is Excel 2003) format
        workbook.save(dataDir + "output.xls")

        # Print message
        print "Tabs are now hidden, please check the output file."

if __name__ == '__main__':        
    DisplayHideTabs()