from asposecells import Settings
from com.aspose.cells import Workbook


class DisplayHideScrollBars:

    def __init__(self):
        dataDir = Settings.dataDir + 'WorkingWithWorksheets/DisplayHideScrollBars/'
     
        workbook = Workbook(dataDir + "Book1.xls")
        
        #Hiding the vertical scroll bar of the Excel file
        workbook.getSettings().setVScrollBarVisible(False)

        #Hiding the horizontal scroll bar of the Excel file
        workbook.getSettings().setHScrollBarVisible(False)

        #Saving the modified Excel file in default (that is Excel 2003) format
        workbook.save(dataDir + "output.xls")

        # Print message
        print "Scroll bars are now hidden, please check the output document."
 
if __name__ == '__main__':        
    DisplayHideScrollBars()