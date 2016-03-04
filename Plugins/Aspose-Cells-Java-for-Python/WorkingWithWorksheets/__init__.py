__author__ = 'fahadadeel'
import jpype

class AddingWorksheetstoNewExcelFile:

    def __init__(self,dataDir):

        self.dataDir = dataDir
        self.Workbook = jpype.JClass("com.aspose.cells.Workbook")
        self.SaveFormat = jpype.JClass("com.aspose.cells.SaveFormat")

    def main(self): 
            
        workbook = self.Workbook(self.dataDir + "Book1.xls")

        #Adding a new worksheet to the Workbook object
        worksheets = workbook.getWorksheets()

        sheetIndex = worksheets.add()
        worksheet = worksheets.get(sheetIndex)

        #Setting the name of the newly added worksheet
        worksheet.setName("My Worksheet")

        #Saving the Excel file
        workbook.save(self.dataDir + "book.out.xls")

        #Print Message
        print "Sheet added successfully."

class CopyingAndMovingWorksheets:
    
    def __init__(self,dataDir):

        self.dataDir = dataDir
        self.Workbook = jpype.JClass("com.aspose.cells.Workbook")
        self.SaveFormat = jpype.JClass("com.aspose.cells.SaveFormat")
    
    def main(self):
        
        # Copy Worksheets within a Workbook
        self.copy_worksheet()

        # Move Worksheets within Workbook
        self.move_worksheet()
        
    def copy_worksheet(self):  
                
        # Instantiating a Workbook object by excel file path
        workbook = self.Workbook(self.dataDir + "Book1.xls")
        

        # Create a Worksheets object with reference to the sheets of the Workbook.
        sheets = workbook.getWorksheets()

        # Copy data to a new sheet from an existing sheet within the Workbook.
        sheets.addCopy("Sheet1")

        # Saving the modified Excel file in default (that is Excel 2003) format
        workbook.save(self.dataDir + "Copy Worksheet.xls")

        print "Copy worksheet, please check the output file."

    

    def move_worksheet(self):
                
        # Instantiating a Workbook object by excel file path
        workbook = self.Workbook(self.dataDir + "Book1.xls")
    

        # Get the first worksheet in the book.
        sheet = workbook.getWorksheets().get(0)

        # Move the first sheet to the third position in the workbook.
        sheet.moveTo(2)

        # Saving the modified Excel file in default (that is Excel 2003) format
        workbook.save(self.dataDir + "Move_Worksheet.xls")

        print "Move worksheet, please check the output file."
        
class DisplayHideGridlines:
    
    def __init__(self,dataDir):

        self.dataDir = dataDir
        self.Workbook = jpype.JClass("com.aspose.cells.Workbook")
        self.SaveFormat = jpype.JClass("com.aspose.cells.SaveFormat")
        
    def main(self):
             
        workbook = self.Workbook(self.dataDir + "Book1.xls")
        
        #Accessing the first worksheet in the Excel file
        worksheets = workbook.getWorksheets()

        worksheet = worksheets.get(0)

        #Hiding the grid lines of the first worksheet of the Excel file
        worksheet.setGridlinesVisible(False)

        #Saving the modified Excel file in default (that is Excel 2000) format
        workbook.save(self.dataDir + "output.xls")

        # Print message
        print "Grid lines are now hidden on sheet 1, please check the output document."
        
class DisplayHideScrollBars:
    
    def __init__(self,dataDir):

        self.dataDir = dataDir
        self.Workbook = jpype.JClass("com.aspose.cells.Workbook")
        
    def main(self):
             
        workbook = self.Workbook(self.dataDir + "Book1.xls")
        
        #Hiding the vertical scroll bar of the Excel file
        workbook.getSettings().setVScrollBarVisible(False)

        #Hiding the horizontal scroll bar of the Excel file
        workbook.getSettings().setHScrollBarVisible(False)

        #Saving the modified Excel file in default (that is Excel 2003) format
        workbook.save(self.dataDir + "output.xls")

        # Print message
        print "Scroll bars are now hidden, please check the output document."
        
class DisplayHideTabs:
    
    def __init__(self,dataDir):

        self.dataDir = dataDir
        self.Workbook = jpype.JClass("com.aspose.cells.Workbook")
        
    def main(self):
             
        workbook = self.Workbook(self.dataDir + "Book1.xls")
        
        #Hiding the tabs of the Excel file
        workbook.getSettings().setShowTabs(False)

        #Saving the modified Excel file in default (that is Excel 2003) format
        workbook.save(self.dataDir + "output.xls")

        # Print message
        print "Tabs are now hidden, please check the output file."
        
class FreezePanes:
    
    def __init__(self,dataDir):

        self.dataDir = dataDir
        self.Workbook = jpype.JClass("com.aspose.cells.Workbook")
        
    def main(self):
             
        workbook = self.Workbook(self.dataDir + "Book1.xls")
        
        #Accessing the first worksheet in the Excel file
        worksheets = workbook.getWorksheets()
        worksheet = worksheets.get(0)

        #Applying freeze panes settings
        worksheet.freezePanes(3,2,3,2)

        #Saving the modified Excel file in default format
        workbook.save(self.dataDir + "book.out.xls")

        #Print Message
        print "Panes freeze successfull."
        
class HideUnhideWorksheet:
    
    def __init__(self,dataDir):

        self.dataDir = dataDir
        self.Workbook = jpype.JClass("com.aspose.cells.Workbook")
        
    def main(self):
             
        workbook = self.Workbook(self.dataDir + "Book1.xls")
        
        #Accessing the first worksheet in the Excel file
        worksheets = workbook.getWorksheets()
        worksheet = worksheets.get(0)

        #Hiding the first worksheet of the Excel file
        worksheet.setVisible(True)

        #Saving the modified Excel file in default (that is Excel 2003) format
        workbook.save(self.dataDir + "output.xls")

        # Print message
        print "Worksheet 1 is now hidden, please check the output document."
        
class ManagingPageBreaks:
    
    def __init__(self,dataDir):

        self.dataDir = dataDir
        self.Workbook = jpype.JClass("com.aspose.cells.Workbook")
        self.SaveFormat = jpype.JClass("com.aspose.cells.SaveFormat")
        
    def main(self):
        
        # Adding Page Breaks
        self.add_page_breaks()

        # Clearing All Page Breaks
        self.clear_all_page_breaks()

        # Removing Specific Page Break
        self.remove_page_break()
        
    def add_page_breaks(self):
                
        # Instantiating a Workbook object
        workbook = self.Workbook(self.dataDir + "Book1.xls")
    
        worksheets = workbook.getWorksheets()
        worksheet = worksheets.get(0)

        h_page_breaks = worksheet.getHorizontalPageBreaks()
        h_page_breaks.add("Y30")

        v_page_breaks = worksheet.getVerticalPageBreaks()
        v_page_breaks.add("Y30")

        # Saving the modified Excel file in default (that is Excel 2003) format
        workbook.save(self.dataDir + "Add Page Breaks.xls")

        print "Add page breaks, please check the output file."

    

    def clear_all_page_breaks(self):
                
        # Instantiating a Workbook object
        workbook = self.Workbook(self.dataDir + "Book1.xls")
    

        workbook.getWorksheets().get(0).getHorizontalPageBreaks().clear()
        workbook.getWorksheets().get(0).getVerticalPageBreaks().clear()
        
        # Saving the modified Excel file in default (that is Excel 2003) format
        workbook.save(self.dataDir + "Clear All Page Breaks.xls")

        print "Clear all page breaks, please check the output file."

    

    def remove_page_break(self):
            
        # Instantiating a Workbook object
        workbook = self.Workbook(self.dataDir + "Book1.xls")

        worksheets = workbook.getWorksheets()
        worksheet = worksheets.get(0)

        h_page_breaks = worksheet.getHorizontalPageBreaks()
        h_page_breaks.removeAt(0)

        v_page_breaks = worksheet.getVerticalPageBreaks()
        v_page_breaks.removeAt(0)

        # Saving the modified Excel file in default (that is Excel 2003) format
        workbook.save(self.dataDir + "Remove Page Break.xls")

        print "Remove page break, please check the output file."
        
class PageBreakPreview:
    
    def __init__(self,dataDir):

        self.dataDir = dataDir
        self.Workbook = jpype.JClass("com.aspose.cells.Workbook")
        
    def main(self):
             
        workbook = self.Workbook(self.dataDir + "Book1.xls")
        
        #Adding a new worksheet to the Workbook object
        worksheets = workbook.getWorksheets()

        worksheet = worksheets.get(0)

        #Displaying the worksheet in page break preview
        worksheet.setPageBreakPreview(True)

        #Saving the modified Excel file in default format
        workbook.save(self.dataDir + "output.xls")

        # Print message
        print "Page break preview is enabled for sheet 1, please check the output document." 
    
class ProtectingWorksheet:
        
    def __init__(self,dataDir):

        self.dataDir = dataDir
        self.Workbook = jpype.JClass("com.aspose.cells.Workbook")
        self.SaveFormat = jpype.JClass("com.aspose.cells.SaveFormat")
        
    def main(self):
                
        #Instantiating a Excel object by excel file path
        excel = self.Workbook(self.dataDir + "Book1.xls")

        #Accessing the first worksheet in the Excel file
        worksheets = excel.getWorksheets()
        worksheet = worksheets.get(0)

        protection = worksheet.getProtection()

        #The following 3 methods are only for Excel 2000 and earlier formats
        protection.setAllowEditingContent(False)
        protection.setAllowEditingObject(False)
        protection.setAllowEditingScenario(False)

        #Protects the first worksheet with a password "1234"
        protection.setPassword("1234")

        #Saving the modified Excel file in default format
        excel.save(self.dataDir + "output.xls")

        #Print Message
        print "Sheet protected successfully."
        
class RemovingWorksheetsusingSheetIndex:
    
    def __init__(self,dataDir):

        self.dataDir = dataDir
        self.Workbook = jpype.JClass("com.aspose.cells.Workbook")
        self.FileInputStream = jpype.JClass("java.io.FileInputStream")
        
    def main(self):
                  
        fstream=self.FileInputStream(self.dataDir + "Book1.xls");

        #Instantiating a Workbook object with the stream
        workbook = self.Workbook(fstream)

        #Removing a worksheet using its sheet index
        workbook.getWorksheets().removeAt(0)

        #Saving the Excel file
        workbook.save(self.dataDir + "book.out.xls")

        #Closing the file stream to free all resources
        fstream.close()


        #Print Message
        print "Sheet removed successfully."
        
class RemovingWorksheetsusingSheetName:
    
    def __init__(self,dataDir):

        self.dataDir = dataDir
        self.Workbook = jpype.JClass("com.aspose.cells.Workbook")
        self.FileInputStream = jpype.JClass("java.io.FileInputStream")
        
    def main(self):
                  
        #Creating a file stream containing the Excel file to be opened
        fstream = self.FileInputStream(self.dataDir + "Book1.xls");

        #Instantiating a Workbook object with the stream
        workbook = self.Workbook(fstream);

        #Removing a worksheet using its sheet name
        workbook.getWorksheets().removeAt("Sheet1");

        #Saving the Excel file
        workbook.save(self.dataDir + "book.out.xls");

        #Closing the file stream to free all resources
        fstream.close();

        #Print Message
        print "Sheet removed successfully.";
        
class SettingPageOptions:
    
    def __init__(self,dataDir):

        self.dataDir = dataDir
        self.Workbook = jpype.JClass("com.aspose.cells.Workbook")
        self.PageOrientationType = jpype.JClass("com.aspose.cells.PageOrientationType")
        self.FileInputStream = jpype.JClass("java.io.FileInputStream")
        
    def main(self):
        
        self.page_orientation()

        self.scaling()
    
    def page_orientation(self):
        
        # Instantiating a Workbook object by excel file path
        workbook = self.Workbook()

        # Accessing the first worksheet in the Excel file
        worksheets = workbook.getWorksheets()
        sheet_index = worksheets.add()
        sheet = worksheets.get(sheet_index)

        # Setting the orientation to Portrait
        page_setup = sheet.getPageSetup()
        page_orientation_type = self.PageOrientationType
        page_setup.setOrientation(page_orientation_type.PORTRAIT)

        # Saving the modified Excel file in default (that is Excel 2003) format
        workbook.save(self.dataDir + "Page_Orientation.xls")

        print "Set page orientation, please check the output file."
    
    def scaling(self):        
        # Instantiating a Workbook object by excel file path
        workbook = self.Workbook(self.dataDir + 'Book1.xls')

        # Accessing the first worksheet in the Excel file
        worksheets = workbook.getWorksheets()
        sheet_index = worksheets.add()
        sheet = worksheets.get(sheet_index)

        # Setting the scaling factor to 100
        page_setup = sheet.getPageSetup()
        page_setup.setZoom(100)

        # Saving the modified Excel file in default (that is Excel 2003) format
        workbook.save(self.dataDir + "Scaling.xls")

        print "Set scaling, please check the output file."
        
class SplitPanes:
    
    def __init__(self,dataDir):

        self.dataDir = dataDir
        self.Workbook = jpype.JClass("com.aspose.cells.Workbook")
        self.SaveFormat = jpype.JClass("com.aspose.cells.SaveFormat")
        
    def main(self):
                
        saveFormat = self.SaveFormat;
     
        workbook = self.Workbook(self.dataDir + "Book1.xls")

        #Set the active cell
        workbook.getWorksheets().get(0).setActiveCell("A20");

        #Split the worksheet window
        workbook.getWorksheets().get(0).split();

        #Save the excel file
        workbook.save(self.dataDir + "book.out.xls", saveFormat.EXCEL_97_TO_2003);

        #Print Message
        print "Panes split successfully."
        
class UnprotectingPasswordProtectedWorksheet:
    
    def __init__(self,dataDir):

        self.dataDir = dataDir
        self.Workbook = jpype.JClass("com.aspose.cells.Workbook")
        self.SaveFormat = jpype.JClass("com.aspose.cells.SaveFormat")
        self.FileFormatType = jpype.JClass("com.aspose.cells.FileFormatType")
        
    def main(self):
                
        filesFormatType = self.FileFormatType

        #Instantiating a Workbook object
        workbook = self.Workbook(self.dataDir + "book1.xls")

        worksheets = workbook.getWorksheets()
        worksheet = worksheets.get(0)

        protection = worksheet.getProtection()

        #Unprotecting the worksheet with a password
        worksheet.unprotect("aspose")

        # Save the excel file.
        workbook.save(self.dataDir + "output.xls", filesFormatType.EXCEL_97_TO_2003)

        #Print Message
        print "Worksheet unprotected successfully."
    
class UnprotectingSimplyProtectedWorksheet:
    
    def __init__(self,dataDir):

        self.dataDir = dataDir
        self.Workbook = jpype.JClass("com.aspose.cells.Workbook")
        self.SaveFormat = jpype.JClass("com.aspose.cells.SaveFormat")
        self.FileFormatType = jpype.JClass("com.aspose.cells.FileFormatType")
        
    def main(self):
                
        filesFormatType = self.FileFormatType

        #Instantiating a Workbook object
        workbook = self.Workbook(self.dataDir + "Book1.xls")

        worksheets = workbook.getWorksheets()
        worksheet = worksheets.get(0)

        protection = worksheet.getProtection()

        #The following 3 methods are only for Excel 2000 and earlier formats
        protection.setAllowEditingContent(False)
        protection.setAllowEditingObject(False)
        protection.setAllowEditingScenario(False)

        #Unprotecting the worksheet
        worksheet.unprotect()

        # Save the excel file.
        workbook.save(self.dataDir + "output.xls", filesFormatType.EXCEL_97_TO_2003)

        #Print Message
        print "Worksheet unprotected successfully."
        
class ZoomFactor:
    
    def __init__(self,dataDir):

        self.dataDir = dataDir
        self.Workbook = jpype.JClass("com.aspose.cells.Workbook")
        self.SaveFormat = jpype.JClass("com.aspose.cells.SaveFormat")
        
    def main(self):
             
        workbook = self.Workbook(self.dataDir + "Book1.xls")

        #Accessing the first worksheet in the Excel file
        worksheets = workbook.getWorksheets()
        worksheet = worksheets.get(0)

        #Setting the zoom factor of the worksheet to 75
        worksheet.setZoom(75)

        #Saving the modified Excel file in default format
        workbook.save(self.dataDir + "output.xls")

        # Print message
        print "Zoom factor set to 75% for sheet 1, please check the output document."