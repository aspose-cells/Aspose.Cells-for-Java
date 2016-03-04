__author__ = 'fahadadeel'
import jpype

class ChartToImage:

    def __init__(self,dataDir):

        self.dataDir = dataDir
        self.Workbook = jpype.JClass("com.aspose.cells.Workbook")
        self.ChartType = jpype.JClass("com.aspose.cells.ChartType")
        self.ImageOrPrintOptions = jpype.JClass("com.aspose.cells.ImageOrPrintOptions")
        self.ImageFormat = jpype.JClass("com.aspose.cells.ImageFormat")
        self.FileOutputStream = jpype.JClass("java.io.FileOutputStream")
        self.Color = jpype.JClass("java.awt.Color")

    def main(self):

        chartType = self.ChartType
        color = self.Color
        imageFormat = self.ImageFormat

        #Create a Workbook+
        workbook = self.Workbook()

        #Get the first worksheet+
        sheet = workbook.getWorksheets().get(0)

        #Set the name of worksheet
        sheet.setName("Data")

        #Get the cells collection in the sheet+
        cells = workbook.getWorksheets().get(0).getCells()

        #Put some values into a cells of the Data sheet+
        cells.get("A1").setValue("Region")
        cells.get("A2").setValue("France")
        cells.get("A3").setValue("Germany")
        cells.get("A4").setValue("England")
        cells.get("A5").setValue("Sweden")
        cells.get("A6").setValue("Italy")
        cells.get("A7").setValue("Spain")
        cells.get("A8").setValue("Portugal")
        cells.get("B1").setValue("Sale")
        cells.get("B2").setValue(70000)
        cells.get("B3").setValue(55000)
        cells.get("B4").setValue(30000)
        cells.get("B5").setValue(40000)
        cells.get("B6").setValue(35000)
        cells.get("B7").setValue(32000)
        cells.get("B8").setValue(10000)

        #Create chart
        chartIndex = sheet.getCharts().add(chartType.COLUMN, 12, 1, 33, 12)
        chart = sheet.getCharts().get(chartIndex)

        #Set properties of chart title
        chart.getTitle().setText("Sales By Region")
        chart.getTitle().getFont().setBold(True)
        chart.getTitle().getFont().setSize(12)

        #Set properties of nseries
        chart.getNSeries().add("Data!B2:B8", True)
        chart.getNSeries().setCategoryData("Data!A2:A8")

        #Set the fill colors for the series's data points (France - Portugal(7 points))
        chartPoints = chart.getNSeries().get(0).getPoints()

        point = chartPoints.get(0)
        #print(self.Color.getWhite())
        
        point.getArea().setForegroundColor(self.Color.white())
        
        point = chartPoints.get(1)
        point.getArea().setForegroundColor(self.Color.getBlue())

        point = chartPoints.get(2)
        point.getArea().setForegroundColor(self.Color.getYellow())

        point = chartPoints.get(3)
        point.getArea().setForegroundColor(self.Color.getRed())

        point = chartPoints.get(4)
        point.getArea().setForegroundColor(self.Color.getBlack())

        point = chartPoints.get(5)
        point.getArea().setForegroundColor(self.Color.getGreen())

        point = chartPoints.get(6)
        point.getArea().setForegroundColor(self.Color.getMaroon())

        #Set the legend invisible
        chart.setShowLegend(false)



        #Get the Chart image
        imgOpts = self.ImageOrPrintOptions()
        imgOpts.setImageFormat(imageFormat.getEmf())

        fs = FileOutputStream(dataDir + "Chart.emf")

        #Save the chart image file+
        chart.toImage(fs, imgOpts)

        fs.close()

        # Print message
        print("<BR>")
        print("Processing performed successfully")
        
class ConvertingExcelFilesToHtml:
    
    def __init__(self,dataDir):

        self.dataDir = dataDir
        self.Workbook = jpype.JClass("com.aspose.cells.Workbook")
        self.SaveFormat = jpype.JClass("com.aspose.cells.SaveFormat")
          
    
    def main(self):
             
        saveFormat = self.SaveFormat

        workbook = self.Workbook(self.dataDir + "Book1.xls")

        #Save the document in PDF format
        workbook.save(self.dataDir + "OutBook1.html", saveFormat.HTML)

        # Print message
        print "\n Excel to HTML conversion performed successfully."

class ConvertingToMhtmlFiles:
    
    def __init__(self,dataDir):
        self.dataDir = dataDir
        self.Workbook = jpype.JClass("com.aspose.cells.Workbook")
        self.HtmlSaveOptions = jpype.JClass("com.aspose.cells.HtmlSaveOptions")
        self.SaveFormat = jpype.JClass("com.aspose.cells.SaveFormat")

    def main(self):
               
        saveFormat = self.SaveFormat

        #Specify the file path
        filePath = self.dataDir + "Book1.xlsx"

        #Specify the HTML saving options
        sv = self.HtmlSaveOptions(saveFormat.M_HTML)

        #Instantiate a workbook and open the template XLSX file
        wb = self.Workbook(filePath)

        #Save the MHT file
        wb.save(filePath + ".out.mht", sv)

        # Print message
        print "Excel to MHTML conversion performed successfully."
        
class ConvertingToXPS:
    
    def __init__(self,dataDir):
        self.dataDir = dataDir
        self.Workbook = jpype.JClass("com.aspose.cells.Workbook")
        self.ImageFormat = jpype.JClass("com.aspose.cells.ImageFormat")
        self.ImageOrPrintOptions = jpype.JClass("com.aspose.cells.ImageOrPrintOptions")
        self.SheetRender = jpype.JClass("com.aspose.cells.SheetRender")
        self.SaveFormat = jpype.JClass("com.aspose.cells.SaveFormat")
        
    def main(self):
               
        saveFormat = self.SaveFormat

        workbook = self.Workbook(self.dataDir + "Book1.xls")

        #Get the first worksheet.
        sheet = workbook.getWorksheets().get(0)

        #Apply different Image and Print options
        options = self.ImageOrPrintOptions()

        #Set the Format
        options.setSaveFormat(saveFormat.XPS)

        # Render the sheet with respect to specified printing options
        sr = self.SheetRender(sheet, options)
        sr.toImage(0, self.dataDir + "out_printingxps.xps")

        #Save the complete Workbook in XPS format
        workbook.save(self.dataDir + "out_whole_printingxps", saveFormat.XPS)

        # Print message
        print "Excel to XPS conversion performed successfully."
        
class ConvertingWorksheetToSVG:
    
    def __init__(self,dataDir):
        self.dataDir = dataDir
        self.Workbook = jpype.JClass("com.aspose.cells.Workbook")
        self.ImageFormat = jpype.JClass("com.aspose.cells.ImageFormat")
        self.ImageOrPrintOptions = jpype.JClass("com.aspose.cells.ImageOrPrintOptions")
        self.SheetRender = jpype.JClass("com.aspose.cells.SheetRender")
        self.SaveFormat = jpype.JClass("com.aspose.cells.SaveFormat")
        
    def main(self):
         
        saveFormat = self.SaveFormat

        workbook = self.Workbook(self.dataDir + "Book1.xls")

        #Convert each worksheet into svg format in a single page.
        imgOptions = ImageOrPrintOptions()
        imgOptions.setSaveFormat(saveFormat.SVG)
        imgOptions.setOnePagePerSheet(True)

        #Convert each worksheet into svg format
        sheetCount = workbook.getWorksheets().getCount()

        #for(i=0; i<sheetCount; i++)
        for i in range(sheetCount):
        
            sheet = workbook.getWorksheets().get(i)

            sr = SheetRender(sheet, imgOptions)

            pageCount = sr.getPageCount()
            #for (k = 0 k < pageCount k++)
            for k in range(pageCount):
            
                #Output the worksheet into Svg image format
                sr.toImage(k, self.dataDir + sheet.getName() + ".out.svg")
            
        

        # Print message
        print "Excel to SVG conversion completed successfully."
        
class ConvertingWorksheetToSVG:
    
    def __init__(self,dataDir):
        self.dataDir = dataDir
        self.Workbook = jpype.JClass("com.aspose.cells.Workbook")
        self.ImageFormat = jpype.JClass("com.aspose.cells.ImageFormat")
        self.ImageOrPrintOptions = jpype.JClass("com.aspose.cells.ImageOrPrintOptions")
        self.SheetRender = jpype.JClass("com.aspose.cells.SheetRender")
        self.SaveFormat = jpype.JClass("com.aspose.cells.SaveFormat")
        
    def main(self):
               
        saveFormat = self.SaveFormat

        workbook = self.Workbook(self.dataDir + "Book1.xls")

        #Convert each worksheet into svg format in a single page.
        imgOptions = self.ImageOrPrintOptions()
        imgOptions.setSaveFormat(saveFormat.SVG)
        imgOptions.setOnePagePerSheet(True)

        #Convert each worksheet into svg format
        sheetCount = workbook.getWorksheets().getCount()

        #for(i=0; i<sheetCount; i++)
        for i in range(sheetCount):
        
            sheet = workbook.getWorksheets().get(i)

            sr = self.SheetRender(sheet, imgOptions)

            pageCount = sr.getPageCount()
            #for (k = 0 k < pageCount k++)
            for k in range(pageCount):
            
                #Output the worksheet into Svg image format
                sr.toImage(k, self.dataDir + sheet.getName() + ".out.svg")
            
        

        # Print message
        print "Excel to SVG conversion completed successfully."
        
class Excel2PdfConversion:
    
    def __init__(self,dataDir):
        self.dataDir = dataDir
        self.Workbook = jpype.JClass("com.aspose.cells.Workbook")
        self.SaveFormat = jpype.JClass("com.aspose.cells.SaveFormat")
    
    def main(self):
                
        saveFormat = self.SaveFormat

        workbook = self.Workbook(self.dataDir + "Book1.xls")

        #Save the document in PDF format
        workbook.save(self.dataDir + "OutBook1.pdf", saveFormat.PDF)

        # Print message
        print "\n Excel to PDF conversion performed successfully."
        
class ManagingDocumentProperties:
    
    def __init__(self,dataDir):
        self.dataDir = dataDir
        self.Workbook = jpype.JClass("com.aspose.cells.Workbook")
        self.SaveFormat = jpype.JClass("com.aspose.cells.SaveFormat")
        
    def main(self):

        workbook = self.Workbook(self.dataDir + "Book1.xls")

        #Retrieve a list of all custom document properties of the Excel file
        customProperties = workbook.getWorksheets().getCustomDocumentProperties()

        #Accessing a custom document property by using the property index
        #customProperty1 = customProperties.get(3)

        #Accessing a custom document property by using the property name
        customProperty2 = customProperties.get("Owner")


        #Adding a custom document property to the Excel file
        publisher = customProperties.add("Publisher", "Aspose")

        #Save the file
        workbook.save(self.dataDir + "Test_Workbook.xls")

        #Removing a custom document property
        customProperties.remove("Publisher")

        #Save the file
        workbook.save(self.dataDir + "Test_Workbook_RemovedProperty.xls")

        # Print message
        print "Excel file's custom properties accessed successfully."
    
class OpeningFiles:
    
    def __init__(self,dataDir):
        self.dataDir = dataDir
        self.Workbook = jpype.JClass("com.aspose.cells.Workbook")
        self.FileFormatType = jpype.JClass("com.aspose.cells.FileFormatType")
        self.LoadOptions = jpype.JClass("com.aspose.cells.LoadOptions")
        self.FileInputStream = jpype.JClass("java.io.FileInputStream")
        
    def main(self):
        
        fileFormatType = self.FileFormatType
        # 1. Opening from path
        # Creatin an Workbook object with an Excel file path
        workbook1 = self.Workbook(self.dataDir + "Book1.xls")

        print "Workbook opened using path successfully.";
        
        # 2 Opening workbook from stream
        
        #Create a Stream object
        fstream = self.FileInputStream(self.dataDir + "Book2.xls")
        #Creating an Workbook object with the stream object
        workbook2 = self.Workbook(fstream)
        fstream.close()

        print ("Workbook opened using stream successfully.");
        
        # 3.
        # Opening Microsoft Excel 97 Files
        #Createing and EXCEL_97_TO_2003 LoadOptions object
        loadOptions1 = self.LoadOptions(fileFormatType.EXCEL_97_TO_2003)
        #Creating an Workbook object with excel 97 file path and the loadOptions object
        workbook3 = self.Workbook(self.dataDir + "Book_Excel97_2003.xls", loadOptions1)
        # Print message
        print("Excel 97 Workbook opened successfully.");
        
        # 4.
        # Opening Microsoft Excel 2007 XLSX Files
        #Createing and XLSX LoadOptions object
        loadOptions2 = self.LoadOptions(fileFormatType.XLSX)
        #Creating an Workbook object with 2007 xlsx file path and the loadOptions object
        workbook4 = self.Workbook(self.dataDir + "Book_Excel2007.xlsx", loadOptions2)
        # Print message
        print ("Excel 2007 Workbook opened successfully.")
        
        
        # 5.
        # Opening SpreadsheetML Files
        #Creating and EXCEL_2003_XML LoadOptions object
        loadOptions3 = self.LoadOptions(fileFormatType.EXCEL_2003_XML)
        #Creating an Workbook object with SpreadsheetML file path and the loadOptions object
        workbook5 = self.Workbook(self.dataDir + "Book3.xml", loadOptions3)
        
        # Print message
        print ("SpreadSheetML format workbook has been opened successfully.");
        
        # 6.
        # Opening CSV Files
        #Creating and CSV LoadOptions object
        loadOptions4 = self.LoadOptions(fileFormatType.CSV)
        #Creating an Workbook object with CSV file path and the loadOptions object
        workbook6 = self.Workbook(self.dataDir + "Book_CSV.csv", loadOptions4)
        # Print message
        print ("CSV format workbook has been opened successfully.")
        
        
        # 7.
        # Opening Tab Delimited Files
        # Creating and TAB_DELIMITED LoadOptions object
        loadOptions5 = self.LoadOptions(fileFormatType.TAB_DELIMITED);

        # Creating an Workbook object with Tab Delimited text file path and the loadOptions object
        workbook7 = self.Workbook(self.dataDir + "Book1TabDelimited.txt", loadOptions5)

        # Print message
        print("<br />");
        print ("Tab Delimited workbook has been opened successfully.");



        # 8.
        # Opening Encrypted Excel Files
        # Creating and EXCEL_97_TO_2003 LoadOptions object
        loadOptions6 = self.LoadOptions(fileFormatType.EXCEL_97_TO_2003)

        # Setting the password for the encrypted Excel file
        loadOptions6.setPassword("1234")

        # Creating an Workbook object with file path and the loadOptions object
        workbook8 = self.Workbook(self.dataDir + "encryptedBook.xls", loadOptions6)

        # Print message
        print("<br />");
        print ("Encrypted workbook has been opened successfully.");
        
class SavingFiles:
    def __init__(self,dataDir):
        self.dataDir = dataDir
        self.Workbook = jpype.JClass("com.aspose.cells.Workbook")
        self.FileFormatType = jpype.JClass("com.aspose.cells.FileFormatType")
        
    def main(self):
        
        fileFormatType = self.FileFormatType
        
        
        #Creating an Workbook object with an Excel file path
        workbook = self.Workbook(self.dataDir + "Book1.xls")
        #Save in default (Excel2003) format
        workbook.save(self.dataDir + "book.default.out.xls")

        #Save in Excel2003 format
        workbook.save(self.dataDir + "book.out.xls", fileFormatType.EXCEL_97_TO_2003)

        #Save in Excel2007 xlsx format
        workbook.save(self.dataDir + "book.out.xlsx", fileFormatType.XLSX)

        #Save in SpreadsheetML format
        workbook.save(self.dataDir + "book.out.xml", fileFormatType.EXCEL_2003_XML)
        
        #Print Message
        print("<BR>")
        print("Worksheets are saved successfully.")
    
class WorksheetToImage:
    
    def __init__(self,dataDir):
        self.dataDir = dataDir
        self.Workbook = jpype.JClass("com.aspose.cells.Workbook")
        self.ImageFormat = jpype.JClass("com.aspose.cells.ImageFormat")
        self.ImageOrPrintOptions = jpype.JClass("com.aspose.cells.ImageOrPrintOptions")
        self.SheetRender = jpype.JClass("com.aspose.cells.SheetRender")
    
    def main(self):
               
        imageFormat = self.ImageFormat
        
        #Instantiate a workbook with path to an Excel file
        book = self.Workbook(self.dataDir + "Book1.xls")

        #Create an object for ImageOptions
        imgOptions = self.ImageOrPrintOptions()

        #Set the image type
        imgOptions.setImageFormat(imageFormat.getPng())

        #Get the first worksheet.
        sheet = book.getWorksheets().get(0)

        #Create a SheetRender object for the target sheet
        sr =self.SheetRender(sheet, imgOptions)
        for i in range(sr.getPageCount()):
        
            #Generate an image for the worksheet
            sr.toImage(i, self.dataDir + "mysheetimg" + ".png")

        
        # Print message
        print "Images generated successfully."