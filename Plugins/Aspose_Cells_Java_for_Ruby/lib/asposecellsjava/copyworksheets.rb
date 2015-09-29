module Asposecellsjava
  module CopyWorksheets
    def initialize()
        @data_dir = File.dirname(File.dirname(File.dirname(__FILE__))) + '/data/'
        
        # Instantiating a Workbook object by excel file path
        workbook = Rjb::import('com.aspose.cells.Workbook').new(@data_dir + 'Book1.xls')
        
        # Copy Worksheets within a Workbook
        copy_worksheet(workbook)

        # Move Worksheets within Workbook
        move_worksheet(workbook)
    end

    def copy_worksheet(workbook)
        # Create a Worksheets object with reference to the sheets of the Workbook.
        sheets = workbook.getWorksheets()

        # Copy data to a new sheet from an existing sheet within the Workbook.
        sheets.addCopy("Sheet1")

        # Saving the modified Excel file in default (that is Excel 2003) format
        workbook.save(@data_dir + "Copy Worksheet.xls")

        puts "Copy worksheet, please check the output file."
    end    

    def move_worksheet(workbook)
        # Get the first worksheet in the book.
        sheet = workbook.getWorksheets().get(0)

        # Move the first sheet to the third position in the workbook.
        sheet.moveTo(2)

        # Saving the modified Excel file in default (that is Excel 2003) format
        workbook.save(@data_dir + "Move Worksheet.xls")

        puts "Move worksheet, please check the output file."
    end    
  end
end
