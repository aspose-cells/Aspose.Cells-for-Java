module Asposecellsjava
  module HideUnhideWorksheet
    def initialize()
        data_dir = File.dirname(File.dirname(File.dirname(__FILE__))) + '/data/'
        
        # Instantiating a Workbook object by excel file path
        workbook = Rjb::import('com.aspose.cells.Workbook').new(data_dir + 'Book1.xls')
        
        # Accessing the first worksheet in the Excel file
        #worksheets = Rjb::import('java.util.ArrayList').new
        worksheets = workbook.getWorksheets()
        worksheet = worksheets.get(0)

        # Hiding the first worksheet of the Excel file
        worksheet.setVisible(false)
        
        # Saving the modified Excel file in default (that is Excel 2003) format
        workbook.save(data_dir + "output.xls")

        puts "Worksheet 1 is now hidden, please check the output document."
    end
  end
end
