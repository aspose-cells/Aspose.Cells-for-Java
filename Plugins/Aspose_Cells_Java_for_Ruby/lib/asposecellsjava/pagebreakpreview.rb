module Asposecellsjava
  module PageBreakPreview
    def initialize()
        data_dir = File.dirname(File.dirname(File.dirname(__FILE__))) + '/data/'
        
        # Instantiating a Workbook object by excel file path
        workbook = Rjb::import('com.aspose.cells.Workbook').new(data_dir + 'Book1.xls')
        
        # Accessing the first worksheet in the Excel file
        worksheets = workbook.getWorksheets()
        sheet_index = worksheets.add()
        worksheet = worksheets.get(sheet_index)

        # Displaying the worksheet in page break preview
        worksheet.setPageBreakPreview(true)
        
        # Saving the modified Excel file in default (that is Excel 2003) format
        workbook.save(data_dir + "output.xls")

        puts "Set page break preview, please check the output file."
    end
  end
end
