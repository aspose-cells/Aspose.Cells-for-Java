module Asposecellsjava
  module SplitPanes
    def initialize()
        data_dir = File.dirname(File.dirname(File.dirname(__FILE__))) + '/data/'
        
        # Instantiating a Workbook object by excel file path
        workbook = Rjb::import('com.aspose.cells.Workbook').new(data_dir + 'Book1.xls')
        
        # Set the active cell
        workbook.getWorksheets().get(0).setActiveCell("A20")

        # Split the worksheet window
        workbook.getWorksheets().get(0).split()
        
        # Saving the modified Excel file in default (that is Excel 2003) format
        workbook.save(data_dir + "SplitPanes output.xls")

        puts "Panes split successfully."
    end
  end
end
