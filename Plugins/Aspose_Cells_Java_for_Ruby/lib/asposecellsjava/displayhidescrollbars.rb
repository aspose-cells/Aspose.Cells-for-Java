module Asposecellsjava
  module DisplayHideScrollBars
    def initialize()
        data_dir = File.dirname(File.dirname(File.dirname(__FILE__))) + '/data/'
        
        # Instantiating a Workbook object by excel file path
        workbook = Rjb::import('com.aspose.cells.Workbook').new(data_dir + 'Book1.xls')
        
        # Hiding the vertical scroll bar of the Excel file
        workbook.getSettings().setVScrollBarVisible(false)

        # Hiding the horizontal scroll bar of the Excel file
        workbook.getSettings().setHScrollBarVisible(false)
        
        # Saving the modified Excel file in default (that is Excel 2003) format
        workbook.save(data_dir + "output.xls")

        puts "Scroll Bars are now hidden, please check the output file."
    end
  end
end
