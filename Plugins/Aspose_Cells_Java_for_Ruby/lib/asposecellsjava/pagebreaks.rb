module Asposecellsjava
  module PageBreaks
    def initialize()
        @data_dir = File.dirname(File.dirname(File.dirname(__FILE__))) + '/data/'
        
        # Instantiating a Workbook object
        workbook = Rjb::import('com.aspose.cells.Workbook').new
        
        # Adding Page Breaks
        add_page_breaks(workbook)

        # Clearing All Page Breaks 
        clear_all_page_breaks(workbook)

        # Removing Specific Page Break
        remove_page_break(workbook)
    end

    def add_page_breaks(workbook)
        worksheets = workbook.getWorksheets()
        worksheet = worksheets.get(0)

        h_page_breaks = worksheet.getHorizontalPageBreaks()
        h_page_breaks.add("Y30")
        
        v_page_breaks = worksheet.getVerticalPageBreaks()
        v_page_breaks.add("Y30")

        # Saving the modified Excel file in default (that is Excel 2003) format
        workbook.save(@data_dir + "Add Page Breaks.xls")

        puts "Add page breaks, please check the output file."
    end    

    def clear_all_page_breaks(workbook)
        workbook.getWorksheets().get(0).getHorizontalPageBreaks().clear()
        workbook.getWorksheets().get(0).getVerticalPageBreaks().clear()

        # Saving the modified Excel file in default (that is Excel 2003) format
        workbook.save(@data_dir + "Clear All Page Breaks.xls")

        puts "Clear all page breaks, please check the output file."
    end  

    def remove_page_break(workbook)
        worksheets = workbook.getWorksheets()
        worksheet = worksheets.get(0)
        
        h_page_breaks = worksheet.getHorizontalPageBreaks()
        h_page_breaks.removeAt(0)
        
        v_page_breaks = worksheet.getVerticalPageBreaks()
        v_page_breaks.removeAt(0)

        # Saving the modified Excel file in default (that is Excel 2003) format
        workbook.save(@data_dir + "Remove Page Break.xls")

        puts "Remove page break, please check the output file."
    end  
  end
end
