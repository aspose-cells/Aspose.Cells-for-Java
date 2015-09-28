module Asposecellsjava
  module Protection
    def initialize()
        # Proctect a Worksheet
        protect_worksheet()

        # Unproctect a Worksheet
        unprotect_worksheet()
    end

    def protect_worksheet()
        data_dir = File.dirname(File.dirname(File.dirname(__FILE__))) + '/data/'
        
        # Instantiating a Workbook object by excel file path
        workbook = Rjb::import('com.aspose.cells.Workbook').new(data_dir + 'Book1.xls')
        
        # Accessing the first worksheet in the Excel file
        worksheets = workbook.getWorksheets()
        worksheet = worksheets.get(0)

        protection = worksheet.getProtection()

        # The following 3 methods are only for Excel 2000 and earlier formats
        protection.setAllowEditingContent(false)
        protection.setAllowEditingObject(false)
        protection.setAllowEditingScenario(false)

        # Protects the first worksheet with a password "1234"
        protection.setPassword("1234")
        
        # Saving the modified Excel file in default (that is Excel 2003) format
        workbook.save(data_dir + "Protect Worksheet.xls")

        puts "Protect a Worksheet, please check the output file."
    end    

    def unprotect_worksheet()
        data_dir = File.dirname(File.dirname(File.dirname(__FILE__))) + '/data/'
        
        # Instantiating a Workbook object by excel file path
        workbook = Rjb::import('com.aspose.cells.Workbook').new(data_dir + 'Book1.xls')
        
        # Accessing the first worksheet in the Excel file
        worksheets = workbook.getWorksheets()
        worksheet = worksheets.get(0)

        protection = worksheet.getProtection()

        # The following 3 methods are only for Excel 2000 and earlier formats
        protection.setAllowEditingContent(false)
        protection.setAllowEditingObject(false)
        protection.setAllowEditingScenario(false)

        # Unprotecting the worksheet with a password
        worksheet.unprotect("1234")
        
        # Saving the modified Excel file in default (that is Excel 2003) format
        workbook.save(data_dir + "Unprotect Worksheet.xls")

        puts "Unprotect a Worksheet, please check the output file."
    end    
  end
end
