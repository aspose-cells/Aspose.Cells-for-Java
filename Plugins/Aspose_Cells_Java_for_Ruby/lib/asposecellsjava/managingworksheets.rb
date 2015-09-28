module Asposecellsjava
  module ManagingWorksheets
    def initialize()
        @data_dir = File.dirname(File.dirname(File.dirname(__FILE__))) + '/data/'

        # Adding Worksheets to a New Excel File
        add_worksheet(workbook)        

        # Adding Worksheets to a Designer Spreadsheet
        add_worksheet_to_designer_spreadsheet()

        # Accessing Worksheets using Sheet Name
        get_worksheet()

        # Removing Worksheets using Sheet Name
        remove_worksheet_by_name()

        # Removing Worksheets using Sheet Name
        remove_worksheet_by_index()
    end

    def add_worksheet()
        # Instantiating a Workbook object
        workbook = Rjb::import('com.aspose.cells.Workbook').new

        # Adding a new worksheet to the Workbook object
        worksheets = workbook.getWorksheets()

        sheet_index = worksheets.add()
        worksheet = worksheets.get(sheet_index)

        # Setting the name of the newly added worksheet
        worksheet.setName("My Worksheet")

        # Saving the modified Excel file in default (that is Excel 2003) format
        workbook.save(@data_dir + "book.out.xls")

        puts "Sheet added successfully."
    end   

    def add_worksheet_to_designer_spreadsheet()
        # Creating a file stream containing the Excel file to be opened
        fstream = IO.sysopen(@data_dir + 'book1.xls', "w")

        # Instantiating a Workbook object with the stream
        workbook = Rjb::import('com.aspose.cells.Workbook').new(fstream)

        # Adding a new worksheet to the Workbook object
        worksheets = workbook.getWorksheets()
        sheet_index = worksheets.add()
        worksheet = worksheets.get(sheet_index)

        # Setting the name of the newly added worksheet
        worksheet.setName("My Worksheet")

        # Saving the modified Excel file in default (that is Excel 2003) format
        workbook.save(@data_dir + "book1.out.xls")
    end   

    def get_worksheet()
        # Creating a file stream containing the Excel file to be opened
        fstream = IO.sysopen(@data_dir + 'book1.xls', "w")

        # Instantiating a Workbook object with the stream
        workbook = Rjb::import('com.aspose.cells.Workbook').new(fstream)

        # Accessing a worksheet using its sheet name
        worksheet = workbook.getWorksheets().get("Sheet1")

        puts worksheet.to_string
    end  

    def remove_worksheet_by_name()
        # Creating a file stream containing the Excel file to be opened
        fstream = IO.sysopen(@data_dir + 'book1.xls', "w")

        # Instantiating a Workbook object with the stream
        workbook = Rjb::import('com.aspose.cells.Workbook').new(fstream)

        # Removing a worksheet using its sheet name
        workbook.getWorksheets().removeAt("Sheet1")
        
        # Saving the Excel file
        workbook.save(@data_dir + "book.out.xls")
        
        # Print Message
        puts "Sheet removed successfully."
    end  

    def remove_worksheet_by_index()
        # Creating a file stream containing the Excel file to be opened
        fstream = IO.sysopen(@data_dir + 'book1.xls', "w")

        # Instantiating a Workbook object with the stream
        workbook = Rjb::import('com.aspose.cells.Workbook').new(fstream)

        # Removing a worksheet using its sheet name
        workbook.getWorksheets().removeAt(0)
        
        # Saving the Excel file
        workbook.save(@data_dir + "book.out.xls")
        
        # Print Message
        puts "Sheet removed successfully."
    end  
  end
end
