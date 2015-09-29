module Asposecellsjava
  module Document
    def initialize()
        # Accessing Document Properties
        get_properties()

        # Adding Custom Property
        add_custom_property()

        # Removing Custom Properties 
        remove_custom_property()
    end

    def get_properties()
        data_dir = File.dirname(File.dirname(File.dirname(__FILE__))) + '/data/'
        
        # Instantiating a Workbook object by excel file path
        workbook = Rjb::import('com.aspose.cells.Workbook').new(data_dir + 'Book1.xls')

        # Retrieve a list of all custom document properties of the Excel file
        custom_properties = workbook.getWorksheets().getCustomDocumentProperties()

        # Accessng a custom document property by using the property index
        puts "Property By Index: " +  custom_properties.get(1).to_string

        # Accessng a custom document property by using the property name
        puts "Property By Name: " + custom_properties.get("Publisher").to_string
    end

    def add_custom_property()
        data_dir = File.dirname(File.dirname(File.dirname(__FILE__))) + '/data/'
        
        # Instantiating a Workbook object by excel file path
        workbook = Rjb::import('com.aspose.cells.Workbook').new(data_dir + 'Book1.xls')

        # Retrieve a list of all custom document properties of the Excel file
        #custom_properties = Rjb::import('java.util.ArrayList').new
        custom_properties = workbook.getWorksheets().getCustomDocumentProperties()

        # Adding a custom document property to the Excel file
        custom_properties.add("Publisher", "Aspose")

        # Save the document in PDF format
        workbook.save(data_dir + "Add_Property.xls")

        puts "Added custom property successfully."
    end    

    def remove_custom_property()
        data_dir = File.dirname(File.dirname(File.dirname(__FILE__))) + '/data/'
        
        # Instantiating a Workbook object by excel file path
        workbook = Rjb::import('com.aspose.cells.Workbook').new(data_dir + 'Book1.xls')

        # Retrieve a list of all custom document properties of the Excel file
        custom_properties = workbook.getWorksheets().getCustomDocumentProperties()

        # Adding a custom document property to the Excel file
        custom_properties.remove("Publisher")

        # Save the document in PDF format
        workbook.save(data_dir + "Removed_Property.xls")

        puts "Removed custom property successfully."
    end    
  end
end
