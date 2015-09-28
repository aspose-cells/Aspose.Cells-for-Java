module Asposecellsjava
  module RowsAndColumns
    def initialize()
        # Inserting a Row
        insert_row()

        # Inserting Multiple Rows
        insert_multiple_rows()

        # Deleting a Row
        delete_row()

        # Deleting Multiple Rows
        delete_multiple_rows()

        # Inseting one or Multiple Columns
        insert_column()

        # Deleting a Column
        delete_column()

        # Hiding Rows and Columns
        hide_rows_columns()

        # Showing Rows and Columns
        unhide_rows_columns()

        # Grouping Rows & Columns
        group_rows_columns()

        # Ungrouping Rows & Columns
        ungroup_rows_columns()

        # Setting the Row Height
        set_row_height()

        # Setting the Width of a Column
        set_column_width()

        # Auto Fit Row
        autofit_row()

        # Auto Fit Column
        autofit_column()

        # Copying Rows
        copy_rows()

        # Copying Columns
        copy_columns()
    end

    def insert_row()
        data_dir = File.dirname(File.dirname(File.dirname(__FILE__))) + '/data/'
        
        # Instantiating a Workbook object by excel file path
        workbook = Rjb::import('com.aspose.cells.Workbook').new(data_dir + 'Book1.xls')

        # Accessing the first worksheet in the Excel file
        worksheet = workbook.getWorksheets().get(0)

        # Inserting a row into the worksheet at 3rd position
        worksheet.getCells().insertRows(2,1)

        # Saving the modified Excel file in default (that is Excel 2003) format
        workbook.save(data_dir + "Insert Row.xls")

        puts "Insert Row Successfully."
    end    

    def insert_multiple_rows()
        data_dir = File.dirname(File.dirname(File.dirname(__FILE__))) + '/data/'
        
        # Instantiating a Workbook object by excel file path
        workbook = Rjb::import('com.aspose.cells.Workbook').new(data_dir + 'Book1.xls')

        # Accessing the first worksheet in the Excel file
        worksheet = workbook.getWorksheets().get(0)

        # Inserting a row into the worksheet at 3rd position
        worksheet.getCells().insertRows(2,10)

        # Saving the modified Excel file in default (that is Excel 2003) format
        workbook.save(data_dir + "Insert Multiple Rows.xls")

        puts "Insert Multiple Rows Successfully."
    end    

    def delete_row()
        data_dir = File.dirname(File.dirname(File.dirname(__FILE__))) + '/data/'
        
        # Instantiating a Workbook object by excel file path
        workbook = Rjb::import('com.aspose.cells.Workbook').new(data_dir + 'Book1.xls')

        # Accessing the first worksheet in the Excel file
        worksheet = workbook.getWorksheets().get(0)

        # Deleting 3rd row from the worksheet
        worksheet.getCells().deleteRows(2,1,true)

        # Saving the modified Excel file in default (that is Excel 2003) format
        workbook.save(data_dir + "Delete Row.xls")

        puts "Delete Row Successfully."
    end    

    def delete_multiple_rows()
        data_dir = File.dirname(File.dirname(File.dirname(__FILE__))) + '/data/'
        
        # Instantiating a Workbook object by excel file path
        workbook = Rjb::import('com.aspose.cells.Workbook').new(data_dir + 'Book1.xls')

        # Accessing the first worksheet in the Excel file
        worksheet = workbook.getWorksheets().get(0)

        # Deleting 10 rows from the worksheet starting from 3rd row
        worksheet.getCells().deleteRows(2,10,true)

        # Saving the modified Excel file in default (that is Excel 2003) format
        workbook.save(data_dir + "Delete Multiple Rows.xls")

        puts "Delete Multiple Rows Successfully."
    end    

    def insert_column()
        data_dir = File.dirname(File.dirname(File.dirname(__FILE__))) + '/data/'
        
        # Instantiating a Workbook object by excel file path
        workbook = Rjb::import('com.aspose.cells.Workbook').new(data_dir + 'Book1.xls')

        # Accessing the first worksheet in the Excel file
        worksheet = workbook.getWorksheets().get(0)

        # Inserting a column into the worksheet at 2nd position
        worksheet.getCells().insertColumns(1,1)

        # Saving the modified Excel file in default (that is Excel 2003) format
        workbook.save(data_dir + "Insert Column.xls")

        puts "Insert Column Successfully."
    end    

    def delete_column()
        data_dir = File.dirname(File.dirname(File.dirname(__FILE__))) + '/data/'
        
        # Instantiating a Workbook object by excel file path
        workbook = Rjb::import('com.aspose.cells.Workbook').new(data_dir + 'Book1.xls')

        # Accessing the first worksheet in the Excel file
        worksheet = workbook.getWorksheets().get(0)

        # Deleting a column from the worksheet at 2nd position
        worksheet.getCells().deleteColumns(1,1,true)

        # Saving the modified Excel file in default (that is Excel 2003) format
        workbook.save(data_dir + "Delete Column.xls")

        puts "Delete Column Successfully."
    end    

    def hide_rows_columns()
        data_dir = File.dirname(File.dirname(File.dirname(__FILE__))) + '/data/'
        
        # Instantiating a Workbook object by excel file path
        workbook = Rjb::import('com.aspose.cells.Workbook').new(data_dir + 'Book1.xls')

        # Accessing the first worksheet in the Excel file
        worksheet = workbook.getWorksheets().get(0)
        cells = worksheet.getCells()

        # Hiding the 3rd row of the worksheet
        cells.hideRow(2)

        # Hiding the 2nd column of the worksheet
        cells.hideColumn(1)

        # Saving the modified Excel file in default (that is Excel 2003) format
        workbook.save(data_dir + "Hide Rows And Columns.xls")

        puts "Hide Rows And Columns Successfully."
    end   

    def unhide_rows_columns()
        data_dir = File.dirname(File.dirname(File.dirname(__FILE__))) + '/data/'
        
        # Instantiating a Workbook object by excel file path
        workbook = Rjb::import('com.aspose.cells.Workbook').new(data_dir + 'Book1.xls')

        # Accessing the first worksheet in the Excel file
        worksheet = workbook.getWorksheets().get(0)
        cells = worksheet.getCells()

        # Unhiding the 3rd row and setting its height to 13.5
        cells.unhideRow(2,13.5)

        # Unhiding the 2nd column and setting its width to 8.5
        cells.unhideColumn(1,8.5)

        # Saving the modified Excel file in default (that is Excel 2003) format
        workbook.save(data_dir + "Unhide Rows And Columns.xls")

        puts "Unhide Rows And Columns Successfully."
    end   

    def group_rows_columns()
        data_dir = File.dirname(File.dirname(File.dirname(__FILE__))) + '/data/'
        
        # Instantiating a Workbook object by excel file path
        workbook = Rjb::import('com.aspose.cells.Workbook').new(data_dir + 'Book1.xls')

        # Accessing the first worksheet in the Excel file
        worksheet = workbook.getWorksheets().get(0)
        cells = worksheet.getCells()

        # Grouping first six rows (from 0 to 5) and making them hidden by passing true
        cells.groupRows(0,5,true)

        # Grouping first three columns (from 0 to 2) and making them hidden by passing true
        cells.groupColumns(0,2,true)

        # Saving the modified Excel file in default (that is Excel 2003) format
        workbook.save(data_dir + "Group Rows And Columns.xls")

        puts "Group Rows And Columns Successfully."
    end   

    def ungroup_rows_columns()
        data_dir = File.dirname(File.dirname(File.dirname(__FILE__))) + '/data/'
        
        # Instantiating a Workbook object by excel file path
        workbook = Rjb::import('com.aspose.cells.Workbook').new(data_dir + 'Group Rows And Columns.xls')

        # Accessing the first worksheet in the Excel file
        worksheet = workbook.getWorksheets().get(0)
        cells = worksheet.getCells()

        # Ungrouping first six rows (from 0 to 5)
        cells.ungroupRows(0,5)

        # Ungrouping first three columns (from 0 to 2)
        cells.ungroupColumns(0,2)

        # Saving the modified Excel file in default (that is Excel 2003) format
        workbook.save(data_dir + "Ungroup Rows And Columns.xls")

        puts "Ungroup Rows And Columns Successfully."
    end   

    def set_row_height()
        data_dir = File.dirname(File.dirname(File.dirname(__FILE__))) + '/data/'
        
        # Instantiating a Workbook object by excel file path
        workbook = Rjb::import('com.aspose.cells.Workbook').new(data_dir + 'Book1.xls')

        # Accessing the first worksheet in the Excel file
        worksheet = workbook.getWorksheets().get(0)
        cells = worksheet.getCells()

        # Setting the height of the second row to 13
        cells.setRowHeight(1, 13)

        # Saving the modified Excel file in default (that is Excel 2003) format
        workbook.save(data_dir + "Set Row Height.xls")

        puts "Set Row Height Successfully."
    end

    def set_column_width()
        data_dir = File.dirname(File.dirname(File.dirname(__FILE__))) + '/data/'
        
        # Instantiating a Workbook object by excel file path
        workbook = Rjb::import('com.aspose.cells.Workbook').new(data_dir + 'Book1.xls')

        # Accessing the first worksheet in the Excel file
        worksheet = workbook.getWorksheets().get(0)
        cells = worksheet.getCells()

        # Setting the width of the second column to 17.5
        cells.setColumnWidth(1, 17.5)

        # Saving the modified Excel file in default (that is Excel 2003) format
        workbook.save(data_dir + "Set Column Width.xls")

        puts "Set Column Width Successfully."
    end

    def autofit_row()
        data_dir = File.dirname(File.dirname(File.dirname(__FILE__))) + '/data/'
        
        # Instantiating a Workbook object by excel file path
        workbook = Rjb::import('com.aspose.cells.Workbook').new(data_dir + 'Book1.xls')

        # Accessing the first worksheet in the Excel file
        worksheet = workbook.getWorksheets().get(0)

        # Auto-fitting the 3rd row of the worksheet
        worksheet.autoFitRow(2)

        # Auto-fitting the 3rd row of the worksheet based on the contents in a range of
        # cells (from 1st to 9th column) within the row
        #worksheet.autoFitRow(2,0,8) # Uncomment this line if you to do AutoFit Row in a Range of Cells. Also, comment line 288.

        # Saving the modified Excel file in default (that is Excel 2003) format
        workbook.save(data_dir + "Autofit Row.xls")

        puts "Autofit Row Successfully."
    end

    def autofit_column()
        data_dir = File.dirname(File.dirname(File.dirname(__FILE__))) + '/data/'
        
        # Instantiating a Workbook object by excel file path
        workbook = Rjb::import('com.aspose.cells.Workbook').new(data_dir + 'Book1.xls')

        # Accessing the first worksheet in the Excel file
        worksheet = workbook.getWorksheets().get(0)

        # Auto-fitting the 4th column of the worksheet
        worksheet.autoFitColumn(3)

        # Auto-fitting the 4th column of the worksheet based on the contents in a range of
        # cells (from 1st to 9th row) within the column
        #worksheet.autoFitColumn(3,0,8) #Uncomment this line if you to do AutoFit Column in a Range of Cells. Also, comment line 310.

        # Saving the modified Excel file in default (that is Excel 2003) format
        workbook.save(data_dir + "Autofit Column.xls")

        puts "Autofit Column Successfully."
    end

    def copy_rows()
        data_dir = File.dirname(File.dirname(File.dirname(__FILE__))) + '/data/'
        
        # Instantiating a Workbook object by excel file path
        workbook = Rjb::import('com.aspose.cells.Workbook').new(data_dir + 'Book1.xls')

        # Accessing the first worksheet in the Excel file
        worksheet = workbook.getWorksheets().get(0)

        # Copy the second row with data, formattings, images and drawing objects
        # to the 12th row in the worksheet.
        worksheet.getCells().copyRow(worksheet.getCells(),1,11);

        # Saving the modified Excel file in default (that is Excel 2003) format
        workbook.save(data_dir + "Copy Rows.xls")

        puts "Copy Rows Successfully."
    end

    def copy_columns()
        data_dir = File.dirname(File.dirname(File.dirname(__FILE__))) + '/data/'
        
        # Instantiating a Workbook object by excel file path
        workbook = Rjb::import('com.aspose.cells.Workbook').new

        # Accessing the first worksheet in the Excel file
        worksheet = workbook.getWorksheets().get(0)

        # Put some data into header rows (A1:A4)
        i = 0
        while i < 5
            worksheet.getCells().get(i, 0).setValue("Header Row #{i}")
            i +=1
        end

        # Put some detail data (A5:A999)
        i = 5
        while i < 1000
            worksheet.getCells().get(i, 0).setValue("Detail Row #{i}")
            i +=1
        end

        # Create another Workbook.
        workbook1 = Rjb::import('com.aspose.cells.Workbook').new

        # Get the first worksheet in the book.
        worksheet1 = workbook1.getWorksheets().get(0)

        # Copy the first column from the first worksheet of the first workbook into
        # the first worksheet of the second workbook.
        worksheet1.getCells().copyColumn(worksheet.getCells(),0,2)

        # Autofit the column.
        worksheet1.autoFitColumn(2)

        # Saving the modified Excel file in default (that is Excel 2003) format
        workbook.save(data_dir + "Copy Columns.xls")

        puts "Copy Columns Successfully."
    end
  end
end
