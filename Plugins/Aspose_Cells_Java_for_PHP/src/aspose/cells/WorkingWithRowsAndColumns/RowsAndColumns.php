<?php

namespace Aspose\Cells\WorkingWithRowsAndColumns;

use com\aspose\cells\Workbook as Workbook;

class RowsAndColumns {

    public static function run($dataDir=null)
    {
        # Inserting a Row
        RowsAndColumns::insert_row($dataDir);

        # Inserting Multiple Rows
        RowsAndColumns::insert_multiple_rows($dataDir);

        # Deleting a Row
        RowsAndColumns::delete_row($dataDir);

        # Deleting Multiple Rows
        RowsAndColumns::delete_multiple_rows($dataDir);

        # Inseting one or Multiple Columns
        RowsAndColumns::insert_column($dataDir);

        # Deleting a Column
        RowsAndColumns::delete_column($dataDir);

        # Hiding Rows and Columns
        RowsAndColumns::hide_rows_columns($dataDir);

        # Showing Rows and Columns
        RowsAndColumns::unhide_rows_columns($dataDir);

        # Grouping Rows & Columns
        RowsAndColumns::group_rows_columns($dataDir);

        # Ungrouping Rows & Columns
        RowsAndColumns::ungroup_rows_columns($dataDir);

        # Setting the Row Height
        RowsAndColumns::set_row_height($dataDir);

        # Setting the Width of a Column
        RowsAndColumns::set_column_width($dataDir);

        # Auto Fit Row
        RowsAndColumns::autofit_row($dataDir);

        # Auto Fit Column
        RowsAndColumns::autofit_column($dataDir);

        # Copying Rows
        RowsAndColumns::copy_rows($dataDir);

        # Copying Columns
        RowsAndColumns::copy_columns($dataDir);
    }

    public static function insert_row($dataDir)
    {
        # Instantiating a Workbook object by excel file path
        $workbook = new Workbook($dataDir . 'Book1.xls');

        # Accessing the first worksheet in the Excel file
        $worksheet = $workbook->getWorksheets()->get(0);

        # Inserting a row into the worksheet at 3rd position
        $worksheet->getCells()->insertRows(2,1);

        # Saving the modified Excel file in default (that is Excel 2003) format
        $workbook->save($dataDir . "Insert Row.xls");

        print "Insert Row Successfully." . PHP_EOL;

    }

    public static function insert_multiple_rows($dataDir)
    {

        # Instantiating a Workbook object by excel file path
        $workbook = new Workbook($dataDir . 'Book1.xls');

        # Accessing the first worksheet in the Excel file
        $worksheet = $workbook->getWorksheets()->get(0);

        # Inserting a row into the worksheet at 3rd position
        $worksheet->getCells()->insertRows(2,10);

        # Saving the modified Excel file in default (that is Excel 2003) format
        $workbook->save($dataDir . "Insert Multiple Rows.xls");

        print "Insert Multiple Rows Successfully." . PHP_EOL;

    }

    public static function delete_row($dataDir)
    {
        # Instantiating a Workbook object by excel file path
        $workbook = new Workbook($dataDir . 'Book1.xls');

        # Accessing the first worksheet in the Excel file
        $worksheet = $workbook->getWorksheets()->get(0);

        # Deleting 3rd row from the worksheet
        $worksheet->getCells()->deleteRows(2,1,true);

        # Saving the modified Excel file in default (that is Excel 2003) format
        $workbook->save($dataDir . "Delete Row.xls");

        print "Delete Row Successfully." . PHP_EOL;

    }

    public static function delete_multiple_rows($dataDir)
    {

        # Instantiating a Workbook object by excel file path
        $workbook = new Workbook($dataDir . 'Book1.xls');

        # Accessing the first worksheet in the Excel file
        $worksheet = $workbook->getWorksheets()->get(0);

        # Deleting 10 rows from the worksheet starting from 3rd row
        $worksheet->getCells()->deleteRows(2,10,true);

        # Saving the modified Excel file in default (that is Excel 2003) format
        $workbook->save($dataDir . "Delete Multiple Rows.xls");

        print "Delete Multiple Rows Successfully." . PHP_EOL;

    }

    public static function insert_column($dataDir)
    {

        # Instantiating a Workbook object by excel file path
        $workbook = new Workbook($dataDir . 'Book1.xls');

        # Accessing the first worksheet in the Excel file
        $worksheet = $workbook->getWorksheets()->get(0);

        # Inserting a column into the worksheet at 2nd position
        $worksheet->getCells()->insertColumns(1,1);

        # Saving the modified Excel file in default (that is Excel 2003) format
        $workbook->save($dataDir . "Insert Column.xls");

        print "Insert Column Successfully." . PHP_EOL;

    }

    public static function delete_column($dataDir)
    {

        # Instantiating a Workbook object by excel file path
        $workbook = new Workbook($dataDir . 'Book1.xls');

        # Accessing the first worksheet in the Excel file
        $worksheet = $workbook->getWorksheets()->get(0);

        # Deleting a column from the worksheet at 2nd position
        $worksheet->getCells()->deleteColumns(1,1,true);

        # Saving the modified Excel file in default (that is Excel 2003) format
        $workbook->save($dataDir . "Delete Column.xls");

        print "Delete Column Successfully." . PHP_EOL;

    }

    public static function hide_rows_columns($dataDir)
    {

        # Instantiating a Workbook object by excel file path
        $workbook = new Workbook($dataDir . 'Book1.xls');

        # Accessing the first worksheet in the Excel file
        $worksheet = $workbook->getWorksheets()->get(0);
        $cells = $worksheet->getCells();;

        # Hiding the 3rd row of the worksheet
        $cells->hideRow(2);

        # Hiding the 2nd column of the worksheet
        $cells->hideColumn(1);

        # Saving the modified Excel file in default (that is Excel 2003) format
        $workbook->save($dataDir . "Hide Rows And Columns.xls");

        print "Hide Rows And Columns Successfully." . PHP_EOL;

    }

    public static function unhide_rows_columns($dataDir)
    {

        # Instantiating a Workbook object by excel file path
        $workbook = new Workbook($dataDir . 'Book1.xls');

        # Accessing the first worksheet in the Excel file
        $worksheet = $workbook->getWorksheets()->get(0);
        $cells = $worksheet->getCells();;

        # Unhiding the 3rd row and setting its height to 13.5
        $cells->unhideRow(2,13.5);

        # Unhiding the 2nd column and setting its width to 8.5
        $cells->unhideColumn(1,8.5);

        # Saving the modified Excel file in default (that is Excel 2003) format
        $workbook->save($dataDir . "Unhide Rows And Columns.xls");

        print "Unhide Rows And Columns Successfully." . PHP_EOL;

    }

    public static function group_rows_columns($dataDir)
    {

        # Instantiating a Workbook object by excel file path
        $workbook = new Workbook($dataDir . 'Book1.xls');

        # Accessing the first worksheet in the Excel file
        $worksheet = $workbook->getWorksheets()->get(0);
        $cells = $worksheet->getCells();;

        # Grouping first six rows (from 0 to 5) and making them hidden by passing true
        $cells->groupRows(0,5,true);

        # Grouping first three columns (from 0 to 2) and making them hidden by passing true
        $cells->groupColumns(0,2,true);

        # Saving the modified Excel file in default (that is Excel 2003) format
        $workbook->save($dataDir . "Group Rows And Columns.xls");

        print "Group Rows And Columns Successfully." . PHP_EOL;

    }

    public static function ungroup_rows_columns($dataDir)
    {

        # Instantiating a Workbook object by excel file path
        $workbook = new Workbook($dataDir . 'Group Rows And Columns.xls');

        # Accessing the first worksheet in the Excel file
        $worksheet = $workbook->getWorksheets()->get(0);
        $cells = $worksheet->getCells();;

        # Ungrouping first six rows (from 0 to 5)
        $cells->ungroupRows(0,5);

        # Ungrouping first three columns (from 0 to 2)
        $cells->ungroupColumns(0,2);

        # Saving the modified Excel file in default (that is Excel 2003) format
        $workbook->save($dataDir . "Ungroup Rows And Columns.xls");

        print "Ungroup Rows And Columns Successfully." . PHP_EOL;

    }

    public static function set_row_height($dataDir)
    {

        # Instantiating a Workbook object by excel file path
        $workbook = new Workbook($dataDir . 'Book1.xls');

        # Accessing the first worksheet in the Excel file
        $worksheet = $workbook->getWorksheets()->get(0);
        $cells = $worksheet->getCells();

        # Setting the height of the second row to 13
        $cells->setRowHeight(1, 13);

        # Saving the modified Excel file in default (that is Excel 2003) format
        $workbook->save($dataDir . "Set Row Height.xls");

        print "Set Row Height Successfully." . PHP_EOL;

    }

    public static function set_column_width($dataDir)
    {

        # Instantiating a Workbook object by excel file path
        $workbook = new Workbook($dataDir . 'Book1.xls');

        # Accessing the first worksheet in the Excel file
        $worksheet = $workbook->getWorksheets()->get(0);
        $cells = $worksheet->getCells();

        # Setting the width of the second column to 17.5
        $cells->setColumnWidth(1, 17.5);

        # Saving the modified Excel file in default (that is Excel 2003) format
        $workbook->save($dataDir . "Set Column Width.xls");

        print "Set Column Width Successfully." . PHP_EOL;

    }

    public static function autofit_row($dataDir)
    {

        # Instantiating a Workbook object by excel file path
        $workbook = new Workbook($dataDir . 'Book1.xls');

        # Accessing the first worksheet in the Excel file
        $worksheet = $workbook->getWorksheets()->get(0);

        # Auto-fitting the 3rd row of the worksheet
        $worksheet->autoFitRow(2);

        # Auto-fitting the 3rd row of the worksheet based on the contents in a range of
        # cells (from 1st to 9th column) within the row
        #$worksheet->autoFitRow(2,0,8) # Uncomment this line if you to do AutoFit Row in a Range of Cells. Also, comment line 288.

        # Saving the modified Excel file in default (that is Excel 2003) format
        $workbook->save($dataDir . "Autofit Row.xls");

        print "Autofit Row Successfully." . PHP_EOL;

    }

    public static function autofit_column($dataDir)
    {

        # Instantiating a Workbook object by excel file path
        $workbook = new Workbook($dataDir . 'Book1.xls');

        # Accessing the first worksheet in the Excel file
        $worksheet = $workbook->getWorksheets()->get(0);

        # Auto-fitting the 4th column of the worksheet
        $worksheet->autoFitColumn(3);

        # Auto-fitting the 4th column of the worksheet based on the contents in a range of
        # cells (from 1st to 9th row) within the column
        #$worksheet->autoFitColumn(3,0,8) #Uncomment this line if you to do AutoFit Column in a Range of Cells. Also, comment line 310.

        # Saving the modified Excel file in default (that is Excel 2003) format
        $workbook->save($dataDir . "Autofit Column.xls");

        print "Autofit Column Successfully." . PHP_EOL;

    }

    public static function copy_rows($dataDir)
    {

        # Instantiating a Workbook object by excel file path
        $workbook = new Workbook($dataDir . 'Book1.xls');

        # Accessing the first worksheet in the Excel file
        $worksheet = $workbook->getWorksheets()->get(0);

        # Copy the second row with data, formattings, images and drawing objects
        # to the 12th row in the $worksheet->
        $worksheet->getCells()->copyRow($worksheet->getCells(),1,11);

        # Saving the modified Excel file in default (that is Excel 2003) format
        $workbook->save($dataDir . "Copy Rows.xls");

        print "Copy Rows Successfully." . PHP_EOL;

    }

    public static function copy_columns($dataDir)
    {

        # Instantiating a Workbook object by excel file path
        $workbook = new Workbook();

        # Accessing the first worksheet in the Excel file
        $worksheet = $workbook->getWorksheets()->get(0);

        # Put some data into header rows (A1:A4)
        $i = 0;
        while($i < 5)
        {
            $worksheet->getCells()->get($i, 0)->setValue("Header Row #$i");
            $i++;

        }


        # Put some detail data (A5:A999)
        $i = 5;
        while ($i < 1000) {

            $worksheet->getCells()->get($i, 0)->setValue("Detail Row #$i");
            $i++;
        }

        # Create another Workbook.
        $workbook1 = new Workbook();

        # Get the first worksheet in the book.
        $worksheet1 = $workbook1->getWorksheets()->get(0);

        # Copy the first column from the first worksheet of the first workbook into
        # the first worksheet of the second workbook.
        $worksheet1->getCells()->copyColumn($worksheet->getCells(),0,2);

        # Autofit the column.
        $worksheet1->autoFitColumn(2);

        # Saving the modified Excel file in default (that is Excel 2003) format
        $workbook->save($dataDir . "Copy Columns.xls");

        print "Copy Columns Successfully." . PHP_EOL;

    }

}