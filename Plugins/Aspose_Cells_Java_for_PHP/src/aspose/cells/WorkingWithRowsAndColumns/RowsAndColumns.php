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


}