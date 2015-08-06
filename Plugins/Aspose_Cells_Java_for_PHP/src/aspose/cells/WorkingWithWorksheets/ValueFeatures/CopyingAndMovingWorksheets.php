<?php

namespace Aspose\Cells\WorkingWithWorksheets\ValueFeatures;

use com\aspose\cells\Workbook as Workbook;


class CopyingAndMovingWorksheets {

    public static function run($dataDir=null)
    {
        # Instantiating a Workbook object by excel file path
        $workbook = new Workbook($dataDir . "Book1.xls");

        # Copy Worksheets within a Workbook
        CopyingAndMovingWorksheets::copy_worksheet($dataDir,$workbook);

        # Move Worksheets within Workbook
        CopyingAndMovingWorksheets::move_worksheet($dataDir,$workbook);
    }

    public static function copy_worksheet($dataDir, $workbook)
    {

        # Create a Worksheets object with reference to the sheets of the Workbook.
        $sheets = $workbook->getWorksheets();

        # Copy data to a new sheet from an existing sheet within the Workbook.
        $sheets->addCopy("Sheet1");

        # Saving the modified Excel file in default (that is Excel 2003) format
        $workbook->save($dataDir . "Copy Worksheet.xls");

        print "Copy worksheet, please check the output file." . PHP_EOL;

    }

    public static function move_worksheet($dataDir=null, $workbook=null)
    {

        # Get the first worksheet in the book.
        $sheet = $workbook->getWorksheets()->get(0);

        # Move the first sheet to the third position in the workbook.
        $sheet->moveTo(2);

        # Saving the modified Excel file in default (that is Excel 2003) format
        $workbook->save($dataDir . "Move Worksheet.xls");

        print "Move worksheet, please check the output file." . PHP_EOL;

    }

}