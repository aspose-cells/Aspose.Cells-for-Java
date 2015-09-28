<?php

namespace Aspose\Cells\QuickStart;

use com\aspose\cells\Workbook as Workbook;
use com\aspose\cells\FileFormatType as FileFormatType;

class HelloWorld {

    public static function run($dataDir=null)
    {

        // Instantiating a Workbook object that represents a Microsoft Excel file.
        $workbook = new Workbook();

        // Note when you create a new workbook, a default worksheet,
        // "Sheet1", is by default added to the workbook.
        // Accessing the first worksheet in the book ("Sheet1").
        $sheet = $workbook->getWorksheets()->get(0);

        # Access cell "A1" in the sheet.
        $cell = $sheet->getCells()->get("A1");

        # Input the "Hello World!" text into the "A1" cell
        $cell->setValue("Hello World!");

        # Saving the modified Excel file in default (that is Excel 2003) format
        $file_format_type = new FileFormatType();
        $workbook->save($dataDir . "HelloWorld.xls", $file_format_type->EXCEL_97_TO_2003);

        print "Document has been saved, please check the output file.";


    }

} 