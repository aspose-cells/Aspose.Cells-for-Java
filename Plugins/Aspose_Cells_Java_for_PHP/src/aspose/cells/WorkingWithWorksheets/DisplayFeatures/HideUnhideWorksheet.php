<?php

namespace Aspose\Cells\WorkingWithWorksheets\DisplayFeatures;

use com\aspose\cells\Workbook as Workbook;

class HideUnhideWorksheet {

    public static function run($dataDir=null)
    {

        //Instantiating a Workbook object by excel file path
        $workbook = new Workbook($dataDir . "book1.xls");

        //Accessing the first worksheet in the Excel file
        $worksheets = $workbook->getWorksheets();
        $worksheet = $worksheets->get(0);

        //Hiding the first worksheet of the Excel file
        $worksheet->setVisible(false);

        //Saving the modified Excel file in default (that is Excel 2003) format
        $workbook->save($dataDir . "output.xls");

        // Print message
        print "Worksheet 1 is now hidden, please check the output document." . PHP_EOL;
    }

}