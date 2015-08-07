<?php

namespace Aspose\Cells\WorkingWithWorksheets\DisplayFeatures;

use com\aspose\cells\Workbook as Workbook;

class DisplayHideGridlines {

    public static function run($dataDir=null)
    {

        //Instantiating a Workbook object by excel file path
        $workbook = new Workbook($dataDir . "book1.xls");

        //Accessing the first worksheet in the Excel file
        $worksheets = $workbook->getWorksheets();

        $worksheet = $worksheets->get(0);

        //Hiding the grid lines of the first worksheet of the Excel file
        $worksheet->setGridlinesVisible(false);

        //Saving the modified Excel file in default (that is Excel 2000) format
        $workbook->save($dataDir . "output.xls");

        // Print message
        print "Grid lines are now hidden on sheet 1, please check the output document." . PHP_EOL;

    }
}