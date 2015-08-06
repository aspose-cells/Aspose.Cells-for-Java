<?php

namespace Aspose\Cells\WorkingWithWorksheets\SecurityFeatures;

use com\aspose\cells\Workbook as Workbook;

class ProtectingWorksheet {

    public static function run($dataDir=null)
    {
        //Instantiating a Excel object by excel file path
        $excel = new Workbook($dataDir . "book1.xls");

        //Accessing the first worksheet in the Excel file
        $worksheets = $excel->getWorksheets();
        $worksheet = $worksheets->get(0);

        $protection = $worksheet->getProtection();

        //The following 3 methods are only for Excel 2000 and earlier formats
        $protection->setAllowEditingContent(false);
        $protection->setAllowEditingObject(false);
        $protection->setAllowEditingScenario(false);

        //Protects the first worksheet with a password "1234"
        $protection->setPassword("1234");

        //Saving the modified Excel file in default format
        $excel->save($dataDir . "output.xls");

        //Print Message
        print "Sheet protected successfully." . PHP_EOL;
    }

}