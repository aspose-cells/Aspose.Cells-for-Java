<?php

namespace Aspose\Cells\WorkingWithWorksheets\SecurityFeatures;

use com\aspose\cells\Workbook as Workbook;
use com\aspose\cells\FileFormatType as FileFormatType;

class UnprotectingSimplyProtectedWorksheet {

    public static function run($dataDir=null)
    {

        $filesFormatType = new FileFormatType();

        //Instantiating a Workbook object
        $workbook = new Workbook($dataDir . "book1.xls");

        $worksheets = $workbook->getWorksheets();
        $worksheet = $worksheets->get(0);

        $protection = $worksheet->getProtection();

        //The following 3 methods are only for Excel 2000 and earlier formats
        $protection->setAllowEditingContent(false);
        $protection->setAllowEditingObject(false);
        $protection->setAllowEditingScenario(false);

        //Unprotecting the worksheet
        $worksheet->unprotect();

        // Save the excel file.
        $workbook->save($dataDir . "output.xls", $filesFormatType->EXCEL_97_TO_2003);

        //Print Message
        print "Worksheet unprotected successfully." . PHP_EOL;

    }

}