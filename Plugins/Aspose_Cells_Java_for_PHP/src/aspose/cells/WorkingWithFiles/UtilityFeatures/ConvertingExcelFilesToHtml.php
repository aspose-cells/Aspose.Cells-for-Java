<?php


namespace Aspose\Cells\WorkingWithFiles\UtilityFeatures;

use com\aspose\cells\Workbook as Workbook;
use com\aspose\cells\SaveFormat as SaveFormat;


class ConvertingExcelFilesToHtml {

    public static function run($dataDir=null)
    {
        $saveFormat = new SaveFormat();

        $workbook = new Workbook($dataDir . "Book1.xls");

        //Save the document in PDF format
        $workbook->save($dataDir . "OutBook1.html", $saveFormat->HTML);

        // Print message
        echo "\n Excel to HTML conversion performed successfully.";
    }

}