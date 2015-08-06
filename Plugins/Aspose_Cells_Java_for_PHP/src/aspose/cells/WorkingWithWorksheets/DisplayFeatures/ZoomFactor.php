<?php


namespace Aspose\Cells\WorkingWithWorksheets\DisplayFeatures;

use com\aspose\cells\Workbook as Workbook;

class ZoomFactor {

    public static function run($dataDir=null)
    {
        //Instantiating a Excel object by excel file path
        $workbook = new Workbook($dataDir . "book1.xls");

        //Accessing the first worksheet in the Excel file
        $worksheets = $workbook->getWorksheets();
        $worksheet = $worksheets->get(0);

        //Setting the zoom factor of the worksheet to 75
        $worksheet->setZoom(75);

        //Saving the modified Excel file in default format
        $workbook->save($dataDir . "output.xls");

        // Print message
        print "Zoom factor set to 75% for sheet 1, please check the output document.";
    }

}