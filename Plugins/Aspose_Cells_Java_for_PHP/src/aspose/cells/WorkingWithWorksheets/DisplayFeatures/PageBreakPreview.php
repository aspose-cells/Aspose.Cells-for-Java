<?php
/**
 * Created by PhpStorm.
 * User: fahadadeel
 * Date: 05/08/15
 * Time: 12:24 PM
 */

namespace Aspose\Cells\WorkingWithWorksheets\DisplayFeatures;

use com\aspose\cells\Workbook as Workbook;

class PageBreakPreview {

    public static function run($dataDir=null)
    {
        //Instantiating a Excel object by excel file path
        $workbook = new Workbook($dataDir . "book1.xls");

        //Adding a new worksheet to the Workbook object
        $worksheets = $workbook->getWorksheets();

        $worksheet = $worksheets->get(0);

        //Displaying the worksheet in page break preview
        $worksheet->setPageBreakPreview(true);

        //Saving the modified Excel file in default format
        $workbook->save($dataDir . "output.xls");

        // Print message
        print "Page break preview is enabled for sheet 1, please check the output document." . PHP_EOL;
    }

}