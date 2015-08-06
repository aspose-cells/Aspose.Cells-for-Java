<?php
/**
 * Created by PhpStorm.
 * User: fahadadeel
 * Date: 05/08/15
 * Time: 12:14 PM
 */

namespace Aspose\Cells\WorkingWithWorksheets\DisplayFeatures;

use com\aspose\cells\Workbook as Workbook;

class DisplayHideTabs {

    public static function run($dataDir=null)
    {

        //Instantiating a Workbook object by excel file path
        $workbook = new Workbook($dataDir . "book1.xls");

        //Hiding the tabs of the Excel file
        $workbook->getSettings()->setShowTabs(false);

        //Saving the modified Excel file in default (that is Excel 2003) format
        $workbook->save($dataDir + "output.xls");

        // Print message
        print "Tabs are now hidden, please check the output file." . PHP_EOL;


    }

}