<?php
/**
 * Created by PhpStorm.
 * User: fahadadeel
 * Date: 05/08/15
 * Time: 12:17 PM
 */

namespace Aspose\Cells\WorkingWithWorksheets\DisplayFeatures;

use com\aspose\cells\Workbook as Workbook;

class DisplayHideScrollBars {

    public static function run($dataDir=null)
    {

        //Instantiating a Excel object by excel file path
        $workbook = new Workbook($dataDir . "book1.xls");

        //Hiding the vertical scroll bar of the Excel file
        $workbook->getSettings()->setVScrollBarVisible(false);

        //Hiding the horizontal scroll bar of the Excel file
        $workbook->getSettings()->setHScrollBarVisible(false);

        //Saving the modified Excel file in default (that is Excel 2003) format
        $workbook->save($dataDir . "output.xls");

        // Print message
        print "Scroll bars are now hidden, please check the output document.";

    }

}