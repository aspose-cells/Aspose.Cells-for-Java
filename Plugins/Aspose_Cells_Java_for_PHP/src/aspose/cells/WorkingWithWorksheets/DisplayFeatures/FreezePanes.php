<?php
/**
 * Created by PhpStorm.
 * User: fahadadeel
 * Date: 05/08/15
 * Time: 12:34 PM
 */

namespace Aspose\Cells\WorkingWithWorksheets\DisplayFeatures;

use com\aspose\cells\Workbook as Workbook;

class FreezePanes {

    public static function run($dataDir=null)
    {
        //Instantiating a Excel object by excel file path
        $workbook = new Workbook($dataDir . "book.xls");

        //Accessing the first worksheet in the Excel file
        $worksheets = $workbook->getWorksheets();
        $worksheet = $worksheets->get(0);

        //Applying freeze panes settings
        $worksheet->freezePanes(3,2,3,2);

        //Saving the modified Excel file in default format
        $workbook->save($dataDir . "book.out.xls");

        //Print Message
        print "Panes freeze successfull." . PHP_EOL;
    }

}