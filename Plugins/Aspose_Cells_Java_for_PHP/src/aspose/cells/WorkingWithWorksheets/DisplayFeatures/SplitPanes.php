<?php

namespace Aspose\Cells\WorkingWithWorksheets\DisplayFeatures;

use com\aspose\cells\Workbook as Workbook;
use com\aspose\cells\SaveFormat as SaveFormat;

class SplitPanes {

    public static function run($dataDir=null)
    {

        $saveFormat = new SaveFormat();
        //Instantiate a new workbook
        //Open a template file
        $book = new Workbook($dataDir . "book.xls");

        //Set the active cell
        $book->getWorksheets()->get(0)->setActiveCell("A20");

        //Split the worksheet window
        $book->getWorksheets()->get(0)->split();

        //Save the excel file
        $book->save($dataDir . "book.out.xls", $saveFormat->EXCEL_97_TO_2003);

        //Print Message
        print "Panes split successfully." . PHP_EOL;

    }

}