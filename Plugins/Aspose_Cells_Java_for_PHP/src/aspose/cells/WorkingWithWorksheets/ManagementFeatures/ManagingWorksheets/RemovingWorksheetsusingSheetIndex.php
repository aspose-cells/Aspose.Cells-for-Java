<?php

namespace Aspose\Cells\WorkingWithWorksheets\ManagementFeatures\ManagingWorksheets;

use com\aspose\cells\Workbook as Workbook;
use java\io\FileInputStream as FileInputStream;


class RemovingWorksheetsusingSheetIndex {

    public static function run($dataDir=null)
    {
        //Creating a file stream containing the Excel file to be opened
        $fstream=new FileInputStream($dataDir . "book.xls");

        //Instantiating a Workbook object with the stream
        $workbook = new Workbook($fstream);

        //Removing a worksheet using its sheet index
        $workbook->getWorksheets()->removeAt(0);

        //Saving the Excel file
        $workbook->save($dataDir . "book.out.xls");

        //Closing the file stream to free all resources
        $fstream->close();


        //Print Message
        print "Sheet removed successfully.";
    }

}