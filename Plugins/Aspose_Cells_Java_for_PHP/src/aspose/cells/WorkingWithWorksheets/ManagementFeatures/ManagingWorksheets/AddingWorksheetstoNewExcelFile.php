<?php

namespace Aspose\Cells\WorkingWithWorksheets\ManagementFeatures\ManagingWorksheets;

use com\aspose\cells\Workbook as Workbook;

class AddingWorksheetstoNewExcelFile {

    public static function run($dataDir=null)
    {

        //Instantiating a Workbook object
        $workbook = new Workbook();

        //Adding a new worksheet to the Workbook object
        $worksheets = $workbook->getWorksheets();

        $sheetIndex = $worksheets->add();
        $worksheet = $worksheets->get($sheetIndex);

        //Setting the name of the newly added worksheet
        $worksheet->setName("My Worksheet");

        //Saving the Excel file
        $workbook->save($dataDir . "book.out.xls");

        //Print Message
        print "Sheet added successfully.";

    }

}