<?php
/**
 * Created by PhpStorm.
 * User: assadmahmood
 * Date: 21/07/15
 * Time: 5:00 PM
 */

namespace Aspose\Cells\WorkingWithFiles\FileHandlingFeatures;

use com\aspose\cells\Workbook as Workbook;
use com\aspose\cells\FileFormatType as FileFormatType;


class SavingFiles {

    public static function run($dataDir=null)
    {

        $fileFormatType = new FileFormatType();
      
        //Creating an Workbook object with an Excel file path
        $workbook = new Workbook($dataDir . "book.xls");
      
        //Save in default (Excel2003) format
        $workbook->save($dataDir . "book.default.out.xls");

        //Save in Excel2003 format
        $workbook->save($dataDir . "book.out.xls",$fileFormatType->EXCEL_97_TO_2003);

        //Save in Excel2007 xlsx format
        $workbook->save($dataDir . "book.out.xlsx", $fileFormatType->XLSX);

        //Save in SpreadsheetML format
        $workbook->save($dataDir . "book.out.xml", $fileFormatType->EXCEL_2003_XML);
        
        //Print Message
        print("<BR>");
        print("Worksheets are saved successfully.");
    }

} 