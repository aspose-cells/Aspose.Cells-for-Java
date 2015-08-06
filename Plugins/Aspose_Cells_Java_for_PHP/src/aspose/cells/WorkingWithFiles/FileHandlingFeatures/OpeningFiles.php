<?php
/**
 * Created by PhpStorm.
 * User: fahadadeel
 * Date: 21/07/15
 * Time: 4:14 PM
 */

namespace Aspose\Cells\WorkingWithFiles\FileHandlingFeatures;

use com\aspose\cells\Workbook as Workbook;
use com\aspose\cells\LoadOptions as LoadOptions;
use com\aspose\cells\FileFormatType as FileFormatType;
use java\io\FileInputStream as FileInputStream;



class OpeningFiles {

    public static function run($dataDir=null)
    {

        $fileFormatType = new FileFormatType();

        // 1.
        // Opening from path.
        //Creating an Workbook object with an Excel file path
        $workbook1 = new Workbook($dataDir . "Book1.xls");

        // Print message

        print("<br />");
        print ("Workbook opened using path successfully.");

        // 2.
        // Opening workbook from stream
        //Create a Stream object
        $fstream = new FileInputStream($dataDir . "Book2.xls");

        //Creating an Workbook object with the stream object
        $workbook2 = new Workbook($fstream);

        $fstream->close();

        // Print message
        print("<br />");
        print ("Workbook opened using stream successfully.");


        // 3.
        // Opening Microsoft Excel 97 Files
        //Createing and EXCEL_97_TO_2003 LoadOptions object
        $loadOptions1 = new LoadOptions($fileFormatType->EXCEL_97_TO_2003);

        //Creating an Workbook object with excel 97 file path and the loadOptions object
        $workbook3 = new Workbook($dataDir . "Book_Excel97_2003.xls", $loadOptions1);

        // Print message
        print("<br />");
        print("Excel 97 Workbook opened successfully.");



        // 4.
        // Opening Microsoft Excel 2007 XLSX Files
        //Createing and XLSX LoadOptions object
        $loadOptions2 = new LoadOptions($fileFormatType->XLSX);

        //Creating an Workbook object with 2007 xlsx file path and the loadOptions object
        $workbook4 = new Workbook($dataDir . "Book_Excel2007.xlsx", $loadOptions2);

        // Print message
        print("<br />");
        print ("Excel 2007 Workbook opened successfully.");



        // 5.
        // Opening SpreadsheetML Files
        //Creating and EXCEL_2003_XML LoadOptions object
        $loadOptions3 = new LoadOptions($fileFormatType->EXCEL_2003_XML);

        //Creating an Workbook object with SpreadsheetML file path and the loadOptions object
        $workbook5 = new Workbook($dataDir . "Book3.xml", $loadOptions3);

        // Print message
        print("<br />");
        print ("SpreadSheetML format workbook has been opened successfully.");

        // 6.
        // Opening CSV Files
        //Creating and CSV LoadOptions object
        $loadOptions4 = new LoadOptions($fileFormatType->CSV);

        //Creating an Workbook object with CSV file path and the loadOptions object
        $workbook6 = new Workbook($dataDir . "Book_CSV.csv", $loadOptions4);

        // Print message
        print("<br />");
        print ("CSV format workbook has been opened successfully.");



        // 7.
        // Opening Tab Delimited Files
        //Creating and TAB_DELIMITED LoadOptions object
        $loadOptions5 = new LoadOptions($fileFormatType->TAB_DELIMITED);

        //Creating an Workbook object with Tab Delimited text file path and the loadOptions object
        $workbook7 = new Workbook($dataDir . "Book1TabDelimited.txt", $loadOptions5);

        // Print message
        print("<br />");
        print ("Tab Delimited workbook has been opened successfully.");



        // 8.
        // Opening Encrypted Excel Files
        //Creating and EXCEL_97_TO_2003 LoadOptions object
        $loadOptions6 = new LoadOptions($fileFormatType->EXCEL_97_TO_2003);

        //Setting the password for the encrypted Excel file
        $loadOptions6->setPassword("1234");

        //Creating an Workbook object with file path and the loadOptions object
        $workbook8 = new Workbook($dataDir . "encryptedBook.xls", $loadOptions6);

        // Print message
        print("<br />");
        print ("Encrypted workbook has been opened successfully.");


    }
} 