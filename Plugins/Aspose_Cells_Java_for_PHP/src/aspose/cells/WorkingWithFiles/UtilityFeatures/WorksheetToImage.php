<?php

namespace Aspose\Cells\WorkingWithFiles\UtilityFeatures;

use com\aspose\cells\Workbook as Workbook;
use com\aspose\cells\ImageFormat as ImageFormat;
use com\aspose\cells\ImageOrPrintOptions as ImageOrPrintOptions;
use com\aspose\cells\SheetRender as SheetRender;



class WorksheetToImage {

    public static function run($dataDir=null)
    {
        $imageFormat = new ImageFormat();
        //Instantiate a new workbook with path to an Excel file
        $book = new Workbook($dataDir . "MyTestBook1.xls");

        //Create an object for ImageOptions
        $imgOptions = new ImageOrPrintOptions();

        //Set the image type
        $imgOptions->setImageFormat($imageFormat->getPng());

        //Get the first worksheet.
        $sheet = $book->getWorksheets()->get(0);

        //Create a SheetRender object for the target sheet
        $sr = new SheetRender($sheet, $imgOptions);
        for ($j = 0; $j < $sr->getPageCount(); $j++)
        {
            //Generate an image for the worksheet
            $sr->toImage($j, $dataDir . "mysheetimg" . $j . ".png");

        }

        // Print message
        print "Images generated successfully." . PHP_EOL;
    }

}