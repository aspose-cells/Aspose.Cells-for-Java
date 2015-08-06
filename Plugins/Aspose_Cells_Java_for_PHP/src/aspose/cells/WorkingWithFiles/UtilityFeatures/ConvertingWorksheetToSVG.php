<?php

namespace Aspose\Cells\WorkingWithFiles\UtilityFeatures;

use com\aspose\cells\Workbook as Workbook;
use com\aspose\cells\ImageFormat as ImageFormat;
use com\aspose\cells\ImageOrPrintOptions as ImageOrPrintOptions;
use com\aspose\cells\SheetRender as SheetRender;
use com\aspose\cells\SaveFormat as SaveFormat;


class ConvertingWorksheetToSVG {

    public static function run($dataDir=null)
    {

        $saveFormat = new SaveFormat();

        $path = $dataDir . "Template.xlsx";

        //Create a workbook object from the template file
        $workbook = new Workbook($path);

        //Convert each worksheet into svg format in a single page.
        $imgOptions = new ImageOrPrintOptions();
        $imgOptions->setSaveFormat($saveFormat->SVG);
        $imgOptions->setOnePagePerSheet(true);

        //Convert each worksheet into svg format
        $sheetCount = $workbook->getWorksheets()->getCount();

        for($i=0; $i < $sheetCount; $i++)
        {
            $sheet = $workbook->getWorksheets()->get($i);

            $sr = new SheetRender($sheet, $imgOptions);

            $pageCount = $sr->getPageCount();
            for ($k = 0; $k < $pageCount; $k++)
            {
                //Output the worksheet into Svg image format
                $sr->toImage($k, $path . $sheet->getName() . $k . ".out.svg");
            }
        }

        // Print message
        print "Excel to SVG conversion completed successfully.";
    }

}