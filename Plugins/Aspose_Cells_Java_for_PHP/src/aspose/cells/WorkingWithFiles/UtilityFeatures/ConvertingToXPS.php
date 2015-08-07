<?php


namespace Aspose\Cells\WorkingWithFiles\UtilityFeatures;

use com\aspose\cells\Workbook as Workbook;
use com\aspose\cells\ImageFormat as ImageFormat;
use com\aspose\cells\ImageOrPrintOptions as ImageOrPrintOptions;
use com\aspose\cells\SheetRender as SheetRender;
use com\aspose\cells\SaveFormat as SaveFormat;


class ConvertingToXPS {

    public static function run($dataDir=null)
    {
        $saveFormat = new SaveFormat();
        $workbook = new Workbook($dataDir . "Book1.xls");

        //Get the first worksheet.
        $sheet = $workbook->getWorksheets()->get(0);

        //Apply different Image and Print options
        $options = new ImageOrPrintOptions();

        //Set the Format
        $options->setSaveFormat($saveFormat->XPS);

        // Render the sheet with respect to specified printing options
        $sr = new SheetRender($sheet, $options);
        $sr->toImage(0, $dataDir . "out_printingxps.xps");

        //Save the complete Workbook in XPS format
        $workbook->save($dataDir . "out_whole_printingxps", $saveFormat->XPS);

        // Print message
        print "Excel to XPS conversion performed successfully.";
    }

}