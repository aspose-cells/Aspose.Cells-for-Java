<?php


namespace Aspose\Cells\WorkingWithFiles\UtilityFeatures;

use com\aspose\cells\Workbook as Workbook;
use com\aspose\cells\HtmlSaveOptions as HtmlSaveOptions;
use com\aspose\cells\SaveFormat as SaveFormat;
//use com\aspose\cells\examples\Utils as Utils;



class ConvertingToMhtmlFiles {

    public static function run($dataDir=null)
    {

        $sveFormat = new SaveFormat();

        //Specify the file path
        $filePath = $dataDir . "Book1.xlsx";

        //Specify the HTML saving options
        $sv = new HtmlSaveOptions($sveFormat->M_HTML);

        //Instantiate a workbook and open the template XLSX file
        $wb = new Workbook($filePath);

        //Save the MHT file
        $wb->save($filePath . ".out.mht", $sv);

        // Print message
        print "<br>";
        print "<br>";
        print "Excel to MHTML conversion performed successfully.";
    }

}