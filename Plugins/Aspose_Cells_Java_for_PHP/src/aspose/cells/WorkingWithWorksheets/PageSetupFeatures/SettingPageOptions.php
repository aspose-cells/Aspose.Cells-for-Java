<?php

namespace Aspose\Cells\WorkingWithWorksheets\PageSetupFeatures;

use com\aspose\cells\Workbook as Workbook;
use com\aspose\cells\FileFormatType as FileFormatType;
use com\aspose\cells\PageOrientationType as PageOrientationType;


class SettingPageOptions {

    public static function run($dataDir=null)
    {

        SettingPageOptions::page_orientation($dataDir);

        SettingPageOptions::scaling($dataDir);

    }

    public static function page_orientation($dataDir=null)
    {

        # Instantiating a Workbook object by excel file path
        $workbook = new Workbook();

        # Accessing the first worksheet in the Excel file
        $worksheets = $workbook->getWorksheets();
        $sheet_index = $worksheets->add();
        $sheet = $worksheets->get($sheet_index);

        # Setting the orientation to Portrait
        $page_setup = $sheet->getPageSetup();
        $page_orientation_type = new PageOrientationType();
        $page_setup->setOrientation($page_orientation_type->PORTRAIT);

        # Saving the modified Excel file in default (that is Excel 2003) format
        $workbook->save($dataDir . "Page Orientation.xls");

        print "Set page orientation, please check the output file." . PHP_EOL;
    }

    public static function scaling($dataDir=null)
    {

        # Instantiating a Workbook object by excel file path
        $workbook = new Workbook($dataDir . 'Book1.xls');

        # Accessing the first worksheet in the Excel file
        $worksheets = $workbook->getWorksheets();
        $sheet_index = $worksheets->add();
        $sheet = $worksheets->get($sheet_index);

        # Setting the scaling factor to 100
        $page_setup = $sheet->getPageSetup();
        $page_setup->setZoom(100);

        # Saving the modified Excel file in default (that is Excel 2003) format
        $workbook->save($dataDir . "Scaling.xls");

        print "Set scaling, please check the output file." . PHP_EOL;

    }

}