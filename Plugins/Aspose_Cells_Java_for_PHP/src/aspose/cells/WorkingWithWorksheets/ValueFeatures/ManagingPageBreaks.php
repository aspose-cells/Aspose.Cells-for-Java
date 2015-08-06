<?php

namespace Aspose\Cells\WorkingWithWorksheets\ValueFeatures;

use com\aspose\cells\Workbook as Workbook;

class ManagingPageBreaks {

    public static function run($dataDir=null)
    {
        # Instantiating a Workbook object
        $workbook = new Workbook();

        # Adding Page Breaks
        ManagingPageBreaks::add_page_breaks($dataDir=null, $workbook);

        # Clearing All Page Breaks
        ManagingPageBreaks::clear_all_page_breaks($dataDir=null, $workbook);

        # Removing Specific Page Break
        ManagingPageBreaks::remove_page_break($dataDir=null, $workbook);

    }

    public static function add_page_breaks($dataDir=null, $workbook)
    {
        $worksheets = $workbook->getWorksheets();
        $worksheet = $worksheets->get(0);

        $h_page_breaks = $worksheet->getHorizontalPageBreaks();
        $h_page_breaks->add("Y30");

        $v_page_breaks = $worksheet->getVerticalPageBreaks();
        $v_page_breaks->add("Y30");

        # Saving the modified Excel file in default (that is Excel 2003) format
        $workbook->save($dataDir . "Add Page Breaks.xls");

        print "Add page breaks, please check the output file." . PHP_EOL;

    }

    public static function clear_all_page_breaks($dataDir=null, $workbook)
    {

        $workbook->getWorksheets()->get(0)->getHorizontalPageBreaks()->clear();
        $workbook->getWorksheets()->get(0)->getVerticalPageBreaks()->clear();

        # Saving the modified Excel file in default (that is Excel 2003) format
        $workbook->save($dataDir . "Clear All Page Breaks.xls");

        print "Clear all page breaks, please check the output file." . PHP_EOL;

    }

    public static function remove_page_break($dataDir=null, $workbook)
    {

        $worksheets = $workbook->getWorksheets();
        $worksheet = $worksheets->get(0);

        $h_page_breaks = $worksheet->getHorizontalPageBreaks();
        $h_page_breaks->removeAt(0);

        $v_page_breaks = $worksheet->getVerticalPageBreaks();
        $v_page_breaks->removeAt(0);

        # Saving the modified Excel file in default (that is Excel 2003) format
        $workbook->save($dataDir . "Remove Page Break.xls");

        print "Remove page break, please check the output file." . PHP_EOL;

    }

}