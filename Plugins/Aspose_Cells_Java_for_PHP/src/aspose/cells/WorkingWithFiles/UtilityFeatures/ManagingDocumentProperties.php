<?php
/**
 * Created by PhpStorm.
 * User: fahadadeel
 * Date: 05/08/15
 * Time: 10:25 AM
 */

namespace Aspose\Cells\WorkingWithFiles\UtilityFeatures;

use com\aspose\cells\Workbook as Workbook;


class ManagingDocumentProperties {

    public static function run($dataDir=null)
    {
        //Instantiate a Workbook object by excel file path
        $workbook = new Workbook($dataDir . "Book1.xls");

        //Retrieve a list of all custom document properties of the Excel file
        $customProperties = $workbook->getWorksheets()->getCustomDocumentProperties();

        //Accessing a custom document property by using the property index
        $customProperty1 = $customProperties->get(3);

        //Accessing a custom document property by using the property name
        $customProperty2 = $customProperties->get("Owner");


        //Adding a custom document property to the Excel file
        $publisher = $customProperties->add("Publisher", "Aspose");

        //Save the file
        $workbook->save($dataDir . "Test_Workbook.xls");

        //Removing a custom document property
        $customProperties->remove("Publisher");

        //Save the file
        $workbook->save($dataDir . "Test_Workbook_RemovedProperty.xls");

        // Print message
        print "Excel file's custom properties accessed successfully." . PHP_EOL;
    }

}