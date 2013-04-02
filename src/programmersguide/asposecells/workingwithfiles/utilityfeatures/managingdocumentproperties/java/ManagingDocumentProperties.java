/* 
 * Copyright 2001-2013 Aspose Pty Ltd. All Rights Reserved.
 *
 * This file is part of Aspose.Cells. The source code in this file
 * is only intended as a supplement to the documentation, and is provided
 * "as is", without warranty of any kind, either expressed or implied.
 */
 
package programmersguide.asposecells.workingwithfiles.utilityfeatures.managingdocumentproperties.java;

import com.aspose.cells.*;

public class ManagingDocumentProperties
{
    public static void main(String[] args) throws Exception
    {
        // The path to the documents directory.
        String dataDir = "src/programmersguide/asposecells/workingwithfiles/utilityfeatures/managingdocumentproperties/data/";

        //Instantiate a Workbook object by excel file path
        Workbook workbook = new Workbook(dataDir + "Book1.xls");

        //Retrieve a list of all custom document properties of the Excel file
        CustomDocumentPropertyCollection customProperties = workbook.getWorksheets().getCustomDocumentProperties();

        //Accessing a custom document property by using the property index
        DocumentProperty customProperty1 = customProperties.get(3);

        //Accessing a custom document property by using the property name
        DocumentProperty customProperty2 = customProperties.get("Owner");


        //Adding a custom document property to the Excel file
        DocumentProperty publisher = customProperties.add("Publisher", "Aspose");

        //Save the file
        workbook.save(dataDir + "Test_Workbook.xls");

        //Removing a custom document property
        customProperties.remove("Publisher");

        //Save the file
        workbook.save(dataDir + "Test_Workbook_RemovedProperty.xls");

        // Print message
        System.out.println("Excel file's custom properties accessed successfully.");
    }
}




