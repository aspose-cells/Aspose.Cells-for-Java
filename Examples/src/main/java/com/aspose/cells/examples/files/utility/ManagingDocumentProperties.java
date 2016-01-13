package com.aspose.cells.examples.files.utility;

import com.aspose.cells.*;
import com.aspose.cells.examples.Utils;

public class ManagingDocumentProperties {

    public static void main(String[] args) throws Exception {
        // The path to the documents directory.
        String dataDir = Utils.getDataDir(ManagingDocumentProperties.class);

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
