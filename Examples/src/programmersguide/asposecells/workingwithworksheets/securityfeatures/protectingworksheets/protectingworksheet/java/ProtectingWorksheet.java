/* 
 * Copyright 2001-2013 Aspose Pty Ltd. All Rights Reserved.
 *
 * This file is part of Aspose.Cells. The source code in this file
 * is only intended as a supplement to the documentation, and is provided
 * "as is", without warranty of any kind, either expressed or implied.
 */
 
package programmersguide.asposecells.workingwithworksheets.securityfeatures.protectingworksheets.protectingworksheet.java;

import com.aspose.cells.*;

import java.io.FileInputStream;

public class ProtectingWorksheet
{
    public static void main(String[] args) throws Exception
    {
        // The path to the documents directory.
        String dataDir = "src/programmersguide/asposecells/workingwithworksheets/securityfeatures/protectingworksheets/protectingworksheet/data/";
        
        //Instantiating a Excel object by excel file path
        Workbook excel = new Workbook(dataDir + "book1.xls");

        //Accessing the first worksheet in the Excel file
        WorksheetCollection worksheets = excel.getWorksheets();
        Worksheet worksheet = worksheets.get(0);

        Protection protection = worksheet.getProtection();

        //The following 3 methods are only for Excel 2000 and earlier formats
        protection.setAllowEditingContent(false);
        protection.setAllowEditingObject(false);
        protection.setAllowEditingScenario(false);

        //Protects the first worksheet with a password "1234"
        protection.setPassword("1234");

        //Saving the modified Excel file in default format
        excel.save(dataDir + "output.xls");
        
        //Print Message
        System.out.println("Sheet protected successfully.");
    }
}




