/* 
 * Copyright 2001-2013 Aspose Pty Ltd. All Rights Reserved.
 *
 * This file is part of Aspose.Cells. The source code in this file
 * is only intended as a supplement to the documentation, and is provided
 * "as is", without warranty of any kind, either expressed or implied.
 */
 
package programmersguide.asposecells.workingwithworksheets.displayfeatures.displayhidegridlines.java;

import com.aspose.cells.*;

public class DisplayHideGridlines
{
    public static void main(String[] args) throws Exception
    {
        // The path to the documents directory.
        String dataDir = "src/programmersguide/asposecells/workingwithworksheets/displayfeatures/displayhidegridlines/data/";

        //Instantiating a Workbook object by excel file path
        Workbook workbook = new Workbook(dataDir + "book1.xls");

        //Accessing the first worksheet in the Excel file
        WorksheetCollection worksheets = workbook.getWorksheets();

        Worksheet worksheet = worksheets.get(0);

        //Hiding the grid lines of the first worksheet of the Excel file
        worksheet.setGridlinesVisible(false);

        //Saving the modified Excel file in default (that is Excel 2000) format
        workbook.save(dataDir + "output.xls");

        // Print message
        System.out.println("Grid lines are now hidden on sheet 1, please check the output document.");
    }
}




