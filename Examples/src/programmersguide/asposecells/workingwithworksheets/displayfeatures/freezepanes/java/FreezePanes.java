/* 
 * Copyright 2001-2013 Aspose Pty Ltd. All Rights Reserved.
 *
 * This file is part of Aspose.Cells. The source code in this file
 * is only intended as a supplement to the documentation, and is provided
 * "as is", without warranty of any kind, either expressed or implied.
 */
 
package programmersguide.asposecells.workingwithworksheets.displayfeatures.freezepanes.java;

import com.aspose.cells.*;

public class FreezePanes
{
    public static void main(String[] args) throws Exception
    {
        // The path to the documents directory.
        String dataDir = "src/programmersguide/asposecells/workingwithworksheets/displayfeatures/freezepanes/data/";
        
      //Instantiating a Excel object by excel file path
        Workbook workbook = new Workbook(dataDir + "book.xls");

        //Accessing the first worksheet in the Excel file
        WorksheetCollection worksheets = workbook.getWorksheets();
        Worksheet worksheet = worksheets.get(0);

        //Applying freeze panes settings
        worksheet.freezePanes(3,2,3,2);

        //Saving the modified Excel file in default format
        workbook.save(dataDir + "book.out.xls");
        
        //Print Message
        System.out.println("Panes freeze successfull.");
    }
}