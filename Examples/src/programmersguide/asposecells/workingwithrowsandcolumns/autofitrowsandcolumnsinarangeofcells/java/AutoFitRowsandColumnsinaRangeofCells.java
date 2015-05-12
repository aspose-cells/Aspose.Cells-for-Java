/* 
 * Copyright 2001-2013 Aspose Pty Ltd. All Rights Reserved.
 *
 * This file is part of Aspose.Cells. The source code in this file
 * is only intended as a supplement to the documentation, and is provided
 * "as is", without warranty of any kind, either expressed or implied.
 */
 
package programmersguide.asposecells.workingwithrowsandcolumns.autofitrowsandcolumnsinarangeofcells.java;

import com.aspose.cells.*;

public class AutoFitRowsandColumnsinaRangeofCells
{
    public static void main(String[] args) throws Exception
    {
        // The path to the documents directory.
        String dataDir = "src/programmersguide/asposecells/workingwithrowsandcolumns/autofitrowsandcolumnsinarangeofcells/data/";
        
        //Instantiating a Workbook object
        Workbook workbook = new Workbook(dataDir + "workbook.xls");

        //Accessing the first worksheet in the Excel file
        Worksheet worksheet = workbook.getWorksheets().get(0);

        //Auto-fitting the 3rd row of the worksheet based on the contents in a range of
        //cells (from 1st to 9th column) within the row
        worksheet.autoFitRow(3,4,10);
        
        //Auto-fitting the 4th column of the worksheet based on the contents in a range of
        //cells (from 1st to 9th row) within the column
        worksheet.autoFitColumn(0,0,8);

        //Saving the modified Excel file in default (that is Excel 2003) format
        workbook.save(dataDir + "workbook.out.xls");
        
        //Print message
        System.out.println("Row and Column auto fit successfully.");        
    }
}




