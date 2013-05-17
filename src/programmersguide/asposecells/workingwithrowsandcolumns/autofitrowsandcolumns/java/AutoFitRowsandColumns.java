/* 
 * Copyright 2001-2013 Aspose Pty Ltd. All Rights Reserved.
 *
 * This file is part of Aspose.Cells. The source code in this file
 * is only intended as a supplement to the documentation, and is provided
 * "as is", without warranty of any kind, either expressed or implied.
 */
 
package programmersguide.asposecells.workingwithrowsandcolumns.autofitrowsandcolumns.java;

import com.aspose.cells.*;

public class AutoFitRowsandColumns
{
    public static void main(String[] args) throws Exception
    {
        // The path to the documents directory.
        String dataDir = "src/programmersguide/asposecells/workingwithrowsandcolumns/autofitrowsandcolumns/data/";
        
        //Instantiating a Workbook object
        Workbook workbook = new Workbook(dataDir + "workbook.xls");

        //Accessing the first worksheet in the Excel file
        Worksheet worksheet = workbook.getWorksheets().get(0);

        //Auto-fitting the 2nd row of the worksheet
        worksheet.autoFitRow(1);
        
        //Auto-fitting the 1st column of the worksheet
        worksheet.autoFitColumn(0);

        //Saving the modified Excel file in default (that is Excel 2003) format
        workbook.save(dataDir + "workbook.out.xls");
        
        //Print message
        System.out.println("Row and Column auto fit successfully.");
    }
}