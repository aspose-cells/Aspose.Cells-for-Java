/* 
 * Copyright 2001-2013 Aspose Pty Ltd. All Rights Reserved.
 *
 * This file is part of Aspose.Cells. The source code in this file
 * is only intended as a supplement to the documentation, and is provided
 * "as is", without warranty of any kind, either expressed or implied.
 */
 
package programmersguide.asposecells.workingwithrowsandcolumns.copyingrowsandcolumns.java;

import com.aspose.cells.*;

public class CopyingRowsandColumns
{
    public static void main(String[] args) throws Exception
    {
        // The path to the documents directory.
        String dataDir = "src/programmersguide/asposecells/workingwithrowsandcolumns/copyingrowsandcolumns/data/";
        
        //Create a new Workbook.
        Workbook excelWorkbook = new Workbook(dataDir + "workbook.xls");

        //Get the first worksheet in the workbook.
        Worksheet wsTemplate = excelWorkbook.getWorksheets().get(0);

        //Copy the second row with data, formating, images and drawing objects
        //to the 12th row in the worksheet.
        wsTemplate.getCells().copyRow(wsTemplate.getCells(),2,10);  
        
        //Copy the first column from the first worksheet of the first workbook into
        //the first worksheet of the second workbook.
        wsTemplate.getCells().copyColumn(wsTemplate.getCells(),1,4);

        //Save the excel file.
        excelWorkbook.save(dataDir + "workbook.out.xls");
        
        //Print message
        System.out.println("Row and Column copied successfully.");        
    }
}




