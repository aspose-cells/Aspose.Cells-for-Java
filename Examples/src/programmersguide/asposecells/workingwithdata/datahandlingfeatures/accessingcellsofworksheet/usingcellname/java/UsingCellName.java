/* 
 * Copyright 2001-2013 Aspose Pty Ltd. All Rights Reserved.
 *
 * This file is part of Aspose.Cells. The source code in this file
 * is only intended as a supplement to the documentation, and is provided
 * "as is", without warranty of any kind, either expressed or implied.
 */
 
package programmersguide.asposecells.workingwithdata.datahandlingfeatures.accessingcellsofworksheet.usingcellname.java;

import com.aspose.cells.FileFormatType;
import com.aspose.cells.LoadOptions;
import com.aspose.cells.Workbook;

import java.io.FileInputStream;

public class UsingCellName
{
    public static void main(String[] args) throws Exception
    {
        // The path to the documents directory.
        String dataDir = "src/programmersguide/asposecells/workingwithdata/datahandlingfeatures/accessingcellsofworksheet/usingcellname/data/";

        //Instantiating a Workbook object
        Workbook workbook = new Workbook(dataDir + "book1.xls");

        //Accessing the worksheet in the Excel file
        com.aspose.cells.Worksheet worksheet = workbook.getWorksheets().get(0);
        com.aspose.cells.Cells cells = worksheet.getCells();

        //Accessing a cell using its name
        com.aspose.cells.Cell cell = cells.get("A1");

        // Print message
        System.out.println("Cell Value: " + cell.getValue());
    }
}




