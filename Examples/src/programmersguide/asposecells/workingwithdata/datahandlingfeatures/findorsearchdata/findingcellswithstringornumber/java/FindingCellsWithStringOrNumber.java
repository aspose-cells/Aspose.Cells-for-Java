/* 
 * Copyright 2001-2013 Aspose Pty Ltd. All Rights Reserved.
 *
 * This file is part of Aspose.Cells. The source code in this file
 * is only intended as a supplement to the documentation, and is provided
 * "as is", without warranty of any kind, either expressed or implied.
 */
 
package programmersguide.asposecells.workingwithdata.datahandlingfeatures.findorsearchdata.findingcellswithstringornumber.java;

import com.aspose.cells.*;

public class FindingCellsWithStringOrNumber
{
    public static void main(String[] args) throws Exception
    {
        // The path to the documents directory.
        String dataDir = "src/programmersguide/asposecells/workingwithdata/datahandlingfeatures/findorsearchdata/findingcellswithstringornumber/data/";

        //Instantiating a Workbook object
        Workbook workbook = new Workbook(dataDir + "book1.xls");

        //Accessing the first worksheet in the Excel file
        Worksheet worksheet = workbook.getWorksheets().get(0);

        //Finding the cell containing the specified formula
        Cells cells = worksheet.getCells();

        //Instantiate FindOptions
        FindOptions findOptions = new FindOptions();

        //Finding the cell containing a string value that starts with "Or"
        findOptions.setLookAtType(LookAtType.START_WITH);

        Cell cell = cells.find("SH",null,findOptions);

        //Printing the name of the cell found after searching worksheet
        System.out.println("Name of the cell containing String: " + cell.getName());
    }
}




