package com.aspose.cells.examples.featurescomparison.cellsrowscolumns.hideunhidecells;

import com.aspose.cells.Cells;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;
import com.aspose.cells.examples.Utils;

public class AsposeHideUnHideCells
{
    public static void main(String[] args) throws Exception
    {
        // The path to the documents directory.
        String dataDir = Utils.getDataDir(AsposeHideUnHideCells.class);

        Workbook workbook = new Workbook(dataDir + "workbook.xls");

        //Accessing the first worksheet in the Excel file
        Worksheet worksheet = workbook.getWorksheets().get(0);
        Cells cells = worksheet.getCells();

        cells.hideRow(2); //Hiding the 3rd row of the worksheet
        cells.hideColumn(1); //Hiding the 2nd column of the worksheet

        //Saving the modified Excel file in default (that is Excel 2003) format
        workbook.save(dataDir + "AsposeHideUnhideCells.xls");

        //Print message
        System.out.println("Rows and Columns hidden successfully.");           

    }
}
