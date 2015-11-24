package com.aspose.cells.examples.featurescomparison.cellsrowscolumns.splitpanes;

import com.aspose.cells.SaveFormat;
import com.aspose.cells.Workbook;
import com.aspose.cells.examples.Utils;

public class AsposeSplitPanes 
{
    public static void main(String[] args) throws Exception 
    {
        // The path to the documents directory.
        String dataDir = Utils.getDataDir(AsposeSplitPanes.class);

        //Instantiate a new workbook / Open a template file
        Workbook book = new Workbook(dataDir + "workbook.xls");

        //Set the active cell
        book.getWorksheets().get(0).setActiveCell("A20");

        //Split the worksheet window
        book.getWorksheets().get(0).split();

        //Save the Excel file
        book.save(dataDir + "AsposeSplitPanes.xls", SaveFormat.EXCEL_97_TO_2003);

        System.out.println("Done.");
    }
}
