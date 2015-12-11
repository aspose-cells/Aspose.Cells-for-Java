package com.aspose.cells.examples.featurescomparison.worksheet.adjustheight;

import com.aspose.cells.Cells;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;
import com.aspose.cells.examples.Utils;

/**
 * @author Shoaib Khan
 */
public class AsposeHeightAdjustment
{
    public static void main(String[] args) throws Exception
    {   
        // The path to the documents directory.
        String dataDir = Utils.getDataDir(AsposeHeightAdjustment.class);

        //Instantiating a Workbook object
        Workbook workbook = new Workbook();

        //Accessing the first worksheet in the Excel file
        Worksheet worksheet = workbook.getWorksheets().get(0);
        Cells cells = worksheet.getCells();

        //Setting the height of all rows in the worksheet to 8
        worksheet.getCells().setStandardHeight(8f);

        //Setting the height of the second row to 40
        cells.setRowHeight(1, 40);

        //Saving the modified Excel file in default (that is Excel 2003) format
        workbook.save(dataDir + "RowHeight-Aspose.xlsx");

        //Print Message
        System.out.println("Worksheet saved successfully.");
    }
}
