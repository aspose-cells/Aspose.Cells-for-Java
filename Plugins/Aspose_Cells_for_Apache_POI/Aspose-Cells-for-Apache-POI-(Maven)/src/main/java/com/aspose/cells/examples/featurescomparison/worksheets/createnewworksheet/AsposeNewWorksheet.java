package com.aspose.cells.examples.featurescomparison.worksheets.createnewworksheet;

import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;
import com.aspose.cells.WorksheetCollection;
import com.aspose.cells.examples.Utils;

public class AsposeNewWorksheet
{
    public static void main(String[] args) throws Exception
    {
        // The path to the documents directory.
        String dataDir = Utils.getDataDir(AsposeNewWorksheet.class);

        //Instantiating a Workbook object
        Workbook workbook = new Workbook();

        //Adding a new worksheet to the Workbook object
        WorksheetCollection worksheets = workbook.getWorksheets();
        Worksheet worksheet = worksheets.add("My Worksheet");

        //Saving the Excel file
        workbook.save(dataDir + "AsposeNewWorksheet.xls");

        //Print Message
        System.out.println("Sheet added successfully.");
    }
}
