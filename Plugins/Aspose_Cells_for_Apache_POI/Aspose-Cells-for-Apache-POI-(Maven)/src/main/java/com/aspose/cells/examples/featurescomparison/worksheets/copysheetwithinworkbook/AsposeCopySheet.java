package com.aspose.cells.examples.featurescomparison.worksheets.copysheetwithinworkbook;

import com.aspose.cells.Workbook;
import com.aspose.cells.WorksheetCollection;
import com.aspose.cells.examples.Utils;

public class AsposeCopySheet
{
    public static void main(String[] args) throws Exception
    {
        // The path to the documents directory.
        String dataDir = Utils.getDataDir(AsposeCopySheet.class);

        //Create a new Workbook by excel file path
        Workbook wb = new Workbook(dataDir + "workbook.xls");

        //Create a Worksheets object with reference to the sheets of the Workbook.
        WorksheetCollection sheets = wb.getWorksheets();

        //Copy data to a new sheet from an existing
        //sheet within the Workbook.
        sheets.addCopy("Sheet1");

        //Save the excel file.
        wb.save(dataDir + "AsposeCopySheet.xls");

        System.out.println("Sheet copied successfully."); // Print Message
    }
}
