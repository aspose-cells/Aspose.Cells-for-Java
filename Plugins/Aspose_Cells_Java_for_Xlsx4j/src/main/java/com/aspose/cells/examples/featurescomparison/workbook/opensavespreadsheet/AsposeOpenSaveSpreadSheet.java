package com.aspose.cells.examples.featurescomparison.workbook.opensavespreadsheet;

import com.aspose.cells.Workbook;
import com.aspose.cells.examples.Utils;

public class AsposeOpenSaveSpreadSheet
{
    public static void main(String[] args) throws Exception
    {
        // The path to the documents directory.
        String dataDir = Utils.getDataDir(AsposeOpenSaveSpreadSheet.class);

        //Creating an Workbook object with an Excel file path
        Workbook workbook = new Workbook(dataDir + "pivot.xlsm");

        //Saving the Excel file
        workbook.save(dataDir + "pivot-rtt-Aspose.xlsm");

        //Print Message
        System.out.println("Worksheet saved successfully.");
    }
}
