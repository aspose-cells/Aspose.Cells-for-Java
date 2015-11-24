package com.aspose.cells.examples.featurescomparison.cellsrowscolumns.autofitrowandcolumn;

import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;
import com.aspose.cells.examples.Utils;

public class AsposeAutoFit
{
    public static void main(String[] args) throws Exception
    {
        // The path to the documents directory.
        String dataDir = Utils.getDataDir(AsposeAutoFit.class);
        
        //Instantiating a Workbook object
        Workbook workbook = new Workbook(dataDir + "workbook.xls");

        //Accessing the first worksheet in the Excel file
        Worksheet worksheet = workbook.getWorksheets().get(0);

        worksheet.autoFitRow(1); //Auto-fitting the 2nd row of the worksheet
        worksheet.autoFitColumn(0); //Auto-fitting the 1st column of the worksheet

        //Saving the modified Excel file in default (that is Excel 2003) format
        workbook.save(dataDir + "AutoFit_Aspose_Out.xls");

        //Print message
        System.out.println("Row and Column auto fit successfully.");
    }
}
